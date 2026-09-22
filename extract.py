"""图片表格识别与汇总导出核心逻辑。

基于 RapidOCR（中文 OCR）+ RapidTable（表格结构识别，PaddleOCR SLANet-Plus 的 ONNX 版），
把图片中的表格识别后汇总到一个 Excel 文件。

可作为库使用，也可命令行调用：
    python extract.py 图片路径或目录... -o 汇总结果.xlsx
"""
from __future__ import annotations

import argparse
import cv2
import html
import io
import logging
import numpy as np
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Sequence

import pandas as pd
from lxml import html as lxml_html

# rapid_table 3.0.2 在导入时会在包目录（site-packages，只读）下创建 models 目录，
# 在 Streamlit Cloud 等只读环境会抛 PermissionError。这里先把 Path.mkdir 变为容错操作。
_original_path_mkdir = Path.mkdir


def _safe_path_mkdir(self, mode=0o777, parents=False, exist_ok=False):
    try:
        return _original_path_mkdir(self, mode=mode, parents=parents, exist_ok=exist_ok)
    except PermissionError:
        return None


Path.mkdir = _safe_path_mkdir

from rapidocr import RapidOCR
from rapid_table import ModelType, RapidTable, RapidTableInput

# 简体中文识别。优先用枚举，找不到时退回字符串 "ch"。
try:
    from rapidocr import LangRec

    _LANG_TYPE = LangRec.CH
except Exception:  # pragma: no cover
    _LANG_TYPE = "ch"

# 减少 OCR / 表格模型的 INFO 日志输出
for _name in ("rapidocr", "RapidOCR", "rapid_table", "RapidTable"):
    logging.getLogger(_name).setLevel(logging.WARNING)

MIN_CONFIDENCE = 30
SUPPORTED_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

# 仓库内自带的本地模型目录（避免运行时联网下载 / 写入只读目录）
MODELS_DIR = Path(__file__).resolve().parent / "models"

# 进度回调签名：(当前序号, 总数, 文件名)
ProgressCallback = Callable[[int, int, str], None]


def build_ocr() -> RapidOCR:
    """创建简体中文 OCR 实例（使用仓库内自带的本地模型）。"""
    return RapidOCR(
        params={
            "Global.model_root_dir": str(MODELS_DIR / "rapidocr"),
            "Rec.lang_type": _LANG_TYPE,
        }
    )


def build_table_engine() -> RapidTable:
    """创建表格结构识别引擎（使用仓库内自带的 SLANet-Plus 本地模型）。"""
    return RapidTable(
        RapidTableInput(
            model_type=ModelType.SLANETPLUS,
            model_dir_or_path=str(MODELS_DIR / "rapid_table" / "slanet-plus.onnx"),
            # RapidTable 内部会再创建一个 OCR 引擎（即使我们外部已传入 OCR 结果），
            # 这里也把它指向本地模型，避免在只读环境尝试联网下载。
            ocr_params={
                "Global.model_root_dir": str(MODELS_DIR / "rapidocr"),
                "Rec.lang_type": _LANG_TYPE,
            },
        )
    )


def _filter_ocr(result, min_confidence: int):
    """按置信度过滤 OCR 结果，返回与 RapidTable 匹配的 (boxes, txts, scores)。"""
    scores = tuple(float(s) for s in result.scores)
    keep = [i for i, s in enumerate(scores) if s * 100 >= min_confidence]
    if not keep:  # 全部低于阈值时保留原样，避免空输入
        return result.boxes, result.txts, result.scores
    boxes = result.boxes[keep]
    txts = tuple(result.txts[i] for i in keep)
    kept_scores = tuple(scores[i] for i in keep)
    return boxes, txts, kept_scores


@dataclass
class CellRef:
    """表格中的一个单元格：逻辑坐标与合并跨度，以及识别出的文字。"""

    row: int
    col: int
    rowspan: int
    colspan: int
    text: str


@dataclass
class TableGrid:
    """还原后的表格结构：网格尺寸 + 有序的单元格列表（含合并信息）。"""

    n_rows: int
    n_cols: int
    cells: list[CellRef]

    def to_frame(self) -> pd.DataFrame:
        """转成 DataFrame（合并单元格文字放在左上角，其余位置留空）。"""
        grid = [[""] * self.n_cols for _ in range(self.n_rows)]
        for cell in self.cells:
            grid[cell.row][cell.col] = cell.text
        return pd.DataFrame(grid)

    def to_html(self) -> str:
        """转成带边框和合并单元格的 HTML 表格，用于网页预览。"""
        td_map = {(cell.row, cell.col): cell for cell in self.cells}
        occupied: set[tuple[int, int]] = set()
        for cell in self.cells:
            for r in range(cell.row, cell.row + cell.rowspan):
                for c in range(cell.col, cell.col + cell.colspan):
                    occupied.add((r, c))

        rows_html: list[str] = []
        for r in range(self.n_rows):
            tds: list[str] = []
            for c in range(self.n_cols):
                cell = td_map.get((r, c))
                if cell is not None:
                    attrs: list[str] = []
                    if cell.rowspan > 1:
                        attrs.append(f'rowspan="{cell.rowspan}"')
                    if cell.colspan > 1:
                        attrs.append(f'colspan="{cell.colspan}"')
                    text = html.escape(cell.text).replace("\n", "<br>") if cell.text else "&nbsp;"
                    attr_str = f" {' '.join(attrs)}" if attrs else ""
                    tds.append(f"<td{attr_str}>{text}</td>")
                elif (r, c) in occupied:
                    continue  # 已被上方/左侧的合并单元格覆盖
                else:
                    tds.append("<td>&nbsp;</td>")
            rows_html.append("<tr>" + "".join(tds) + "</tr>")

        return (
            '<table border="1" cellspacing="0" cellpadding="0" '
            'style="border-collapse:collapse;font-size:13px;table-layout:auto;">'
            + "".join(rows_html)
            + "</table>"
        )

def _parse_table_html(html: str, logic_points) -> TableGrid:
    """把 rapid_table 输出的 HTML 与逻辑坐标还原成带合并信息的表格结构。

    logic_points 每一行是 [起始行, 结束行, 起始列, 结束列]，与 HTML 里的 <td> 一一对应，
    因此能精确还原 rowspan / colspan 的合并单元格。
    """
    try:
        tree = lxml_html.fromstring(html)
        tds = tree.xpath("//td")
    except Exception:
        return TableGrid(0, 0, [])

    cells: list[CellRef] = []
    n_rows = 0
    n_cols = 0
    for index, td in enumerate(tds):
        text = "".join(td.itertext()).strip()
        try:
            rowspan = int(td.get("rowspan", 1) or 1)
            colspan = int(td.get("colspan", 1) or 1)
        except (TypeError, ValueError):
            rowspan = colspan = 1

        if logic_points is not None and index < len(logic_points):
            r0, r1, c0, c1 = (int(v) for v in logic_points[index][:4])
            rowspan = max(rowspan, r1 - r0 + 1)
            colspan = max(colspan, c1 - c0 + 1)
            row, col = r0, c0
        else:
            # 兜底：没有逻辑坐标时按顺序平铺（不合并）。
            row, col = index // max(n_cols, 1), index % max(n_cols, 1)
            if n_cols == 0:
                n_cols = len(tds)

        cells.append(CellRef(row, col, rowspan, colspan, text))
        n_rows = max(n_rows, row + rowspan)
        n_cols = max(n_cols, col + colspan)

    return TableGrid(n_rows, n_cols, cells)


def _load_image_array(path: str | Path) -> np.ndarray | None:
    """读取图片为 BGR 数组（兼容非 ASCII 路径）。"""
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


def _clusters(indices: np.ndarray, gap: int) -> list[int]:
    """把离散像素坐标聚类为线段位置。"""
    groups: list[list[int]] = []
    for x in indices:
        if groups and x - groups[-1][-1] <= gap:
            groups[-1].append(int(x))
        else:
            groups.append([int(x)])
    return [int(sum(g) / len(g)) for g in groups]


def _detect_grid_lines(img: np.ndarray) -> tuple[list[int], list[int], np.ndarray, np.ndarray, int, int]:
    """检测表格的列/行边界线，返回 (col_edges, row_edges, vline_mask, hline_mask, h, w)。"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if img.ndim == 3 else img
    h, w = gray.shape[:2]
    bw = cv2.adaptiveThreshold(
        255 - gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 8
    )

    # 长竖线 → 列边界
    vkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, h // 10)))
    vline = cv2.morphologyEx(bw, cv2.MORPH_OPEN, vkernel)
    # 短竖线/横线，用于判断某个单元格是否被合并（局部是否存在分隔线）
    vline_any = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 15))
    )
    hkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, w // 6), 1))
    hline = cv2.morphologyEx(bw, cv2.MORPH_OPEN, hkernel)
    hline_any = cv2.morphologyEx(
        bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (15, 1))
    )

    col_edges = _clusters(np.where(vline.sum(axis=0) > h * 0.15)[0], 10)
    row_edges = _clusters(np.where(hline.sum(axis=1) > w * 0.5)[0], 8)
    return col_edges, row_edges, vline_any, hline_any, h, w


def _join_cell_texts(texts: list[tuple[float, float, str]]) -> str:
    """把同一单元格内的多条 OCR 文本按行拼接，行内用空格、换行用换行符。"""
    if not texts:
        return ""
    texts.sort()
    lines: list[list[str]] = []
    current: list[str] = []
    current_y: float | None = None
    for cy, _cx, txt in texts:
        if current_y is None or abs(cy - current_y) <= 8:
            current.append(txt)
            if current_y is None:
                current_y = cy
        else:
            lines.append(current)
            current = [txt]
            current_y = cy
    if current:
        lines.append(current)
    return "\n".join(" ".join(line) for line in lines)


def _rebuild_grid(
    col_edges: list[int],
    row_edges: list[int],
    vline_any: np.ndarray,
    hline_any: np.ndarray,
    boxes: Sequence,
    txts: Sequence[str],
    h: int,
    w: int,
) -> TableGrid:
    """按表格线把 OCR 文本归位到网格，并还原 rowspan / colspan 合并。"""
    ncols = len(col_edges) - 1
    nrows = len(row_edges) - 1
    covered = [[False] * ncols for _ in range(nrows)]

    centers: list[tuple[float, float]] = []
    for box in boxes:
        centers.append(
            ((box[:, 0].min() + box[:, 0].max()) / 2, (box[:, 1].min() + box[:, 1].max()) / 2)
        )

    cells: list[CellRef] = []
    for r in range(nrows):
        for c in range(ncols):
            if covered[r][c]:
                continue

            rowspan = 1
            while r + rowspan < nrows:
                y = row_edges[r + rowspan]
                x0, x1 = int(col_edges[c]) + 3, int(col_edges[c + 1]) - 3
                if x1 <= x0 or hline_any[max(0, y - 3): min(h, y + 4), x0:x1].any():
                    break
                rowspan += 1

            colspan = 1
            while c + colspan < ncols:
                x = col_edges[c + colspan]
                y0, y1 = int(row_edges[r]) + 3, int(row_edges[r + rowspan]) - 3
                if y1 <= y0 or vline_any[y0:y1, max(0, int(x) - 3): min(w, int(x) + 4)].any():
                    break
                colspan += 1

            x0, y0 = col_edges[c], row_edges[r]
            x1, y1 = col_edges[c + colspan], row_edges[r + rowspan]
            texts: list[tuple[float, float, str]] = []
            for (cx, cy), txt in zip(centers, txts):
                if x0 <= cx <= x1 and y0 <= cy <= y1:
                    texts.append((cy, cx, txt))
            cells.append(CellRef(r, c, rowspan, colspan, _join_cell_texts(texts)))

            for rr in range(r, r + rowspan):
                for cc in range(c, c + colspan):
                    covered[rr][cc] = True

    return TableGrid(nrows, ncols, cells)


def _try_grid_reconstruction(
    img: np.ndarray,
    ocr_result,
    min_confidence: int,
) -> TableGrid | None:
    """优先用表格线重建（针对长图效果好）；失败返回 None 以回退到模型。"""
    boxes = getattr(ocr_result, "boxes", None)
    txts = getattr(ocr_result, "txts", None)
    scores = getattr(ocr_result, "scores", None)
    if boxes is None or txts is None:
        return None

    try:
        col_edges, row_edges, vline_any, hline_any, h, w = _detect_grid_lines(img)
    except Exception:
        return None

    if len(col_edges) < 3 or len(row_edges) < 4:  # 至少 2 列、3 行
        return None

    keep = [i for i, s in enumerate(scores) if float(s) * 100 >= min_confidence]
    if not keep:
        keep = list(range(len(txts)))
    filtered_boxes = [boxes[i] for i in keep]
    filtered_txts = [txts[i] for i in keep]

    grid = _rebuild_grid(
        col_edges, row_edges, vline_any, hline_any, filtered_boxes, filtered_txts, h, w
    )
    if grid.n_rows < 2 or grid.n_cols < 2:
        return None
    return grid


def extract_tables_from_image(
    image_path: str | Path,
    ocr: RapidOCR,
    table_engine: RapidTable,
    min_confidence: int = MIN_CONFIDENCE,
) -> list[TableGrid]:
    """识别单张图片中的表格，返回还原后的 TableGrid 列表。"""
    ocr_result = ocr(str(image_path))

    # 竖长图（高 > 宽）会被结构模型压扁、导致列错位，优先用表格线重建。
    img = _load_image_array(image_path)
    if img is not None and img.shape[0] > img.shape[1]:
        grid = _try_grid_reconstruction(img, ocr_result, min_confidence)
        if grid is not None:
            return [grid]

    boxes, txts, scores = _filter_ocr(ocr_result, min_confidence)
    results = table_engine(str(image_path), ocr_results=[(boxes, txts, scores)])
    html_list = getattr(results, "pred_htmls", None) or []
    logic_points = getattr(results, "logic_points", None) or []
    if not html_list:
        return []

    grids: list[TableGrid] = []
    for index, html in enumerate(html_list):
        points = logic_points[index] if index < len(logic_points) else None
        grid = _parse_table_html(html, points)
        if grid.n_rows and grid.n_cols:
            grids.append(grid)
    return grids


def _sanitize_sheet_name(name: str, used: set[str]) -> str:
    """生成合法的 Excel 工作表名：最长 31 字符，去掉非法字符并去重。"""
    name = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "Sheet"
    base, index = name, 2
    while name in used:
        suffix = f"_{index}"
        name = f"{base[: 31 - len(suffix)]}{suffix}"
        index += 1
    used.add(name)
    return name


def _display_width(text: str) -> int:
    """粗略估算单元格文字显示宽度：中文字符按 2 个宽度计。"""
    return sum(2 if ord(char) > 127 else 1 for char in text)


def _write_grid_to_sheet(
    workbook,
    worksheet,
    grid: TableGrid,
    first_row_header: bool = False,
) -> None:
    """把 TableGrid 写入工作表，保留合并单元格、边框、列宽，可选首行表头样式。"""
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#4472C4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    cell_fmt = workbook.add_format(
        {"border": 1, "valign": "top", "text_wrap": True}
    )

    # 依据单元格内容估算列宽（合并单元格的文字分摊到各列）。
    widths = [2.0] * grid.n_cols
    for cell in grid.cells:
        if not cell.text:
            continue
        span = max(cell.colspan, 1)
        per_col = _display_width(cell.text) / span
        for c in range(cell.col, min(cell.col + span, grid.n_cols)):
            widths[c] = max(widths[c], per_col)
    for c in range(grid.n_cols):
        worksheet.set_column(c, c, max(4.0, min(widths[c] + 2, 60.0)))

    covered: set[tuple[int, int]] = set()
    for cell in grid.cells:
        fmt = header_fmt if (first_row_header and cell.row == 0) else cell_fmt
        r1 = cell.row + cell.rowspan - 1
        c1 = cell.col + cell.colspan - 1
        if cell.rowspan > 1 or cell.colspan > 1:
            region = [
                (r, c)
                for r in range(cell.row, r1 + 1)
                for c in range(cell.col, c1 + 1)
            ]
            if any(rc in covered for rc in region):
                # 模型偶发输出重叠的合并区域时，退化为普通单元格，避免写入失败。
                if (cell.row, cell.col) not in covered:
                    worksheet.write(cell.row, cell.col, cell.text or "", fmt)
                    covered.add((cell.row, cell.col))
                continue
            worksheet.merge_range(cell.row, cell.col, r1, c1, cell.text or "", fmt)
            covered.update(region)
        else:
            if (cell.row, cell.col) in covered:
                continue
            worksheet.write(cell.row, cell.col, cell.text or "", fmt)
            covered.add((cell.row, cell.col))

    if first_row_header and grid.n_rows:
        worksheet.freeze_panes(1, 0)


def table_to_excel_bytes(grid: TableGrid, first_row_header: bool = False) -> bytes:
    """把单个表格写到 Excel 文件（bytes），保留合并单元格、边框与列宽。"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("表1")
        _write_grid_to_sheet(workbook, worksheet, grid, first_row_header)
    buffer.seek(0)
    return buffer.read()


@dataclass
class ImageResult:
    """一张图片的识别结果：还原出的表格列表，或失败原因。"""

    path: Path
    grids: list[TableGrid] = field(default_factory=list)
    error: str | None = None


def images_to_grids(
    image_paths: Sequence[str | Path],
    min_confidence: int = MIN_CONFIDENCE,
    ocr: RapidOCR | None = None,
    table_engine: RapidTable | None = None,
    progress: ProgressCallback | None = None,
) -> list[ImageResult]:
    """识别多张图片，返回每张图片还原出的表格结构（供预览与导出复用）。"""
    ocr = ocr or build_ocr()
    table_engine = table_engine or build_table_engine()
    results: list[ImageResult] = []
    for index, path in enumerate(image_paths):
        path = Path(path)
        if progress:
            progress(index, len(image_paths), path.name)
        try:
            grids = extract_tables_from_image(
                path, ocr=ocr, table_engine=table_engine, min_confidence=min_confidence
            )
        except Exception as exc:  # 单张失败不影响其它图片
            results.append(ImageResult(path=path, error=str(exc)))
            continue
        results.append(ImageResult(path=path, grids=grids))
    return results


def grids_to_excel_bytes(
    results: Sequence[ImageResult],
    first_row_header: bool = False,
) -> bytes:
    """把识别结果写进一个 Excel 文件（返回 bytes）。

    每张识别出的表格一个工作表（命名「图片名_序号」），另有一个「汇总」工作表。
    """
    buffer = io.BytesIO()
    used_names: set[str] = {"汇总"}
    manifest: list[dict] = []

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        for result in results:
            if result.error:
                manifest.append({"文件": result.path.name, "状态": f"识别失败：{result.error}"})
                continue

            if not result.grids:
                manifest.append({"文件": result.path.name, "状态": "未检测到表格"})
                continue

            for table_index, grid in enumerate(result.grids, start=1):
                sheet = _sanitize_sheet_name(f"{result.path.stem}_{table_index}", used_names)
                worksheet = workbook.add_worksheet(sheet)
                _write_grid_to_sheet(workbook, worksheet, grid, first_row_header)
                manifest.append(
                    {
                        "文件": result.path.name,
                        "工作表": sheet,
                        "状态": "成功",
                        "行数": grid.n_rows,
                        "列数": grid.n_cols,
                    }
                )

        pd.DataFrame(manifest).to_excel(writer, sheet_name="汇总", index=False)

    buffer.seek(0)
    return buffer.read()


def images_to_excel_bytes(
    image_paths: Sequence[str | Path],
    min_confidence: int = MIN_CONFIDENCE,
    ocr: RapidOCR | None = None,
    table_engine: RapidTable | None = None,
    progress: ProgressCallback | None = None,
    first_row_header: bool = False,
) -> bytes:
    """把多张图片里的表格汇总到一个 Excel 文件（返回 bytes）。"""
    results = images_to_grids(
        image_paths,
        min_confidence=min_confidence,
        ocr=ocr,
        table_engine=table_engine,
        progress=progress,
    )
    return grids_to_excel_bytes(results, first_row_header=first_row_header)


def _collect_images(inputs: Sequence[str]) -> list[Path]:
    """展开输入：目录取其中所有受支持图片，文件则直接加入。"""
    files: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            files.extend(
                sorted(p for p in path.iterdir() if p.suffix.lower() in SUPPORTED_EXTS)
            )
        elif path.suffix.lower() in SUPPORTED_EXTS:
            files.append(path)
    return files


def main() -> None:
    parser = argparse.ArgumentParser(description="把多张表格图片识别并汇总到一个 Excel")
    parser.add_argument("inputs", nargs="+", help="图片路径，或包含图片的目录（可混用）")
    parser.add_argument("-o", "--output", default="汇总结果.xlsx", help="输出 Excel 路径")
    parser.add_argument("--min-confidence", type=int, default=MIN_CONFIDENCE, help="OCR 最低置信度（0–99）")
    parser.add_argument("--header", action="store_true", help="把每张表的第一行作为表头")
    args = parser.parse_args()

    files = _collect_images(args.inputs)
    if not files:
        print("未找到受支持的图片文件", file=sys.stderr)
        sys.exit(1)

    data = images_to_excel_bytes(
        files,
        min_confidence=args.min_confidence,
        first_row_header=args.header,
    )
    Path(args.output).write_bytes(data)
    print(f"已生成 {args.output}，共处理 {len(files)} 张图片")


if __name__ == "__main__":
    main()
