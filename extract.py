"""图片表格识别与汇总导出核心逻辑。

基于 RapidOCR（中文 OCR）+ RapidTable（表格结构识别，PaddleOCR SLANet-Plus 的 ONNX 版），
把图片中的表格识别后汇总到一个 Excel 文件。

可作为库使用，也可命令行调用：
    python extract.py 图片路径或目录... -o 汇总结果.xlsx
"""
from __future__ import annotations

import argparse
import io
import logging
import re
import sys
from pathlib import Path
from typing import Callable, Sequence

import pandas as pd

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


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """删除完全为空的行，保留表格列结构。"""
    return df.dropna(axis=0, how="all").reset_index(drop=True)


def extract_tables_from_image(
    image_path: str | Path,
    ocr: RapidOCR,
    table_engine: RapidTable,
    min_confidence: int = MIN_CONFIDENCE,
) -> list[pd.DataFrame]:
    """识别单张图片中的表格，返回 DataFrame 列表。"""
    boxes, txts, scores = _filter_ocr(ocr(str(image_path)), min_confidence)
    results = table_engine(str(image_path), ocr_results=[(boxes, txts, scores)])
    html_list = getattr(results, "pred_htmls", None) or []
    if not html_list:
        return []

    frames: list[pd.DataFrame] = []
    for html in html_list:
        try:
            frames.extend(pd.read_html(io.StringIO(html), header=None))
        except ValueError:  # HTML 中没有表格
            continue
    return [_clean_df(df) for df in frames]


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


def _promote_header(df: pd.DataFrame) -> pd.DataFrame:
    """把第一行提升为列名；空值列名用「列N」兜底。"""
    if df.empty:
        return df
    columns = []
    for index, value in enumerate(df.iloc[0], start=1):
        if value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == "":
            columns.append(f"列{index}")
        else:
            columns.append(str(value).strip())
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = columns
    return df


def images_to_excel_bytes(
    image_paths: Sequence[str | Path],
    min_confidence: int = MIN_CONFIDENCE,
    ocr: RapidOCR | None = None,
    table_engine: RapidTable | None = None,
    progress: ProgressCallback | None = None,
    first_row_header: bool = False,
) -> bytes:
    """把多张图片里的表格汇总到一个 Excel 文件（返回 bytes）。

    每张识别出的表格一个工作表（命名「图片名_序号」），另有一个「汇总」工作表。
    """
    ocr = ocr or build_ocr()
    table_engine = table_engine or build_table_engine()
    buffer = io.BytesIO()
    used_names: set[str] = set()
    manifest: list[dict] = []

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for index, path in enumerate(image_paths):
            path = Path(path)
            if progress:
                progress(index, len(image_paths), path.name)

            try:
                frames = extract_tables_from_image(
                    path, ocr=ocr, table_engine=table_engine, min_confidence=min_confidence
                )
            except Exception as exc:  # 单张失败不影响其它图片
                manifest.append({"文件": path.name, "状态": f"识别失败：{exc}"})
                continue

            if not frames:
                manifest.append({"文件": path.name, "状态": "未检测到表格"})
                continue

            for table_index, df in enumerate(frames, start=1):
                if first_row_header:
                    df = _promote_header(df)
                sheet = _sanitize_sheet_name(f"{path.stem}_{table_index}", used_names)
                df.to_excel(writer, sheet_name=sheet, index=False)
                manifest.append(
                    {
                        "文件": path.name,
                        "工作表": sheet,
                        "状态": "成功",
                        "行数": len(df),
                        "列数": len(df.columns),
                    }
                )

        pd.DataFrame(manifest).to_excel(writer, sheet_name="汇总", index=False)

    buffer.seek(0)
    return buffer.read()


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
