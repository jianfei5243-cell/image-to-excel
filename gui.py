"""图片表格识别 → Excel：macOS 桌面可视化工具。

双击「启动图片转Excel.command」即可打开本窗口。
"""
from __future__ import annotations

import io
import os
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from extract import (
    MIN_CONFIDENCE,
    SUPPORTED_EXTS,
    build_ocr,
    build_table_engine,
    extract_tables_from_image,
    images_to_excel_bytes,
    table_to_excel_bytes,
)

FILE_TYPES = [("图片文件", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"), ("所有文件", "*.*")]
THUMB_W, THUMB_H = 100, 75


class DropZone(tk.Frame):
    """提示区域（已移除拖拽功能）。"""

    def __init__(self, parent, on_add_files):
        super().__init__(parent, bg="#f0f4fa", relief="solid", borderwidth=1)
        self.pack(fill="x", padx=16, pady=(16, 8))
        ttk.Label(
            self, text="点击下方按钮添加图片", font=("", 11), foreground="#666"
        ).pack(fill="x", pady=10)
        ttk.Button(self, text="选择图片文件", command=on_add_files).pack(pady=(0, 10))


class ThumbView(tk.Frame):
    """缩略图预览区域。"""

    def __init__(self, parent):
        super().__init__(parent, bg="#f5f5f5")
        self.pack(fill="both", expand=True, padx=16, pady=(0, 8))

    def show(self, paths):
        for child in self.winfo_children():
            child.destroy()
        if not paths:
            ttk.Label(
                self, text="尚未添加图片", foreground="#999", font=("", 11)
            ).pack(pady=20)
            return
        cols = 5
        for idx, p in enumerate(paths):
            row, col = divmod(idx, cols)
            try:
                thumb = self._make_thumb(p)
                if thumb is not None and thumb.height() > 0:
                    lbl = ttk.Label(self, image=thumb, cursor="hand2")
                    lbl.image = thumb
                    lbl.grid(row=row, column=col, padx=3, pady=3)
            except Exception:
                pass

    @staticmethod
    def _make_thumb(path):
        try:
            from PIL import Image

            with open(path, "rb") as f:
                img = Image.open(f)
                img.thumbnail((THUMB_W, THUMB_H))
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                return tk.PhotoImage(data=buf.read())
        except Exception:
            pass
        return None


class ImageToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片表格识别 → Excel")
        self.root.geometry("680x520")
        self.root.minsize(560, 380)

        self.files: list[Path] = []
        self.ocr = None
        self.table_engine = None
        self.queue = queue.Queue()

        self._build_ui()

    # ---- UI 构建 ----

    def _build_ui(self):
        # 提示标签
        ttk.Label(
            self.root, text="添加图片：点击下方按钮，或使用文件夹功能一次性添加",
            font=("", 11), foreground="#666",
        ).pack(fill="x", padx=16, pady=(16, 8))

        # 缩略图
        self.thumb_view = ThumbView(self.root)
        self.thumb_view.show(self.files)

        # 文件列表
        list_frame = ttk.LabelFrame(self.root, text="已添加的图片")
        list_frame.pack(fill="x", padx=16, pady=(0, 4))
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill="x", padx=6, pady=4)

        self.tree = ttk.Treeview(tree_frame, columns=("path",), show="headings", height=6)
        self.tree.heading("path", text="文件路径")
        self.tree.column("path", width=480, anchor="w")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="x", expand=True)
        scroll.pack(side="right", fill="y")

        # 按钮行
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=16, pady=(0, 4))
        ttk.Button(btn_frame, text="添加图片", command=self.add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_folder).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="清空全部", command=self.clear).pack(side="left")

        # 选项
        opts = ttk.LabelFrame(self.root, text="选项")
        opts.pack(fill="x", padx=16, pady=(4, 0))

        row1 = ttk.Frame(opts)
        row1.pack(fill="x", padx=8, pady=(8, 4))
        self.header_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row1, text="首行作为表头", variable=self.header_var).pack(side="left", padx=(0, 16))

        conf_frame = ttk.Frame(row1)
        conf_frame.pack(side="left")
        ttk.Label(conf_frame, text="OCR 置信度").pack(side="left")
        self.conf_var = tk.IntVar(value=MIN_CONFIDENCE)
        ttk.Spinbox(conf_frame, from_=0, to=99, textvariable=self.conf_var, width=5).pack(side="left", padx=(4, 0))

        row2 = ttk.Frame(opts)
        row2.pack(fill="x", padx=8, pady=(4, 8))
        self.merge_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="合并为单个 Excel 文件（多工作表）", variable=self.merge_var).pack(side="left")

        # 状态 + 进度
        self.status_var = tk.StringVar(value="请先添加图片")
        ttk.Label(self.root, textvariable=self.status_var, foreground="#666", wraplength=600).pack(
            anchor="w", padx=16, pady=(4, 0)
        )
        self.progress = ttk.Progressbar(self.root, mode="determinate", length=500)
        self.progress.pack(padx=16, pady=(4, 4))
        self.start_btn = ttk.Button(self.root, text="开始识别并导出 Excel", command=self.start)
        self.start_btn.pack(pady=(4, 16), ipadx=16, ipady=4)

    # ---- 文件管理 ----

    def add_files(self):
        selected = filedialog.askopenfilenames(title="选择表格图片", filetypes=FILE_TYPES)
        for name in selected:
            path = Path(name)
            if path not in self.files:
                self.files.append(path)
        self._refresh()

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择包含图片的文件夹")
        if not folder:
            return
        folder_path = Path(folder)
        added = 0
        for ext in SUPPORTED_EXTS:
            for p in folder_path.glob(f"*{ext}"):
                if p not in self.files:
                    self.files.append(p)
                    added += 1
        self._refresh()
        if added:
            self.status_var.set(f"已从文件夹添加了 {added} 张图片")

    def remove_selected(self):
        selection = self.tree.selection()
        for item in reversed(selection):
            idx = self.tree.index(item)
            if 0 <= idx < len(self.files):
                del self.files[idx]
        self._refresh()

    def clear(self):
        self.files = []
        self._refresh()

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for path in self.files:
            self.tree.insert("", "end", values=(str(path),))
        self.thumb_view.show(self.files)
        count = len(self.files)
        if count == 0:
            self.status_var.set("请先添加图片")
        elif count == 1:
            self.status_var.set("已添加 1 张图片")
        else:
            self.status_var.set(f"已添加 {count} 张图片")

    # ---- 识别 + 导出 ----

    def start(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加至少一张图片")
            return

        self.start_btn.configure(state="disabled")
        self.status_var.set("正在加载识别引擎…")
        self.progress.configure(value=0, maximum=max(len(self.files), 1))

        files = list(self.files)
        min_conf = int(self.conf_var.get())
        first_row_header = self.header_var.get()
        merge = self.merge_var.get()

        t = threading.Thread(target=self._run, args=(files, min_conf, first_row_header, merge), daemon=True)
        t.start()
        self.root.after(100, self._poll_queue)

    def _run(self, files, min_conf, first_row_header, merge):
        try:
            if self.ocr is None:
                self.ocr = build_ocr()
            if self.table_engine is None:
                self.table_engine = build_table_engine()

            def progress(done, total, name):
                self.queue.put(("progress", done, total, name))

            if merge:
                data = images_to_excel_bytes(
                    files,
                    min_confidence=min_conf,
                    first_row_header=first_row_header,
                    ocr=self.ocr,
                    table_engine=self.table_engine,
                    progress=progress,
                )
                self.queue.put(("merge_done", data))
            else:
                self.queue.put(("batch_start", len(files)))
                success, fail = 0, 0
                for i, path in enumerate(files):
                    self.queue.put(("progress", i, len(files), path.name))
                    try:
                        grids = extract_tables_from_image(
                            path,
                            ocr=self.ocr,
                            table_engine=self.table_engine,
                            min_confidence=min_conf,
                        )
                        if not grids:
                            fail += 1
                            continue
                        for j, grid in enumerate(grids, start=1):
                            buf = table_to_excel_bytes(grid, first_row_header=first_row_header)
                            out = Path(f"{path.stem}_{j + 1}.xlsx")
                            out.write_bytes(buf)
                        success += 1
                    except Exception:
                        fail += 1
                self.queue.put(("batch_done", success, fail))
        except Exception as exc:
            self.queue.put(("error", exc))

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    _, done, total, name = msg
                    self.progress.configure(value=done + 1)
                    self.status_var.set(f"识别中 {done + 1}/{total}：{name}")
                elif kind == "merge_done":
                    data = msg[1]
                    out = filedialog.asksaveasfilename(
                        title="保存汇总 Excel",
                        defaultextension=".xlsx",
                        initialfile="汇总结果.xlsx",
                        filetypes=[("Excel 文件", "*.xlsx")],
                    )
                    if out:
                        Path(out).write_bytes(data)
                        self._on_done(Path(out))
                    else:
                        self.start_btn.configure(state="normal")
                        self.status_var.set("已取消保存")
                    return
                elif kind == "batch_start":
                    total = msg[1]
                    self.progress.configure(value=0, maximum=total)
                    self.status_var.set(f"开始处理 {total} 张图片…")
                elif kind == "batch_done":
                    success, fail = msg[1], msg[2]
                    self.start_btn.configure(state="normal")
                    self.progress.configure(value=0)
                    msg_text = f"完成！成功 {success} 张"
                    if fail:
                        msg_text += f"，失败 {fail} 张"
                    self.status_var.set(msg_text)
                    messagebox.showinfo("完成", msg_text)
                    return
                elif kind == "error":
                    self.start_btn.configure(state="normal")
                    self.progress.configure(value=0)
                    self.status_var.set("识别失败")
                    messagebox.showerror("出错了", str(msg[1]))
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_done(self, out: Path):
        self.start_btn.configure(state="normal")
        self.progress.configure(value=0)
        self.status_var.set(f"完成：{out}")
        if messagebox.askyesno("识别完成", f"已生成：\n{out}\n\n是否在 Finder 中显示？"):
            subprocess.Popen(["open", "-R", str(out)])
def main():
    root = tk.Tk()
    ImageToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
