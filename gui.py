"""图片表格识别 → Excel：macOS 桌面可视化界面（无需命令行）。

双击「启动图片转Excel.command」即可打开本窗口。
"""
from __future__ import annotations

import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from extract import MIN_CONFIDENCE, SUPPORTED_EXTS, build_ocr, build_table_engine, images_to_excel_bytes

FILE_TYPES = [("图片文件", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"), ("所有文件", "*.*")]


class ImageToExcelApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("图片表格识别 → Excel")
        root.geometry("780x580")
        root.minsize(680, 480)

        self.files: list[Path] = []
        self.ocr = None
        self.table_engine = None
        self.queue: "queue.Queue[tuple]" = queue.Queue()
        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(
            self.root,
            text="把多张表格图片识别汇总成一个 Excel 文件",
            font=("", 14, "bold"),
        ).pack(anchor="w", padx=16, pady=(16, 2))

        ttk.Label(
            self.root,
            text="支持中文识别 · 图片格式：PNG / JPG / BMP / TIFF / WEBP",
            foreground="#666666",
        ).pack(anchor="w", padx=16, pady=(0, 12))

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=16)
        ttk.Button(btn_frame, text="➕ 添加图片", command=self.add_files).pack(side="left")
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected).pack(side="left", padx=(8, 0))
        ttk.Button(btn_frame, text="清空", command=self.clear).pack(side="left", padx=(8, 0))

        list_frame = ttk.Frame(self.root)
        list_frame.pack(fill="both", expand=True, padx=16, pady=12)
        self.listbox = tk.Listbox(list_frame, selectmode="extended", height=8)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        opts = ttk.LabelFrame(self.root, text="识别选项")
        opts.pack(fill="x", padx=16)
        self.header_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="首行作为表头（简单表格可勾选）", variable=self.header_var).pack(anchor="w", padx=8, pady=(6, 8))

        conf_row = ttk.Frame(opts)
        conf_row.pack(fill="x", padx=8, pady=(6, 8))
        ttk.Label(conf_row, text="OCR 最低置信度（0-99）").pack(side="left")
        self.conf_var = tk.IntVar(value=MIN_CONFIDENCE)
        ttk.Spinbox(conf_row, from_=0, to=99, textvariable=self.conf_var, width=6).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="请先添加图片")
        ttk.Label(self.root, textvariable=self.status_var).pack(anchor="w", padx=16, pady=(4, 0))

        self.start_btn = ttk.Button(self.root, text="开始识别并导出 Excel", command=self.start)
        self.start_btn.pack(pady=(8, 16), ipadx=12, ipady=4)

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(title="选择表格图片（可多选）", filetypes=FILE_TYPES)
        for name in selected:
            path = Path(name)
            if path not in self.files:
                self.files.append(path)
        self._refresh_list()

    def remove_selected(self) -> None:
        for index in reversed(self.listbox.curselection()):
            del self.files[index]
        self._refresh_list()

    def clear(self) -> None:
        self.files = []
        self._refresh_list()

    def _refresh_list(self) -> None:
        self.listbox.delete(0, "end")
        for path in self.files:
            self.listbox.insert("end", str(path))
        self.status_var.set(f"已选择 {len(self.files)} 张图片" if self.files else "请先添加图片")

    def start(self) -> None:
        if not self.files:
            messagebox.showwarning("提示", "请先添加至少一张图片")
            return
        out = filedialog.asksaveasfilename(
            title="保存汇总 Excel",
            defaultextension=".xlsx",
            initialfile="汇总结果.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not out:
            return

        # 主线程读取参数并复制文件列表，避免后台线程直接操作 tk 变量
        files = list(self.files)
        min_conf = int(self.conf_var.get())
        first_row_header = self.header_var.get()

        self.start_btn.configure(state="disabled")
        self.status_var.set("正在加载识别引擎（首次约几秒）…")
        threading.Thread(target=self._run, args=(Path(out), files, min_conf, first_row_header), daemon=True).start()
        self.root.after(100, self._poll_queue)

    def _run(self, out: Path, files: list[Path], min_conf: int, first_row_header: bool) -> None:
        try:
            if self.ocr is None:
                self.ocr = build_ocr()
            if self.table_engine is None:
                self.table_engine = build_table_engine()

            def progress(done: int, total: int, name: str) -> None:
                self.queue.put(("progress", done, total, name))

            data = images_to_excel_bytes(
                files,
                min_confidence=min_conf,
                first_row_header=first_row_header,
                ocr=self.ocr,
                table_engine=self.table_engine,
                progress=progress,
            )
            out.write_bytes(data)
            self.queue.put(("done", out))
        except Exception as exc:  # 捕获处理错误，避免窗口崩溃
            self.queue.put(("error", exc))

    def _poll_queue(self) -> None:
        """主线程轮询后台消息队列，线程安全地更新界面。"""
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    _, done, total, name = msg
                    self.status_var.set(f"识别中 {done + 1}/{total}：{name}")
                elif kind == "done":
                    self._finished(msg[1])
                    return
                elif kind == "error":
                    self._failed(msg[1])
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _finished(self, out: Path) -> None:
        self.start_btn.configure(state="normal")
        self.status_var.set(f"完成：{out}")
        if messagebox.askyesno("识别完成", f"已生成：\n{out}\n\n是否在 Finder 中显示？"):
            subprocess.Popen(["open", "-R", str(out)])

    def _failed(self, exc: Exception) -> None:
        self.start_btn.configure(state="normal")
        self.status_var.set("识别失败")
        messagebox.showerror("出错了", str(exc))


def main() -> None:
    root = tk.Tk()
    ImageToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
