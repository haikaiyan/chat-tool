"""「合并」选项卡：按 姓名 + 对方姓名 + 日期 把多条聊天合并成一条。"""
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from merger import DEFAULT_SHEET, inspect_sheet, merge_sheet

RULE_TEXT = (
    "合并规则：① 取第 1 个工作表（表 1）；"
    "② 当「姓名」「对方姓名」「消息时间的年月日」三者都相同时合并为一行；"
    "③ 消息内容按原顺序拼接，每条前带 [2026-09-01 07:32:13 说话人] 时间戳"
    "（发送标注姓名、接收标注对方姓名）；"
    "④ 消息时间只留年月日；⑤ 其余列取该组第一条；⑥ 空行跳过，原数据不动。"
)


class MergeTab(ttk.Frame):
    def __init__(self, master, app=None, notebook=None):
        super().__init__(master)
        self.app = app
        self.notebook = notebook
        self.file_path = ""
        self.merging = False
        self.cancel_event = threading.Event()
        self.q = queue.Queue()
        self.save_mode = tk.StringVar(value="inplace")  # inplace / newfile
        self.file_var = tk.StringVar()
        self.sheet_var = tk.StringVar(value=DEFAULT_SHEET)
        self.dest_var = tk.StringVar()
        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 5}

        f1 = ttk.LabelFrame(self, text="① 原始聊天记录 Excel")
        f1.pack(fill="x", **pad)
        ttk.Entry(f1, textvariable=self.file_var).pack(side="left", fill="x", expand=True, padx=(8, 4), pady=8)
        ttk.Button(f1, text="浏览…", command=self.choose_file).pack(side="left", padx=4)
        self.file_info = ttk.Label(f1, text="未选择文件", foreground="#666666")
        self.file_info.pack(side="left", padx=8)

        f2 = ttk.LabelFrame(self, text="② 合并规则")
        f2.pack(fill="x", **pad)
        ttk.Label(f2, text=RULE_TEXT, wraplength=980, justify="left", foreground="#555555").pack(
            anchor="w", padx=10, pady=(6, 4)
        )
        self.col_info = ttk.Label(f2, text="识别到的列：—", foreground="#1a7f37")
        self.col_info.pack(anchor="w", padx=10, pady=(0, 4))
        bar = ttk.Frame(f2)
        bar.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(bar, text="结果工作表名称：").pack(side="left")
        ttk.Entry(bar, textvariable=self.sheet_var, width=20).pack(side="left", padx=4)
        ttk.Label(bar, text="（重名会自动加 _2，不会覆盖已有工作表）", foreground="#888888").pack(side="left", padx=8)

        f3 = ttk.LabelFrame(self, text="③ 保存到")
        f3.pack(fill="x", **pad)
        ttk.Radiobutton(
            f3,
            text="在该文件里新建一个工作表（原文件其他内容不变）",
            variable=self.save_mode,
            value="inplace",
            command=self._toggle_dest,
        ).pack(anchor="w", padx=10, pady=(6, 2))
        row = ttk.Frame(f3)
        row.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Radiobutton(
            row,
            text="另存为新文件：",
            variable=self.save_mode,
            value="newfile",
            command=self._toggle_dest,
        ).pack(side="left")
        self.dest_entry = ttk.Entry(row, textvariable=self.dest_var, state="disabled")
        self.dest_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.dest_btn = ttk.Button(row, text="浏览…", command=self.choose_dest, state="disabled")
        self.dest_btn.pack(side="left")

        f4 = ttk.Frame(self)
        f4.pack(fill="x", **pad)
        self.btn_merge = ttk.Button(f4, text="开始合并", command=self.start_merge)
        self.btn_merge.pack(side="left", padx=8)
        self.btn_stop = ttk.Button(f4, text="停止", command=self.stop_merge, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.progress = ttk.Progressbar(f4, mode="determinate", length=360)
        self.progress.pack(side="left", padx=12)
        self.status = ttk.Label(f4, text="就绪")
        self.status.pack(side="left", padx=6)

        f5 = ttk.Frame(self)
        f5.pack(fill="x", **pad)
        self.result_label = ttk.Label(f5, text="尚未合并", foreground="#666666")
        self.result_label.pack(side="left", padx=8)
        self.btn_open = ttk.Button(f5, text="用筛查打开该文件", command=self.open_in_scan, state="disabled")
        self.btn_open.pack(side="right", padx=8)

        self._toggle_dest()

    # ---------------- 文件 ----------------
    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择原始聊天记录 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.set_file(path)

    def set_file(self, path):
        self.file_path = path
        self.file_var.set(path)
        self.file_info.configure(text="识别中…", foreground="#666666")
        self.update_idletasks()
        try:
            info = inspect_sheet(path)
            cols = info["columns"]
            self.file_info.configure(
                text=f"表1「{info['sheet']}」 约 {info['rows']} 行", foreground="#1a7f37"
            )
            self.col_info.configure(
                text=f"识别到的列：姓名={cols['name']}、对方姓名={cols['peer']}、"
                f"消息时间={cols['time']}、消息内容={cols['content']}、消息角色={cols['role']}"
            )
            if not self.dest_var.get().strip():
                src = Path(path)
                self.dest_var.set(str(src.with_name(src.stem + "_合并.xlsx")))
        except Exception as e:
            self.file_info.configure(text=f"读取失败：{e}", foreground="#c00000")
            self.col_info.configure(text="识别到的列：—")

    def _toggle_dest(self):
        on = self.save_mode.get() == "newfile"
        self.dest_entry.configure(state="normal" if on else "disabled")
        self.dest_btn.configure(state="normal" if on else "disabled")

    def choose_dest(self):
        initialdir = str(Path(self.file_path).parent) if self.file_path else os.path.expanduser("~")
        dest = filedialog.asksaveasfilename(
            title="合并后保存到",
            defaultextension=".xlsx",
            initialfile="合并结果.xlsx",
            initialdir=initialdir,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if dest:
            self.dest_var.set(dest)

    # ---------------- 合并 ----------------
    def start_merge(self):
        if self.merging:
            return
        if not self.file_path or not os.path.exists(self.file_path):
            messagebox.showwarning("提示", "请先选择原始聊天记录 Excel。")
            return
        dest = None
        if self.save_mode.get() == "newfile":
            dest = self.dest_var.get().strip()
            if not dest:
                messagebox.showwarning("提示", "请选择另存为的文件路径。")
                return

        self.merging = True
        self.cancel_event.clear()
        self.btn_merge.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_open.configure(state="disabled")
        self.progress["value"] = 0
        self.status.configure(text="正在读取…")

        thread = threading.Thread(
            target=self._worker,
            args=(self.file_path, dest, self.sheet_var.get().strip() or DEFAULT_SHEET),
            daemon=True,
        )
        thread.start()
        self.after(100, self._poll)

    def _worker(self, path, dest, sheet_title):
        def progress(done, total, stage):
            self.q.put(("progress", done, total, stage))

        try:
            res = merge_sheet(
                path,
                dest_path=dest,
                sheet_title=sheet_title,
                progress=progress,
                should_cancel=self.cancel_event.is_set,
            )
            self.q.put(("done", res))
        except Exception as e:
            self.q.put(("error", f"{type(e).__name__}: {e}"))

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, done, total, stage = item
                    self.progress["value"] = (done / total * 100) if total else 0
                    self.status.configure(text=f"{stage} {done}/{total}")
                elif kind == "done":
                    self._on_done(item[1])
                    return
                elif kind == "error":
                    self._finish()
                    self.status.configure(text="合并失败")
                    messagebox.showerror("合并失败", item[1])
                    return
        except queue.Empty:
            pass
        if self.merging:
            self.after(100, self._poll)

    def _on_done(self, res):
        self._finish()
        if not res.rows_after:
            msg = "没有可合并的数据（已取消或表 1 为空）。" if res.cancelled else "没有可合并的数据。"
            self.status.configure(text="完成")
            self.result_label.configure(text=msg, foreground="#666666")
            messagebox.showwarning("提示", msg)
            return

        note = f"已停止：已写入 {res.rows_after} 行" if res.cancelled else res.summary
        text = f"{note}，结果在「{res.sheet}」工作表"
        if res.overflow:
            text += f"；⚠ {len(res.overflow)} 行超过 32767 字（最长 {res.longest} 字），Excel 会截断"
        self.status.configure(text="完成")
        self.result_label.configure(text=text, foreground="#1a7f37")
        self.btn_open.configure(state="normal")
        self.last_dest = res.dest
        messagebox.showinfo("合并完成", f"{text}\n\n文件：{res.dest}")

    def _finish(self):
        self.merging = False
        self.btn_merge.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.progress["value"] = 100

    def stop_merge(self):
        if self.merging:
            self.cancel_event.set()
            self.status.configure(text="正在停止…")

    def open_in_scan(self):
        """把合并后的文件填进筛查页的 ① 待检查文件，并切到筛查页。"""
        if not getattr(self, "last_dest", ""):
            return
        if self.app is not None:
            self.app.set_file(self.last_dest)
            if self.notebook is not None:
                self.notebook.select(0)
