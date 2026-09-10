from __future__ import annotations

import argparse
import json
import os
import random
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from extractor import APP_VERSION, extract_files


SETTINGS_PATH = Path(os.getenv("APPDATA", Path.home())) / "病历自动抽取工具" / "settings.json"


def load_settings(path: Path = SETTINGS_PATH) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        date.fromisoformat(data["start_date"])
        date.fromisoformat(data["end_date"])
        return data
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        return {}


def save_settings(start: date, end: date, path: Path = SETTINGS_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"start_date": start.isoformat(), "end_date": end.isoformat()}, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_run_parameters(start: date, end: date, count: int, output_dir: str) -> None:
    if start > end:
        raise ValueError("检查开始日期不能晚于结束日期")
    if count < 1:
        raise ValueError("每文件抽取条数必须大于 0")
    if not output_dir.strip():
        raise ValueError("请选择输出目录")


def run_cli(args) -> None:
    output = extract_files(args.input, date.fromisoformat(args.start), date.fromisoformat(args.end), args.count, args.seed, args.output_dir)
    print(output)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"病历自动抽取工具 V{APP_VERSION}")
        self.geometry("820x600")
        self.minsize(720, 520)
        self.files: list[str] = []
        today = date.today()
        settings = load_settings()
        self.start_var = tk.StringVar(value=settings.get("start_date", f"{today.year}-01-01"))
        self.end_var = tk.StringVar(value=settings.get("end_date", today.isoformat()))
        self.count_var = tk.IntVar(value=5)
        self.seed_var = tk.StringVar(value=str(random.SystemRandom().randint(100000, 999999999)))
        self.output_var = tk.StringVar(value=str(Path.cwd() / "输出结果"))
        self.status_var = tk.StringVar(value="请选择 Excel 文件。")
        self._build()

    def _build(self):
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="病历自动抽取工具", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(frame, text="源文件只读；姓名完整保留；个人编号和身份证/证件号不输出。", foreground="#555555").pack(anchor="w", pady=(4, 12))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="添加 Excel", command=self.add_files).pack(side="left")
        ttk.Button(buttons, text="添加文件夹", command=self.add_folder).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="移除选中", command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(buttons, text="清空", command=self.clear_files).pack(side="left")
        self.listbox = tk.Listbox(frame, height=10, selectmode="extended")
        self.listbox.pack(fill="both", expand=True, pady=8)

        options = ttk.LabelFrame(frame, text="抽取参数", padding=10)
        options.pack(fill="x", pady=8)
        labels = (("检查开始日", self.start_var), ("检查结束日", self.end_var), ("每文件条数", self.count_var), ("随机种子", self.seed_var))
        for index, (label, variable) in enumerate(labels):
            ttk.Label(options, text=label).grid(row=index // 2, column=(index % 2) * 2, sticky="e", padx=5, pady=5)
            ttk.Entry(options, textvariable=variable, width=22).grid(row=index // 2, column=(index % 2) * 2 + 1, sticky="ew", padx=5, pady=5)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(3, weight=1)
        out = ttk.Frame(frame)
        out.pack(fill="x", pady=6)
        ttk.Label(out, text="输出目录").pack(side="left")
        ttk.Entry(out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out, text="选择", command=self.choose_output).pack(side="left")
        self.run_button = ttk.Button(frame, text="开始抽取", command=self.start_run)
        self.run_button.pack(anchor="e", pady=8)
        ttk.Label(frame, textvariable=self.status_var, foreground="#1F4E78").pack(anchor="w")

    def add_files(self):
        chosen = filedialog.askopenfilenames(filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        for file in chosen:
            if file not in self.files:
                self.files.append(file)
                self.listbox.insert("end", file)

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]

    def add_folder(self):
        chosen = filedialog.askdirectory()
        if not chosen:
            return
        for file in sorted(Path(chosen).glob("*.xlsx")):
            text = str(file)
            if text not in self.files:
                self.files.append(text)
                self.listbox.insert("end", text)

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")

    def choose_output(self):
        chosen = filedialog.askdirectory()
        if chosen:
            self.output_var.set(chosen)

    def start_run(self):
        try:
            if not self.files:
                raise ValueError("请至少选择一个 Excel 文件")
            start = date.fromisoformat(self.start_var.get().strip())
            end = date.fromisoformat(self.end_var.get().strip())
            count = int(self.count_var.get())
            seed = int(self.seed_var.get().strip())
            output_dir = self.output_var.get().strip()
            validate_run_parameters(start, end, count, output_dir)
            save_settings(start, end)
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.run_button.configure(state="disabled")
        self.status_var.set("正在抽取，请稍候……")
        threading.Thread(target=self._run, args=(tuple(self.files), start, end, count, seed, output_dir), daemon=True).start()

    def _run(self, files, start, end, count, seed, output_dir):
        try:
            output, errors = extract_files(files, start, end, count, seed, output_dir, return_errors=True)
            self.after(0, lambda: self._finished(output, errors))
        except Exception as exc:
            self.after(0, lambda error=exc: self._failed(error))

    def _finished(self, output: Path, errors: list[tuple[str, str]]):
        self.run_button.configure(state="normal")
        if errors:
            self.status_var.set(f"完成，但有 {len(errors)} 个文件失败：{output}")
            messagebox.showwarning("抽取完成（有异常）", f"结果已保存到：\n{output}\n\n有 {len(errors)} 个文件处理失败，请查看“异常明细”工作表。")
            return
        self.status_var.set(f"完成：{output}")
        messagebox.showinfo("抽取完成", f"结果已保存到：\n{output}")

    def _failed(self, exc: Exception):
        self.run_button.configure(state="normal")
        self.status_var.set("抽取失败。")
        messagebox.showerror("抽取失败", str(exc))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cli", action="store_true")
    parser.add_argument("--input", nargs="+")
    parser.add_argument("--start")
    parser.add_argument("--end")
    parser.add_argument("--count", type=int, default=5)
    parser.add_argument("--seed", type=int, default=20260826)
    parser.add_argument("--output-dir", default="输出结果")
    args = parser.parse_args()
    if args.cli:
        run_cli(args)
    else:
        App().mainloop()


if __name__ == "__main__":
    main()
