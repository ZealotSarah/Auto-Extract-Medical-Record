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

from extractor import APP_VERSION, EXTRACTION_MODES, MODE_DEPARTMENT_TOP10, MODE_RANDOM, extract_files


SETTINGS_PATH = Path(os.getenv("APPDATA", Path.home())) / "病历自动抽取工具" / "settings.json"


def load_settings(path: Path = SETTINGS_PATH) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        start = date.fromisoformat(data["start_date"])
        end = date.fromisoformat(data["end_date"])
        if start > end:
            raise ValueError("检查开始日期不能晚于结束日期")
        extraction_mode = data.get("extraction_mode", MODE_RANDOM)
        if extraction_mode not in EXTRACTION_MODES:
            raise ValueError("抽取方式无效")
        data["extraction_mode"] = extraction_mode
        return data
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        return {}


def save_settings(start: date, end: date, input_dir: str = "", output_dir: str = "", path: Path = SETTINGS_PATH, extraction_mode: str = MODE_RANDOM) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "input_dir": input_dir,
        "output_dir": output_dir,
        "extraction_mode": extraction_mode,
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_run_parameters(start: date, end: date, count: int, output_dir: str, extraction_mode: str = MODE_RANDOM) -> None:
    if start > end:
        raise ValueError("检查开始日期不能晚于结束日期")
    if extraction_mode not in EXTRACTION_MODES:
        raise ValueError("请选择有效的抽取方式")
    if extraction_mode == MODE_RANDOM and count < 1:
        raise ValueError("每文件抽取条数必须大于 0")
    if not output_dir.strip():
        raise ValueError("请选择输出目录")


def normalize_excel_paths(paths) -> list[str]:
    normalized = []
    seen = set()
    for value in paths:
        path = Path(value).resolve()
        key = str(path).casefold()
        if path.suffix.casefold() == ".xlsx" and key not in seen:
            normalized.append(str(path))
            seen.add(key)
    return normalized


def run_cli(args) -> None:
    output = extract_files(args.input, date.fromisoformat(args.start), date.fromisoformat(args.end), args.count, args.seed, args.output_dir, extraction_mode=args.mode)
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
        self.input_dir = settings.get("input_dir") or str(Path.cwd())
        self.count_var = tk.IntVar(value=5)
        self.seed_var = tk.StringVar(value=str(random.SystemRandom().randint(100000, 999999999)))
        self.output_var = tk.StringVar(value=settings.get("output_dir") or str(Path.cwd() / "输出结果"))
        self.mode_var = tk.StringVar(value=settings.get("extraction_mode", MODE_RANDOM))
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
        ttk.Label(options, text="抽取方式").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        mode_box = ttk.Combobox(options, textvariable=self.mode_var, values=EXTRACTION_MODES, state="readonly")
        mode_box.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=5)
        mode_box.bind("<<ComboboxSelected>>", self._mode_changed)
        labels = (("检查开始日", self.start_var), ("检查结束日", self.end_var), ("每文件条数", self.count_var), ("随机种子", self.seed_var))
        for index, (label, variable) in enumerate(labels):
            ttk.Label(options, text=label).grid(row=index // 2 + 1, column=(index % 2) * 2, sticky="e", padx=5, pady=5)
            entry = ttk.Entry(options, textvariable=variable, width=22)
            entry.grid(row=index // 2 + 1, column=(index % 2) * 2 + 1, sticky="ew", padx=5, pady=5)
            if variable in (self.count_var, self.seed_var):
                setattr(self, "count_entry" if variable is self.count_var else "seed_entry", entry)
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
        self._mode_changed()

    def _mode_changed(self, _event=None):
        state = "disabled" if self.mode_var.get() == MODE_DEPARTMENT_TOP10 else "normal"
        self.count_entry.configure(state=state)
        self.seed_entry.configure(state=state)

    def add_files(self):
        chosen = filedialog.askopenfilenames(initialdir=self.input_dir, filetypes=[("Excel 文件", "*.xlsx")])
        if chosen:
            self.input_dir = str(Path(chosen[0]).resolve().parent)
        self._add_paths(chosen)

    def _add_paths(self, paths):
        existing = {str(Path(file).resolve()).casefold() for file in self.files}
        for file in normalize_excel_paths(paths):
            if file.casefold() not in existing:
                self.files.append(file)
                self.listbox.insert("end", file)
                existing.add(file.casefold())

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]

    def add_folder(self):
        chosen = filedialog.askdirectory(initialdir=self.input_dir)
        if not chosen:
            return
        self.input_dir = str(Path(chosen).resolve())
        self._add_paths(sorted(Path(chosen).iterdir()))

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")

    def choose_output(self):
        chosen = filedialog.askdirectory(initialdir=self.output_var.get().strip() or str(Path.cwd()))
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
            extraction_mode = self.mode_var.get()
            validate_run_parameters(start, end, count, output_dir, extraction_mode)
            save_settings(start, end, self.input_dir, output_dir, extraction_mode=extraction_mode)
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.run_button.configure(state="disabled")
        self.status_var.set("正在抽取，请稍候……")
        threading.Thread(target=self._run, args=(tuple(self.files), start, end, count, seed, output_dir, extraction_mode), daemon=True).start()

    def _run(self, files, start, end, count, seed, output_dir, extraction_mode):
        try:
            output, errors = extract_files(files, start, end, count, seed, output_dir, return_errors=True, extraction_mode=extraction_mode)
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
    parser.add_argument("--mode", choices=EXTRACTION_MODES, default=MODE_RANDOM)
    parser.add_argument("--output-dir", default="输出结果")
    args = parser.parse_args()
    if args.cli:
        run_cli(args)
    else:
        App().mainloop()


if __name__ == "__main__":
    main()
