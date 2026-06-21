from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .pipeline import run_pipeline


class BalanceGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("MRP 平衡表工具")
        self.root.geometry("860x520")
        self.root.minsize(760, 460)

        self.input_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.reply_sources_var = tk.StringVar(value="未选择旧平衡表")
        self.status_var = tk.StringVar(value="请选择文件后开始生成。")
        self.apply_suggestion_exclusions_var = tk.BooleanVar(value=True)
        self.reply_source_paths: list[str] = []

        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._started_at = 0.0

        self._build_layout()
        self.root.after(200, self._poll_queue)

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(container, text="MRP 平衡表生成", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        hint = ttk.Label(
            container,
            text="选择输入文件、模板文件和输出路径后，点击“开始生成”。",
        )
        hint.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))

        self._add_path_row(
            container,
            row=2,
            label="输入文件",
            variable=self.input_var,
            browse_command=self._choose_input,
            button_text="选择输入",
        )
        self._add_path_row(
            container,
            row=3,
            label="模板文件",
            variable=self.template_var,
            browse_command=self._choose_template,
            button_text="选择模板",
        )
        self._add_path_row(
            container,
            row=4,
            label="输出文件",
            variable=self.output_var,
            browse_command=self._choose_output,
            button_text="选择输出",
        )
        self._add_path_row(
            container,
            row=5,
            label="旧平衡表",
            variable=self.reply_sources_var,
            browse_command=self._choose_reply_sources,
            button_text="选择旧表",
            readonly=True,
        )

        ttk.Checkbutton(
            container,
            text="寤鸿鎺掍骇搴旂敤鎺掗櫎娓呭崟",
            variable=self.apply_suggestion_exclusions_var,
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 8))

        exclude_hint = ttk.Label(
            container,
            text="排除清单不用单独选文件。请直接在输入工作簿里新增工作表“物料编码排除清单”，"
            "把要排除的料号填进去即可，程序会自动读取。"
            "如果选择了旧平衡表，程序会只回填今天之后日期列的“采购答复”。",
            wraplength=760,
            foreground="#666666",
        )
        exclude_hint.grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 10))

        action_frame = ttk.Frame(container)
        action_frame.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        action_frame.columnconfigure(0, weight=1)

        self.run_button = ttk.Button(action_frame, text="开始生成", command=self._start)
        self.run_button.grid(row=0, column=0, sticky="w")

        self.open_output_button = ttk.Button(
            action_frame,
            text="打开输出目录",
            command=self._open_output_dir,
        )
        self.open_output_button.grid(row=0, column=1, padx=(8, 0), sticky="w")

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(4, 8))

        status = ttk.Label(container, textvariable=self.status_var)
        status.grid(row=10, column=0, columnspan=3, sticky="w")

        log_label = ttk.Label(container, text="运行日志")
        log_label.grid(row=11, column=0, columnspan=3, sticky="w", pady=(16, 6))

        self.log_text = tk.Text(container, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=12, column=0, columnspan=3, sticky="nsew")
        container.rowconfigure(12, weight=1)

    def _add_path_row(
        self,
        parent: ttk.Frame,
        *,
        row: int,
        label: str,
        variable: tk.StringVar,
        browse_command,
        button_text: str,
        readonly: bool = False,
    ) -> None:
        ttk.Label(parent, text=label, width=10).grid(row=row, column=0, sticky="w", pady=6)
        entry = ttk.Entry(parent, textvariable=variable, state="readonly" if readonly else "normal")
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=6)
        ttk.Button(parent, text=button_text, command=browse_command).grid(row=row, column=2, sticky="ew", pady=6)

    def _choose_input(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择输入工作簿",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_var.set(file_path)
        if not self.output_var.get():
            suggested = Path(file_path).with_name(f"{Path(file_path).stem}_平衡表.xlsx")
            self.output_var.set(str(suggested))

    def _choose_template(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择模板工作簿",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if file_path:
            self.template_var.set(file_path)

    def _choose_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="选择输出路径",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if file_path:
            self.output_var.set(file_path)

    def _choose_reply_sources(self) -> None:
        file_paths = filedialog.askopenfilenames(
            title="选择已填写的旧平衡表",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not file_paths:
            return
        self.reply_source_paths = list(file_paths)
        if len(file_paths) == 1:
            summary = Path(file_paths[0]).name
        else:
            summary = f"已选择 {len(file_paths)} 个旧平衡表"
        self.reply_sources_var.set(summary)

    def _merge_selected_paths(self, existing_paths: list[str], new_paths) -> list[str]:
        merged: list[str] = []
        seen: set[str] = set()
        for path in list(existing_paths or []) + list(new_paths or []):
            normalized = str(Path(path).expanduser().resolve()).lower()
            if normalized in seen:
                continue
            seen.add(normalized)
            merged.append(str(path))
        return merged

    def _choose_reply_sources(self) -> None:
        file_paths = filedialog.askopenfilenames(
            title="选择已填写的旧平衡表",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not file_paths:
            return
        self.reply_source_paths = self._merge_selected_paths(self.reply_source_paths, file_paths)
        self.reply_sources_var.set(f"已选择 {len(self.reply_source_paths)} 个旧平衡表")

    def _start(self) -> None:
        if self._worker and self._worker.is_alive():
            return

        input_path = Path(self.input_var.get().strip())
        template_path = Path(self.template_var.get().strip()) if self.template_var.get().strip() else None
        output_path = Path(self.output_var.get().strip())
        reply_source_paths = [Path(path) for path in self.reply_source_paths]
        apply_suggestion_exclusions = bool(self.apply_suggestion_exclusions_var.get())

        if not input_path.exists():
            messagebox.showerror("输入文件错误", "请选择有效的输入工作簿。")
            return
        if template_path and not template_path.exists():
            messagebox.showerror("模板文件错误", "请选择有效的模板工作簿，或清空模板路径。")
            return
        if not output_path.name:
            messagebox.showerror("输出路径错误", "请选择有效的输出文件路径。")
            return

        self._started_at = time.time()
        self._set_running(True)
        self._append_log(f"开始生成：{input_path}")
        if template_path:
            self._append_log(f"使用模板：{template_path}")
        self._append_log(f"输出路径：{output_path}")
        if reply_source_paths:
            self._append_log(f"旧平衡表：{len(reply_source_paths)} 个，只带入今天之后的采购答复")

        self._worker = threading.Thread(
            target=self._run_job,
            args=(input_path, output_path, template_path, reply_source_paths, apply_suggestion_exclusions),
            daemon=True,
        )
        self._worker.start()

    def _run_job(
        self,
        input_path: Path,
        output_path: Path,
        template_path: Path | None,
        reply_source_paths: list[Path],
        apply_suggestion_exclusions: bool,
    ) -> None:
        try:
            result = run_pipeline(
                input_path,
                output_path,
                template_path,
                carry_forward_paths=reply_source_paths,
                apply_suggestion_exclusions=apply_suggestion_exclusions,
                progress_callback=lambda message: self._queue.put(("progress", message)),
            )
            self._queue.put(("success", {"output_path": str(output_path), "result": result}))
        except Exception:
            self._queue.put(("error", traceback.format_exc()))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "progress":
                    self.status_var.set(str(payload))
                    self._append_log(str(payload))
                elif kind == "success":
                    elapsed = time.time() - self._started_at
                    self._set_running(False)
                    output_path = payload["output_path"]
                    result = payload["result"]
                    carried_count = int(getattr(result, "carried_reply_cell_count", 0) or 0)
                    carried_material_count = int(getattr(result, "carried_reply_material_count", 0) or 0)
                    carried_file_count = int(getattr(result, "carried_reply_file_count", 0) or 0)
                    summary = f"生成完成，用时 {elapsed:.1f} 秒。"
                    if carried_count:
                        summary = f"{summary} 已带入采购答复 {carried_count} 格。"
                    self.status_var.set(summary)
                    self._append_log(f"生成完成：{output_path}")
                    if carried_count:
                        self._append_log(
                            f"已带入采购答复 {carried_count} 格，覆盖料号 {carried_material_count} 个，来源文件 {carried_file_count} 个"
                        )
                    messagebox.showinfo("完成", f"文件已生成：\n{output_path}")
                elif kind == "error":
                    self._set_running(False)
                    self.status_var.set("生成失败，请查看日志。")
                    self._append_log(payload)
                    messagebox.showerror("生成失败", "运行时出错，详细信息已写入日志。")
        except queue.Empty:
            pass
        finally:
            self.root.after(200, self._poll_queue)

    def _set_running(self, running: bool) -> None:
        if running:
            self.run_button.configure(state="disabled")
            self.progress.start(10)
            self.status_var.set("正在生成，请等待。")
        else:
            self.run_button.configure(state="normal")
            self.progress.stop()

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_output_dir(self) -> None:
        output_text = self.output_var.get().strip()
        target = Path(output_text).parent if output_text else Path.cwd()
        target.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(target)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(target)], check=True)
            else:
                subprocess.run(["xdg-open", str(target)], check=True)
        except Exception as exc:
            messagebox.showerror("打开目录失败", str(exc))


def main() -> int:
    root = tk.Tk()
    ttk.Style(root).theme_use("aqua" if sys.platform == "darwin" else "vista")
    BalanceGuiApp(root)
    root.mainloop()
    return 0
