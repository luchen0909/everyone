import os
import traceback
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Set, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


PARENT_ALIASES = [
    "母件料号",
    "母料号",
    "母件编码",
    "母器件料号",
    "母件",
    "上层料号",
    "parent",
]
CHILD_ALIASES = [
    "子件料号",
    "物料号",
    "物料编码",
    "料号",
    "子件编码",
    "child",
]
QTY_ALIASES = [
    "用量",
    "单位用量",
    "数量",
    "单台用量",
    "组成数量",
    "qty",
]
NAME_ALIASES = ["品名", "名称", "物料名称", "name"]
SPEC_ALIASES = ["规格", "规格型号", "型号", "spec"]


def _norm_col(col: str) -> str:
    return str(col).strip().replace(" ", "").lower()


def _find_col(columns: List[str], aliases: List[str]) -> str:
    norm_to_raw = {_norm_col(c): c for c in columns}
    for alias in aliases:
        key = _norm_col(alias)
        if key in norm_to_raw:
            return norm_to_raw[key]
    return ""


def _to_float(val) -> float:
    try:
        if pd.isna(val):
            return 0.0
        s = str(val).strip().replace(",", "")
        if not s:
            return 0.0
        return float(s)
    except Exception:
        return 0.0


@dataclass
class RowItem:
    parent: str
    child: str
    qty: float
    name: str = ""
    spec: str = ""


class BomDiffEngine:
    def __init__(self) -> None:
        self.rows: List[RowItem] = []
        self.graph: Dict[str, List[RowItem]] = defaultdict(list)
        self.meta: Dict[str, Tuple[str, str]] = {}

    def load_file(self, file_path: str) -> Tuple[int, str]:
        ext = Path(file_path).suffix.lower()
        if ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path)
        elif ext in [".csv", ".txt"]:
            df = pd.read_csv(file_path, encoding="utf-8", engine="python")
        else:
            raise ValueError("仅支持 xlsx/xls/csv/txt 文件")

        if df.empty:
            raise ValueError("文件为空，未读取到 BOM 数据")

        cols = [str(c) for c in df.columns]
        parent_col = _find_col(cols, PARENT_ALIASES)
        child_col = _find_col(cols, CHILD_ALIASES)
        qty_col = _find_col(cols, QTY_ALIASES)
        name_col = _find_col(cols, NAME_ALIASES)
        spec_col = _find_col(cols, SPEC_ALIASES)

        if not parent_col or not child_col or not qty_col:
            raise ValueError(
                "无法识别关键列，请至少包含：母件料号/子件料号(物料号)/用量。"
                f"\n当前列名: {cols}"
            )

        self.rows.clear()
        self.graph.clear()
        self.meta.clear()

        for _, r in df.iterrows():
            parent = str(r.get(parent_col, "")).strip()
            child = str(r.get(child_col, "")).strip()
            qty = _to_float(r.get(qty_col, 0))
            name = str(r.get(name_col, "")).strip() if name_col else ""
            spec = str(r.get(spec_col, "")).strip() if spec_col else ""

            if not parent or not child:
                continue
            if qty == 0:
                continue

            item = RowItem(parent=parent, child=child, qty=qty, name=name, spec=spec)
            self.rows.append(item)
            self.graph[parent].append(item)
            if child not in self.meta:
                self.meta[child] = (name, spec)

        return len(self.rows), f"已加载 {len(self.rows)} 条 BOM 关系"

    def explode(self, root: str) -> Tuple[Dict[str, float], List[str]]:
        root = str(root).strip()
        if not root:
            raise ValueError("物料编码不能为空")

        total_qty: Dict[str, float] = defaultdict(float)
        warnings: List[str] = []
        call_stack: Set[str] = set()

        def dfs(parent: str, factor: float, depth: int) -> None:
            if depth > 80:
                warnings.append(f"超过最大递归层级，已截断: {parent}")
                return
            children = self.graph.get(parent, [])
            for item in children:
                child_factor = factor * item.qty
                total_qty[item.child] += child_factor
                if item.child in call_stack:
                    warnings.append(f"检测到循环引用，跳过继续展开: {' -> '.join(list(call_stack) + [item.child])}")
                    continue
                if item.child in self.graph:
                    call_stack.add(item.child)
                    dfs(item.child, child_factor, depth + 1)
                    call_stack.remove(item.child)

        call_stack.add(root)
        dfs(root, 1.0, 0)
        call_stack.remove(root)
        return total_qty, warnings

    def compare(self, code_a: str, code_b: str):
        qty_a, warn_a = self.explode(code_a)
        qty_b, warn_b = self.explode(code_b)

        set_a = set(qty_a.keys())
        set_b = set(qty_b.keys())

        only_a = sorted(set_a - set_b)
        only_b = sorted(set_b - set_a)
        common = sorted(set_a & set_b)

        rows_only_a = []
        rows_only_b = []
        rows_diff = []

        for k in only_a:
            name, spec = self.meta.get(k, ("", ""))
            rows_only_a.append(
                {"物料编码": k, "品名": name, "规格": spec, "A总用量": round(qty_a[k], 6), "B总用量": 0.0}
            )
        for k in only_b:
            name, spec = self.meta.get(k, ("", ""))
            rows_only_b.append(
                {"物料编码": k, "品名": name, "规格": spec, "A总用量": 0.0, "B总用量": round(qty_b[k], 6)}
            )
        for k in common:
            a = round(qty_a[k], 6)
            b = round(qty_b[k], 6)
            if abs(a - b) < 1e-12:
                continue
            name, spec = self.meta.get(k, ("", ""))
            rows_diff.append(
                {
                    "物料编码": k,
                    "品名": name,
                    "规格": spec,
                    "A总用量": a,
                    "B总用量": b,
                    "差值(A-B)": round(a - b, 6),
                }
            )

        return (
            pd.DataFrame(rows_only_a),
            pd.DataFrame(rows_only_b),
            pd.DataFrame(rows_diff),
            warn_a + warn_b,
        )


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("BOM差异对比工具")
        self.root.geometry("1100x760")

        self.engine = BomDiffEngine()
        self.file_path = ""
        self.df_a = pd.DataFrame()
        self.df_b = pd.DataFrame()
        self.df_d = pd.DataFrame()

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")

        ttk.Button(top, text="选择BOM文件", command=self.on_choose_file).pack(side="left")
        self.lbl_file = ttk.Label(top, text="未选择文件")
        self.lbl_file.pack(side="left", padx=8)

        input_bar = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        input_bar.pack(fill="x")

        ttk.Label(input_bar, text="物料编码A:").pack(side="left")
        self.ent_a = ttk.Entry(input_bar, width=24)
        self.ent_a.pack(side="left", padx=(4, 12))

        ttk.Label(input_bar, text="物料编码B:").pack(side="left")
        self.ent_b = ttk.Entry(input_bar, width=24)
        self.ent_b.pack(side="left", padx=(4, 12))

        ttk.Button(input_bar, text="开始对比", command=self.on_compare).pack(side="left")
        ttk.Button(input_bar, text="导出Excel", command=self.on_export).pack(side="left", padx=8)

        self.lbl_status = ttk.Label(self.root, text="就绪")
        self.lbl_status.pack(fill="x", padx=12, pady=(0, 8))

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tv_a = self._make_tab(notebook, "仅A有")
        self.tv_b = self._make_tab(notebook, "仅B有")
        self.tv_d = self._make_tab(notebook, "共有但用量不同")

    def _make_tab(self, notebook: ttk.Notebook, title: str) -> ttk.Treeview:
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=title)

        tv = ttk.Treeview(frame, show="headings")
        tv.pack(side="left", fill="both", expand=True)
        ybar = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        ybar.pack(side="right", fill="y")
        tv.configure(yscrollcommand=ybar.set)
        return tv

    def _show_df(self, tv: ttk.Treeview, df: pd.DataFrame) -> None:
        tv.delete(*tv.get_children())
        if df.empty:
            tv["columns"] = ()
            return

        columns = list(df.columns)
        tv["columns"] = columns
        for c in columns:
            tv.heading(c, text=c)
            tv.column(c, width=140, anchor="center")
        for _, row in df.iterrows():
            tv.insert("", "end", values=[row[c] for c in columns])

    def on_choose_file(self) -> None:
        fp = filedialog.askopenfilename(
            title="选择BOM文件",
            filetypes=[
                ("Excel/CSV", "*.xlsx *.xls *.csv *.txt"),
                ("All Files", "*.*"),
            ],
        )
        if not fp:
            return
        try:
            count, msg = self.engine.load_file(fp)
            self.file_path = fp
            self.lbl_file.config(text=f"{Path(fp).name}  (记录数: {count})")
            self.lbl_status.config(text=msg)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))

    def on_compare(self) -> None:
        if not self.file_path:
            messagebox.showwarning("提示", "请先选择BOM文件")
            return
        code_a = self.ent_a.get().strip()
        code_b = self.ent_b.get().strip()
        if not code_a or not code_b:
            messagebox.showwarning("提示", "请同时输入物料编码A和B")
            return

        try:
            self.df_a, self.df_b, self.df_d, warns = self.engine.compare(code_a, code_b)
            self._show_df(self.tv_a, self.df_a)
            self._show_df(self.tv_b, self.df_b)
            self._show_df(self.tv_d, self.df_d)

            summary = (
                f"完成: 仅A有 {len(self.df_a)} 条, 仅B有 {len(self.df_b)} 条, "
                f"共有但用量不同 {len(self.df_d)} 条"
            )
            if warns:
                summary += f" | 警告 {len(warns)} 条"
            self.lbl_status.config(text=summary)
        except Exception as e:
            messagebox.showerror("对比失败", f"{e}\n\n{traceback.format_exc()}")

    def on_export(self) -> None:
        if self.df_a.empty and self.df_b.empty and self.df_d.empty:
            messagebox.showwarning("提示", "暂无可导出结果，请先执行对比")
            return
        default_name = f"BOM差异对比_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        fp = filedialog.asksaveasfilename(
            title="导出Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not fp:
            return
        try:
            with pd.ExcelWriter(fp) as writer:
                self.df_a.to_excel(writer, index=False, sheet_name="仅A有")
                self.df_b.to_excel(writer, index=False, sheet_name="仅B有")
                self.df_d.to_excel(writer, index=False, sheet_name="共有但用量不同")
            self.lbl_status.config(text=f"导出完成: {fp}")
            messagebox.showinfo("完成", f"导出成功:\n{fp}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
