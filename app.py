import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core import DataError, LEVELS, export_summary, summarize_excel


class StatisticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("月度等级统计")
        self.root.geometry("760x520")
        self.root.minsize(680, 420)
        self.file_path = tk.StringVar()
        self.status = tk.StringVar(value="请选择 Excel 文件。")
        self.rows = []
        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        file_frame = ttk.Frame(container)
        file_frame.pack(fill="x")
        ttk.Entry(file_frame, textvariable=self.file_path, state="readonly").pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(file_frame, text="选择 Excel", command=self.choose_file).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(file_frame, text="开始统计", command=self.run_statistics).pack(
            side="left", padx=(8, 0)
        )

        columns = ("月份", *LEVELS, "合计")
        table_frame = ttk.Frame(container)
        table_frame.pack(fill="both", expand=True, pady=16)
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, anchor="center", width=100)
        self.tree.column("月份", width=120)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        bottom = ttk.Frame(container)
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status).pack(side="left")
        self.export_button = ttk.Button(
            bottom, text="导出 Excel", command=self.export, state="disabled"
        )
        self.export_button.pack(side="right")

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择统计文件",
            filetypes=[("Excel 文件", "*.xls *.xlsx")],
        )
        if path:
            self.file_path.set(path)
            self.status.set("文件已选择，点击“开始统计”。")

    def run_statistics(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("提示", "请先选择 Excel 文件。")
            return

        try:
            rows = summarize_excel(path)
        except DataError as exc:
            messagebox.showerror("统计失败", str(exc))
            self.status.set("统计失败，请检查文件。")
            return

        self.rows = rows
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in rows:
            self.tree.insert("", "end", values=[row[key] for key in ("月份", *LEVELS, "合计")])

        total = sum(row["合计"] for row in rows)
        self.status.set(f"统计完成：{len(rows)} 个月，共 {total} 条记录。")
        self.export_button.configure(state="normal")

    def export(self):
        source = Path(self.file_path.get())
        default_name = f"{source.stem}_月度等级统计.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出统计结果",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_summary(self.rows, path)
        except Exception as exc:
            messagebox.showerror("导出失败", f"无法保存文件：{exc}")
            return
        messagebox.showinfo("导出完成", f"统计结果已保存到：\n{path}")


def main():
    root = tk.Tk()
    StatisticsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
