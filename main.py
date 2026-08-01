# -*- coding: utf-8 -*-
"""入口：拖拽/选择 Excel 文件 -> 选工作表 -> 统计时长 -> 日志展示 + 写回合计。

支持一次导入/拖入多个文件，结果直接显示在下方日志区，无需点确定即可继续导入。
"""

import os
import re

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import calc
import reader

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    HAS_DND = False


def _extract_paths(data):
    paths = []
    for m in re.finditer(r"\{([^{}]+)\}|\S+", data or ""):
        if m.group(1) is not None:
            paths.append(m.group(1).strip())
        else:
            paths.append(m.group(0).strip())
    return [p for p in paths if p]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 视频时长统计工具")
        self.root.geometry("480x720")
        self.root.minsize(420, 560)

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)

        hint = ("把 Excel 文件拖到这里（可一次拖多个）\n"
                "（支持 .xlsx / .xls / .csv）")
        tk.Label(self.root, text=hint, font=("Microsoft YaHei", 11),
                 pady=12).pack()

        tk.Button(self.root, text="选择文件…（可多选）", width=18,
                  command=self.choose_file).pack(pady=4)

        row = tk.Frame(self.root)
        row.pack(fill=tk.X, padx=16, pady=(0, 4))
        tk.Label(row, text="已导入：", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.var_import = tk.StringVar()
        self.om_import = tk.OptionMenu(row, self.var_import, "")
        self.om_import.config(width=24)
        self.om_import.pack(side=tk.LEFT, padx=4)
        tk.Button(row, text="打开", width=5, command=self.open_import).pack(side=tk.LEFT)

        self.imported = []
        self.import_labels = {}

        tk.Label(self.root, text="── 统计结果 ──",
                 font=("Microsoft YaHei", 9), fg="#999").pack()

        self.txt_log = scrolledtext.ScrolledText(
            self.root, height=9, font=("Consolas", 10),
            wrap=tk.WORD, relief=tk.GROOVE, bd=1, state=tk.DISABLED)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=16, pady=4)
        self.txt_log.tag_configure("error", foreground="#c00")
        self.txt_log.tag_configure("ok", foreground="#080")

        tk.Label(self.root,
                 text="── 或直接粘贴时长数据（每行一条，如 3:40 / 90）──",
                 font=("Microsoft YaHei", 9), fg="#999").pack(pady=(4, 4))

        self.txt_input = scrolledtext.ScrolledText(
            self.root, height=5, font=("Consolas", 11),
            wrap=tk.WORD, relief=tk.GROOVE, bd=1)
        self.txt_input.pack(fill=tk.BOTH, expand=True, padx=16)

        tk.Button(self.root, text="统计粘贴的时长", width=18,
                  command=self.process_text).pack(pady=6)

        self.var_status = tk.StringVar(value="")
        tk.Label(self.root, textvariable=self.var_status,
                 font=("Microsoft YaHei", 9), fg="#666").pack(pady=(0, 8))

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _log(self, text, tag=None):
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.insert(tk.END, text + "\n", tag or ())
        self.txt_log.configure(state=tk.DISABLED)
        self.txt_log.see(tk.END)

    def on_drop(self, event):
        paths = _extract_paths(event.data)
        for p in paths:
            self.process_file(p)

    def choose_file(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件（可多选）",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv"),
                       ("所有文件", "*.*")])
        for p in paths:
            self.process_file(p)

    def _remember(self, path):
        if path in self.imported:
            return
        self.imported.append(path)
        base = os.path.basename(path)
        label = base
        i = 2
        while label in self.import_labels:
            label = "%s (%d)" % (base, i)
            i += 1
        self.import_labels[label] = path
        menu = self.om_import["menu"]
        menu.delete(0, tk.END)
        for lb in self.import_labels:
            menu.add_command(label=lb, command=lambda v=lb: self.var_import.set(v))
        self.var_import.set(label)

    def open_import(self):
        path = self.import_labels.get(self.var_import.get())
        if not path:
            self._log("请先导入文件，再点「打开」。", "error")
            return
        try:
            os.startfile(path)
        except Exception as e:
            self._log("打开失败：%s" % e, "error")

    def process_text(self):
        text = self.txt_input.get("1.0", tk.END).strip()
        if not text:
            self._log("提示：请先粘贴时长数据。", "error")
            return
        values = [(line, False) for line in text.splitlines()
                  if line.strip()]
        total, parsed, skipped = calc.total_from_values(values)
        if parsed == 0:
            self._log("粘贴数据：没有解析到任何时长，请检查输入内容。"
                      "（支持 3:40 / 3:00 / 90）", "error")
            return
        self._log("【粘贴数据】")
        for line in calc.build_result_lines(total, parsed, skipped):
            self._log(line, "ok")
        self.var_status.set("已统计 %d 条粘贴数据" % parsed)

    def process_file(self, path):
        name = os.path.basename(path)
        self._log("=" * 44)
        self._log("处理：%s" % name)
        try:
            handle, sheets, kind = reader.load_file(path)
        except Exception as e:
            self._log("读取失败：%s" % e, "error")
            return
        self._remember(path)

        if kind in ("xlsx", "xls") and len(sheets) > 1:
            sheet = self.choose_sheet(sheets)
            if sheet is None:
                self._log("已取消（未选择工作表）", "error")
                return
        else:
            sheet = sheets[0]

        try:
            values = reader.read_column_a(handle, kind, sheet)
            total, parsed, skipped = calc.total_from_values(values)
        except Exception as e:
            self._log("计算失败：%s" % e, "error")
            return

        if parsed == 0:
            self._log("A 列没有解析到任何时长，请检查表格内容。", "error")
            return

        for line in calc.build_result_lines(total, parsed, skipped):
            self._log(line, "ok")

        note = ""
        if kind in ("xlsx", "csv"):
            try:
                ok = reader.write_total(handle, kind, sheet,
                                        calc.build_total_text(total), path)
                if ok:
                    note = "合计已写入：%s" % name
                else:
                    note = "该格式暂不支持写回合计"
            except Exception as e:
                note = "写回失败：%s" % e
        else:
            note = ".xls 文件暂不支持写回合计，仅显示结果"

        self._log(note, "error" if note.startswith("写回失败") else None)
        self.var_status.set("已处理：%s" % name)

    def choose_sheet(self, sheets):
        top = tk.Toplevel(self.root)
        top.title("选择工作表")
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)

        tk.Label(top, text="该文件有多个工作表，请选择要统计的那一个：",
                 font=("Microsoft YaHei", 9)).pack(padx=12, pady=(12, 4))

        lb = tk.Listbox(top, width=40, height=min(8, len(sheets)),
                        font=("Microsoft YaHei", 10))
        for s in sheets:
            lb.insert(tk.END, s)
        lb.selection_set(0)
        lb.pack(padx=12, pady=4)

        result = {}

        def ok():
            sel = lb.curselection()
            if sel:
                result["name"] = sheets[sel[0]]
            top.destroy()

        def cancel():
            top.destroy()

        frame = tk.Frame(top)
        frame.pack(pady=8)
        tk.Button(frame, text="确定", width=8, command=ok).pack(side=tk.LEFT, padx=6)
        tk.Button(frame, text="取消", width=8, command=cancel).pack(side=tk.LEFT, padx=6)

        self._center(top)
        top.wait_window()
        return result.get("name")

    def _center(self, win):
        win.update_idletasks()
        x = self.root.winfo_rootx() + 60
        y = self.root.winfo_rooty() + 60
        win.geometry("+%d+%d" % (x, y))

    def on_close(self):
        self.root.destroy()


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        messagebox.showwarning("提示", "拖拽支持不可用，请使用「选择文件」按钮。")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
