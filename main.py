# -*- coding: utf-8 -*-
"""入口：拖拽/选择 Excel 文件 -> 选工作表 -> 统计时长 -> 弹窗 + 写回合计。"""

import os
import re

import tkinter as tk
from tkinter import filedialog, messagebox

import calc
import reader
import tencent

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


class LineNumberText(tk.Frame):
    """带 Excel 式行号的文本框。"""

    def __init__(self, parent, **kw):
        tk.Frame.__init__(self, parent)
        self.skipped_rows = set()
        self.ln = tk.Canvas(self, width=42, highlightthickness=0,
                            bd=0, bg="#f2f2f2")
        self.vsb = tk.Scrollbar(self, orient=tk.VERTICAL,
                                command=self._scroll_text)
        self.text = tk.Text(self, yscrollcommand=self._scroll_bar, **kw)
        self.ln.pack(side=tk.LEFT, fill=tk.Y)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.text.bind("<KeyRelease>", self._on_key)
        self.text.bind("<Configure>", lambda e: self._redraw())
        self.text.bind("<<Modified>>", self._on_modified)
        self.text.bind("<MouseWheel>", self._wheel)
        self.ln.bind("<MouseWheel>", self._wheel)
        self.vsb.bind("<MouseWheel>", self._wheel)
        self._redraw()

    def _on_key(self, event):
        self.clear_skipped()
        self._redraw()

    def _wheel(self, event):
        self.text.yview_scroll(int(-event.delta / 120), "units")
        self._redraw()
        return "break"

    def _scroll_bar(self, *args):
        self.vsb.set(*args)
        self.ln.yview_moveto(args[0])

    def _scroll_text(self, *args):
        self.text.yview(*args)
        self._redraw()

    def _on_modified(self, event):
        if self.text.edit_modified():
            self._redraw()

    def _redraw(self):
        self.text.edit_modified(False)
        self.ln.delete("all")
        i = self.text.index("@0,0")
        while True:
            dline = self.text.dlineinfo(i)
            if dline is None:
                break
            y = dline[1]
            n = int(str(i).split(".")[0])
            fill = "#c00" if n in self.skipped_rows else "#999"
            self.ln.create_text(36, y, anchor="ne", text=n,
                                font=self.text.cget("font"), fill=fill)
            i = self.text.index("%s+1line" % i)

    def set_skipped(self, rows):
        self.skipped_rows = set(rows)
        self._redraw()

    def clear_skipped(self):
        if self.skipped_rows:
            self.skipped_rows = set()
            self._redraw()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel视频时长统计工具（Beta版）")
        self.root.geometry("480x640")
        self.root.minsize(440, 600)

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)

        hint = ("把 Excel 文件拖到这里（可一次拖多个）\n"
                "（支持 .xlsx / .xls / .csv）")
        tk.Label(self.root, text=hint, font=("Microsoft YaHei", 12),
                 pady=16).pack()

        tk.Button(self.root, text="选择文件…（可多选）", width=18,
                  command=self.choose_file).pack(pady=6)

        row = tk.Frame(self.root)
        row.pack(fill=tk.X, padx=16, pady=(0, 4))
        tk.Label(row, text="已导入：", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.var_import = tk.StringVar()
        self.om_import = tk.OptionMenu(row, self.var_import, "")
        self.om_import.config(width=24)
        self.om_import.pack(side=tk.LEFT, padx=4)
        tk.Button(row, text="打开", width=5, command=self.open_import).pack(side=tk.LEFT)
        tk.Button(row, text="✕", width=2, fg="#c00", relief=tk.FLAT,
                  command=self.remove_import).pack(side=tk.LEFT, padx=(6, 0))

        self.imported = []
        self.import_labels = {}

        tk.Label(self.root,
                 text="── 或粘贴在线文档链接（腾讯文档 / WPS，一行一个）──",
                 font=("Microsoft YaHei", 9), fg="#999").pack(pady=(10, 4))

        self.txt_url = tk.Entry(self.root, font=("Consolas", 10))
        self.txt_url.pack(fill=tk.X, padx=16)

        tk.Button(self.root, text="获取在线文档数据", width=18,
                  command=self.process_url).pack(pady=4)

        self.lb_loading = tk.Label(self.root, text="", fg="#1a73e8",
                                   font=("Microsoft YaHei", 9))
        self._loading_job = None

        tk.Label(self.root,
                 text="── 或直接粘贴时长数据（每行一条，如 3:40 / 90）──",
                 font=("Microsoft YaHei", 9), fg="#999").pack(pady=(10, 4))

        inp = LineNumberText(self.root, height=6, font=("Consolas", 11),
                             wrap=tk.WORD, relief=tk.GROOVE, bd=1)
        inp.pack(fill=tk.BOTH, expand=True, padx=16)
        self.txt_input = inp.text
        self.txt_input_ln = inp

        tk.Button(self.root, text="统计粘贴的时长", width=18,
                  command=self.process_text).pack(pady=8)

        self.var_status = tk.StringVar(value="")
        tk.Label(self.root, textvariable=self.var_status,
                 font=("Microsoft YaHei", 9), fg="#666").pack(pady=(0, 8))

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

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

    def process_text(self):
        text = self.txt_input.get("1.0", tk.END)
        if not text.strip():
            messagebox.showinfo("提示", "请先粘贴时长数据。")
            return
        lines = text.splitlines()
        values = []
        skipped_rows = []
        for idx, line in enumerate(lines, 1):
            if not line.strip():
                continue
            values.append((line, False))
            if calc.parse_seconds(line, False) is None:
                skipped_rows.append(idx)
        self.txt_input_ln.set_skipped(skipped_rows)
        total, parsed, skipped = calc.total_from_values(values)
        if parsed == 0:
            messagebox.showwarning(
                "没有有效数据",
                "没有解析到任何时长，请检查输入内容。\n"
                "支持格式：3:40 / 3:00 / 90（秒）")
            return
        msg = "\n".join(calc.build_result_lines(total, parsed, skipped))
        messagebox.showinfo("统计结果", msg)
        self.var_status.set("已统计 %d 条粘贴数据" % parsed)

    def process_url(self):
        text = self.txt_url.get().strip()
        if not text:
            messagebox.showinfo("提示", "请先粘贴在线文档链接。")
            return
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        for link in lines:
            self.process_online(link)

    def _start_loading(self, text):
        self.lb_loading.config(text=text)
        self._loading_job = self.root.after(50, lambda: self._animate_dots(text))

    def _animate_dots(self, base):
        dots = "." * (((self._loading_dots % 3) + 1))
        self.lb_loading.config(text=base + dots)
        self._loading_dots += 1
        self._loading_job = self.root.after(400, lambda: self._animate_dots(base))

    def _stop_loading(self):
        if self._loading_job:
            self.root.after_cancel(self._loading_job)
            self._loading_job = None
        self.lb_loading.config(text="")

    def process_online(self, url):
        self._loading_dots = 0
        self._start_loading("正在获取数据中")
        self.var_status.set("正在获取在线文档，请稍候…")
        self._disable_buttons(True)

        def worker():
            result = None
            err = None
            try:
                result = tencent.fetch_workbook_with_browser(
                    url, progress=self._on_progress)
            except Exception as e:
                err = str(e)
            self.root.after(0, lambda: self._online_done(url, result, err))

        import threading
        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _on_progress(self, msg):
        self._stop_loading()
        self._loading_dots = 0
        self._start_loading("「%s」" % msg)

    def _online_done(self, url, result, err):
        self._stop_loading()
        self._disable_buttons(False)
        if err:
            messagebox.showerror("获取失败", err)
            self.var_status.set("获取失败")
            return
        if not result:
            messagebox.showerror("获取失败", "没有获取到任何数据，请检查链接。")
            self.var_status.set("获取失败")
            return

        # 1) 智能检测每个工作簿有哪些列有数据
        sheets_info = []
        all_cols = set()
        for wb in result:
            name = wb["name"]
            columns = wb.get("columns", {})
            usable_cols = []
            for col in sorted(columns):
                vals = columns[col]
                total, parsed, _ = calc.total_from_values(
                    [(c, False) for c in vals])
                if parsed > 0:
                    usable_cols.append(col)
                    all_cols.add(col)
            sheets_info.append({
                "name": name,
                "columns": columns,
                "usable_cols": usable_cols,
            })

        # 2) 弹出工作簿多选框（只列有可用数据的）
        selectable = [s for s in sheets_info if s["usable_cols"]]
        if not selectable:
            messagebox.showwarning(
                "没有有效数据",
                "所有工作簿都没有解析到任何时长，请检查文档内容。")
            return
        picked_sheets = self._choose_sheets(selectable)
        if picked_sheets is None:
            self.var_status.set("已取消")
            return

        # 3) 弹出列选择框（统一固定选一列，只列有数据的列）
        usable_cols = sorted(all_cols)
        pick_col = self._choose_column(usable_cols)
        if pick_col is None:
            self.var_status.set("已取消")
            return

        # 4) 对每个选中的工作簿，统计选中的那一列
        parts = ["在线文档：%s" % url, ""]
        valid = 0
        for name in picked_sheets:
            info = next(s for s in sheets_info if s["name"] == name)
            cols = info["columns"]
            vals = cols.get(pick_col, [])
            total, parsed, skipped = calc.total_from_values(
                [(c, False) for c in vals])
            parts.append("──── %s（%s 列）────" % (name, pick_col))
            if parsed == 0:
                parts.append("  该工作簿的 %s 列没有时长数据" % pick_col)
                parts.append("")
                continue
            for line in calc.build_result_lines(total, parsed, skipped):
                parts.append("  " + line)
            parts.append("")
            valid += 1

        if valid == 0:
            messagebox.showwarning(
                "没有有效数据",
                "所选工作簿的 %s 列都没有解析到任何时长，请检查文档内容。"
                % pick_col)
            return

        messagebox.showinfo("统计结果", "\n".join(parts))
        self.var_status.set("已统计 %d 个工作簿的 %s 列" % (valid, pick_col))

    def _choose_sheets(self, sheets_info):
        """弹出工作簿多选框，可勾选多个工作簿。

        sheets_info: [{"name":..., "columns":..., "usable_cols":[...]}, ...]
        返回选中的工作簿名列表，或 None（取消）。
        """
        top = tk.Toplevel(self.root)
        top.title("选择工作簿")
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)

        tk.Label(top,
                 text="该文档有多个工作簿，请勾选要统计的工作簿（可多选）：",
                 font=("Microsoft YaHei", 9)).pack(padx=12, pady=(12, 4))

        box = tk.Frame(top)
        box.pack(padx=12, pady=4)
        vars_map = {}
        for info in sheets_info:
            name = info["name"]
            cols_desc = ", ".join(info["usable_cols"]) if info["usable_cols"] else "无"
            var = tk.BooleanVar(value=True)
            vars_map[name] = var
            cb = tk.Checkbutton(
                box, text="工作簿 %s（列：%s）" % (name, cols_desc),
                font=("Microsoft YaHei", 9), variable=var, anchor="w")
            cb.pack(fill=tk.X, padx=8, pady=2)

        result = {}

        def ok():
            picked = [n for n, v in vars_map.items() if v.get()]
            if not picked:
                messagebox.showinfo("提示", "请至少勾选一个工作簿。")
                return
            result["names"] = picked
            top.destroy()

        def cancel():
            top.destroy()

        btn = tk.Frame(top)
        btn.pack(pady=8)
        tk.Button(btn, text="确定", width=8, command=ok).pack(side=tk.LEFT, padx=6)
        tk.Button(btn, text="取消", width=8, command=cancel).pack(side=tk.LEFT, padx=6)

        self._center(top)
        top.wait_window()
        return result.get("names")

    def _choose_column(self, usable_cols):
        """弹出列选择框，只列出有数据的列（统一固定选一列）。

        usable_cols: [col, ...] 有数据的列名列表。
        返回选中的列名，或 None（取消）。
        """
        if not usable_cols:
            return None
        top = tk.Toplevel(self.root)
        top.title("选择统计列")
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)

        tk.Label(top,
                 text="请选择要统计的列（将统计所有选中工作簿的该列）：",
                 font=("Microsoft YaHei", 9)).pack(padx=12, pady=(12, 4))

        frame = tk.Frame(top)
        frame.pack(padx=12, pady=4)
        labels = ["列 %s" % col for col in usable_cols]

        var = tk.StringVar(value=labels[0])
        om = tk.OptionMenu(frame, var, *labels)
        om.config(width=24)
        om.pack()

        result = {}

        def ok():
            result["col"] = usable_cols[labels.index(var.get())]
            top.destroy()

        def cancel():
            top.destroy()

        btn = tk.Frame(top)
        btn.pack(pady=8)
        tk.Button(btn, text="确定", width=8, command=ok).pack(side=tk.LEFT, padx=6)
        tk.Button(btn, text="取消", width=8, command=cancel).pack(side=tk.LEFT, padx=6)

        self._center(top)
        top.wait_window()
        return result.get("col")

    def _disable_buttons(self, disabled):
        for w in self.root.winfo_children():
            if isinstance(w, tk.Button):
                w.config(state=tk.DISABLED if disabled else tk.NORMAL)

    def process_file(self, path):
        try:
            handle, sheets, kind = reader.load_file(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        self._remember(path)

        if kind in ("xlsx", "xls") and len(sheets) > 1:
            sheet = self.choose_sheet(sheets)
            if sheet is None:
                self.var_status.set("已取消")
                return
        else:
            sheet = sheets[0]

        try:
            values = reader.read_column_a(handle, kind, sheet)
            total, parsed, skipped = calc.total_from_values(values)
        except Exception as e:
            messagebox.showerror("计算失败", str(e))
            return

        if parsed == 0:
            messagebox.showwarning("没有有效数据",
                                   "A 列没有解析到任何时长，请检查表格内容。")
            return

        lines = calc.build_result_lines(total, parsed, skipped)
        msg = "\n".join(lines)

        note = ""
        if kind in ("xlsx", "csv"):
            try:
                ok = reader.write_total(handle, kind, sheet,
                                        calc.build_total_text(total), path)
                if ok:
                    note = "\n\n合计已写入：%s" % os.path.basename(path)
                else:
                    note = "\n\n（该格式暂不支持写回合计）"
            except Exception as e:
                note = "\n\n写回失败：%s" % e
        else:
            note = "\n\n（.xls 文件暂不支持写回合计，仅显示结果）"

        messagebox.showinfo("统计结果", msg + note)
        self.var_status.set("已处理：%s" % os.path.basename(path))

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
        self._refresh_import_menu()
        self.var_import.set(label)

    def _refresh_import_menu(self):
        menu = self.om_import["menu"]
        menu.delete(0, tk.END)
        if not self.import_labels:
            self.var_import.set("")
            return
        for lb in self.import_labels:
            menu.add_command(label=lb, command=lambda v=lb: self.var_import.set(v))

    def remove_import(self):
        label = self.var_import.get()
        if not label or label not in self.import_labels:
            messagebox.showinfo("提示", "请先选择一个要移除的文件。")
            return
        if not messagebox.askyesno("确认移除",
                                   "确定从列表移除「%s」吗？" % label):
            return
        del self.import_labels[label]
        self.imported = list(self.import_labels.values())
        self._refresh_import_menu()
        if self.import_labels:
            self.var_import.set(next(iter(self.import_labels)))
        self.var_status.set("已移除：%s" % label)

    def open_import(self):
        path = self.import_labels.get(self.var_import.get())
        if not path:
            messagebox.showinfo("提示", "请先导入文件，再点「打开」。")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

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
