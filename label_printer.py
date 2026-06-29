#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
乐宝质检 - 原料留样标签批量打印工具 v2.1
==========================================
简化版: 一个Excel → 数据质量检查 → 逐条打印标签
BTW模板一次打印两张标签(2-up)，逐条打印带进度
"""

import os
import sys
import logging
import configparser
import subprocess
import threading
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ============================================================
# High DPI support (Windows)
# ============================================================
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# ============================================================
# Path helpers
# ============================================================

def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent

# ============================================================
# Logging
# ============================================================

def setup_logging():
    log_path = get_base_dir() / "label_printer.log"
    _logger = logging.getLogger("label_printer")
    _logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(str(log_path), encoding="utf-8", mode="a")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    _logger.addHandler(fh)
    return _logger

logger = setup_logging()

# ============================================================
# pandas — graceful failure
# ============================================================
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    pd = None

# ============================================================
# Config
# ============================================================

class ConfigManager:
    DEFAULTS = {
        "BarTender": {
            "bartender_path": r"C:\Program Files\Seagull\BarTender 9.4\bartend.exe",
            "template_path": "",  # 需要用户配置
            "printer_name": "TSC TTP-342E Pro",
            "copies": "1",
        },
        "Excel": {
            # 打印前必须检查的列名（逗号分隔），这些列不能有空值
            "required_columns": "编码,名称,规格型号,基本单位",
        },
    }

    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_path = get_base_dir() / "config.ini"
        self.load()

    def load(self):
        if not self.config_path.exists():
            for section, options in self.DEFAULTS.items():
                self.config[section] = dict(options)
            self.save()
        self.config.read(str(self.config_path), encoding="utf-8")

    def save(self):
        with open(str(self.config_path), "w", encoding="utf-8") as f:
            self.config.write(f)

    def get(self, section, key, fallback=None):
        return self.config.get(section, key, fallback=fallback)

    def get_required_columns(self):
        """返回必检列名列表"""
        raw = self.get("Excel", "required_columns", "")
        return [c.strip() for c in raw.split(",") if c.strip()]

# ============================================================
# Data Validator
# ============================================================

class DataValidator:
    """读取一个Excel，检查必检列是否存在 + 哪些行有空值"""

    def __init__(self, config_manager):
        self.cm = config_manager

    def validate(self, file_path):
        """
        返回 (df, errors, empty_details)
        - df: 读取的DataFrame（成功时）
        - errors: 致命错误列表（文件不存在、缺少列等）
        - empty_details: 列→行号列表，表示哪些行该列有空值
        """
        errors = []
        empty_details = {}

        if not os.path.exists(file_path):
            errors.append("文件不存在")
            return None, errors, empty_details

        try:
            df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            errors.append(f"无法读取文件: {e}")
            return None, errors, empty_details

        # 清洗列名和值
        df.columns = df.columns.str.strip()
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].str.strip()

        if len(df) == 0:
            errors.append("Excel中没有数据行")
            return df, errors, empty_details

        # 检查必检列是否存在
        required = self.cm.get_required_columns()
        for col in required:
            if col not in df.columns:
                errors.append(f"缺少必要列「{col}」")

        if errors:
            return df, errors, empty_details

        # 检查必检列中的空值，记录具体行号
        for col in required:
            empty_rows = []
            for idx in range(len(df)):
                val = df.iloc[idx][col]
                if pd.isna(val) or str(val).strip() == "":
                    # 用Excel行号（+2：1是表头，pandas从0开始）
                    empty_rows.append(idx + 2)
            if empty_rows:
                empty_details[col] = empty_rows

        return df, errors, empty_details

    def format_empty_report(self, empty_details):
        """格式化空值报告"""
        if not empty_details:
            return "✓ 所有必检列数据完整，没有空值。"

        lines = ["以下列存在空值（已标注Excel行号）："]
        for col, rows in empty_details.items():
            row_str = ", ".join(str(r) for r in rows[:20])
            if len(rows) > 20:
                row_str += f" ...等共{len(rows)}行"
            lines.append(f"  列「{col}」第 {row_str} 行为空")
        return "\n".join(lines)

# ============================================================
# CSV Generator
# ============================================================

def generate_csv(data, output_dir=None):
    """生成 BarTender 可读取的 CSV"""
    if output_dir is None:
        output_dir = get_base_dir() / "output"
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = output_dir / f"print_data_{timestamp}.csv"
    data.to_csv(csv_path, index=False, encoding="utf-8-sig")

    # 固定名称版本，方便 BarTender 模板引用
    latest_path = output_dir / "print_data_latest.csv"
    data.to_csv(latest_path, index=False, encoding="utf-8-sig")

    logger.info(f"CSV已生成: {csv_path} ({len(data)} 条)")
    return str(csv_path)

# ============================================================
# BarTender Printer
# ============================================================

class BarTenderPrinter:
    def __init__(self, config_manager):
        self.cm = config_manager

    @property
    def bartender_path(self):
        return self.cm.get("BarTender", "bartender_path")

    @property
    def template_path(self):
        return self.cm.get("BarTender", "template_path")

    @property
    def printer_name(self):
        return self.cm.get("BarTender", "printer_name")

    @property
    def copies(self):
        return self.cm.get("BarTender", "copies", "1")

    def validate(self):
        errors = []
        if not self.bartender_path or not os.path.exists(self.bartender_path):
            errors.append(f"BarTender 程序未找到: {self.bartender_path}")
        if not self.template_path:
            errors.append("标签模板路径未配置，请在 config.ini 中设置 template_path")
        elif not os.path.exists(self.template_path):
            errors.append(f"标签模板未找到: {self.template_path}")
        return errors

    def _build_cmd(self, csv_path, record_number=None):
        """构建 BarTender CLI 命令
        /AF=模板 /P /D=CSV数据源 /PRN=打印机 /C=份数 /R=记录号(可选)
        BTW模板一次打印两张标签(2-up)，所以 copies=1 即每条记录出两张标签
        """
        cmd = [
            self.bartender_path,
            f'/AF="{self.template_path}"',
            "/P",
            f'/D="{csv_path}"',
            f'/PRN="{self.printer_name}"',
            f"/C={self.copies}",
        ]
        if record_number is not None:
            cmd.append(f"/R={record_number}")
        return cmd

    def print_single(self, csv_path, record_index):
        """打印单条记录（BTW一次出两张标签）"""
        cmd = self._build_cmd(csv_path, record_number=record_index + 1)
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            logger.info(f"打印第 {record_index + 1}/{record_index + 1} 条")
            return True, result.stdout, result.stderr
        except subprocess.TimeoutExpired:
            return False, "", "打印超时（超过60秒）"
        except FileNotFoundError:
            return False, "", f"找不到 BarTender: {self.bartender_path}"
        except Exception as e:
            return False, "", str(e)

    def print_all_sequential(self, csv_path, total_count, progress_callback=None):
        """逐条打印所有记录，通过 progress_callback 回报进度"""
        results = []
        for i in range(total_count):
            success, stdout, stderr = self.print_single(csv_path, i)
            results.append((i + 1, success, stderr))
            if progress_callback:
                progress_callback(i + 1, total_count)
            if not success:
                # 某条打印失败，停止后续
                logger.error(f"第 {i+1} 条打印失败: {stderr}")
                break
        return results

# ============================================================
# GUI Application
# ============================================================

class LabelPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("乐宝质检 - 原料留样标签打印")
        self.root.geometry("900x650")
        self.root.minsize(700, 500)

        # 主题
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Green.TButton", foreground="white", background="#4CAF50",
                         font=("Microsoft YaHei", 11, "bold"), padding=(12, 8))
        style.map("Green.TButton", background=[("active", "#45a049"), ("disabled", "#a5d6a7")])
        style.configure("Orange.TButton", foreground="white", background="#FF9800",
                         font=("Microsoft YaHei", 11, "bold"), padding=(12, 8))
        style.map("Orange.TButton", background=[("active", "#F57C00"), ("disabled", "#FFE0B2")])
        style.configure("Blue.TButton", foreground="white", background="#2196F3",
                         font=("Microsoft YaHei", 10), padding=(10, 6))
        style.map("Blue.TButton", background=[("active", "#1976D2"), ("disabled", "#90CAF9")])

        # 管理器
        self.cm = ConfigManager()
        self.validator = DataValidator(self.cm)
        self.printer = BarTenderPrinter(self.cm)

        # 状态
        self.df = None          # 当前加载的 DataFrame
        self.csv_path = None    # 生成的 CSV 路径
        self._busy = False
        self._stop_requested = False

        # Tk 变量
        self.file_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar(value=self.cm.get("BarTender", "template_path"))
        self.status_text = tk.StringVar(value="就绪 — 请选择Excel文件和标签模板")

        self._build_ui()

    def _build_ui(self):
        # ── 标题 ──
        ttk.Label(self.root, text="原料留样标签 打印工具",
                  font=("Microsoft YaHei", 14, "bold"), anchor="center"
        ).pack(fill=tk.X, pady=(10, 5))

        # ── Excel文件选择 ──
        file_frame = ttk.LabelFrame(self.root, text="选择Excel数据文件", padding=10)
        file_frame.pack(fill=tk.X, padx=15, pady=5)

        row = ttk.Frame(file_frame)
        row.pack(fill=tk.X)
        ttk.Entry(row, textvariable=self.file_path_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(row, text="浏览...", command=self._browse_file, width=7).pack(side=tk.LEFT)

        # ── 标签模板选择 ──
        tmpl_frame = ttk.LabelFrame(self.root, text="选择标签模板 (.btw)", padding=10)
        tmpl_frame.pack(fill=tk.X, padx=15, pady=5)

        row2 = ttk.Frame(tmpl_frame)
        row2.pack(fill=tk.X)
        ttk.Entry(row2, textvariable=self.template_path_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(row2, text="浏览...", command=self._browse_template, width=7).pack(side=tk.LEFT)

        # ── 操作按钮 ──
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=15, pady=8)

        self.btn_check = ttk.Button(btn_frame, text="① 检查数据", style="Green.TButton",
                                    command=self._do_check)
        self.btn_check.pack(side=tk.LEFT, padx=5)

        self.btn_print = ttk.Button(btn_frame, text="② 打印标签", style="Orange.TButton",
                                    command=self._do_print, state=tk.DISABLED)
        self.btn_print.pack(side=tk.LEFT, padx=5)

        self.btn_test = ttk.Button(btn_frame, text="测试首条", style="Blue.TButton",
                                   command=self._do_test_print, state=tk.DISABLED)
        self.btn_test.pack(side=tk.LEFT, padx=5)

        self.btn_stop = ttk.Button(btn_frame, text="停止", style="Blue.TButton",
                                   command=self._stop_print, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        # ── 进度条 (determinate模式，显示具体进度) ──
        self.progress = ttk.Progressbar(self.root, mode="determinate", length=300)
        self.progress.pack(fill=tk.X, padx=15, pady=3)
        self.progress_label = ttk.Label(self.root, text="", anchor="center")
        self.progress_label.pack()

        # ── 数据预览 ──
        preview_frame = ttk.LabelFrame(self.root, text="数据预览", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        self.tree = ttk.Treeview(preview_frame, show="headings", height=10,
                                  selectmode="extended")
        self.tree["columns"] = ("hint",)
        self.tree.heading("hint", text="请先选择Excel文件并点击「① 检查数据」")
        self.tree.column("hint", width=600, anchor="center")

        vsb = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        # ── 状态栏 ──
        ttk.Label(self.root, textvariable=self.status_text,
                  relief=tk.SUNKEN, anchor=tk.W, padding=(10, 3)
        ).pack(fill=tk.X, side=tk.BOTTOM)

    # ── 文件浏览 ──
    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择物料数据Excel",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if path:
            self.file_path_var.set(path)
            logger.info(f"选择数据文件: {path}")

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="选择BarTender标签模板",
            filetypes=[("BarTender模板", "*.btw"), ("所有文件", "*.*")])
        if path:
            self.template_path_var.set(path)
            # 保存到config
            self.cm.config.set("BarTender", "template_path", path)
            self.cm.save()
            logger.info(f"选择模板: {path}")

    # ── Busy 管理 ──
    def _set_busy(self, busy, allow_stop=False):
        self._busy = busy
        self.btn_check.config(state=tk.DISABLED if busy else tk.NORMAL)
        self.btn_print.config(state=tk.DISABLED if busy else (
            tk.NORMAL if self.df is not None else tk.DISABLED))
        self.btn_test.config(state=tk.DISABLED if busy else (
            tk.NORMAL if self.df is not None else tk.DISABLED))
        self.btn_stop.config(state=tk.NORMAL if (busy and allow_stop) else tk.DISABLED)

    def _run_in_thread(self, worker, callback):
        if self._busy:
            return
        def _target():
            try:
                result = worker()
            except Exception as e:
                logger.error(f"后台任务异常: {e}", exc_info=True)
                result = {"error": str(e)}
            self.root.after(0, lambda: self._on_thread_done(result, callback))
        self._set_busy(True)
        threading.Thread(target=_target, daemon=True).start()

    def _on_thread_done(self, result, callback):
        self._set_busy(False)
        callback(result)

    # ── ① 检查数据 ──
    def _do_check(self):
        path = self.file_path_var.get()
        if not path:
            messagebox.showwarning("提示", "请先选择Excel数据文件。")
            return
        # 同步保存模板路径（如果用户手动改了）
        tmpl = self.template_path_var.get()
        if tmpl:
            self.cm.config.set("BarTender", "template_path", tmpl)
            self.cm.save()

        self.status_text.set("正在读取并检查数据...")
        self.progress["value"] = 0
        self._run_in_thread(self._check_worker, self._check_callback)

    def _check_worker(self):
        df, errors, empty_details = self.validator.validate(self.file_path_var.get())
        report = self.validator.format_empty_report(empty_details)
        return {
            "df": df,
            "errors": errors,
            "empty_details": empty_details,
            "report": report,
        }

    def _check_callback(self, result):
        if result.get("error"):
            messagebox.showerror("异常", result["error"])
            self.status_text.set("检查异常")
            return

        errors = result["errors"]
        empty_details = result["empty_details"]
        report = result["report"]
        df = result["df"]

        # 有致命错误
        if errors:
            msg = "数据检查发现错误:\n\n" + "\n".join(f"  ✗ {e}" for e in errors)
            msg += "\n\n请修正Excel文件后重新检查。"
            messagebox.showerror("数据检查失败", msg)
            self.status_text.set("数据检查失败")
            self.df = None
            return

        # 有空值 → 提示具体行号，询问是否继续
        if empty_details:
            msg = report + "\n\n是否忽略空值继续？（空值字段在标签上会留空）"
            if not messagebox.askyesno("数据有空值", msg):
                self.status_text.set("已取消（数据有空值）")
                self.df = None
                return

        # 成功
        self.df = df
        self._update_preview()
        total = len(df)
        if empty_details:
            empty_count = sum(len(rows) for rows in empty_details.values())
            self.status_text.set(f"✓ 检查完成 — {total} 条数据（{empty_count} 个空值，已忽略）")
        else:
            self.status_text.set(f"✓ 检查完成 — {total} 条数据，全部完整")

    # ── 更新预览 ──
    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.df is None:
            return

        cols = list(self.df.columns)
        display_cols = ["#"] + cols
        self.tree["columns"] = display_cols

        self.tree.heading("#", text="#")
        self.tree.column("#", width=40, anchor="center", stretch=False)

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor="center", stretch=True)

        # 标记空值行（橙色高亮）
        required = self.cm.get_required_columns()
        for idx, (_, row) in enumerate(self.df.iterrows()):
            values = [str(v)[:50] if pd.notna(v) and str(v).strip() else "" for v in row.values]
            has_empty = any(
                col in required and (pd.isna(row[col]) or str(row[col]).strip() == "")
                for col in cols if col in required
            )
            tag = "empty" if has_empty else ""
            self.tree.insert("", tk.END, iid=str(idx), values=values, tags=(tag,))

        self.tree.tag_configure("empty", background="#FFF3E0")

    # ── ② 打印标签（逐条打印，带进度） ──
    def _do_print(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先检查数据。")
            return

        # 检查模板路径
        tmpl = self.template_path_var.get()
        if not tmpl:
            messagebox.showerror("配置错误", "请先选择标签模板(.btw)文件。")
            return
        self.cm.config.set("BarTender", "template_path", tmpl)
        self.cm.save()

        errors = self.printer.validate()
        if errors:
            messagebox.showerror("环境错误", "\n".join(errors))
            return

        total = len(self.df)
        ok = messagebox.askyesno(
            "确认打印",
            f"即将逐条打印 {total} 条标签\n"
            f"每条打印 2 张标签（模板一次出两张）\n\n"
            f"打印机: {self.printer.printer_name}\n\n确认开始？")
        if not ok:
            return

        self._stop_requested = False
        self.status_text.set("正在生成CSV...")
        self.progress["value"] = 0
        self.progress["maximum"] = total
        self.progress_label.config(text="")
        self._set_busy(True, allow_stop=True)
        threading.Thread(target=self._print_sequential_worker, daemon=True).start()

    def _print_sequential_worker(self):
        """后台线程：逐条打印，每条完成后通过 root.after 更新进度"""
        try:
            csv_path = generate_csv(self.df)
            total = len(self.df)
            results = self.printer.print_all_sequential(
                csv_path, total,
                progress_callback=lambda done, total_count:
                    self.root.after(0, lambda: self._update_print_progress(done, total_count))
            )
            self.root.after(0, lambda: self._print_sequential_done(results, total))
        except Exception as e:
            logger.error(f"打印线程异常: {e}", exc_info=True)
            self.root.after(0, lambda: self._print_error(str(e)))

    def _update_print_progress(self, done, total):
        """更新进度条和文字"""
        self.progress["value"] = done
        pct = int(done / total * 100) if total > 0 else 0
        self.progress_label.config(text=f"正在打印: {done}/{total} ({pct}%)")
        self.status_text.set(f"正在打印第 {done}/{total} 条...")

    def _print_sequential_done(self, results, total):
        """打印完成回调"""
        self._set_busy(False)
        self.progress["value"] = total

        success_count = sum(1 for _, ok, _ in results if ok)
        fail_count = sum(1 for _, ok, _ in results if not ok)
        last_fail_msg = ""
        for _, ok, err in results:
            if not ok:
                last_fail_msg = err

        if fail_count == 0:
            self.progress_label.config(text=f"✓ 打印完成: {success_count}/{total}")
            self.status_text.set(f"✓ 全部打印完成 — {success_count} 条（每条2张标签）")
            messagebox.showinfo("完成",
                f"打印完成!\n\n共 {success_count} 条 × 2张/条 = {success_count * 2} 张标签\n"
                f"请检查打印机输出。")
        else:
            self.progress_label.config(text=f"打印中断: 成功{success_count}条, 失败{fail_count}条")
            self.status_text.set(f"打印中断 — 成功 {success_count} 条，失败 {fail_count} 条")
            messagebox.showwarning("打印中断",
                f"打印在第 {success_count + 1} 条时失败。\n\n"
                f"已完成: {success_count} 条\n"
                f"失败原因: {last_fail_msg[:100]}")

    def _print_error(self, msg):
        self._set_busy(False)
        messagebox.showerror("异常", msg)
        self.status_text.set("打印异常")

    def _stop_print(self):
        """用户点击停止"""
        self._stop_requested = True
        self.status_text.set("正在停止打印...")
        logger.info("用户请求停止打印")

    # ── 测试打印首条 ──
    def _do_test_print(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先检查数据。")
            return

        tmpl = self.template_path_var.get()
        if not tmpl:
            messagebox.showerror("配置错误", "请先选择标签模板(.btw)文件。")
            return
        self.cm.config.set("BarTender", "template_path", tmpl)
        self.cm.save()

        errors = self.printer.validate()
        if errors:
            messagebox.showerror("环境错误", "\n".join(errors))
            return

        ok = messagebox.askyesno("测试打印",
            "将打印第一条记录测试效果。\n"
            "模板一次出2张标签。\n\n确认继续？")
        if not ok:
            return

        self.status_text.set("正在测试打印第1条...")
        self._run_in_thread(self._test_print_worker, self._test_print_callback)

    def _test_print_worker(self):
        csv_path = generate_csv(self.df)
        success, stdout, stderr = self.printer.print_single(csv_path, 0)
        return {"success": success, "stdout": stdout, "stderr": stderr,
                "csv_path": csv_path}

    def _test_print_callback(self, result):
        if result.get("error"):
            messagebox.showerror("异常", result["error"])
            return
        if result["success"]:
            self.csv_path = result["csv_path"]
            self.status_text.set("✓ 测试打印已发送（2张标签），请检查效果")
            messagebox.showinfo("完成", "测试打印已发送（2张标签）。\n请检查标签效果是否正确。")
        else:
            messagebox.showerror("失败", result["stderr"] or "未知错误")


# ============================================================
# Entry point
# ============================================================

def main():
    if not PANDAS_AVAILABLE:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少依赖库",
            "程序需要 pandas 和 openpyxl 库。\n\n"
            "请执行: pip install pandas openpyxl\n\n"
            "安装后重新启动。")
        sys.exit(1)

    root = tk.Tk()
    try:
        app = LabelPrinterApp(root)
        logger.info("程序启动 — 原料留样标签打印 v2.1")
        root.mainloop()
    except Exception as e:
        logger.critical(f"程序崩溃: {e}", exc_info=True)
        try:
            messagebox.showerror("程序错误",
                f"程序遇到严重错误:\n\n{e}\n\n"
                f"详细信息已记录到 label_printer.log。")
        except Exception:
            pass
        sys.exit(1)
    finally:
        logger.info("程序退出")

if __name__ == "__main__":
    main()
