"""
开题报告格式校验助手 - 桌面 GUI 版本

基于 customtkinter 构建的现代化图形界面，封装 Format_verify_tool 的核心校验逻辑。
用法：python gui_main.py
"""

import os
import sys
import threading
import webbrowser
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 确保同目录下的模块可被导入
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from Format_verify_tool import analyze_proposal, HTMLReporter, _default_config


# ─────────────────────────────────────────────────────────────
# 全局主题配置
# ─────────────────────────────────────────────────────────────

ctk.set_appearance_mode("light")    # 强制明亮模式，保证清新感
ctk.set_default_color_theme("blue") # 蓝色主色调

# 设计常量
BG_COLOR      = "#F0F4F8"   # 主窗口背景：极淡灰蓝
CARD_BG       = "#FFFFFF"   # 卡片背景：纯白
CARD_RADIUS   = 20          # 卡片圆角
BTN_BLUE      = "#2563EB"   # 按钮主色：现代亮蓝
BTN_BLUE_HOV  = "#1D4ED8"   # 按钮悬停色：深蓝
BTN_DISABLED  = "#CBD5E1"   # 按钮禁用色
TEXT_PRIMARY   = "#1E293B"   # 主文字色：深灰蓝
TEXT_SECONDARY = "#64748B"   # 次文字色：中灰
TEXT_MUTED     = "#94A3B8"   # 弱文字色：浅灰
DIVIDER_COLOR  = "#E2E8F0"   # 分割线色
STATUS_RUNNING = "#F59E0B"   # 状态：运行中（琥珀）
STATUS_OK      = "#10B981"   # 状态：成功（翡翠绿）
STATUS_ERR     = "#EF4444"   # 状态：失败（红）
FOOTER_COLOR   = "#9CA3AF"   # 页脚文字色


# ─────────────────────────────────────────────────────────────
# 主窗口
# ─────────────────────────────────────────────────────────────

class App(ctk.CTk):
    """开题报告格式校验助手主窗口"""

    WINDOW_W, WINDOW_H = 560, 440
    SUPPORTED_EXT = ".docx"

    def __init__(self):
        super().__init__()

        # ── 窗口基础属性 ──
        self.title("培正学院 - 开题报告格式校验助手")
        self.geometry(f"{self.WINDOW_W}x{self.WINDOW_H}")
        self.minsize(self.WINDOW_W, self.WINDOW_H)
        self.resizable(False, False)
        self.configure(fg_color=BG_COLOR)

        # 居中显示
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.WINDOW_W) // 2
        y = (self.winfo_screenheight() - self.WINDOW_H) // 2
        self.geometry(f"+{x}+{y}")

        # 内部状态
        self._selected_file: str | None = None
        self._is_running: bool = False

        # 构建 UI
        self._build_ui()

    # ─────────────────────────────────────────────────────────
    # UI 构建
    # ─────────────────────────────────────────────────────────

    def _build_ui(self):
        """搭建全部界面组件：白色居中卡片 + 底部署名"""

        # ══════════════════════════════════════════════════════
        # 根容器：垂直排列卡片和页脚
        # ══════════════════════════════════════════════════════
        root = ctk.CTkFrame(self, fg_color="transparent")
        root.place(relx=0.5, rely=0.5, anchor="center")

        # ══════════════════════════════════════════════════════
        # 白色卡片主框架 —— 悬浮感核心
        # ══════════════════════════════════════════════════════
        card = ctk.CTkFrame(
            root,
            fg_color=CARD_BG,
            corner_radius=CARD_RADIUS,
            border_width=1,
            border_color=DIVIDER_COLOR,
        )
        card.pack(padx=30, pady=(0, 16))

        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(padx=40, pady=36, fill="both", expand=True)

        # ── 标题区 ──
        ctk.CTkLabel(
            card_inner,
            text="🎓 培正学院开题报告校验助手",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=TEXT_PRIMARY,
        ).pack(anchor="center")

        ctk.CTkLabel(
            card_inner,
            text="毕业季神器 · 本地运行 · 保护隐私",
            font=ctk.CTkFont(size=13),
            text_color=TEXT_MUTED,
        ).pack(anchor="center", pady=(6, 0))

        # ── 分割线 ──
        ctk.CTkFrame(card_inner, height=1, fg_color=DIVIDER_COLOR).pack(
            fill="x", pady=(22, 0)
        )

        # ── 交互区 ──
        action = ctk.CTkFrame(card_inner, fg_color="transparent")
        action.pack(fill="x", pady=(22, 0))

        # 选择文件按钮
        self._btn_select = ctk.CTkButton(
            action,
            text="📂  选择 .docx 文件",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            corner_radius=10,
            fg_color=BTN_BLUE,
            hover_color=BTN_BLUE_HOV,
            command=self._on_select_file,
        )
        self._btn_select.pack(fill="x")

        # 文件路径标签
        self._lbl_file = ctk.CTkLabel(
            action,
            text="尚未选择文件",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
            wraplength=440,
        )
        self._lbl_file.pack(pady=(12, 0))

        # ── 高级配置区 ──
        ctk.CTkFrame(action, height=1, fg_color=DIVIDER_COLOR).pack(
            fill="x", pady=(16, 0)
        )
        ctk.CTkLabel(
            action,
            text="⚙️ 高级自定义参数配置（可选）",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_SECONDARY,
        ).pack(anchor="w", pady=(10, 6))

        cfg_row = ctk.CTkFrame(action, fg_color="transparent")
        cfg_row.pack(fill="x")

        # 配置项：预期行距
        self._ent_spacing = self._make_cfg_field(cfg_row, "预期行距", "1.5")
        # 配置项：文献数量下限
        self._ent_min_refs = self._make_cfg_field(cfg_row, "文献数量≥", "10")
        # 配置项：综述字数下限
        self._ent_min_words = self._make_cfg_field(cfg_row, "字数≥", "1800")
        # 配置项：综述字数上限
        self._ent_max_words = self._make_cfg_field(cfg_row, "字数≤", "2200")

        # ── 规则开关控制区 ──
        ctk.CTkFrame(action, height=1, fg_color=DIVIDER_COLOR).pack(
            fill="x", pady=(16, 0)
        )
        ctk.CTkLabel(
            action,
            text="🔘 规则开关（可选禁用特定检查项）",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_SECONDARY,
        ).pack(anchor="w", pady=(10, 6))

        toggle_row = ctk.CTkFrame(action, fg_color="transparent")
        toggle_row.pack(fill="x")

        self._chk_tutor_space = ctk.CTkCheckBox(
            toggle_row,
            text="导师职称空格规范",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_PRIMARY,
            fg_color=BTN_BLUE,
            hover_color=BTN_BLUE_HOV,
            border_color=DIVIDER_COLOR,
            corner_radius=4,
            checkbox_width=18,
            checkbox_height=18,
        )
        self._chk_tutor_space.pack(side="left", expand=True, padx=(0, 8))
        self._chk_tutor_space.select()  # 默认开启

        self._chk_indent = ctk.CTkCheckBox(
            toggle_row,
            text="参考文献悬挂缩进",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_PRIMARY,
            fg_color=BTN_BLUE,
            hover_color=BTN_BLUE_HOV,
            border_color=DIVIDER_COLOR,
            corner_radius=4,
            checkbox_width=18,
            checkbox_height=18,
        )
        self._chk_indent.pack(side="left", expand=True, padx=(0, 8))
        self._chk_indent.select()  # 默认开启

        self._chk_timeline = ctk.CTkCheckBox(
            toggle_row,
            text="进度安排时间线顺叙",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_PRIMARY,
            fg_color=BTN_BLUE,
            hover_color=BTN_BLUE_HOV,
            border_color=DIVIDER_COLOR,
            corner_radius=4,
            checkbox_width=18,
            checkbox_height=18,
        )
        self._chk_timeline.pack(side="left", expand=True)
        self._chk_timeline.select()  # 默认开启

        # ── 开始校验按钮 ──
        self._btn_run = ctk.CTkButton(
            action,
            text="🚀  开始校验",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            corner_radius=12,
            state="disabled",
            fg_color=BTN_DISABLED,
            hover_color=BTN_DISABLED,
            command=self._on_run_verify,
        )
        self._btn_run.pack(fill="x", pady=(18, 0))

        # ── 状态标签 ──
        self._lbl_status = ctk.CTkLabel(
            action,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
        )
        self._lbl_status.pack(pady=(12, 0))

        # ══════════════════════════════════════════════════════
        # 页脚署名
        # ══════════════════════════════════════════════════════
        ctk.CTkLabel(
            root,
            text="Developed with ❤️ by 林格  |  谨以此工具献给被格式折磨的同学们",
            font=ctk.CTkFont(size=12),
            text_color=FOOTER_COLOR,
        ).pack(pady=(0, 4))

    # ─────────────────────────────────────────────────────────
    # 事件处理
    # ─────────────────────────────────────────────────────────

    def _on_select_file(self):
        """弹出文件选择对话框，仅允许 .docx"""
        if self._is_running:
            return

        path = filedialog.askopenfilename(
            title="选择开题报告文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if not path:
            return  # 用户取消

        # 校验扩展名
        if not path.lower().endswith(self.SUPPORTED_EXT):
            messagebox.showwarning("文件格式错误", "请选择 .docx 格式的 Word 文件。")
            return

        self._selected_file = path
        self._update_file_label(path)
        self._enable_run_button()

    def _on_run_verify(self):
        """点击「开始校验」：在后台线程中执行校验，防止 UI 冻结"""
        if self._is_running or not self._selected_file:
            return

        self._set_running(True)

        # 提取用户自定义配置
        custom_config = self._build_config()

        # 后台线程执行耗时操作
        thread = threading.Thread(
            target=self._run_verify_task, args=(custom_config,), daemon=True
        )
        thread.start()

    # ─────────────────────────────────────────────────────────
    # 后台任务
    # ─────────────────────────────────────────────────────────

    def _run_verify_task(self, custom_config: dict):
        """
        在子线程中执行完整校验流程：
        analyze_proposal → HTMLReporter.generate → webbrowser.open
        """
        try:
            file_path = self._selected_file

            # 1. 执行核心校验（传入用户自定义配置）
            result = analyze_proposal(file_path, config=custom_config)

            # 2. 生成 HTML 报告
            doc_name = os.path.basename(file_path)
            reporter = HTMLReporter(result, doc_name=doc_name)
            report_path = reporter.generate("report.html")

            # 3. 在浏览器中打开
            webbrowser.open("file://" + report_path)

            # 4. 回到主线程更新 UI
            summary = result.get("summary", {})
            msg = (
                f"校验完成！共 {summary.get('total_checks', 0)} 项检查，"
                f"通过 {summary.get('passed', 0)} 项，"
                f"失败 {summary.get('failed', 0)} 项。"
                f"报告已在浏览器中打开。"
            )
            self.after(0, self._on_verify_done, True, msg)

        except Exception as e:
            self.after(0, self._on_verify_done, False, f"校验出错：{e}")

    def _on_verify_done(self, success: bool, message: str):
        """校验结束后在主线程恢复 UI 状态"""
        self._set_running(False)
        if success:
            self._set_status(message, color=STATUS_OK)
            messagebox.showinfo("校验完成", message)
        else:
            self._set_status(message, color=STATUS_ERR)
            messagebox.showerror("校验失败", message)

    # ─────────────────────────────────────────────────────────
    # UI 状态辅助方法
    # ─────────────────────────────────────────────────────────

    @staticmethod
    def _make_cfg_field(parent, label: str, default: str) -> ctk.CTkEntry:
        """在一行中创建一个带标签的小输入框"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", expand=True, padx=(0, 6))
        ctk.CTkLabel(
            frame, text=label,
            font=ctk.CTkFont(size=11),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")
        entry = ctk.CTkEntry(
            frame,
            width=70, height=28,
            corner_radius=6,
            font=ctk.CTkFont(size=12),
            border_width=1,
            border_color=DIVIDER_COLOR,
            fg_color="#F8FAFC",
            text_color=TEXT_PRIMARY,
            justify="center",
        )
        entry.pack(fill="x", pady=(2, 0))
        entry.insert(0, default)
        return entry

    def _build_config(self) -> dict:
        """从输入框提取用户配置，进行类型转换与异常保护"""
        defaults = _default_config()
        cfg = {}

        # 预期行距
        try:
            cfg["line_spacing"] = float(self._ent_spacing.get())
        except (ValueError, TypeError):
            cfg["line_spacing"] = defaults["line_spacing"]

        # 文献数量下限
        try:
            cfg["min_refs"] = int(self._ent_min_refs.get())
        except (ValueError, TypeError):
            cfg["min_refs"] = defaults["min_refs"]

        # 综述字数范围
        try:
            cfg["min_words"] = int(self._ent_min_words.get())
        except (ValueError, TypeError):
            cfg["min_words"] = defaults["min_words"]

        try:
            cfg["max_words"] = int(self._ent_max_words.get())
        except (ValueError, TypeError):
            cfg["max_words"] = defaults["max_words"]

        # 确保 min <= max
        if cfg["min_words"] > cfg["max_words"]:
            cfg["min_words"], cfg["max_words"] = cfg["max_words"], cfg["min_words"]

        # 规则开关（勾选框布尔值）
        cfg["check_tutor_space"] = self._chk_tutor_space.get() == 1
        cfg["check_indent"] = self._chk_indent.get() == 1
        cfg["check_timeline"] = self._chk_timeline.get() == 1

        return cfg

    def _update_file_label(self, path: str):
        """更新文件路径标签，过长时截断显示"""
        display = self._truncate_path(path, max_len=55)
        self._lbl_file.configure(text=display, text_color=TEXT_SECONDARY)

    def _enable_run_button(self):
        """启用「开始校验」按钮并切换为活跃色"""
        self._btn_run.configure(
            state="normal",
            fg_color=BTN_BLUE,
            hover_color=BTN_BLUE_HOV,
        )

    def _set_running(self, running: bool):
        """切换运行中状态：禁用按钮、显示提示"""
        self._is_running = running
        if running:
            self._btn_select.configure(state="disabled")
            self._btn_run.configure(
                state="disabled",
                text="⏳  正在校验中，请稍候...",
                fg_color=BTN_DISABLED,
                hover_color=BTN_DISABLED,
            )
            self._set_status("正在校验中，请稍候...", color=STATUS_RUNNING)
        else:
            self._btn_select.configure(state="normal")
            self._btn_run.configure(
                state="normal" if self._selected_file else "disabled",
                text="🚀  开始校验",
                fg_color=BTN_BLUE if self._selected_file else BTN_DISABLED,
                hover_color=BTN_BLUE_HOV if self._selected_file else BTN_DISABLED,
            )

    def _set_status(self, text: str, color: str = "gray"):
        """更新底部状态标签"""
        self._lbl_status.configure(text=text, text_color=color)

    @staticmethod
    def _truncate_path(path: str, max_len: int = 55) -> str:
        """
        截断过长的路径，保留盘符/前缀和文件名，中间用 ... 连接。
        例：E:\\Very\\Long\\Path\\...\\report.docx
        """
        if len(path) <= max_len:
            return path
        name = os.path.basename(path)
        prefix_len = max_len - len(name) - 5  # 5 = len(" ...\\")
        if prefix_len < 6:
            # 路径名本身就很长，直接截断尾部
            return path[:max_len - 3] + "..."
        return path[:prefix_len] + " ...\\" + name


# ─────────────────────────────────────────────────────────────
# 启动入口
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
