#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能文本自动断行工具1.2.0 (纯黑深色主题 - 单元格卡片流架构)
功能特性：
1. 视觉设计：全套沉浸式科技暗黑/深色主题（Dark Theme），高对比度护眼配色。
2. 窗口自适应防遮挡：固定合理的推荐尺寸并锁定安全最小尺寸，工具栏分行自适应布局，窗口缩放时按钮文字绝不被挤压遮挡。
3. 单元格架构：左侧一个单元格放一个文本，支持动态添加、删除、重命名单元格。
4. 嵌套盒子展示：右侧外层大盒子（单元格X）里面包裹内层小盒子展示断行结果。
5. 智能断行算法：支持两种模式切换 —— ①句子完整优先（默认，以完整句子为单位分行，优先保证一句话不被拆断）；②贴近字符数优先（就近任意标点断行，更贴近设定字符数上限，可能拆开长句）；每行断开后空一行隔开。
6. 批量处理支持：支持一次性导入多个 TXT 文件为独立单元格，支持按单元格批量导出独立文件或合并导出。
7. 便捷操作：每个单元格支持单独清理换行、单独复制结果、删除；顶部支持一键全部断行与全部复制。
"""

import os
import re
import sys
import time
import csv
import io
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Windows 系统下自动隐藏后台黑色控制台窗口
if sys.platform.startswith("win"):
    try:
        import ctypes
        console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if console_hwnd != 0:
            ctypes.windll.user32.ShowWindow(console_hwnd, 0)
    except Exception:
        pass

# 句子结束标点：分句的“硬边界”。正常情况下只会在这里断行，绝不会把一句话从中间截断
SENTENCE_END_CHARS = set("。！？…!?.")
# 从句/语气停顿标点：仅当单个句子本身长度已超过 max_chars 时，才会退化到在这里就近断行
CLAUSE_PAUSE_CHARS = set("，、；：,;:—~")
# 标点符号后紧随的闭合字符（避免引号、括号与前面的标点分离）
CLOSING_CHARS = set("”’'\"）】》」』〕)]}>")

# ===== 暗黑/深色主题色彩规范 =====
COLOR_MAIN_BG = "#0b0f19"       # 应用程序主背景（深蓝黑）
COLOR_PANEL_BG = "#0f172a"      # 滚动容器与面板底色
COLOR_CARD_BG = "#1e293b"       # 单元格卡片与外层大盒子背景
COLOR_INPUT_BG = "#0f172a"      # 文本框内层输入区与小盒子背景
COLOR_BORDER = "#334155"        # 卡片与控件边框线条颜色
COLOR_TEXT_MAIN = "#f8fafc"     # 主要高亮文字颜色（纯白）
COLOR_TEXT_SUB = "#94a3b8"      # 次要信息文字颜色（淡浅灰）
COLOR_ACCENT_BLUE = "#3b82f6"   # 科技蓝重点强调色
COLOR_ACCENT_HOVER = "#2563eb"  # 按钮悬浮高亮色
COLOR_BTN_BG = "#334155"        # 次级功能按钮深色背景
COLOR_BTN_FG = "#f1f5f9"        # 次级功能按钮文字颜色
COLOR_DELETE_FG = "#f87171"     # 删除操作醒目红字


# 断行模式标识
WRAP_MODE_SENTENCE_FIRST = "sentence"  # 句子完整优先：绝不把一句话拆开，只有单句超长才退化断行
WRAP_MODE_STRICT_LIMIT = "strict"      # 贴近字符数优先：就近选任意标点（含逗号）断行，尽量贴近上限

# 两种模式共用：所有可作为断行候选点的标点（句子结束标点 + 从句停顿标点）
ALL_PUNCTUATION_CHARS = SENTENCE_END_CHARS | CLAUSE_PAUSE_CHARS

# 单个导入文件的大小上限（字节）：防止误选超大文件导致内存暴涨、界面卡死
MAX_IMPORT_BYTES = 20 * 1024 * 1024


def _sanitize_filename(name: str, fallback: str, max_len: int = 80) -> str:
    """
    将单元格标题转换为安全的文件名主体：
    替换非法字符与控制字符，去掉首尾空格和句点（Windows 不允许以句点/空格结尾），
    限制长度；结果为空时使用 fallback。
    """
    name = re.sub(r'[\x00-\x1f\\/:*?"<>|]', "_", name or "")
    name = name.strip().strip(".").strip()
    name = name[:max_len].rstrip(" .")
    return name if name else fallback


def _split_into_sentences(para: str):
    """
    将一段文本按句子结束标点（。！？…!? 等）切分成若干个完整的“句子”片段。
    句末标点紧跟的引号/括号（如“他说完了。”）会保留在同一句结尾，不会被拆开。
    """
    sentences = []
    start = 0
    n = len(para)
    i = 0
    while i < n:
        if para[i] in SENTENCE_END_CHARS:
            end = i + 1
            while end < n and para[end] in CLOSING_CHARS:
                end += 1
            sentences.append(para[start:end])
            start = end
            i = end
        else:
            i += 1
    if start < n:
        sentences.append(para[start:])
    return [s for s in sentences if s]


def _find_best_cut(text: str, cur: int, limit: int, total_len: int, punctuation_set):
    """
    在 [cur, limit) 范围内，从给定标点集合中寻找最贴近 limit 的一个断行位置。
    标点后紧随的引号/括号会一并保留在行末，避免与前面的标点分离。
    找不到则返回 -1。
    """
    min_prefer = max(1, int((limit - cur) * 0.3))
    candidates = []
    for i in range(cur, limit):
        if text[i] in punctuation_set:
            cut_pos = i + 1
            valid_quote = True
            while cut_pos < total_len and text[cut_pos] in CLOSING_CHARS:
                if cut_pos + 1 <= limit:
                    cut_pos += 1
                else:
                    valid_quote = False
                    break
            if valid_quote and cur < cut_pos <= limit:
                candidates.append(cut_pos)

    preferred = [c for c in candidates if c >= cur + min_prefer]
    if preferred:
        return max(preferred)
    if candidates:
        return max(candidates)
    return -1


def _break_long_sentence(sentence: str, max_chars: int):
    """
    专门处理“单个句子本身长度就超过 max_chars”的极端情况：
    在不超过 max_chars 的前提下，就近寻找逗号、顿号等次级停顿标点断行；
    实在找不到时退化为按空格或直接硬切分，保证每一行都不超过 max_chars。
    只有一句话确实装不进一行时才会调用到这里，因此“拆句子”是不得已而为之。
    """
    lines = []
    cur = 0
    total_len = len(sentence)

    while cur < total_len:
        while cur < total_len and sentence[cur] in (" ", "\t"):
            cur += 1
        if cur >= total_len:
            break

        remaining_len = total_len - cur
        if remaining_len <= max_chars:
            lines.append(sentence[cur:])
            break

        limit = cur + max_chars
        best_cut = _find_best_cut(sentence, cur, limit, total_len, CLAUSE_PAUSE_CHARS)

        if best_cut != -1:
            lines.append(sentence[cur:best_cut].rstrip())
            cur = best_cut
        else:
            cut_pos = limit
            space_pos = sentence.rfind(" ", cur + int(max_chars * 0.5), limit)
            if space_pos != -1 and space_pos > cur:
                cut_pos = space_pos + 1
            lines.append(sentence[cur:cut_pos].rstrip())
            cur = cut_pos

    return lines


def _wrap_paragraph_sentence_first(para: str, max_chars: int):
    """
    “句子完整优先”模式（默认）：
    1. 先把段落按句子结束标点切分成一个个完整的句子。
    2. 从前到后把整句整句地攒进同一行：只要加入下一句后总长度仍不超过 max_chars，
       就继续累加；一旦加入下一句会超过 max_chars，就把已攒好的内容整体作为一行输出，
       再从这一句开始重新累加。断行基本只会发生在两句话之间，
       不会出现"一句话被从中间拦腰截断、看起来变成两句话"的不自然情况。
    3. 只有当单独一句话本身长度就已经超过 max_chars（极端长句）时，才会退化为按逗号、
       顿号、分号等次级停顿标点就近断行，尽量断在语气停顿处，实在没有的话再按空格或硬切。
    """
    sentences = _split_into_sentences(para)
    lines = []
    line_buf = ""

    for sent in sentences:
        stripped_sent = sent.lstrip(" \t")
        piece = stripped_sent if not line_buf else sent
        candidate = line_buf + piece

        if len(candidate) <= max_chars:
            line_buf = candidate
            continue

        # 加入这一整句会超出限制：先把已经攒好的完整句子输出为一行
        if line_buf:
            lines.append(line_buf.rstrip())
            line_buf = ""

        if len(stripped_sent) <= max_chars:
            # 单独这一句还能装进一行，作为新一行的起点继续往后攒
            line_buf = stripped_sent
        else:
            # 单句本身就超过 max_chars，退化为逗号等次级标点就近断行
            lines.extend(_break_long_sentence(stripped_sent, max_chars))
            line_buf = ""

    if line_buf:
        lines.append(line_buf.rstrip())

    return lines


def _wrap_paragraph_strict_limit(para: str, max_chars: int):
    """
    “贴近字符数优先”模式：
    在不超过 max_chars 的前提下，就近选择任意标点（含逗号、句号等）断行，
    尽量让每一行的字符数都贴近设定的上限，但可能会把长句子从中间拆开，
    因此行数通常比“句子完整优先”模式更少、每行更饱满。
    """
    lines = []
    cur = 0
    total_len = len(para)

    while cur < total_len:
        while cur < total_len and para[cur] in (" ", "\t"):
            cur += 1
        if cur >= total_len:
            break

        remaining_len = total_len - cur
        if remaining_len <= max_chars:
            lines.append(para[cur:])
            break

        limit = cur + max_chars
        best_cut = _find_best_cut(para, cur, limit, total_len, ALL_PUNCTUATION_CHARS)

        if best_cut != -1:
            lines.append(para[cur:best_cut].rstrip())
            cur = best_cut
        else:
            cut_pos = limit
            space_pos = para.rfind(" ", cur + int(max_chars * 0.5), limit)
            if space_pos != -1 and space_pos > cur:
                cut_pos = space_pos + 1
            lines.append(para[cur:cut_pos].rstrip())
            cur = cut_pos

    return lines


def smart_punctuation_wrap(text: str, max_chars: int = 150, mode: str = WRAP_MODE_SENTENCE_FIRST) -> str:
    """
    智能断行核心算法，支持两种可切换的模式：
    - "sentence"（句子完整优先，默认）：以完整句子为最小单位分行，优先保证一句话不被从中间
      拆断，只有单句本身超过 max_chars 时才退化为逗号等次级标点就近断行。
    - "strict"（贴近字符数优先）：在不超过 max_chars 的前提下，就近选择任意标点（含逗号）
      断行，尽量让每行更贴近设定的字符数上限，但可能会把长句子从中间拆开。
    断行后的结果中，每行之间空一行隔开。
    """
    if not text:
        return ""

    paragraphs = text.splitlines()
    wrapped_lines = []

    for para in paragraphs:
        if not para.strip():
            wrapped_lines.append("")
            continue

        if mode == WRAP_MODE_STRICT_LIMIT:
            wrapped_lines.extend(_wrap_paragraph_strict_limit(para, max_chars))
        else:
            wrapped_lines.extend(_wrap_paragraph_sentence_first(para, max_chars))

    valid_lines = [line for line in wrapped_lines if line.strip()]
    return "\n\n".join(valid_lines)



def clean_newlines(text: str) -> str:
    """
    清除文本中的换行符：
    若换行处两侧均为英文字母或数字，保留一个空格避免单词粘连；
    其它情况直接移除换行符。
    """
    lines = text.splitlines()
    cleaned_parts = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if not cleaned_parts:
            cleaned_parts.append(stripped)
        else:
            prev = cleaned_parts[-1]
            if (
                prev
                and stripped
                and prev[-1].isalnum()
                and stripped[0].isalnum()
                and ord(prev[-1]) < 128
                and ord(stripped[0]) < 128
            ):
                cleaned_parts.append(" " + stripped)
            else:
                cleaned_parts.append(stripped)
    return "".join(cleaned_parts)


class ScrollableFrame(tk.Frame):
    """深色极简平滑滚动容器组件，支持鼠标滚轮垂直平滑惯性滚动"""

    def __init__(self, container, *args, **kwargs):
        super().__init__(container, bg=COLOR_MAIN_BG, *args, **kwargs)

        self.canvas = tk.Canvas(
            self,
            borderwidth=0,
            highlightthickness=0,
            bg=COLOR_MAIN_BG,
        )
        # 采用极简暗黑现代滚动条样式
        self.scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.canvas.yview,
            style="Dark.Vertical.TScrollbar",
        )
        self.scrollable_content = tk.Frame(self.canvas, bg=COLOR_MAIN_BG)

        self.scrollable_content.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_content, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 随画布宽度自适应扩展内容宽度
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width)
        )

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 平滑惯性滚动状态
        self._velocity = 0.0       # 当前速度（yview 分数/帧）
        self._animating = False    # 动画帧是否正在运行
        self._after_id = None      # after 回调任务 ID
        self._in_area = False      # 鼠标是否在滚动区域内

        # 进入/离开区域时激活或释放全局滚轮监听
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._schedule_leave_check)
        self.scrollable_content.bind("<Enter>", self._on_enter)

    def _on_enter(self, event=None):
        """鼠标进入容器，使用 bind_all 拦截全窗口滚轮事件"""
        self._in_area = True
        # bind_all 使本 ScrollableFrame 优先捕获，return 'break' 阻止子控件默认滚动
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _schedule_leave_check(self, event=None):
        """延迟 80ms 再确认是否真正离开，避免鼠标移向子控件时误判"""
        self.after(80, self._check_real_leave)

    def _check_real_leave(self):
        """检查鼠标是否仍在本容器内，若已真正离开则解除全局滚轮监听"""
        try:
            rx = self.winfo_rootx()
            ry = self.winfo_rooty()
            rw = self.winfo_width()
            rh = self.winfo_height()
            px = self.winfo_pointerx()
            py = self.winfo_pointery()
            if rx <= px <= rx + rw and ry <= py <= ry + rh:
                # 仍在区域内，保持拦截
                return
        except Exception:
            pass
        # 确认离开，解除拦截，恢复子控件默认滚轮行为
        self._in_area = False
        try:
            self.unbind_all("<MouseWheel>")
        except Exception:
            pass

    def _on_mousewheel(self, event):
        """全局滚轮事件处理：叠加速度冲量并启动惯性动画"""
        # 每次滚轮施加速度冲量（delta/120 = ±1 步，乘以 0.018 转为 yview 分数）
        impulse = -1 * (event.delta / 120) * 0.018
        self._velocity += impulse
        # 速度限幅，防止过快飞出
        self._velocity = max(-0.09, min(0.09, self._velocity))
        if not self._animating:
            self._animating = True
            self._animate()
        # 阻止 Text 等子控件消耗滚轮，确保 Canvas 响应
        return "break"

    def _animate(self):
        """惯性动画：每 16ms（≈60fps）按速度推进位置并施加摩擦系数减速"""
        if abs(self._velocity) < 0.0003:
            self._velocity = 0.0
            self._animating = False
            return

        try:
            cur = self.canvas.yview()[0]
        except Exception:
            self._animating = False
            return

        nxt = cur + self._velocity

        # 到达顶部或底部边界：夹紧并停止
        if nxt <= 0.0:
            self.canvas.yview_moveto(0.0)
            self._velocity = 0.0
            self._animating = False
            return
        if nxt >= 1.0:
            self.canvas.yview_moveto(1.0)
            self._velocity = 0.0
            self._animating = False
            return

        self.canvas.yview_moveto(nxt)
        # 摩擦系数 0.82：每帧衰减 18%，约 0.5s 内停止，顺滑自然
        self._velocity *= 0.82
        # 16ms ≈ 60fps 帧率
        self._after_id = self.after(16, self._animate)


def _find_scrollable_ancestor(widget):
    """从给定控件向上查找最近的 ScrollableFrame 祖先容器，找不到则返回 None"""
    w = widget.master
    while w is not None:
        if isinstance(w, ScrollableFrame):
            return w
        try:
            w = w.master
        except Exception:
            break
    return None


class InputCell(tk.Frame):
    """深色待换行文本单元格卡片控件：一个单元格放一个文本"""

    def __init__(self, parent, cell_index: int, on_delete_callback=None, on_change_callback=None, default_title=None, default_content=""):
        super().__init__(parent, bg=COLOR_MAIN_BG, padx=4, pady=4)
        self.on_delete_callback = on_delete_callback
        self.on_change_callback = on_change_callback
        self.cell_index = cell_index

        # 外层单元格卡片容器（深色背景、精致边框）
        self.card = tk.Frame(
            self,
            bg=COLOR_CARD_BG,
            bd=1,
            relief=tk.SOLID,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.card.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部标题控制栏
        top_bar = tk.Frame(self.card, bg=COLOR_CARD_BG)
        top_bar.pack(fill=tk.X, pady=(0, 6))

        # 单元格序号与自定义标题
        self.idx_label = tk.Label(
            top_bar,
            text=f"单元格 {cell_index}:",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_ACCENT_BLUE,
        )
        self.idx_label.pack(side=tk.LEFT)

        self.title_var = tk.StringVar(value=default_title if default_title else f"文本 {cell_index}")
        self.title_entry = tk.Entry(
            top_bar,
            textvariable=self.title_var,
            width=16,
            font=("Microsoft YaHei UI", 9),
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_MAIN,
            insertbackground=COLOR_ACCENT_BLUE,
            relief=tk.SOLID,
            bd=1,
        )
        self.title_entry.pack(side=tk.LEFT, padx=(6, 8))

        # 实时字数显示
        self.count_label = tk.Label(
            top_bar,
            text="字数: 0",
            font=("Microsoft YaHei UI", 8),
            bg=COLOR_CARD_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.count_label.pack(side=tk.LEFT)

        # 删除按钮
        self.btn_del = tk.Button(
            top_bar,
            text="❌ 删除",
            font=("Microsoft YaHei UI", 8, "bold"),
            bg=COLOR_BTN_BG,
            fg=COLOR_DELETE_FG,
            activebackground="#450a0a",
            activeforeground="#fca5a5",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._delete_self,
        )
        self.btn_del.pack(side=tk.RIGHT, padx=(4, 0))

        # 单独清理当前单元格换行按钮
        self.btn_clean = tk.Button(
            top_bar,
            text="🧹 清理换行",
            font=("Microsoft YaHei UI", 8),
            bg=COLOR_BTN_BG,
            fg=COLOR_BTN_FG,
            activebackground=COLOR_BORDER,
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._clean_cell_newlines,
        )
        self.btn_clean.pack(side=tk.RIGHT, padx=(0, 4))

        # 2. 文本输入区域
        self.text_widget = tk.Text(
            self.card,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 10),
            height=6,
            padx=8,
            pady=8,
            relief=tk.SOLID,
            bd=1,
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_MAIN,
            insertbackground=COLOR_ACCENT_BLUE,
            selectbackground=COLOR_ACCENT_HOVER,
            selectforeground="#ffffff",
            undo=True,
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        self.text_widget.bind("<KeyRelease>", self._on_text_change)
        # 普通滚轮：不滚动本单元格内部文本，转发给外层容器整体滚动
        self.text_widget.bind("<MouseWheel>", self._on_wheel_plain)
        # 按住 Ctrl + 滚轮：仅滚动本单元格内部文本内容
        self.text_widget.bind("<Control-MouseWheel>", self._on_wheel_ctrl)

        if default_content:
            self.text_widget.insert("1.0", default_content)
            self._on_text_change()

    def update_index(self, new_index: int):
        """更新单元格编号"""
        self.cell_index = new_index
        self.idx_label.config(text=f"单元格 {new_index}:")
        cur_title = self.title_var.get()
        if re.match(r"^文本 \d+$", cur_title):
            self.title_var.set(f"文本 {new_index}")

    def get_title(self) -> str:
        t = self.title_var.get().strip()
        return t if t else f"文本 {self.cell_index}"

    def get_content(self) -> str:
        return self.text_widget.get("1.0", tk.END).rstrip("\n")

    def set_content(self, text: str, title: str = None):
        if title:
            self.title_var.set(title)
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.insert("1.0", text)
        self._on_text_change()

    def _clean_cell_newlines(self):
        """清除当前单元格中的换行符"""
        raw = self.get_content()
        if raw.strip():
            cleaned = clean_newlines(raw)
            self.text_widget.delete("1.0", tk.END)
            self.text_widget.insert("1.0", cleaned)
            self._on_text_change()

    def _delete_self(self):
        if self.on_delete_callback:
            self.on_delete_callback(self)

    def _on_wheel_plain(self, event):
        """普通滚轮悬停在单元格文本上：不滚动本单元格内容，转发给外层容器整体滚动"""
        frame = _find_scrollable_ancestor(self)
        if frame is not None:
            frame._on_mousewheel(event)
        return "break"

    def _on_wheel_ctrl(self, event):
        """Ctrl + 滚轮悬停在单元格文本上：仅滚动本单元格内部文本内容"""
        self.text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _on_text_change(self, event=None):
        content = self.get_content()
        self.count_label.config(text=f"字数: {len(content)}")
        if self.on_change_callback:
            self.on_change_callback()


class OutputBox(tk.Frame):
    """深色断行结果嵌套盒子控件：外层大盒子里面包着内层小盒子展示断行结果"""

    def __init__(self, parent, cell_index: int, title: str, wrapped_content: str, on_copy_callback=None):
        super().__init__(parent, bg=COLOR_MAIN_BG, padx=4, pady=4)
        self.cell_index = cell_index
        self.title = title
        self.wrapped_content = wrapped_content
        self.on_copy_callback = on_copy_callback

        # 1. 外层大盒子（卡片容器：深色面板、微亮描边）
        self.outer_box = tk.Frame(
            self,
            bg=COLOR_CARD_BG,
            bd=1,
            relief=tk.SOLID,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        self.outer_box.pack(fill=tk.BOTH, expand=True)

        # 外层盒子顶部信息栏
        top_bar = tk.Frame(self.outer_box, bg=COLOR_CARD_BG)
        top_bar.pack(fill=tk.X, pady=(0, 6))

        # 标题标签：明确标出外层盒子所对应的单元格编号与名称
        box_title = f"📦 单元格 {cell_index} 结果: {title}"
        self.title_label = tk.Label(
            top_bar,
            text=box_title,
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_ACCENT_BLUE,
        )
        self.title_label.pack(side=tk.LEFT)

        # 统计信息（行数与字符数）
        valid_lines = [l for l in wrapped_content.splitlines() if l.strip()]
        lines_count = len(valid_lines)
        chars_count = len(wrapped_content)
        self.stat_label = tk.Label(
            top_bar,
            text=f"({lines_count} 行 | {chars_count} 字符)",
            font=("Microsoft YaHei UI", 8),
            bg=COLOR_CARD_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.stat_label.pack(side=tk.LEFT, padx=(6, 0))

        # 外层盒子右上角：一键复制该单元格结果
        self.btn_copy = tk.Button(
            top_bar,
            text="📋 复制该文本",
            font=("Microsoft YaHei UI", 8, "bold"),
            bg=COLOR_ACCENT_BLUE,
            fg="#ffffff",
            activebackground=COLOR_ACCENT_HOVER,
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            padx=10,
            pady=2,
            command=self._copy_content,
        )
        self.btn_copy.pack(side=tk.RIGHT)

        # 2. 内层小盒子（嵌套在内，纯净深黑展示区）
        self.inner_box = tk.Frame(
            self.outer_box,
            bg=COLOR_INPUT_BG,
            bd=1,
            relief=tk.SOLID,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        self.inner_box.pack(fill=tk.BOTH, expand=True)

        # 内层盒子中的多行文本展示控件（高度按断行结果实际行数自适应，支持手动编辑）
        text_height = max(2, lines_count)
        self.text_widget = tk.Text(
            self.inner_box,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 10),
            height=text_height,
            padx=6,
            pady=6,
            relief=tk.FLAT,
            bd=0,
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_MAIN,
            insertbackground=COLOR_ACCENT_BLUE,
            selectbackground=COLOR_ACCENT_HOVER,
            selectforeground="#ffffff",
            undo=True,
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        self.text_widget.insert("1.0", wrapped_content)
        # 手动编辑后实时刷新行数/字符数统计，并重新自适应高度
        self.text_widget.bind("<KeyRelease>", self._on_text_edited)
        # 宽度发生变化（如窗口缩放）时，按真实换行后的显示行数重新自适应高度，
        # 这样断行结果始终整段显示完整，不需要再在单元格内部上下滑动查看
        self.text_widget.bind("<Configure>", self._auto_fit_height)
        # 普通滚轮：不滚动本结果盒子内部文本，转发给外层容器整体滚动
        self.text_widget.bind("<MouseWheel>", self._on_wheel_plain)
        # 按住 Ctrl + 滚轮：仅滚动本结果盒子内部文本内容
        self.text_widget.bind("<Control-MouseWheel>", self._on_wheel_ctrl)
        # 创建后延迟到布局完成，再按真实渲染宽度校正一次高度
        self.after(30, self._auto_fit_height)

    def _auto_fit_height(self, event=None):
        """让文本框高度刚好等于断行结果实际渲染后的行数：末尾内容不可见就增高，多余空白就收缩"""
        if getattr(self, "_fitting_in_progress", False):
            return  # 防止调整过程中自身触发的 Configure 事件递归重入
        self._fitting_in_progress = True
        try:
            widget = self.text_widget
            widget.update_idletasks()

            if widget.index("1.0") == widget.index("end-1c"):
                return  # 空内容，保持默认最小高度即可

            def last_char_visible():
                return widget.bbox("end-1c") is not None

            current_h = int(widget.cget("height"))
            h = current_h

            if last_char_visible():
                # 末尾内容已完整可见，尝试逐步收缩，去掉多余的空白行，直到刚好还能看到末尾为止
                while h > 2:
                    widget.configure(height=h - 1)
                    widget.update_idletasks()
                    if not last_char_visible():
                        widget.configure(height=h)
                        widget.update_idletasks()
                        break
                    h -= 1
            else:
                # 末尾内容还看不到，持续增高直到完整显示为止
                while not last_char_visible() and h < 1000:
                    h += 1
                    widget.configure(height=h)
                    widget.update_idletasks()
        except Exception:
            pass
        finally:
            self._fitting_in_progress = False

    def get_content(self) -> str:
        """获取当前文本框中的实时内容（包含用户手动编辑后的修改）"""
        return self.text_widget.get("1.0", tk.END).rstrip("\n")

    def _on_text_edited(self, event=None):
        """用户手动编辑结果内容后，实时刷新顶部的行数/字符数统计，并重新自适应高度"""
        content = self.get_content()
        valid_lines = [l for l in content.splitlines() if l.strip()]
        self.stat_label.config(text=f"({len(valid_lines)} 行 | {len(content)} 字符)")
        self._auto_fit_height()

    def _on_wheel_plain(self, event):
        """普通滚轮悬停在结果文本上：不滚动本盒子内容，转发给外层容器整体滚动"""
        frame = _find_scrollable_ancestor(self)
        if frame is not None:
            frame._on_mousewheel(event)
        return "break"

    def _on_wheel_ctrl(self, event):
        """Ctrl + 滚轮悬停在结果文本上：仅滚动本盒子内部文本内容"""
        self.text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _copy_content(self):
        if self.on_copy_callback:
            self.on_copy_callback(self.get_content(), self.title)


class AutoWrapApp(tk.Tk):
    """可视化自动断行窗口应用程序（纯黑深色主题 + 窗口自适应防遮挡）"""

    def __init__(self):
        super().__init__()
        self.title("智能文本自动断行工具1.2.0")

        # 窗口推荐固定与安全最小尺寸适配，避免缩小遮挡按钮文字
        window_width = 1200
        window_height = 760
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        pos_x = max(0, (screen_w - window_width) // 2)
        pos_y = max(0, (screen_h - window_height) // 2)

        self.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
        self.minsize(1050, 620)  # 设置安全底线最小尺寸，杜绝按钮文字被压缩

        # 单元格列表与结果列表
        self.input_cells = []
        self.output_boxes = []
        self.current_results_data = []

        # 尝试设置现代 ttk 主题
        self.style = ttk.Style(self)
        available_themes = self.style.theme_names()
        if "clam" in available_themes:
            self.style.theme_use("clam")

        self._configure_styles()

        # 构建主界面结构
        self._create_header()
        self._create_main_content()
        self._create_footer_controls()
        self._create_status_bar()

        # 默认初始化 2 个待换行单元格
        self.add_input_cell()
        self.add_input_cell()

    def _configure_styles(self):
        """配置深色模式主题样式与控件色彩"""
        self.configure(bg=COLOR_MAIN_BG)

        self.style.configure(".", background=COLOR_MAIN_BG, font=("Microsoft YaHei UI", 10))
        self.style.configure("Header.TLabel", font=("Microsoft YaHei UI", 13, "bold"), foreground=COLOR_TEXT_MAIN, background=COLOR_MAIN_BG)
        self.style.configure("SubHeader.TLabel", font=("Microsoft YaHei UI", 9), foreground=COLOR_TEXT_SUB, background=COLOR_MAIN_BG)
        self.style.configure("Section.TLabel", font=("Microsoft YaHei UI", 10, "bold"), foreground=COLOR_TEXT_MAIN, background=COLOR_MAIN_BG)

        # 1. 彻底移除按钮默认的高亮/浅色描边环与边框
        self.style.layout(
            "Primary.TButton",
            [("Button.padding", {"children": [("Button.label", {"sticky": "nswe"})], "sticky": "nswe"})],
        )
        self.style.configure(
            "Primary.TButton",
            font=("Microsoft YaHei UI", 11, "bold"),
            foreground="#ffffff",
            background=COLOR_ACCENT_BLUE,
            padding=(14, 6),
            borderwidth=0,
            relief="flat",
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", COLOR_ACCENT_HOVER), ("pressed", COLOR_ACCENT_HOVER)],
        )

        self.style.layout(
            "Action.TButton",
            [("Button.padding", {"children": [("Button.label", {"sticky": "nswe"})], "sticky": "nswe"})],
        )
        self.style.configure(
            "Action.TButton",
            font=("Microsoft YaHei UI", 9),
            foreground=COLOR_BTN_FG,
            background=COLOR_BTN_BG,
            padding=(8, 4),
            borderwidth=0,
            relief="flat",
        )
        self.style.map(
            "Action.TButton",
            background=[("active", "#475569"), ("pressed", "#64748b")],
        )

        # 2. 定制极简深色现代滚动条（8px 宽度，无箭头，滑槽完全透明融入背景）
        self.style.layout(
            "Dark.Vertical.TScrollbar",
            [
                (
                    "Vertical.Scrollbar.trough",
                    {
                        "children": [
                            (
                                "Vertical.Scrollbar.thumb",
                                {"expand": "1", "sticky": "nswe"},
                            )
                        ],
                        "sticky": "ns",
                    },
                )
            ],
        )
        self.style.configure(
            "Dark.Vertical.TScrollbar",
            troughcolor="#1e293b",           # 滑轨轨道：深蓝灰（比主背景亮一档，轨道带可见）
            background="#94a3b8",            # 滑块：亮灰蓝（与深底形成高对比）
            bordercolor="#1e293b",
            darkcolor="#94a3b8",             # clam 阴影色：与滑块同色
            lightcolor="#cbd5e1",            # clam 高光色：较亮，微弱立体感
            arrowcolor="#1e293b",
            arrowsize=0,
            width=10,                        # 加宽到 10px 更易点击
            relief="flat",
        )
        self.style.map(
            "Dark.Vertical.TScrollbar",
            background=[("active", "#cbd5e1"), ("pressed", "#e2e8f0")],
            darkcolor=[("active", "#cbd5e1"), ("pressed", "#e2e8f0")],
            lightcolor=[("active", "#cbd5e1"), ("pressed", "#e2e8f0")],
        )

        # 3. 断行模式切换单选按钮（深色主题配色）
        self.style.configure(
            "Mode.TRadiobutton",
            font=("Microsoft YaHei UI", 9),
            foreground=COLOR_TEXT_MAIN,
            background=COLOR_MAIN_BG,
            focuscolor=COLOR_MAIN_BG,
        )
        self.style.map(
            "Mode.TRadiobutton",
            foreground=[("active", COLOR_ACCENT_BLUE), ("selected", COLOR_ACCENT_BLUE)],
            background=[("active", COLOR_MAIN_BG), ("pressed", COLOR_MAIN_BG), ("selected", COLOR_MAIN_BG), ("!active", COLOR_MAIN_BG)],
            indicatorcolor=[("selected", COLOR_ACCENT_BLUE), ("!selected", COLOR_INPUT_BG)],
        )

    def _create_header(self):
        """创建顶部标题与功能提示条"""
        header_frame = tk.Frame(self, bg=COLOR_MAIN_BG, padx=16, pady=10)
        header_frame.pack(side=tk.TOP, fill=tk.X)

        title_lbl = tk.Label(
            header_frame,
            text="📄 智能文本自动断行工具1.2.0",
            font=("Microsoft YaHei UI", 13, "bold"),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_MAIN,
        )
        title_lbl.pack(side=tk.LEFT, anchor="w")


    def _create_main_content(self):
        """创建中间核心区域：左侧单元格输入流与右侧结果流（采用双行工具栏防遮挡）"""
        content_frame = tk.Frame(self, bg=COLOR_MAIN_BG, padx=16, pady=4)
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # ===== 左侧：待换行文本单元格列表 =====
        left_panel = tk.Frame(content_frame, bg=COLOR_MAIN_BG)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left_panel.rowconfigure(1, weight=1)
        left_panel.columnconfigure(0, weight=1)

        # 左侧顶部两行自适应工具栏（防止按钮受挤压文字看不见）
        left_header_container = tk.Frame(left_panel, bg=COLOR_MAIN_BG)
        left_header_container.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        # 第一行：标题与统计信息
        left_row1 = tk.Frame(left_header_container, bg=COLOR_MAIN_BG)
        left_row1.pack(fill=tk.X, pady=(0, 4))
        tk.Label(
            left_row1,
            text="待换行文本单元格",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_MAIN,
        ).pack(side=tk.LEFT)

        self.input_summary_lbl = tk.Label(
            left_row1,
            text="共 0 个单元格 | 0 字",
            font=("Microsoft YaHei UI", 9),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.input_summary_lbl.pack(side=tk.LEFT, padx=(8, 0))

        # 第二行：按钮操作栏（横向充裕排布，绝对不被压缩）
        left_row2 = tk.Frame(left_header_container, bg=COLOR_MAIN_BG)
        left_row2.pack(fill=tk.X)

        ttk.Button(left_row2, text="➕ 添加单元格", style="Action.TButton", command=self.add_input_cell).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(left_row2, text="📂 导入文件为单元格", style="Action.TButton", command=self._import_files_as_cells).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(left_row2, text="🧹 清理所有换行", style="Action.TButton", command=self._clean_all_newlines).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(left_row2, text="🗑️ 清空全部", style="Action.TButton", command=self._clear_all_cells).pack(
            side=tk.LEFT, padx=(0, 4)
        )

        # 第三行：Google 表格一键粘贴按钮（单独一行，避免与其它按钮拥挤）
        left_row3 = tk.Frame(left_header_container, bg=COLOR_MAIN_BG)
        left_row3.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(
            left_row3,
            text="📥 一键粘贴(Google表格)",
            style="Action.TButton",
            command=self._paste_table_as_cells,
        ).pack(side=tk.LEFT, padx=(0, 4))

        # 左侧可滚动单元格列表容器
        self.left_scroll_frame = ScrollableFrame(left_panel)
        self.left_scroll_frame.grid(row=1, column=0, sticky="nsew")

        # ===== 右侧：断行结果列表 =====
        right_panel = tk.Frame(content_frame, bg=COLOR_MAIN_BG)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        right_panel.rowconfigure(1, weight=1)
        right_panel.columnconfigure(0, weight=1)

        # 右侧顶部两行自适应工具栏
        right_header_container = tk.Frame(right_panel, bg=COLOR_MAIN_BG)
        right_header_container.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        # 第一行：标题与统计信息
        right_row1 = tk.Frame(right_header_container, bg=COLOR_MAIN_BG)
        right_row1.pack(fill=tk.X, pady=(0, 4))
        tk.Label(
            right_row1,
            text="断行后结果展示",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_MAIN,
        ).pack(side=tk.LEFT)

        self.output_summary_lbl = tk.Label(
            right_row1,
            text="共 0 个结果盒子",
            font=("Microsoft YaHei UI", 9),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.output_summary_lbl.pack(side=tk.LEFT, padx=(8, 0))

        # 第二行：按钮操作栏
        right_row2 = tk.Frame(right_header_container, bg=COLOR_MAIN_BG)
        right_row2.pack(fill=tk.X)

        ttk.Button(right_row2, text="📋 复制全部结果", style="Action.TButton", command=self._copy_all_output).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(right_row2, text="💾 批量/合并导出...", style="Action.TButton", command=self._export_results).pack(
            side=tk.LEFT, padx=(0, 6)
        )

        # 右侧可滚动结果盒子容器
        self.right_scroll_frame = ScrollableFrame(right_panel)
        self.right_scroll_frame.grid(row=1, column=0, sticky="nsew")

        # 右侧初始占位提示
        self.placeholder_lbl = tk.Label(
            self.right_scroll_frame.scrollable_content,
            text="👉 在左侧各单元格输入文本后，点击下方【⚡ 一键全部断行】\n断行结果将在此处以【大盒子包小盒子】的形式一一呈现！",
            font=("Microsoft YaHei UI", 10),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_SUB,
            pady=60,
        )
        self.placeholder_lbl.pack(fill=tk.BOTH, expand=True)

    def _create_footer_controls(self):
        """创建底部控制栏：一键断行按钮 + 自定义每行最大字符数设置"""
        controls_frame = tk.Frame(self, bg=COLOR_MAIN_BG, padx=16, pady=10)
        controls_frame.pack(side=tk.TOP, fill=tk.X)

        # 1. 一键全部断行主按钮
        self.btn_wrap = ttk.Button(
            controls_frame,
            text="⚡ 一键全部断行",
            style="Primary.TButton",
            command=self.perform_wrap_all,
            cursor="hand2",
        )
        self.btn_wrap.pack(side=tk.LEFT, padx=(0, 16))

        # 2. 自定义每行最大字符数（位于一键断行右侧）
        param_frame = tk.Frame(controls_frame, bg=COLOR_MAIN_BG)
        param_frame.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(
            param_frame,
            text="每行最大字符数：",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_MAIN,
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.width_var = tk.IntVar(value=150)
        self.spin_width = tk.Spinbox(
            param_frame,
            from_=10,
            to=2000,
            increment=10,
            textvariable=self.width_var,
            width=7,
            font=("Microsoft YaHei UI", 10),
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_MAIN,
            buttonbackground=COLOR_CARD_BG,
            buttoncursor="hand2",
            insertbackground=COLOR_ACCENT_BLUE,
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_ACCENT_BLUE,
        )
        self.spin_width.pack(side=tk.LEFT, padx=(0, 12))

        # 断行模式切换：句子完整优先 / 贴近字符数优先
        self.wrap_mode_var = tk.StringVar(value=WRAP_MODE_SENTENCE_FIRST)

        tk.Label(
            param_frame,
            text="断行模式：",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_MAIN,
        ).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Radiobutton(
            param_frame,
            text="📖 句子完整优先",
            variable=self.wrap_mode_var,
            value=WRAP_MODE_SENTENCE_FIRST,
            style="Mode.TRadiobutton",
            command=self._on_wrap_mode_change,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Radiobutton(
            param_frame,
            text="📏 贴近字符数优先",
            variable=self.wrap_mode_var,
            value=WRAP_MODE_STRICT_LIMIT,
            style="Mode.TRadiobutton",
            command=self._on_wrap_mode_change,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.mode_desc_lbl = tk.Label(
            param_frame,
            text="",
            font=("Microsoft YaHei UI", 9),
            bg=COLOR_MAIN_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.mode_desc_lbl.pack(side=tk.LEFT)
        self._update_mode_desc()

    def _on_wrap_mode_change(self):
        """断行模式切换后，更新下方说明文字"""
        self._update_mode_desc()

    def _update_mode_desc(self):
        """根据当前选中的断行模式，刷新对应的说明文字"""
        if self.wrap_mode_var.get() == WRAP_MODE_STRICT_LIMIT:
            desc = "（严格不超此字数，就近任意标点断行，更贴近字符数上限，但可能拆开长句子）"
        else:
            desc = "（以完整句子为单位分行，优先保证一句话不被拆开，仅超长单句才会拆分）"
        self.mode_desc_lbl.config(text=desc)

    def _create_status_bar(self):
        """创建底部状态栏"""
        status_frame = tk.Frame(self, bg=COLOR_PANEL_BG, padx=16, pady=6)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_lbl = tk.Label(
            status_frame,
            text="状态：就绪。一个单元格放一个文本，点击【➕ 添加单元格】可继续增加卡片。",
            font=("Microsoft YaHei UI", 9),
            bg=COLOR_PANEL_BG,
            fg=COLOR_TEXT_SUB,
        )
        self.status_lbl.pack(side=tk.LEFT)

    def add_input_cell(self, title=None, content=""):
        """向左侧列表追加一个待换行文本单元格"""
        new_index = len(self.input_cells) + 1
        cell = InputCell(
            self.left_scroll_frame.scrollable_content,
            cell_index=new_index,
            on_delete_callback=self.delete_input_cell,
            on_change_callback=self.update_input_summary,
            default_title=title,
            default_content=content,
        )
        cell.pack(fill=tk.X, expand=True, pady=4, padx=2)
        self.input_cells.append(cell)
        self.update_input_summary()
        return cell

    def delete_input_cell(self, cell_to_delete: InputCell):
        """删除指定的待换行单元格并重新排列序号"""
        if cell_to_delete in self.input_cells:
            self.input_cells.remove(cell_to_delete)
            cell_to_delete.destroy()

            # 重新编号
            for idx, cell in enumerate(self.input_cells, 1):
                cell.update_index(idx)

            # 保证至少留有 1 个单元格
            if not self.input_cells:
                self.add_input_cell()

            self.update_input_summary()

    def _update_status(self, text: str):
        """统一更新底部状态栏文本"""
        self.status_lbl.config(text=f"状态：{text}")

    def update_input_summary(self):
        """更新左侧单元格总数与总字数统计"""
        total_cells = len(self.input_cells)
        total_chars = sum(len(c.get_content()) for c in self.input_cells)
        non_empty = sum(1 for c in self.input_cells if c.get_content().strip())
        self.input_summary_lbl.config(
            text=f"共 {total_cells} 个单元格 (有效 {non_empty}) | {total_chars} 字"
        )

    def perform_wrap_all(self):
        """执行一键断行：根据左侧所有单元格在右侧生成嵌套盒子"""
        try:
            max_chars = int(self.width_var.get())
            if max_chars < 5:
                raise ValueError
        except Exception:
            messagebox.showwarning("参数提示", "每行最大字符数必须为不小于 5 的正整数！")
            self.spin_width.focus_set()
            return

        valid_cells = [c for c in self.input_cells if c.get_content().strip()]
        if not valid_cells:
            messagebox.showinfo("提示", "左侧单元格中暂无待断行的有效文本！")
            return

        t_start = time.perf_counter()

        # 清空右侧之前的占位文本和旧盒子
        if self.placeholder_lbl.winfo_exists():
            self.placeholder_lbl.pack_forget()

        for box in self.output_boxes:
            box.destroy()
        self.output_boxes.clear()
        self.current_results_data.clear()

        total_lines_all = 0
        total_chars_all = 0
        wrap_mode = self.wrap_mode_var.get()

        # 为每一个单元格在右侧生成对应的大盒子（包着内层小盒子）
        for idx, cell in enumerate(valid_cells, 1):
            raw_text = cell.get_content()
            title = cell.get_title()

            # 按当前选择的断行模式执行智能断行（严格不超过字符上限，空行隔开）
            wrapped = smart_punctuation_wrap(raw_text, max_chars, wrap_mode)

            # 构造大盒子包小盒子的嵌套结果控件
            box = OutputBox(
                self.right_scroll_frame.scrollable_content,
                cell_index=idx,
                title=title,
                wrapped_content=wrapped,
                on_copy_callback=self._copy_single_cell_result,
            )
            box.pack(fill=tk.X, pady=6, padx=2)
            self.output_boxes.append(box)

            self.current_results_data.append({
                "index": idx,
                "title": title,
                "content": wrapped,
            })

            valid_lines = [l for l in wrapped.splitlines() if l.strip()]
            total_lines_all += len(valid_lines)
            total_chars_all += len(wrapped)

        t_elapsed = (time.perf_counter() - t_start) * 1000

        self.output_summary_lbl.config(
            text=f"共生成 {len(self.output_boxes)} 个结果盒子 | 总行数: {total_lines_all} | {total_chars_all} 字符"
        )
        mode_name = "句子完整优先" if wrap_mode == WRAP_MODE_SENTENCE_FIRST else "贴近字符数优先"
        self.status_lbl.config(
            text=f"状态：处理完成！耗时 {t_elapsed:.2f} 毫秒，断行模式「{mode_name}」，已生成 {len(self.output_boxes)} 个断行结果盒子。"
        )

    def _copy_single_cell_result(self, content: str, title: str):
        """单独复制某个单元格的断行结果"""
        self.clipboard_clear()
        self.clipboard_append(content)
        self.update()
        messagebox.showinfo("复制成功", f"已成功将【{title}】的断行结果复制到剪贴板！")

    def _get_live_results(self):
        """从右侧所有结果盒子中实时读取当前文本内容（包含用户手动编辑后的修改）"""
        return [
            {"index": box.cell_index, "title": box.title, "content": box.get_content()}
            for box in self.output_boxes
        ]

    def _copy_all_output(self):
        """一键复制所有单元格断行结果（以规整分隔线串联，读取实时编辑后的内容）"""
        if not self.output_boxes:
            messagebox.showinfo("提示", "右侧尚无断行结果可复制！")
            return

        live_results = self._get_live_results()
        combined_blocks = []
        for item in live_results:
            header = f"====================【{item['title']}】===================="
            combined_blocks.append(f"{header}\n\n{item['content']}")

        full_text = "\n\n\n".join(combined_blocks)
        self.clipboard_clear()
        self.clipboard_append(full_text)
        self.update()
        messagebox.showinfo("成功", f"已将全部 {len(live_results)} 个单元格的断行结果复制到剪贴板！")

    def _clean_all_newlines(self):
        """一键清理所有单元格内部的硬回车"""
        cleaned_count = 0
        for cell in self.input_cells:
            content = cell.get_content()
            if content.strip():
                cell.set_content(clean_newlines(content))
                cleaned_count += 1
        self.update_input_summary()
        self.status_lbl.config(text=f"状态：已成功对 {cleaned_count} 个单元格完成换行清理！")

    def _clear_all_cells(self):
        """清空全部单元格，同步清除右侧断行结果"""
        if messagebox.askyesno("清空确认", "确定要清空左侧全部待换行单元格吗？"):
            # 清除左侧全部输入单元格
            for cell in self.input_cells:
                cell.destroy()
            self.input_cells.clear()
            # 同步清除右侧全部断行结果盒子
            for box in self.output_boxes:
                box.destroy()
            self.output_boxes.clear()
            self.current_results_data.clear()
            # 恢复右侧占位提示文字
            if self.placeholder_lbl.winfo_exists():
                self.placeholder_lbl.pack(fill=tk.BOTH, expand=True)
            # 重置状态栏
            self._update_status("已清空全部单元格与断行结果。")
            # 默认补充两个空单元格
            self.add_input_cell()
            self.add_input_cell()

    def _paste_table_as_cells(self):
        """
        一键粘贴 Google 表格（或 Excel 等）内容：
        从剪贴板读取以制表符分隔列、换行分隔行的表格文本，
        按“一个表格单元格对应软件里一个单元格”的规则依次填入，
        若软件当前单元格数量不够，会自动新增单元格补全到匹配数量。
        """
        try:
            clip_text = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning(
                "粘贴提示",
                "剪贴板中没有可粘贴的文本内容！\n请先在 Google 表格中选中并复制单元格。",
            )
            return

        if not clip_text or not clip_text.strip("\r\n"):
            messagebox.showwarning(
                "粘贴提示",
                "剪贴板中没有可粘贴的文本内容！\n请先在 Google 表格中选中并复制单元格。",
            )
            return

        # Google表格/Excel 复制多个单元格时，采用制表符分隔列、换行分隔行的规则；
        # 若某个单元格内容本身含有制表符/换行/引号，会自动用英文双引号包裹并转义，
        # 用标准 csv 模块按 \t 分隔解析，可以正确区分"单元格内部换行"与"单元格边界换行"。
        normalized = clip_text.replace("\r\n", "\n").replace("\r", "\n")
        normalized = normalized.rstrip("\n")  # 去掉复制时末尾多余的空行

        try:
            rows = list(csv.reader(io.StringIO(normalized), delimiter="\t"))
        except Exception as e:
            messagebox.showerror("解析失败", f"解析剪贴板表格内容失败：\n{e}")
            return

        # 按行优先顺序，展开为一个个独立的表格单元格文本
        cell_texts = [cell for row in rows for cell in row]

        if not cell_texts:
            messagebox.showwarning("粘贴提示", "未能从剪贴板中识别出任何单元格内容！")
            return

        total = len(cell_texts)
        current_total = len(self.input_cells)
        need_new = max(0, total - current_total)

        extra_hint = f"并自动新增 {need_new} 个单元格补全。" if need_new else "现有单元格数量已足够，无需新增。"
        if not messagebox.askyesno(
            "一键粘贴确认",
            f"检测到剪贴板中共有 {total} 个表格单元格文本。\n\n"
            f"将按顺序依次填入左侧第 1 ~ {total} 个单元格（会覆盖对应位置的原有内容），{extra_hint}\n\n"
            "是否继续？",
        ):
            return

        # 数量不够时自动新增单元格，直到与剪贴板中的单元格数量一致
        for _ in range(need_new):
            self.add_input_cell()

        for idx, text in enumerate(cell_texts):
            self.input_cells[idx].set_content(text)

        self.update_input_summary()
        self._update_status(f"已从剪贴板一键粘贴 {total} 个 Google 表格单元格文本，共填入 {total} 个单元格。")

    def _import_files_as_cells(self):
        """多选导入多个 TXT 文件，每个文件自动创建一个独立的单元格"""
        file_paths = filedialog.askopenfilenames(
            title="选择要导入的文本文件 (支持多选)",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if not file_paths:
            return

        imported_count = 0
        for path in file_paths:
            base_name = os.path.basename(path)
            content = ""
            try:
                if os.path.getsize(path) > MAX_IMPORT_BYTES:
                    messagebox.showwarning(
                        "文件过大",
                        f"文件 {base_name} 超过 {MAX_IMPORT_BYTES // (1024 * 1024)} MB，已跳过。",
                    )
                    continue
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        content = f.read()
                except UnicodeDecodeError:
                    with open(path, "r", encoding="gbk", errors="ignore") as f:
                        content = f.read()
            except Exception as e:
                messagebox.showerror("打开错误", f"读取文件 {base_name} 失败：\n{e}")
                continue

            first_empty = None
            for c in self.input_cells:
                if not c.get_content().strip():
                    first_empty = c
                    break

            if first_empty:
                first_empty.set_content(content, title=base_name)
            else:
                self.add_input_cell(title=base_name, content=content)

            imported_count += 1

        self.update_input_summary()
        self.status_lbl.config(text=f"状态：已成功将 {imported_count} 个文件导入为独立文本单元格！")

    def _export_results(self):
        """导出断行结果：支持按单元格批量导出独立文件或合并导出（读取实时编辑后的内容）"""
        if not self.output_boxes:
            messagebox.showinfo("提示", "右侧尚无断行结果可导出！")
            return

        live_results = self._get_live_results()
        count = len(live_results)
        if count == 1:
            save_path = filedialog.asksaveasfilename(
                title="保存断行结果",
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if save_path:
                try:
                    with open(save_path, "w", encoding="utf-8") as f:
                        f.write(live_results[0]["content"])
                    messagebox.showinfo("成功", f"断行结果已保存至：\n{save_path}")
                except Exception as e:
                    messagebox.showerror("保存失败", f"文件保存失败：\n{e}")
            return

        choice = messagebox.askyesnocancel(
            "多单元格导出选项",
            f"检测到当前有 {count} 个单元格的断行结果！\n\n"
            "【是 (Yes)】：批量保存为独立文件（保存至指定文件夹）\n"
            "【否 (No)】：合并保存为一个 TXT 文件\n"
            "【取消】：取消操作",
        )

        if choice is None:
            return

        if choice is True:
            folder = filedialog.askdirectory(title="选择批量保存的目标文件夹")
            if not folder:
                return

            saved_ok = 0
            failures = []
            used_names = set()
            for item in live_results:
                safe_title = _sanitize_filename(item["title"], f"单元格_{item['index']}")
                filename = f"{safe_title}_断行.txt"
                # 标题重名时追加序号，避免后一个文件静默覆盖前一个（Windows 文件名不区分大小写）
                if filename.lower() in used_names:
                    filename = f"{safe_title}_{item['index']}_断行.txt"
                used_names.add(filename.lower())
                out_path = os.path.join(folder, filename)
                try:
                    with open(out_path, "w", encoding="utf-8") as f:
                        f.write(item["content"])
                    saved_ok += 1
                except Exception as e:
                    failures.append(f"{filename}: {e}")

            summary = f"已成功将 {saved_ok}/{count} 个单元格结果分别导出至：\n{folder}"
            if failures:
                summary += "\n\n以下文件写入失败：\n" + "\n".join(failures[:5])
                if len(failures) > 5:
                    summary += f"\n……（另有 {len(failures) - 5} 个）"
                messagebox.showwarning("批量导出完成（部分失败）", summary)
            else:
                messagebox.showinfo("批量导出完成", summary)
        else:
            save_path = filedialog.asksaveasfilename(
                title="合并保存全部断行结果",
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if save_path:
                combined_text = "\n\n\n".join(
                    f"====================【{item['title']}】====================\n\n{item['content']}"
                    for item in live_results
                )
                try:
                    with open(save_path, "w", encoding="utf-8") as f:
                        f.write(combined_text)
                    messagebox.showinfo("成功", f"合并断行结果已保存至：\n{save_path}")
                except Exception as e:
                    messagebox.showerror("保存失败", f"文件保存失败：\n{e}")


def main():
    app = AutoWrapApp()
    app.mainloop()


if __name__ == "__main__":
    main()
