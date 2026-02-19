# -*- coding: utf-8 -*-
"""
教程模块 - 显示使用教程（从外部Markdown文件读取）
"""

import os
import tkinter as tk
from tkinter import scrolledtext
import ttkbootstrap as ttk
import re


class TutorialPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.create_widgets()
        self.load_tutorial()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="使用教程",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """教程内容使用Markdown格式编写。

文件位置：项目根目录下的 tutorial.md 文件
编辑保存后，点击"刷新"按钮即可更新显示内容。"""
        ttk.Label(
            info_frame,
            text=info_text,
            font=("Microsoft YaHei UI", 10),
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        toolbar_frame = ttk.Frame(main_frame)
        toolbar_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            toolbar_frame,
            text="教程内容",
            font=("Microsoft YaHei UI", 10, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Button(
            toolbar_frame,
            text="刷新",
            command=self.load_tutorial,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.RIGHT)

        self.content_text = scrolledtext.ScrolledText(
            main_frame, font=("Microsoft YaHei UI", 11), wrap=tk.WORD, padx=15, pady=15
        )
        self.content_text.pack(fill=tk.BOTH, expand=True)

        self.setup_tags()

    def setup_tags(self):
        self.content_text.tag_configure(
            "h1",
            font=("Microsoft YaHei UI", 20, "bold"),
            spacing3=15,
            foreground="#2c3e50",
        )
        self.content_text.tag_configure(
            "h2",
            font=("Microsoft YaHei UI", 16, "bold"),
            spacing3=10,
            foreground="#34495e",
        )
        self.content_text.tag_configure(
            "h3",
            font=("Microsoft YaHei UI", 13, "bold"),
            spacing3=8,
            foreground="#7f8c8d",
        )
        self.content_text.tag_configure("bold", font=("Microsoft YaHei UI", 11, "bold"))
        self.content_text.tag_configure(
            "code", font=("Consolas", 10), background="#ecf0f1", foreground="#e74c3c"
        )
        self.content_text.tag_configure(
            "code_block",
            font=("Consolas", 10),
            background="#2c3e50",
            foreground="#ecf0f1",
            spacing1=5,
            spacing3=5,
        )
        self.content_text.tag_configure("list", lmargin1=20, lmargin2=30)
        self.content_text.tag_configure(
            "quote",
            lmargin1=20,
            lmargin2=30,
            foreground="#7f8c8d",
            font=("Microsoft YaHei UI", 11, "italic"),
        )
        self.content_text.tag_configure("hr", foreground="#bdc3c7")
        self.content_text.tag_configure("link", foreground="#3498db", underline=True)

    def load_tutorial(self):
        self.content_text.config(state=tk.NORMAL)
        self.content_text.delete(1.0, tk.END)

        md_path = os.path.join(self.base_dir, "tutorial.md")

        if not os.path.exists(md_path):
            self.show_default_tutorial()
            return

        try:
            with open(md_path, "r", encoding="utf-8") as f:
                content = f.read()
            self.render_markdown(content)
        except Exception as e:
            self.content_text.insert(tk.END, f"加载教程文件失败: {e}")

        self.content_text.config(state=tk.DISABLED)

    def show_default_tutorial(self):
        default_content = """# 网络工具箱 - 使用教程

## 快速开始

### 工作流程

```
生成模板 → SSH采集 → LLDP解析 → 生成配置 → 拓扑可视化
```

## 功能模块说明

### 1. 生成模板

生成设备清单Excel模板，用于填写设备信息（名称、IP、用户名、密码等）。

**使用方法：**
1. 点击"生成设备清单模板"
2. 选择保存位置
3. 填写设备信息

### 2. SSH采集

通过SSH批量连接网络设备，采集LLDP邻居信息。

**使用方法：**
1. 准备设备清单Excel
2. 选择文件，设置并发数
3. 点击开始采集

**配置文件：** `config/lldp_commands.txt`

### 3. LLDP解析

解析采集的配置文件，自动提取设备互联信息，生成Excel布线表。

**使用方法：**
1. 选择配置文件目录
2. 点击开始解析
3. 查看生成的布线表

### 4. 生成配置

根据Excel布线表和Jinja2模板，批量生成设备配置文件。

**使用方法：**
1. 选择互联数据Excel
2. 选择模板目录
3. 点击生成配置

**模板文件：** `templates/华三.txt`, `templates/华为.txt`, `templates/锐捷.txt`

### 5. PDF拓扑

根据Excel布线表生成PDF格式的网络拓扑图。

**依赖：** 需要安装Graphviz软件

### 6. HTML拓扑

生成可交互的HTML网络拓扑图，支持缩放、筛选、搜索。

---

## 自定义配置

- **采集命令：** 编辑 `config/lldp_commands.txt`
- **配置模板：** 编辑 `templates/*.txt`
- **教程内容：** 编辑 `tutorial.md`（Markdown格式）
"""
        self.render_markdown(default_content)
        self.content_text.config(state=tk.DISABLED)

    def render_markdown(self, content):
        lines = content.split("\n")
        in_code_block = False
        code_content = []

        for line in lines:
            if line.strip().startswith("```"):
                if in_code_block:
                    self.content_text.insert(
                        tk.END, "\n".join(code_content) + "\n", "code_block"
                    )
                    code_content = []
                    in_code_block = False
                else:
                    in_code_block = True
                continue

            if in_code_block:
                code_content.append(line)
                continue

            if line.startswith("# "):
                self.content_text.insert(tk.END, line[2:] + "\n", "h1")
            elif line.startswith("## "):
                self.content_text.insert(tk.END, line[3:] + "\n", "h2")
            elif line.startswith("### "):
                self.content_text.insert(tk.END, line[4:] + "\n", "h3")
            elif line.strip() == "---":
                self.content_text.insert(tk.END, "─" * 50 + "\n", "hr")
            elif line.startswith("> "):
                self.content_text.insert(tk.END, line[2:] + "\n", "quote")
            elif line.startswith("- ") or line.startswith("* "):
                self.content_text.insert(tk.END, "• " + line[2:] + "\n", "list")
            elif line.startswith("  - ") or line.startswith("  * "):
                self.content_text.insert(tk.END, "  ◦ " + line[4:] + "\n", "list")
            elif re.match(r"^\d+\.", line):
                self.content_text.insert(tk.END, line + "\n", "list")
            else:
                processed = self.process_inline_formatting(line)
                self.content_text.insert(tk.END, processed + "\n")

    def process_inline_formatting(self, line):
        return line
