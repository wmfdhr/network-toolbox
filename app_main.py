# -*- coding: utf-8 -*-
"""
网络工具箱 - 主框架
仅包含UI框架，功能模块从外部动态加载
启动时预加载所有模块，确保切换流畅
"""

import os
import sys
import threading
import importlib.util
import tkinter as tk
import ttkbootstrap as ttk


MODULES_TO_LOAD = [
    ("excel_generator", "ExcelGeneratorPanel", "生成模板"),
    ("ssh_collector", "SSHCollectorPanel", "SSH采集"),
    ("lldp_parser", "LLDPParserPanel", "LLDP解析"),
    ("config_generator", "ConfigGeneratorPanel", "生成配置"),
    ("topo_pdf", "TopoPDFPanel", "PDF拓扑"),
    ("topo_html", "TopoHTMLPanel", "HTML拓扑"),
    ("tutorial", "TutorialPanel", "使用教程"),
]


class LoadingOverlay:
    def __init__(self, parent):
        self.parent = parent
        self.overlay = tk.Toplevel(parent)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        self.overlay.configure(bg="#1a1a2e")

        self.width = 600
        self.height = 200

        self.parent.update_idletasks()
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()
        x = (sw - self.width) // 2
        y = (sh - self.height) // 2
        self.overlay.geometry(f"{self.width}x{self.height}+{x}+{y}")

        self.create_widgets()

    def create_widgets(self):
        container = tk.Frame(self.overlay, bg="#1a1a2e")
        container.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(
            container,
            text="网络工具箱 v1.0",
            bg="#1a1a2e",
            fg="#ffffff",
            font=("Consolas", 24, "bold"),
        )
        title_label.pack(pady=(40, 30))

        self.progress_label = tk.Label(
            container,
            text="  " + "░" * 50 + "  0%",
            bg="#1a1a2e",
            fg="#00d4ff",
            font=("Consolas", 16),
        )
        self.progress_label.pack(pady=(0, 15))

        self.status_label = tk.Label(
            container,
            text="正在初始化...",
            bg="#1a1a2e",
            fg="#888899",
            font=("Consolas", 13),
        )
        self.status_label.pack()

    def update_progress(self, percent, module_name):
        if percent < 0:
            self.status_label.configure(text=module_name, fg="#ff4444")
            self.overlay.update()
            return

        bar_width = 50
        filled = int(bar_width * percent / 100)
        bar = "█" * filled + "░" * (bar_width - filled)

        self.progress_label.configure(text=f"  {bar}  {percent}%")

        if percent < 100:
            self.status_label.configure(text=f"正在加载: {module_name}...")
        else:
            self.status_label.configure(text="加载完成，正在启动...")

        self.overlay.update()

    def destroy(self):
        self.overlay.destroy()


class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网络工具箱 v1.0")
        self.root.geometry("1280x800")
        self.root.minsize(1000, 700)

        self.base_dir = self.get_base_dir()
        self.current_module = None
        self.loaded_modules = {}

        self.colors = {
            "sidebar_bg": "#2c3e50",
            "sidebar_text": "#ecf0f1",
            "sidebar_hover": "#34495e",
            "sidebar_active": "#3498db",
            "content_bg": "#f5f6fa",
            "card_bg": "#ffffff",
            "primary": "#3498db",
            "text_dark": "#2c3e50",
            "text_gray": "#7f8c8d",
        }

        self.loading = LoadingOverlay(self.root)
        self.setup_style()

    def get_base_dir(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def setup_style(self):
        self.style = ttk.Style()

        self.style.configure("Sidebar.TFrame", background=self.colors["sidebar_bg"])
        self.style.configure(
            "SidebarTitle.TLabel",
            background=self.colors["sidebar_bg"],
            foreground=self.colors["sidebar_text"],
            font=("Microsoft YaHei UI", 16, "bold"),
        )
        self.style.configure(
            "Sidebar.TButton",
            background=self.colors["sidebar_bg"],
            foreground=self.colors["sidebar_text"],
            font=("Microsoft YaHei UI", 11),
            anchor="w",
            padding=(20, 12),
        )
        self.style.map(
            "Sidebar.TButton",
            background=[
                ("active", self.colors["sidebar_hover"]),
                ("!active", self.colors["sidebar_bg"]),
            ],
            foreground=[
                ("active", self.colors["sidebar_text"]),
                ("!active", self.colors["sidebar_text"]),
            ],
        )
        self.style.configure(
            "SidebarActive.TButton",
            background=self.colors["sidebar_active"],
            foreground="#ffffff",
            font=("Microsoft YaHei UI", 11, "bold"),
            anchor="w",
            padding=(20, 12),
        )
        self.style.configure("Content.TFrame", background=self.colors["content_bg"])
        self.style.configure("Card.TFrame", background=self.colors["card_bg"])
        self.style.configure(
            "Title.TLabel",
            background=self.colors["content_bg"],
            foreground=self.colors["text_dark"],
            font=("Microsoft YaHei UI", 28, "bold"),
        )
        self.style.configure(
            "Subtitle.TLabel",
            background=self.colors["content_bg"],
            foreground=self.colors["text_gray"],
            font=("Microsoft YaHei UI", 13),
        )
        self.style.configure(
            "CardTitle.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_dark"],
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        self.style.configure(
            "Step.TLabel",
            background=self.colors["content_bg"],
            foreground=self.colors["text_dark"],
            font=("Microsoft YaHei UI", 11),
        )

    def preload_modules(self, progress_callback):
        total = len(MODULES_TO_LOAD)

        for i, (module_name, class_name, display_name) in enumerate(MODULES_TO_LOAD):
            try:
                percent = int(((i + 1) / total) * 100)
                progress_callback(percent, display_name)

                module_path = os.path.join(
                    self.base_dir, "modules", f"{module_name}.py"
                )

                if not os.path.exists(module_path):
                    raise FileNotFoundError(f"模块文件不存在: {module_path}")

                spec = importlib.util.spec_from_file_location(module_name, module_path)
                if spec is None or spec.loader is None:
                    raise ImportError(f"无法创建模块规范: {module_path}")

                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                if not hasattr(module, class_name):
                    raise AttributeError(f"模块中未找到类: {class_name}")

                self.loaded_modules[module_name] = getattr(module, class_name)

            except Exception as e:
                error_msg = f"加载 {display_name} 失败:\n{str(e)}"
                progress_callback(-1, error_msg)
                return False

        progress_callback(100, "加载完成")
        return True

    def create_ui(self):
        self.loading.destroy()

        self.paned = tk.PanedWindow(
            self.root,
            orient=tk.HORIZONTAL,
            bg="#f5f6fa",
            sashwidth=6,
            sashrelief=tk.FLAT,
            opaqueresize=True,
        )
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.create_sidebar()
        self.create_content_area()

        self.paned.add(self.sidebar, width=260, minsize=200)
        self.paned.add(self.content_frame)

    def create_sidebar(self):
        self.sidebar = ttk.Frame(self.paned, style="Sidebar.TFrame")

        title_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        title_frame.pack(fill=tk.X, pady=(30, 25), padx=20)
        ttk.Label(title_frame, text="网络工具箱", style="SidebarTitle.TLabel").pack(
            anchor=tk.W
        )

        self.menu_items = [
            ("首页", "home", "欢迎使用网络工具箱"),
            ("生成模板", "excel_generator", "生成设备清单Excel模板"),
            ("SSH采集", "ssh_collector", "批量采集设备LLDP信息"),
            ("LLDP解析", "lldp_parser", "解析生成互联Excel表"),
            ("生成配置", "config_generator", "批量生成设备配置"),
            ("PDF拓扑", "topo_pdf", "生成PDF网络拓扑图"),
            ("HTML拓扑", "topo_html", "生成交互式HTML拓扑"),
        ]

        self.nav_buttons = {}

        menu_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        menu_frame.pack(fill=tk.X, padx=5)

        for text, module_name, desc in self.menu_items:
            btn = ttk.Button(
                menu_frame,
                text=text,
                style="Sidebar.TButton",
                command=lambda m=module_name, t=text: self.switch_module(m, t),
                width=15,
            )
            btn.pack(fill=tk.X, pady=2)
            self.nav_buttons[module_name] = btn

        ttk.Separator(self.sidebar, bootstyle="secondary").pack(
            fill=tk.X, padx=20, pady=20
        )

        tutorial_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        tutorial_frame.pack(fill=tk.X, padx=5)

        tutorial_btn = ttk.Button(
            tutorial_frame,
            text="使用教程",
            style="Sidebar.TButton",
            command=lambda: self.switch_module("tutorial", "使用教程"),
            width=15,
        )
        tutorial_btn.pack(fill=tk.X, pady=2)
        self.nav_buttons["tutorial"] = tutorial_btn

        version_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        version_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=15)
        ttk.Label(
            version_frame,
            text="v1.0",
            background=self.colors["sidebar_bg"],
            foreground=self.colors["text_gray"],
            font=("Microsoft YaHei UI", 9),
        ).pack()

    def create_content_area(self):
        self.content_frame = ttk.Frame(self.paned, style="Content.TFrame")
        self.show_home()

    def switch_module(self, module_name, button_text):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        self.current_module = None

        for name, btn in self.nav_buttons.items():
            btn.configure(style="Sidebar.TButton")
        if module_name in self.nav_buttons:
            self.nav_buttons[module_name].configure(style="SidebarActive.TButton")

        if module_name == "home":
            self.show_home()
        else:
            self.load_module(module_name)

    def load_module(self, module_name):
        if module_name not in self.loaded_modules:
            self.show_error(f"模块未加载: {module_name}")
            return

        try:
            panel_class = self.loaded_modules[module_name]
            self.current_module = panel_class(self.content_frame, self.base_dir)
        except Exception as e:
            self.show_error(f"初始化模块失败:\n{str(e)}")

    def show_home(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        for name, btn in self.nav_buttons.items():
            btn.configure(style="Sidebar.TButton")
        self.nav_buttons["home"].configure(style="SidebarActive.TButton")

        home_frame = ttk.Frame(self.content_frame, style="Content.TFrame", padding=50)
        home_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(home_frame, text="网络工具箱", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            home_frame, text="欢迎使用网络设备配置自动化工具", style="Subtitle.TLabel"
        ).pack(anchor=tk.W, pady=(5, 40))

        workflow_frame = ttk.Frame(home_frame, style="Content.TFrame")
        workflow_frame.pack(fill=tk.X, pady=(0, 30))

        ttk.Label(workflow_frame, text="快速开始", style="CardTitle.TLabel").pack(
            anchor=tk.W, pady=(0, 15)
        )

        steps = [
            ("1", "生成模板", "创建设备清单Excel"),
            ("2", "SSH采集", "批量采集设备LLDP信息"),
            ("3", "LLDP解析", "生成互联Excel表"),
            ("4", "生成配置", "批量生成设备配置"),
            ("5", "拓扑图", "可视化网络结构"),
        ]

        for num, title, desc in steps:
            step_frame = ttk.Frame(workflow_frame, style="Content.TFrame")
            step_frame.pack(fill=tk.X, pady=4)

            num_label = ttk.Label(
                step_frame,
                text=num,
                background=self.colors["primary"],
                foreground="#ffffff",
                font=("Microsoft YaHei UI", 10, "bold"),
                width=3,
                anchor="center",
            )
            num_label.pack(side=tk.LEFT, padx=(0, 15))

            ttk.Label(step_frame, text=f"{title}  —  {desc}", style="Step.TLabel").pack(
                side=tk.LEFT
            )

        ttk.Separator(home_frame, bootstyle="light").pack(fill=tk.X, pady=25)

        quick_frame = ttk.Frame(home_frame, style="Content.TFrame")
        quick_frame.pack(fill=tk.X, pady=10)

        ttk.Label(quick_frame, text="常用功能", style="CardTitle.TLabel").pack(
            anchor=tk.W, pady=(0, 15)
        )

        quick_actions = [
            ("生成模板", "excel_generator"),
            ("SSH采集", "ssh_collector"),
            ("LLDP解析", "lldp_parser"),
        ]

        cards_container = ttk.Frame(quick_frame, style="Content.TFrame")
        cards_container.pack(fill=tk.X)

        for text, module in quick_actions:
            card = ttk.Frame(cards_container, style="Card.TFrame", padding=20)
            card.pack(side=tk.LEFT, padx=(0, 15), pady=5)

            ttk.Label(card, text=text, style="CardTitle.TLabel").pack(
                anchor=tk.W, pady=(0, 10)
            )

            ttk.Button(
                card,
                text="立即使用",
                bootstyle="primary",
                command=lambda m=module, t=text: self.switch_module(m, t),
                width=12,
            ).pack()

        hint_frame = ttk.Frame(home_frame, style="Content.TFrame")
        hint_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=20)
        ttk.Label(
            hint_frame,
            text="详细教程请点击左侧「使用教程」",
            background=self.colors["content_bg"],
            foreground=self.colors["text_gray"],
            font=("Microsoft YaHei UI", 10),
        ).pack(anchor=tk.W)

    def show_error(self, message):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        error_frame = ttk.Frame(self.content_frame, style="Content.TFrame", padding=50)
        error_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            error_frame,
            text="加载失败",
            background=self.colors["content_bg"],
            foreground="#e74c3c",
            font=("Microsoft YaHei UI", 20, "bold"),
        ).pack(anchor=tk.W, pady=(0, 20))

        ttk.Label(
            error_frame,
            text=message,
            background=self.colors["content_bg"],
            foreground=self.colors["text_gray"],
            font=("Consolas", 10),
            justify=tk.LEFT,
        ).pack(anchor=tk.W)


def main():
    root = ttk.Window(themename="cosmo")
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"1280x800+{(sw - 1280) // 2}+{(sh - 800) // 2}")

    app = MainApp(root)

    def preload():
        success = app.preload_modules(app.loading.update_progress)
        root.after(0, lambda: finish(success))

    def finish(success):
        if success:
            app.create_ui()
        else:
            app.loading.destroy()
            root.destroy()
            sys.exit(1)

    threading.Thread(target=preload, daemon=True).start()
    root.mainloop()


if __name__ == "__main__":
    main()
