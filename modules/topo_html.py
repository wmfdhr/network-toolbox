# -*- coding: utf-8 -*-
"""
HTML拓扑模块 - 生成交互式HTML网络拓扑图
"""

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pyvis.network import Network
from datetime import datetime
import ttkbootstrap as ttk


class InteractiveTopo:
    def __init__(self, excel_path, output_dir):
        self.excel_path = excel_path
        self.output_dir = os.path.join(output_dir, "graphs")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def generate(self):
        try:
            df = pd.read_excel(self.excel_path)
            net = Network(
                height="900px",
                width="100%",
                bgcolor="#222222",
                font_color="white",
                select_menu=True,
                filter_menu=True,
                cdn_resources="in_line",
            )

            net.set_options("""
            var options = {
              "nodes": { "font": { "size": 16, "face": "microsoft yahei" }, "shape": "box", "margin": 10, "borderWidth": 2 },
              "edges": { "color": { "inherit": true }, "font": { "size": 12, "align": "top" }, "smooth": { "enabled": true, "type": "curvedCW" } },
              "layout": { "hierarchical": { "enabled": true, "levelSeparation": 400, "nodeSpacing": 600, "treeSpacing": 600, "blockShifting": true, "edgeMinimization": false, "parentCentralization": true, "direction": "UD", "sortMethod": "directed" } },
              "physics": { "enabled": false },
              "interaction": { "hover": true, "navigationButtons": true }
            }
            """)

            def get_level(dev):
                name = str(dev).upper()
                if "WER" in name:
                    return 0
                if "WBS" in name:
                    return 1
                if "WDS" in name:
                    return 2
                if "WAS" in name:
                    return 3
                return 4

            def get_color(level):
                colors = {
                    0: "#ff6666",
                    1: "#ffff66",
                    2: "#66ff66",
                    3: "#66ccff",
                    4: "#cccccc",
                }
                return colors.get(level, "#cccccc")

            added_nodes = set()
            for _, row in df.iterrows():
                for dev_col in ["本端设备", "对端设备"]:
                    dev = str(row.get(dev_col, "")).strip()
                    if dev and dev != "nan" and dev not in added_nodes:
                        lvl = get_level(dev)
                        net.add_node(dev, label=dev, level=lvl, color=get_color(lvl))
                        added_nodes.add(dev)

            pair_counters = {}
            for index, row in df.iterrows():
                l_dev = str(row.get("本端设备", "")).strip()
                r_dev = str(row.get("对端设备", "")).strip()
                l_phys = str(row.get("本端物理接口", row.get("本端接口", ""))).strip()
                r_phys = str(row.get("对端物理接口", row.get("对端接口", ""))).strip()
                l_agg = str(row.get("本端聚合接口", row.get("本端聚合口", ""))).strip()
                r_agg = str(row.get("对端聚合接口", row.get("对端聚合口", ""))).strip()

                if not l_dev or l_dev == "nan" or not r_dev or r_dev == "nan":
                    continue

                l_phys = "" if l_phys == "nan" else l_phys
                r_phys = "" if r_phys == "nan" else r_phys
                l_agg_str = f"[{l_agg}]" if l_agg and l_agg != "nan" else ""
                r_agg_str = f"[{r_agg}]" if r_agg and r_agg != "nan" else ""

                pair_key = tuple(sorted([l_dev, r_dev]))
                link_index = pair_counters.get(pair_key, 0)
                pair_counters[pair_key] = link_index + 1

                edge_label = f"{l_phys}{l_agg_str} - {r_phys}{r_agg_str}"
                edge_obj = {
                    "from": l_dev,
                    "to": r_dev,
                    "id": f"link_{index}_{l_dev}_{r_dev}",
                    "label": edge_label,
                    "width": 2,
                    "color": "#888888",
                    "smooth": {
                        "enabled": True,
                        "type": "curvedCW",
                        "roundness": 0.2 + (link_index * 0.3),
                    },
                }
                net.edges.append(edge_obj)

            if not added_nodes:
                return None
            timestamp = datetime.now().strftime("%m%d_%H%M")
            output_file = os.path.join(
                self.output_dir, f"interactive_topo_{timestamp}.html"
            )
            html_content = net.generate_html()
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(html_content)
            return output_file
        except Exception as e:
            raise e


class TopoHTMLPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.output_dir = os.path.join(base_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        self.path_var = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="生成交互式HTML拓扑图",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """根据Excel互联表生成交互式HTML网络拓扑图。

特点：
- 可在浏览器中打开，支持缩放、拖拽
- 支持节点筛选、搜索功能
- 显示详细的接口连接信息
- 生成的HTML文件可直接分享给他人查看"""
        ttk.Label(
            info_frame, text=info_text, font=("Microsoft YaHei UI", 10), justify=tk.LEFT
        ).pack(anchor=tk.W)

        input_frame = ttk.Labelframe(main_frame, text=" 数据源设置 ", padding=20)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        path_row = ttk.Frame(input_frame)
        path_row.pack(fill=tk.X, pady=5)
        ttk.Label(path_row, text="互联表文件:", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Entry(path_row, textvariable=self.path_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)
        )
        ttk.Button(
            path_row,
            text="选择文件",
            command=self.select_file,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="生成HTML拓扑图",
            command=self.run,
            bootstyle="warning",
            width=18,
        ).pack()

    def select_file(self):
        f = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")], initialdir=self.base_dir
        )
        if f:
            self.path_var.set(f)

    def run(self):
        if not self.path_var.get():
            messagebox.showwarning("提示", "请选择Excel文件")
            return

        try:
            topo = InteractiveTopo(self.path_var.get(), self.output_dir)
            result = topo.generate()
            if result and os.path.exists(result):
                messagebox.showinfo("成功", f"交互式拓扑图已生成！\n\n文件: {result}")
            else:
                messagebox.showerror("错误", "生成失败，请检查Excel文件格式")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")
