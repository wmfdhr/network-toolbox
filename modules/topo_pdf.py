# -*- coding: utf-8 -*-
"""
PDF拓扑模块 - 生成PDF网络拓扑图
"""

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from graphviz import Digraph
from datetime import datetime
import ttkbootstrap as ttk


class TopoGrapher:
    def __init__(self, excel_path, output_dir):
        self.excel_path = excel_path
        self.output_dir = os.path.join(output_dir, "graphs")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def generate(self):
        try:
            df = pd.read_excel(self.excel_path)

            dot = Digraph(comment="Network Topology", engine="dot")
            dot.encoding = "utf-8"

            dot.attr(rankdir="TB", splines="polyline", nodesep="1.0", ranksep="1.5")
            dot.attr(
                "node",
                shape="box",
                style="filled",
                fillcolor="lightblue",
                fontname="Microsoft YaHei",
                fixedsize="false",
                width="2.0",
            )
            dot.attr("edge", fontsize="9", fontname="Arial")

            all_devices = set()
            for _, row in df.iterrows():
                all_devices.add(str(row["本端设备"]))
                all_devices.add(str(row["对端设备"]))

            layers = {"WER": [], "WBS": [], "WDS": [], "WAS": [], "OTHER": []}

            for dev in all_devices:
                dev_upper = dev.upper()
                if "WER" in dev_upper:
                    layers["WER"].append(dev)
                elif "WBS" in dev_upper:
                    layers["WBS"].append(dev)
                elif "WDS" in dev_upper:
                    layers["WDS"].append(dev)
                elif "WAS" in dev_upper:
                    layers["WAS"].append(dev)
                else:
                    layers["OTHER"].append(dev)

            for layer_name in ["WER", "WBS", "WDS", "WAS", "OTHER"]:
                devices = layers[layer_name]
                if not devices:
                    continue
                with dot.subgraph() as s:
                    s.attr(rank="same")
                    for dev in devices:
                        color = "lightgray"
                        if layer_name == "WER":
                            color = "#ff9999"
                        elif layer_name == "WDS":
                            color = "#99ff99"
                        elif layer_name == "WBS":
                            color = "#ffff99"
                        elif layer_name == "WAS":
                            color = "#99ccff"
                        s.node(dev, fillcolor=color)

            for _, row in df.iterrows():
                l_dev = str(row["本端设备"])
                r_dev = str(row["对端设备"])

                l_phys = str(row.get("本端物理接口", row.get("本端接口", "")))
                r_phys = str(row.get("对端物理接口", row.get("对端接口", "")))
                l_logi = str(row.get("本端逻辑接口", ""))
                r_logi = str(row.get("对端逻辑接口", ""))

                def format_label(phys, logi):
                    phys = str(phys) if str(phys) != "nan" else ""
                    logi = str(logi) if str(logi) != "nan" else ""
                    if logi and logi != phys:
                        return f"{phys}\n({logi})"
                    return phys

                l_label = format_label(l_phys, l_logi)
                r_label = format_label(r_phys, r_logi)

                edge_label = f"{l_label}  <->  {r_label}"
                dot.edge(l_dev, r_dev, label=edge_label)

            timestamp = datetime.now().strftime("%m%d_%H%M")
            output_filename = f"topo_{timestamp}"
            output_path = os.path.join(self.output_dir, output_filename)

            try:
                dot.render(output_path, format="pdf", cleanup=True)
            except Exception as e:
                try:
                    if hasattr(e, "stderr") and e.stderr:
                        msg = e.stderr.decode("gbk")
                        return f"GRAPHVIZ_ERROR: {msg}"
                except:
                    pass
                raise e

            return f"{output_path}.pdf"

        except Exception as e:
            err_msg = str(e)
            if "codec" in err_msg:
                return "ENCODING_ERROR"
            raise e


class TopoPDFPanel:
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
            text="生成PDF网络拓扑图",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """根据Excel互联表生成PDF格式的网络拓扑图。

设备会根据名称自动分层显示：
- WER (城域网/POP) → 红色，最上层
- WBS (边界) → 黄色
- WDS (核心) → 绿色
- WAS (接入) → 蓝色，最下层

依赖：需要系统已安装 Graphviz 软件包。"""
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
            text="生成PDF拓扑图",
            command=self.run,
            bootstyle="success",
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
            messagebox.showwarning("提示", "请先选择互联表Excel文件")
            return

        try:
            grapher = TopoGrapher(self.path_var.get(), self.output_dir)
            result_path = grapher.generate()

            if result_path and "GRAPHVIZ_ERROR" in result_path:
                messagebox.showerror(
                    "Graphviz 渲染错误",
                    f"绘图引擎返回错误：\n{result_path.replace('GRAPHVIZ_ERROR: ', '')}",
                )
            elif result_path == "ENCODING_ERROR":
                messagebox.showerror(
                    "编码冲突",
                    "检测到系统字符集冲突。请尝试：\n1. 确保Excel中没有非法特殊字符\n2. 重新运行程序试试",
                )
            else:
                messagebox.showinfo(
                    "成功", f"拓扑图已生成！\n\n文件已存至: {result_path}"
                )
        except Exception as e:
            messagebox.showerror("生成失败", f"Graphviz 运行出错：\n{e}")
