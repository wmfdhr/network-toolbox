# -*- coding: utf-8 -*-
"""
Excel模板生成模块 - 生成设备清单Excel模板
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk


class ExcelGeneratorPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.create_widgets()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="生成设备清单Excel模板",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """此模板用于填写设备信息，包含两个工作表：

1. 设备清单
   - 设备名称、厂商、管理IP、用户名、密码、启用状态
   - 用于SSH批量采集

2. 互联数据
   - 本端/对端设备、接口、IP地址等信息
   - 用于生成设备配置

填写完成后可用于后续的SSH采集和配置生成功能。"""
        ttk.Label(
            info_frame, text=info_text, font=("Microsoft YaHei UI", 10), justify=tk.LEFT
        ).pack(anchor=tk.W)

        input_frame = ttk.Labelframe(main_frame, text=" 保存设置 ", padding=20)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        self.path_var = tk.StringVar()

        path_row = ttk.Frame(input_frame)
        path_row.pack(fill=tk.X, pady=5)
        ttk.Label(path_row, text="保存位置:", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Entry(
            path_row, textvariable=self.path_var, font=("Microsoft YaHei UI", 10)
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(
            path_row,
            text="选择位置",
            command=self.select_path,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="生成设备清单模板",
            command=self.generate_device_list,
            bootstyle="success",
            width=18,
        ).pack(side=tk.LEFT, padx=8)
        ttk.Button(
            btn_frame,
            text="生成互联数据模板",
            command=self.generate_link_template,
            bootstyle="primary",
            width=18,
        ).pack(side=tk.LEFT, padx=8)

    def select_path(self):
        file_path = filedialog.asksaveasfilename(
            title="保存Excel模板",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="设备清单模板.xlsx",
            initialdir=self.base_dir,
        )
        if file_path:
            self.path_var.set(file_path)

    def generate_device_list(self):
        file_path = self.path_var.get()
        if not file_path:
            file_path = filedialog.asksaveasfilename(
                title="保存设备清单模板",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="设备清单模板.xlsx",
                initialdir=self.base_dir,
            )
            if not file_path:
                return

        df_devices = pd.DataFrame(
            columns=["设备名称", "厂商", "管理IP", "用户名", "密码", "启用"]
        )

        df_links = pd.DataFrame(
            columns=[
                "本端设备",
                "本端接口",
                "本端IPv4地址",
                "本端IPv6地址",
                "本端VPN实例",
                "对端VPN实例",
                "对端IPv6地址",
                "对端IPv4地址",
                "对端接口",
                "对端设备",
                "备注",
            ]
        )

        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df_devices.to_excel(writer, sheet_name="设备清单", index=False)
                df_links.to_excel(writer, sheet_name="连线信息", index=False)
            messagebox.showinfo("成功", f"模板已生成至：\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    def generate_link_template(self):
        file_path = filedialog.asksaveasfilename(
            title="保存互联数据模板",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="互联数据模板.xlsx",
            initialdir=self.base_dir,
        )

        if not file_path:
            return

        df_links = pd.DataFrame(
            columns=[
                "本端设备",
                "本端接口",
                "本端IPv4地址",
                "本端IPv6地址",
                "本端VPN实例",
                "对端VPN实例",
                "对端IPv6地址",
                "对端IPv4地址",
                "对端接口",
                "对端设备",
                "备注",
            ]
        )

        df_devices = pd.DataFrame(columns=["设备名称", "厂商", "管理IP", "Loopback0"])

        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df_links.to_excel(writer, sheet_name="连线信息", index=False)
                df_devices.to_excel(writer, sheet_name="设备信息", index=False)
            messagebox.showinfo("成功", f"模板已生成至：\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")
