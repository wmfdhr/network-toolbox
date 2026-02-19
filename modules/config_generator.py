# -*- coding: utf-8 -*-
"""
配置生成模块 - 根据Excel互联表生成设备配置
"""

import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
import ttkbootstrap as ttk


def strip_all_string_columns(df):
    return df.map(lambda x: str(x).strip() if pd.notnull(x) else "")


def parse_ip_mask(ip_str):
    s = str(ip_str).strip()
    if not s or s.lower() == "nan" or "/" not in s:
        return s if s.lower() != "nan" else "", ""
    parts = s.split("/")
    return parts[0], parts[1]


class ConfigGeneratorPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar(value=os.path.join(base_dir, "templates"))
        self.create_widgets()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="根据互联表生成设备配置",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """根据Excel互联表和Jinja2模板，批量生成设备配置文件。

使用方法：
1. 选择互联数据Excel文件（LLDP解析生成的布线表）
2. 选择Jinja2模板目录
3. 点击"开始生成配置"

模板文件：templates/华三.txt、华为.txt、锐捷.txt
输出文件：output/全部设备配置汇总_时间戳.txt"""
        ttk.Label(
            info_frame, text=info_text, font=("Microsoft YaHei UI", 10), justify=tk.LEFT
        ).pack(anchor=tk.W)

        input_frame = ttk.Labelframe(main_frame, text=" 数据源设置 ", padding=20)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        row1 = ttk.Frame(input_frame)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(
            row1, text="互联数据Excel:", width=15, font=("Microsoft YaHei UI", 10)
        ).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.excel_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)
        )
        ttk.Button(
            row1,
            text="选择文件",
            command=self.select_excel,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        row2 = ttk.Frame(input_frame)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(
            row2, text="Jinja2模板目录:", width=15, font=("Microsoft YaHei UI", 10)
        ).pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.template_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)
        )
        ttk.Button(
            row2,
            text="选择目录",
            command=self.select_template,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        log_frame = ttk.Labelframe(main_frame, text=" 生成日志 ", padding=15)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=12, font=("Consolas", 10), bg="#1e1e1e", fg="#00ff00"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)

        self.run_btn = ttk.Button(
            btn_frame,
            text="开始生成配置",
            command=self.start_generate,
            bootstyle="success",
            width=18,
        )
        self.run_btn.pack()

    def select_excel(self):
        f = filedialog.askopenfilename(
            title="选择互联数据Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialdir=self.base_dir,
        )
        if f:
            self.excel_path.set(f)

    def select_template(self):
        d = filedialog.askdirectory(
            title="选择Jinja2模板目录", initialdir=self.base_dir
        )
        if d:
            self.template_path.set(d)

    def log(self, msg):
        def update():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)

        self.parent_frame.after(0, update)

    def start_generate(self):
        excel = self.excel_path.get()
        template_dir = self.template_path.get()

        if not excel:
            messagebox.showwarning("提示", "请选择互联数据Excel文件")
            return
        if not template_dir:
            messagebox.showwarning("提示", "请选择Jinja2模板目录")
            return

        self.log_text.delete(1.0, tk.END)
        self.run_btn.config(state=tk.DISABLED)

        import threading

        threading.Thread(
            target=self.run_generate, args=(excel, template_dir), daemon=True
        ).start()

    def run_generate(self, excel_path, template_dir):
        try:
            output_dir = os.path.join(self.base_dir, "output")
            os.makedirs(output_dir, exist_ok=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"全部设备配置汇总_{timestamp}.txt")

            df_links = pd.read_excel(
                excel_path, sheet_name="连线信息", dtype=str
            ).dropna(how="all")
            df_links = strip_all_string_columns(df_links)

            try:
                df_devices = pd.read_excel(
                    excel_path, sheet_name="设备信息", dtype=str
                ).dropna(how="all")
                df_devices = strip_all_string_columns(df_devices)
                device_dict = df_devices.set_index("设备名称").to_dict("index")
            except:
                device_dict = {}

            processed_interfaces = []
            for index, row in df_links.iterrows():
                item = row.to_dict()
                if not item.get("本端设备") or not item.get("本端接口"):
                    continue

                item["本端IPv4"], item["本端IPv4掩码"] = parse_ip_mask(
                    item.get("本端IPv4地址", "")
                )
                item["本端IPv6"], item["本端IPv6掩码"] = parse_ip_mask(
                    item.get("本端IPv6地址", "")
                )
                item["对端IPv4_仅IP"], _ = parse_ip_mask(item.get("对端IPv4地址", ""))
                item["对端IPv6_仅IP"], _ = parse_ip_mask(item.get("对端IPv6地址", ""))
                item["对端管理IP"] = str(
                    device_dict.get(item.get("对端设备", ""), {}).get(
                        "管理IP", "未知IP"
                    )
                ).strip()

                processed_interfaces.append(item)

            df_processed = pd.DataFrame(processed_interfaces)
            env = Environment(
                loader=FileSystemLoader(template_dir),
                trim_blocks=True,
                lstrip_blocks=True,
            )

            count = 0
            all_configs = []

            self.log("=" * 50)
            self.log("开始生成配置...")
            self.log("=" * 50)

            for device_name, group in df_processed.groupby("本端设备"):
                vendor = str(device_dict.get(device_name, {}).get("厂商", "")).strip()
                if not vendor:
                    for col in ["华为", "华三", "锐捷", "H3C", "Huawei"]:
                        if (
                            col in device_name.upper()
                            or col.upper() in device_name.upper()
                        ):
                            vendor = col
                            break
                    if not vendor:
                        vendor = "华三"

                context = {
                    "device_info": device_dict.get(
                        device_name, {"设备名称": device_name, "厂商": vendor}
                    ),
                    "interfaces": group.to_dict("records"),
                }
                context["device_info"]["设备名称"] = device_name

                self.log(f"\n>>> 正在渲染设备: {device_name} (厂商: {vendor})")

                try:
                    template = env.get_template(vendor + ".txt")
                    all_configs.append(template.render(context).strip() + "\n\n")
                    count += 1
                except Exception as e:
                    self.log(f"  渲染出错: {e}")

            if all_configs:
                with open(output_file, "w", encoding="utf-8") as f:
                    f.writelines(all_configs)
                self.log(f"\n生成成功！共 {count} 台设备")
                self.log(f"文件保存至: {output_file}")
                self.parent_frame.after(
                    0,
                    lambda: messagebox.showinfo(
                        "完成",
                        f"生成成功！\n共生成 {count} 台设备配置\n\n保存至: {output_file}",
                    ),
                )
            else:
                self.log("\n未生成任何配置")
                self.parent_frame.after(
                    0, lambda: messagebox.showwarning("警告", "未生成任何配置文件")
                )

        except Exception as e:
            self.log(f"\n错误: {e}")
            self.parent_frame.after(
                0, lambda: messagebox.showerror("错误", f"生成失败：{e}")
            )
        finally:
            self.parent_frame.after(0, lambda: self.run_btn.config(state=tk.NORMAL))
