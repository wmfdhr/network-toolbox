# -*- coding: utf-8 -*-
"""
LLDP解析模块 - 解析配置文件生成Excel布线表
"""

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
import ttkbootstrap as ttk

from net_inspect import NetInspect


class LLDPTextParser:
    def __init__(self, input_folder, output_dir, log_callback=None):
        self.input_folder = os.path.abspath(input_folder)
        self.output_dir = output_dir
        self.log_callback = log_callback
        os.makedirs(self.output_dir, exist_ok=True)

    def log(self, msg):
        if self.log_callback:
            self.log_callback(msg)
        else:
            print(f"[DEBUG] {msg}")

    def parse_all(self):
        self.log("=" * 60)
        self.log(f"开始解析文件夹: {self.input_folder}")

        if not os.path.isdir(self.input_folder):
            return False, f"文件夹不存在: {self.input_folder}"

        txt_files = [f for f in os.listdir(self.input_folder) if f.endswith(".txt")]
        if not txt_files:
            return False, f"文件夹中未找到txt文件: {self.input_folder}"

        self.log(f"找到 {len(txt_files)} 个设备文件")

        net = NetInspect()
        net.set_plugins(input_plugin="console")

        self.log("正在调用 net_inspect 解析...")
        net.run(input_path=self.input_folder)

        self.log(f"解析完成，设备数量: {len(net.cluster.devices)}")

        all_devices = []
        all_links = []

        for device in net.cluster.devices:
            hostname = device.info.hostname
            vendor = device.info.vendor
            ip = device.info.ip
            model = device.info.model
            version = device.info.version

            self.log(f"\n处理设备: {hostname}")
            self.log(f"  厂商: {vendor}, 型号: {model}, IP: {ip}")

            device_info = {
                "hostname": hostname,
                "vendor": vendor,
                "ip": ip,
                "model": model,
                "version": version,
                "loopback0": "",
            }
            all_devices.append(device_info)

            intf_ip_map = self._extract_interface_ip(device)

            lldp_cmds = [
                "display lldp neighbor brief",
                "display lldp neighbor-information list",
                "show lldp neighbors",
                "display lldp neighbor",
            ]

            for lldp_cmd in lldp_cmds:
                try:
                    parse_result = device.parse_result(lldp_cmd)
                    if parse_result:
                        self.log(
                            f"  找到LLDP命令: {lldp_cmd}, {len(parse_result)} 条记录"
                        )
                        self.log(f"  第一条记录字段: {list(parse_result[0].keys())}")
                        self.log(f"  第一条记录内容: {parse_result[0]}")
                        links = self._extract_lldp_links(
                            hostname, parse_result, intf_ip_map
                        )
                        all_links.extend(links)
                        self.log(f"  提取链路: {len(links)} 条")
                        break
                except Exception as e:
                    self.log(f"  命令 '{lldp_cmd}' 解析失败: {e}")
                    continue

        if not all_devices:
            return False, "未能成功解析任何设备文件"

        all_links = self._deduplicate_links(all_links)

        success, excel_path = self._save_to_excel(all_devices, all_links)

        if success:
            self.log(f"\n解析完成: {len(all_devices)} 设备, {len(all_links)} 链路")
            return True, excel_path
        else:
            return False, excel_path

    def _extract_interface_ip(self, device):
        intf_ip_map = {}

        intf_cmds = [
            "display ip interface brief",
            "display interface brief",
            "show ip interface brief",
        ]

        for cmd_name in intf_cmds:
            try:
                parse_result = device.parse_result(cmd_name)
                if parse_result:
                    for row in parse_result:
                        intf_name = (
                            row.get("interface")
                            or row.get("port")
                            or row.get("name", "")
                        )
                        ip = row.get("ip_address") or row.get("ip") or ""
                        ipv6 = row.get("ipv6") or ""
                        vrf = row.get("vrf") or row.get("vpn_instance") or ""

                        if intf_name:
                            intf_ip_map[intf_name] = {
                                "ipv4": ip,
                                "ipv6": ipv6,
                                "vrf": vrf,
                            }
                    if intf_ip_map:
                        break
            except:
                continue

        return intf_ip_map

    def _extract_lldp_links(self, hostname, parse_result, intf_ip_map):
        links = []

        for row in parse_result:
            local_port = (
                row.get("local_interface")
                or row.get("local_port")
                or row.get("interface")
                or row.get("port")
                or ""
            ).strip()

            neighbor_dev = (
                row.get("neighbor")
                or row.get("neighbor_name")
                or row.get("remote_device")
                or row.get("remote_host")
                or row.get("system_name")
                or ""
            ).strip()

            neighbor_port = (
                row.get("neighbor_port_id")
                or row.get("neighbor_interface")
                or row.get("remote_port")
                or row.get("remote_interface")
                or row.get("port")
                or ""
            ).strip()

            mgmt_ip = (
                row.get("management_address")
                or row.get("remote_ip")
                or row.get("ip")
                or ""
            )

            if not local_port or not neighbor_dev:
                continue

            ip_info = intf_ip_map.get(local_port, {})

            link = {
                "本端设备": hostname,
                "本端接口": local_port,
                "本端IPv4地址": ip_info.get("ipv4", ""),
                "本端IPv6地址": ip_info.get("ipv6", ""),
                "本端VPN实例": ip_info.get("vrf", ""),
                "对端VPN实例": "",
                "对端IPv6地址": "",
                "对端IPv4地址": mgmt_ip,
                "对端接口": neighbor_port,
                "对端设备": neighbor_dev,
                "备注": "LLDP自动解析",
            }
            links.append(link)

        return links

    def _deduplicate_links(self, links):
        if not links:
            return links

        df = pd.DataFrame(links)

        def link_key(row):
            a = f"{row['本端设备']}_{row['本端接口']}"
            b = f"{row['对端设备']}_{row['对端接口']}"
            return "-".join(sorted([a, b]))

        df["_key"] = df.apply(link_key, axis=1)
        df = df.drop_duplicates(subset="_key").drop(columns="_key")

        return df.to_dict("records")

    def _save_to_excel(self, devices, links):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(self.output_dir, f"布线表_{timestamp}.xlsx")

        df_devices = pd.DataFrame(devices)
        df_devices = df_devices.rename(
            columns={
                "hostname": "设备名称",
                "vendor": "厂商",
                "ip": "管理IP",
                "model": "设备型号",
                "version": "软件版本",
                "loopback0": "Loopback0",
            }
        )

        df_links = pd.DataFrame(links)

        if df_links.empty:
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
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                df_links.to_excel(writer, sheet_name="连线信息", index=False)
                df_devices.to_excel(writer, sheet_name="设备信息", index=False)
            return True, excel_path
        except Exception as e:
            return False, f"保存Excel失败: {e}"


class LLDPParserPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.output_dir = os.path.join(base_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        self.create_widgets()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="LLDP智能解析工具",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """解析SSH采集的配置文件，自动提取设备互联信息，生成Excel布线表。

使用方法：
1. 选择存放配置文件的文件夹（SSH采集的输出目录）
2. 点击"开始解析"
3. 查看生成的布线表

输出文件：output/布线表_时间戳.xlsx"""
        ttk.Label(
            info_frame, text=info_text, font=("Microsoft YaHei UI", 10), justify=tk.LEFT
        ).pack(anchor=tk.W)

        input_frame = ttk.Labelframe(main_frame, text=" 数据源设置 ", padding=20)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        self.path_var = tk.StringVar(value=self.base_dir)

        path_row = ttk.Frame(input_frame)
        path_row.pack(fill=tk.X, pady=5)
        ttk.Label(path_row, text="配置文件夹:", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Entry(
            path_row, textvariable=self.path_var, font=("Microsoft YaHei UI", 10)
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(
            path_row,
            text="选择文件夹",
            command=self.select_dir,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        log_frame = ttk.Labelframe(main_frame, text=" 解析日志 ", padding=15)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=15, font=("Consolas", 10), bg="#1e1e1e", fg="#00ff00"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)

        ttk.Button(
            btn_frame,
            text="开始解析",
            command=self.start_parse,
            bootstyle="success",
            width=18,
        ).pack()

    def log(self, msg):
        def update():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.parent_frame.after(0, update)

    def select_dir(self):
        d = filedialog.askdirectory(
            title="请选择存放配置文件的文件夹", initialdir=self.base_dir
        )
        if d:
            self.path_var.set(d)

    def start_parse(self):
        path = self.path_var.get()
        if not path:
            messagebox.showwarning("提示", "请先选择配置文件文件夹")
            return

        self.log("开始解析...")
        parser = LLDPTextParser(path, self.output_dir, log_callback=self.log)

        def run_parse():
            success, info = parser.parse_all()
            self.parent_frame.after(0, lambda: self.on_parse_complete(success, info))

        import threading

        threading.Thread(target=run_parse, daemon=True).start()

    def on_parse_complete(self, success, info):
        if success:
            messagebox.showinfo("成功", f"布线表生成完成！\n{info}")
        else:
            messagebox.showerror("失败", info)
