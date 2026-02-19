# -*- coding: utf-8 -*-
"""
SSH采集模块 - 批量采集设备LLDP信息
"""

import pandas as pd
import paramiko
import socks
import os
import time
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
import threading
from concurrent.futures import ThreadPoolExecutor
import ttkbootstrap as ttk


class LLDPSSHCollector:
    def __init__(self, base_dir, log_callback):
        self.base_dir = base_dir
        self.output_dir = os.path.join(base_dir, "lldp_data")
        self.config_dir = os.path.join(base_dir, "config")
        self.log_callback = log_callback
        for d in [self.output_dir, self.config_dir]:
            if not os.path.exists(d):
                os.makedirs(d)
        self.commands = {}
        self.socks_host, self.socks_port = "localhost", 1999
        self.max_workers = 50
        self.load_commands()
        self.stats = {"success": 0, "failed": 0}
        self._lock = threading.Lock()

    def load_commands(self):
        cmd_file = os.path.join(self.config_dir, "lldp_commands.txt")
        if not os.path.exists(cmd_file):
            return
        try:
            with open(cmd_file, "r", encoding="utf-8") as f:
                v = None
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if line.startswith("[") and line.endswith("]"):
                        v = line[1:-1]
                        self.commands[v] = []
                    elif v:
                        self.commands[v].append(line)
        except Exception as e:
            self.log(f"加载命令文件失败: {e}")

    def log(self, message):
        self.log_callback(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")

    def collect_single_device(self, dev):
        name, ip, vendor = dev["设备名称"], dev["管理IP"], dev["厂商"].strip()
        self.log(f"准备连接: {name} ({ip})")

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            sock = socks.socksocket()
            sock.set_proxy(socks.SOCKS5, self.socks_host, self.socks_port)
            sock.connect((ip, 22))
            ssh.connect(
                hostname=ip,
                username=dev["用户名"],
                password=dev["密码"],
                sock=sock,
                timeout=30,
                look_for_keys=False,
                allow_agent=False,
            )

            shell = ssh.invoke_shell(width=200, height=1000)
            shell.settimeout(1)

            time.sleep(2)
            if shell.recv_ready():
                shell.recv(65535)

            disable_cmds = {
                "华为": "screen-length 0 temporary\n",
                "华三": "screen-length disable\n",
                "锐捷": "terminal length 0\n",
            }
            if vendor in disable_cmds:
                shell.send(disable_cmds[vendor].encode())
                time.sleep(1)
                if shell.recv_ready():
                    shell.recv(65535)

            output_content = [f"<{name}>"]
            for cmd in self.commands.get(vendor, []):
                shell.send((cmd + "\n").encode())
                buffer, start_time, last_recv = "", time.time(), time.time()
                prompt_regex = re.compile(r"[>#\]]\s*$")

                while True:
                    if time.time() - start_time > 300:
                        break
                    if shell.recv_ready():
                        chunk = shell.recv(65535).decode("utf-8", errors="ignore")
                        if chunk:
                            buffer += chunk
                            last_recv = time.time()
                    lines = buffer.strip().splitlines()
                    if lines and prompt_regex.search(lines[-1]):
                        time.sleep(0.5)
                        if not shell.recv_ready():
                            break
                    if time.time() - last_recv > 5 and len(buffer) > 0:
                        break
                    time.sleep(0.2)

                output_content.append(f"<{name}>{cmd}")
                output_content.append(buffer)

            with open(
                os.path.join(self.output_dir, f"{name}.txt"), "w", encoding="utf-8"
            ) as f:
                f.write("\n".join(output_content))

            with self._lock:
                self.stats["success"] += 1
            self.log(f"  {name} 采集成功。")
            return True
        except Exception as e:
            with self._lock:
                self.stats["failed"] += 1
            self.log(f"  {name} 失败: {e}")
            return False
        finally:
            ssh.close()

    def collect_batch(self, excel_path, concurrent_limit):
        self.stats = {"success": 0, "failed": 0}
        try:
            df = pd.read_excel(excel_path, sheet_name="设备清单", dtype=str)
            devices = df[df["启用"].str.strip() == "是"].to_dict("records")
        except Exception as e:
            self.log(f"错误: 无法读取设备清单 - {e}")
            return

        if not devices:
            self.log("警告: 设备清单中没有已启用的设备。")
            return

        self.log(
            f"开始并发采集，并发上限: {concurrent_limit}, 总设备数: {len(devices)}"
        )

        with ThreadPoolExecutor(max_workers=concurrent_limit) as executor:
            list(executor.map(self.collect_single_device, devices))

        self.log(f"\n批量采集任务结束：")
        self.log(f"成功: {self.stats['success']} 台")
        self.log(f"失败: {self.stats['failed']} 台")


class SSHCollectorPanel:
    def __init__(self, parent_frame, base_dir):
        self.parent_frame = parent_frame
        self.base_dir = base_dir
        self.path_var = tk.StringVar()
        self.concurrent_var = tk.IntVar(value=50)
        self.collector = LLDPSSHCollector(base_dir, self.append_log)
        self.create_widgets()

    def create_widgets(self):
        header_frame = ttk.Frame(self.parent_frame, bootstyle="dark")
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text="SSH批量采集设备LLDP信息",
            bootstyle="inverse-dark",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(pady=15)

        main_frame = ttk.Frame(self.parent_frame, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Labelframe(main_frame, text=" 使用说明 ", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        info_text = """批量通过SSH连接网络设备，采集LLDP邻居信息。

使用方法：
1. 准备包含设备信息的Excel清单（设备名称、IP、用户名、密码）
2. 选择设备清单文件
3. 设置并发数量（推荐30-100）
4. 点击"开始执行并发采集"

配置文件：config/lldp_commands.txt"""
        ttk.Label(
            info_frame, text=info_text, font=("Microsoft YaHei UI", 10), justify=tk.LEFT
        ).pack(anchor=tk.W)

        input_frame = ttk.Labelframe(main_frame, text=" 任务设置 ", padding=20)
        input_frame.pack(fill=tk.X, pady=(0, 15))

        file_row = ttk.Frame(input_frame)
        file_row.pack(fill=tk.X, pady=5)
        ttk.Label(file_row, text="设备清单:", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        path_entry = ttk.Entry(
            file_row, textvariable=self.path_var, font=("Microsoft YaHei UI", 10)
        )
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(
            file_row,
            text="选择文件",
            command=self.select_file,
            bootstyle="primary",
            width=12,
        ).pack(side=tk.LEFT)

        config_row = ttk.Frame(input_frame)
        config_row.pack(fill=tk.X, pady=5)
        ttk.Label(config_row, text="并发上限:", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        spinbox = ttk.Spinbox(
            config_row,
            from_=1,
            to=200,
            textvariable=self.concurrent_var,
            width=15,
            font=("Microsoft YaHei UI", 10),
        )
        spinbox.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(
            config_row,
            text="(推荐: 30-100)",
            foreground="gray",
            font=("Microsoft YaHei UI", 9),
        ).pack(side=tk.LEFT)

        log_frame = ttk.Labelframe(main_frame, text=" 实时日志 ", padding=15)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            font=("Consolas", 10),
            bg="#1e1e1e",
            fg="#00ff00",
            insertbackground="white",
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        self.run_btn = ttk.Button(
            btn_frame,
            text="开始执行并发采集",
            command=self.start_task,
            bootstyle="success",
            width=18,
        )
        self.run_btn.pack()

    def append_log(self, text):
        def update():
            self.log_area.insert(tk.END, text)
            self.log_area.see(tk.END)

        self.parent_frame.after(0, update)

    def select_file(self):
        f = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")], initialdir=self.base_dir
        )
        if f:
            self.path_var.set(f)

    def start_task(self):
        if not self.path_var.get():
            messagebox.showwarning("提示", "请先选择设备清单文件！")
            return
        self.run_btn.config(state=tk.DISABLED, text="并发任务执行中...")
        self.log_area.delete(1.0, tk.END)
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.collector.collect_batch(self.path_var.get(), self.concurrent_var.get())
        finally:
            self.parent_frame.after(0, self.finish_task)

    def finish_task(self):
        self.run_btn.config(state=tk.NORMAL, text="开始执行并发采集")
        messagebox.showinfo(
            "完成",
            f"并发采集结束！\n成功: {self.collector.stats['success']}\n失败: {self.collector.stats['failed']}",
        )
