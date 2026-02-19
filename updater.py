# -*- coding: utf-8 -*-
"""
自动更新检查模块 - 启动时检查GitHub是否有新版本
"""

import os
import sys
import json
import urllib.request
import urllib.error
import tempfile
import zipfile
import shutil
import tkinter as tk
from tkinter import messagebox
import threading
import subprocess


GITHUB_REPO = "wmfdhr/network-toolbox"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/version.json"
RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"


def get_local_version(base_dir):
    """获取本地版本号"""
    version_file = os.path.join(base_dir, "version.json")
    if os.path.exists(version_file):
        try:
            with open(version_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("version", "0.0.0")
        except:
            pass
    return "0.0.0"


def get_remote_version():
    """获取远程版本信息"""
    try:
        req = urllib.request.Request(
            VERSION_URL, headers={"User-Agent": "network-toolbox-updater"}
        )
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode("utf-8"))
            return data
    except:
        return None


def compare_versions(v1, v2):
    """比较版本号，返回 True 表示 v2 更新"""
    try:
        parts1 = [int(x) for x in v1.split(".")]
        parts2 = [int(x) for x in v2.split(".")]
        for i in range(max(len(parts1), len(parts2))):
            p1 = parts1[i] if i < len(parts1) else 0
            p2 = parts2[i] if i < len(parts2) else 0
            if p2 > p1:
                return True
            elif p2 < p1:
                return False
        return False
    except:
        return False


def check_update(base_dir):
    """检查是否有更新"""
    local_ver = get_local_version(base_dir)
    remote_data = get_remote_version()

    if not remote_data:
        return None, None, None

    remote_ver = remote_data.get("version", "0.0.0")
    changelog = remote_data.get("changelog", "")
    download_url = remote_data.get("download_url", "")

    if compare_versions(local_ver, remote_ver):
        return remote_ver, changelog, download_url

    return None, None, None


def show_update_dialog(new_version, changelog, download_url):
    """显示更新对话框"""
    root = tk.Tk()
    root.withdraw()

    message = f"发现新版本: {new_version}\n\n"
    message += f"更新内容:\n{changelog}\n\n"
    message += "是否前往下载页面？"

    result = messagebox.askyesno("发现新版本", message)

    if result:
        import webbrowser

        webbrowser.open(RELEASES_URL)

    root.destroy()
    return result


def check_and_prompt_update(base_dir):
    """检查更新并提示用户"""
    new_version, changelog, download_url = check_update(base_dir)

    if new_version:
        show_update_dialog(new_version, changelog, download_url)


def main():
    """主函数 - 启动时检查更新"""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    check_and_prompt_update(base_dir)


if __name__ == "__main__":
    main()
