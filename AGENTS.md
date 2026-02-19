# 网络工具箱 (Network Toolbox)

## 项目介绍

一个用于网络设备配置自动化的桌面应用程序，支持批量采集设备LLDP信息、生成布线表、设备配置、拓扑图可视化。

- 技术栈：Python 3.x + tkinter + ttkbootstrap
- UI框架：ttkbootstrap (基于tkinter的现代UI库)

## 目录结构

```
网络工具箱/
├── app_main.py           # 主程序入口
├── 启动.bat             # 启动脚本（双击运行）
├── updater.py           # 自动更新检查模块
├── version.json         # 版本信息文件
├── requirements.txt     # Python依赖列表
├── modules/            # 功能模块目录（动态加载）
│   ├── __init__.py
│   ├── excel_generator.py    # Excel模板生成
│   ├── ssh_collector.py      # SSH批量采集
│   ├── lldp_parser.py       # LLDP解析
│   ├── config_generator.py   # 配置生成
│   ├── topo_pdf.py          # PDF拓扑图
│   ├── topo_html.py         # HTML拓扑图
│   └── tutorial.py          # 使用教程
├── config/             # 配置文件目录
│   ├── lldp_commands.txt    # SSH采集命令配置
│   └── device_list_template.xlsx  # 设备清单模板
├── templates/          # Jinja2模板目录
│   ├── 华为.txt
│   ├── 华三.txt
│   └── 锐捷.txt
├── lib/               # 前端库（HTML拓扑图用）
└── tutorial.md        # 教程文档
```

## 核心原则

### 最核心原则

**不能破坏"随便一个用户下载就能使用"的原则**

- 所有路径必须使用相对路径或动态获取，禁止硬编码绝对路径
- 所有功能对话框初始目录必须指向项目根目录

### 模块化原则

- UI框架(app_main.py)与功能模块(modules/)完全分离
- 功能模块从外部.py文件动态加载，不直接import
- 切换模块时卡顿解决方案：启动时预加载所有模块到缓存

### UI统一原则

- main_frame padding=25
- info_frame/input_frame 使用 Labelframe，padding=20
- 按钮宽度统一为18-20
- 字体统一使用 Microsoft YaHei UI

### 线程安全原则

- 所有UI更新必须使用 parent_frame.after(0, callback)
- 禁止在子线程中直接修改UI组件

## 注意事项

- 用户需要将Python添加到系统环境变量（安装时勾选选项）
- 首次使用需运行：pip install -r requirements.txt
- 双击启动.bat即可运行程序
- 自动更新功能依赖GitHub仓库

## 发布更新流程

1. 修改 version.json 中的版本号
2. 打包项目为zip文件
3. 在GitHub上创建Release，上传zip包
4. 执行：git add . && git commit -m "更新" && git push
