#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Power Automate Desktop 剪贴板输入工具
通过复制和粘贴方式向 PAD 中输入命令
"""

import time
import os
import pyautogui
import pyperclip
import requests
import json
import subprocess
import threading
from pywinauto import Application, Desktop
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

class PADClipboardInputTool:
    def __init__(self):
        # 创建 UI
        self.create_ui()
        # Gemini API 密钥
        self.gemini_api_key = "AIzaSyDu1uAkjzDs8TNJXe2eUkjjHqJRTwWaDI8"
        # PAD 路径
        self.pad_path = r"C:\Program Files (x86)\Power Automate Desktop\PAD.exe"
    
    def create_ui(self):
        """创建用户界面"""
        self.root = tk.Tk()
        self.root.title("PAD 剪贴板输入工具")
        self.root.geometry("800x600")
        
        # 设置样式
        style = ttk.Style()
        style.configure("TButton", padding=5, font=("Arial", 10))
        style.configure("TLabel", font=("Arial", 10))
        style.configure("Title.TLabel", font=("Arial", 14, "bold"))
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Power Automate Desktop 剪贴板输入工具", style="Title.TLabel")
        title_label.pack(pady=10)
        
        # 创建笔记本控件(选项卡)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 标准模式选项卡
        standard_tab = ttk.Frame(notebook)
        notebook.add(standard_tab, text="标准模式")
        
        # AI模式选项卡
        ai_tab = ttk.Frame(notebook)
        notebook.add(ai_tab, text="AI转换模式")
        
        # =========== 标准模式界面 ===========
        # 左右分栏
        content_frame = ttk.Frame(standard_tab)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 左侧 - 命令输入区域
        left_frame = ttk.LabelFrame(content_frame, text="命令输入", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 命令输入框 - 使用多行文本框替代单行输入框
        ttk.Label(left_frame, text="请输入要粘贴到 PAD 的命令:").pack(anchor=tk.W, pady=(0, 5))
        
        # 创建滚动文本框
        self.command_text = scrolledtext.ScrolledText(left_frame, height=10, width=40, wrap=tk.WORD, font=("Consolas", 11))
        self.command_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.command_text.insert(tk.END, "WAIT 3")
        
        # 预设命令下拉框
        preset_frame = ttk.Frame(left_frame)
        preset_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(preset_frame, text="常用命令预设:").pack(side=tk.LEFT)
        
        self.preset_commands = [
            "WAIT 3", 
            "TEXT ASSIGN Value1 =\"Hello World\"", 
            "MOUSE CLICK LeftClick, 1", 
            "KEYBOARD SEND ~", 
            "IF Value1 = \"Hello World\" THEN",
            "END",
            "WebAutomation.LaunchEdge.LaunchEdge Url: $'''https://www.bilibili.com/''' WindowState: WebAutomation.BrowserWindowState.Normal ClearCache: False ClearCookies: False WaitForPageToLoadTimeout: 60 Timeout: 60 TargetDesktop: $'''{{\"DisplayName\":\"本地计算机\",\"Route\":{{\"ServerType\":\"Local\",\"ServerAddress\":\"\"}},\"DesktopType\":\"local\"}}''' BrowserInstance=> Browser"
        ]
        
        self.preset_var = tk.StringVar()
        preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_var, values=self.preset_commands, width=30)
        preset_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_selected)
        
        # 插入预设按钮
        insert_preset_button = ttk.Button(preset_frame, text="插入", command=self.insert_preset)
        insert_preset_button.pack(side=tk.LEFT, padx=5)
        
        # 清空按钮
        clear_text_button = ttk.Button(preset_frame, text="清空", command=self.clear_command_text)
        clear_text_button.pack(side=tk.LEFT, padx=5)
        
        # 右侧 - 操作区域
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
        
        # 操作区域
        operation_frame = ttk.LabelFrame(right_frame, text="操作", padding="10")
        operation_frame.pack(fill=tk.X, pady=5)
        
        # 选择点击方式
        click_frame = ttk.Frame(operation_frame)
        click_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(click_frame, text="点击方式:").pack(anchor=tk.W)
        
        self.click_method_var = tk.StringVar(value="current")
        ttk.Radiobutton(click_frame, text="自动查找编辑区", variable=self.click_method_var, value="auto").pack(anchor=tk.W, padx=15)
        ttk.Radiobutton(click_frame, text="使用鼠标当前位置", variable=self.click_method_var, value="current").pack(anchor=tk.W, padx=15)
        
        # 操作延迟
        delay_frame = ttk.Frame(operation_frame)
        delay_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(delay_frame, text="点击后延迟(秒):").pack(anchor=tk.W)
        
        self.delay_var = tk.DoubleVar(value=1.0)
        delay_spin = ttk.Spinbox(delay_frame, from_=0.1, to=5.0, increment=0.1, textvariable=self.delay_var, width=5)
        delay_spin.pack(anchor=tk.W, padx=15, pady=5)
        
        # 鼠标位置
        pos_frame = ttk.LabelFrame(right_frame, text="鼠标位置", padding="10")
        pos_frame.pack(fill=tk.X, pady=5)
        
        # 记录的位置
        self.mouse_pos_var = tk.StringVar(value="未记录")
        ttk.Label(pos_frame, textvariable=self.mouse_pos_var).pack(anchor=tk.W, pady=5)
        
        # 记录鼠标位置按钮
        record_pos_button = ttk.Button(pos_frame, text="记录鼠标位置(5秒)", command=self.record_mouse_position)
        record_pos_button.pack(fill=tk.X, pady=5)
        
        # 启动PAD按钮
        pad_button = ttk.Button(right_frame, text="启动 Power Automate Desktop", command=self.launch_pad)
        pad_button.pack(fill=tk.X, pady=5)
        
        # =========== AI转换模式界面 ===========
        # AI转换框架
        ai_content_frame = ttk.Frame(ai_tab)
        ai_content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 需求输入区域
        ai_input_frame = ttk.LabelFrame(ai_content_frame, text="需求输入", padding="10")
        ai_input_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        ttk.Label(ai_input_frame, text="请用自然语言描述您的自动化需求:").pack(anchor=tk.W, pady=(0, 5))
        
        # API密钥输入
        api_frame = ttk.Frame(ai_input_frame)
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="Gemini API密钥:").pack(side=tk.LEFT)
        
        self.api_key_var = tk.StringVar(value="AIzaSyDu1uAkjzDs8TNJXe2eUkjjHqJRTwWaDI8")
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=40, show="*")
        api_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 需求输入文本框
        self.requirement_text = scrolledtext.ScrolledText(ai_input_frame, height=10, width=40, wrap=tk.WORD, font=("Consolas", 11))
        self.requirement_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.requirement_text.insert(tk.END, "请输入您的需求，例如：打开Edge浏览器，访问百度网站并搜索Power Automate")
        
        # 转换按钮和执行按钮框架
        ai_buttons_frame = ttk.Frame(ai_input_frame)
        ai_buttons_frame.pack(fill=tk.X, pady=5)
        
        # 转换按钮
        convert_button = ttk.Button(ai_buttons_frame, text="转换为PAD命令", command=self.convert_to_pad_commands)
        convert_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 一键执行按钮
        execute_ai_button = ttk.Button(ai_buttons_frame, text="一键执行（转换+启动+执行）", command=self.one_key_execute)
        execute_ai_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 转换结果区域
        result_frame = ttk.LabelFrame(ai_content_frame, text="转换结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 转换结果文本框
        self.ai_result_text = scrolledtext.ScrolledText(result_frame, height=10, width=40, wrap=tk.WORD, font=("Consolas", 11))
        self.ai_result_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # AI模式操作说明
        ai_tips_frame = ttk.LabelFrame(ai_content_frame, text="使用说明", padding="10")
        ai_tips_frame.pack(fill=tk.X, pady=5)
        
        ai_tips = """
1. API密钥已默认设置
2. 用自然语言描述您想要自动化的任务
3. 点击"转换为PAD命令"将需求转换为PAD命令
4. 检查转换结果，并根据需要修改
5. 点击"一键执行"启动PAD并输入命令
        """
        ttk.Label(ai_tips_frame, text=ai_tips, justify=tk.LEFT).pack(anchor=tk.W)
        
        # =========== 底部共享区域 ===========
        # 执行按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 执行按钮 - 改为大按钮
        execute_button = ttk.Button(button_frame, text="执行复制粘贴", command=self.execute_clipboard_input)
        execute_button.pack(fill=tk.X, pady=5, ipady=5)  # 使按钮更高
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(fill=tk.X, pady=(5, 0))
        
        # 创建日志区域
        log_frame = ttk.LabelFrame(main_frame, text="执行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=5, width=50, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 鼠标位置
        self.recorded_position = None

        # 默认选择AI转换模式选项卡
        notebook.select(ai_tab)
    
    def log(self, message):
        """向日志区域添加消息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        print(message)
    
    def on_preset_selected(self, event):
        """当选择预设命令时更新命令变量"""
        selected_preset = self.preset_var.get()
        if selected_preset:
            self.preset_var.set("")  # 清空下拉菜单选择
    
    def insert_preset(self):
        """将选中的预设命令插入到命令输入框"""
        selected_preset = self.preset_var.get()
        if selected_preset:
            # 在光标位置插入文本
            self.command_text.insert(tk.INSERT, selected_preset)
            self.preset_var.set("")  # 清空下拉菜单选择
    
    def clear_command_text(self):
        """清空命令输入框"""
        self.command_text.delete(1.0, tk.END)
    
    def record_mouse_position(self):
        """记录鼠标当前位置（延迟5秒）"""
        self.status_var.set("准备记录鼠标位置...5秒后记录")
        self.log("准备记录鼠标位置，请将鼠标移动到 PAD 编辑区...")
        
        # 倒计时
        for i in range(5, 0, -1):
            self.status_var.set(f"准备记录鼠标位置...{i}秒后记录")
            self.root.update()
            time.sleep(1)
        
        # 记录位置
        x, y = pyautogui.position()
        self.recorded_position = (x, y)
        self.mouse_pos_var.set(f"记录的位置: X={x}, Y={y}")
        self.status_var.set(f"已记录鼠标位置: X={x}, Y={y}")
        self.log(f"已记录鼠标位置: X={x}, Y={y}")
    
    def launch_pad(self):
        """启动 Power Automate Desktop"""
        try:
            if os.path.exists(self.pad_path):
                self.log("正在启动 Power Automate Desktop...")
                subprocess.Popen(self.pad_path)
                self.log("Power Automate Desktop 已启动")
                self.status_var.set("已启动 PAD")
            else:
                error_msg = "未找到 Power Automate Desktop 程序，请检查路径"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
        except Exception as e:
            error_msg = f"启动 PAD 时出错: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def convert_to_pad_commands(self):
        """将自然语言需求转换为PAD命令"""
        # 获取API密钥
        api_key = self.api_key_var.get()
        if not api_key:
            messagebox.showwarning("警告", "请输入Gemini API密钥")
            return
        
        # 保存API密钥
        self.gemini_api_key = api_key
        
        # 获取需求文本
        requirement = self.requirement_text.get(1.0, tk.END).strip()
        if not requirement or requirement == "请输入您的需求，例如：打开Edge浏览器，访问百度网站并搜索Power Automate":
            messagebox.showwarning("警告", "请输入自动化需求")
            return
        
        self.log(f"正在将需求转换为PAD命令...")
        self.status_var.set("正在转换...")
        
        # 启动转换线程
        threading.Thread(target=self.do_convert_to_pad_commands, args=(requirement,), daemon=True).start()
    
    def do_convert_to_pad_commands(self, requirement):
        """在后台线程中调用Gemini API将需求转换为PAD命令"""
        try:
            # 准备API请求
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.gemini_api_key}"
            headers = {
                "Content-Type": "application/json"
            }
            
            # 构建提示
            prompt = f"""你是一个Power Automate Desktop (PAD) 专家，精通将用户的自然语言需求转换为准确可执行的PAD命令。

请将以下自然语言需求转换为PAD的Robin语言命令。你必须使用正确的Robin语言语法，这是PAD特有的语法。请只生成你确定正确的命令，不要猜测或添加可能不存在的命令。

示例1 - 等待3秒:
WAIT 3

示例2 - 启动Edge浏览器访问百度:
WebAutomation.LaunchEdge.LaunchEdge Url: $'''https://www.baidu.com/''' WindowState: WebAutomation.BrowserWindowState.Normal ClearCache: False ClearCookies: False WaitForPageToLoadTimeout: 60 Timeout: 60 TargetDesktop: $'''{{\"DisplayName\":\"本地计算机\",\"Route\":{{\"ServerType\":\"Local\",\"ServerAddress\":\"\"}},\"DesktopType\":\"local\"}}''' BrowserInstance=> Browser

极其重要的注意事项:
1. 只生成有把握正确的命令，不要猜测不确定的命令
2. 不要使用可能不存在的模块或操作，如不确定某个操作是否存在，请只使用基本命令
3. 对于复杂需求，应分解为多个基本步骤而不是使用猜测的高级命令
4. 确保所有用到的模块都是PAD中实际存在的
5. 如果一个需求只需一条命令就能完成，请只生成一条命令

用户需求: {requirement}

请仅输出Robin语言命令，不要包含任何解释、注释或markdown格式。只返回可以直接复制粘贴到PAD的命令，确保每条命令都能在PAD中正确执行。"""
            
            data = {
                "contents": [{
                    "parts": [{"text": prompt}]
                }]
            }
            
            # 发送请求
            self.root.after(0, self.log, "正在调用Gemini API...")
            response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            
            # 解析响应
            result = response.json()
            # Gemini API响应格式与DeepSeek不同，需要适配
            pad_commands = result["candidates"][0]["content"]["parts"][0]["text"].strip()
            
            # 更新UI（在主线程中）
            self.root.after(0, self.update_conversion_result, pad_commands)
            
        except Exception as e:
            error_msg = f"转换过程中出错: {str(e)}"
            self.root.after(0, self.log, error_msg)
            self.root.after(0, self.status_var.set, "转换失败")
            self.root.after(0, messagebox.showerror, "错误", error_msg)
    
    def update_conversion_result(self, pad_commands):
        """更新转换结果到UI"""
        # 处理命令，过滤掉可能不正确的命令
        cleaned_commands = self.clean_pad_commands(pad_commands)
        
        self.ai_result_text.delete(1.0, tk.END)
        self.ai_result_text.insert(tk.END, cleaned_commands)
        
        self.log("需求已转换为PAD命令")
        self.status_var.set("转换完成")
        
        # 同时也更新到标准模式的命令输入框
        self.command_text.delete(1.0, tk.END)
        self.command_text.insert(tk.END, cleaned_commands)
    
    def clean_pad_commands(self, commands):
        """清理和验证PAD命令，移除可能不正确的命令"""
        # 分行处理命令
        command_lines = commands.strip().split('\n')
        valid_commands = []
        
        # 已知有效的PAD命令前缀和模块
        valid_prefixes = [
            # 基本控制命令
            "WAIT", "TEXT ASSIGN", "MOUSE CLICK", "KEYBOARD", "IF", "END", "ELSE", 
            "WHILE", "FOR EACH", "ON ERROR", "SUB", "FUNCTION", "EXIT", "RETURN",
            
            # 变量和数据处理
            "VARIABLE", "SET ", "INCREMENT", "DECREMENT", "DATA", "DATETIME", "DISPLAY",
            
            # 文件操作
            "FILES.", "FOLDER.", "PATH.", "EXCEL.", "WORD.", "ZIP.", "XML.", "CSV.",
            
            # 浏览器自动化
            "WebAutomation.", "WebAutomation.LaunchEdge", "WebAutomation.LaunchChrome", 
            "WebAutomation.BrowserInstance", "WebAutomation.Navigate", "WebAutomation.Close", 
            "WebAutomation.GetPageTitle", "WebAutomation.ExecuteJavascript", "WebAutomation.GetUrl",
            
            # 系统操作
            "SYSTEM.", "SERVICE.", "EXCHANGE.", "PRINTER.", "PROCESS.",
            
            # 网络操作
            "MAIL.", "FTP.", "HTTP.", "PING", "REST.",
            
            # UI自动化
            "UIAutomation.", "OCR."
        ]
        
        # 检查每行命令
        first_valid_command_found = False
        
        for line in command_lines:
            line = line.strip()
            if not line:
                continue
                
            # 检查命令是否有效
            is_valid = False
            for prefix in valid_prefixes:
                if line.startswith(prefix):
                    is_valid = True
                    first_valid_command_found = True
                    break
            
            # 如果是第一条有效命令，添加到结果中
            if is_valid:
                valid_commands.append(line)
            # 如果已经找到了有效命令，但当前命令不是已知有效命令，则停止添加
            elif first_valid_command_found:
                self.log(f"忽略可能不正确的命令: {line}")
                # 如果我们找到了一个有效命令后面跟着无效命令，则停止处理
                break
        
        # 如果没有有效命令，返回原始命令
        if not valid_commands:
            self.log("警告: 未找到有效的PAD命令，返回原始内容")
            return commands
            
        return "\n".join(valid_commands)
    
    def one_key_execute(self):
        """一键执行：转换+启动+执行"""
        # 先转换
        api_key = self.api_key_var.get()
        requirement = self.requirement_text.get(1.0, tk.END).strip()
        
        if not api_key:
            messagebox.showwarning("警告", "请输入Gemini API密钥")
            return
            
        if not requirement or requirement == "请输入您的需求，例如：打开Edge浏览器，访问百度网站并搜索Power Automate":
            messagebox.showwarning("警告", "请输入自动化需求")
            return
        
        self.log("开始一键执行流程...")
        self.status_var.set("正在执行一键流程...")
        
        # 启动转换线程
        threading.Thread(target=self.do_one_key_execute, args=(requirement,), daemon=True).start()
    
    def do_one_key_execute(self, requirement):
        """在后台线程中执行一键流程"""
        try:
            # 1. 转换
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.gemini_api_key}"
            headers = {
                "Content-Type": "application/json"
            }
            
            prompt = f"""你是一个Power Automate Desktop (PAD) 专家，精通将用户的自然语言需求转换为准确可执行的PAD命令。

请将以下自然语言需求转换为PAD的Robin语言命令。你必须使用正确的Robin语言语法，这是PAD特有的语法。请只生成你确定正确的命令，不要猜测或添加可能不存在的命令。

示例1 - 等待3秒:
WAIT 3

示例2 - 启动Edge浏览器访问百度:
WebAutomation.LaunchEdge.LaunchEdge Url: $'''https://www.baidu.com/''' WindowState: WebAutomation.BrowserWindowState.Normal ClearCache: False ClearCookies: False WaitForPageToLoadTimeout: 60 Timeout: 60 TargetDesktop: $'''{{\"DisplayName\":\"本地计算机\",\"Route\":{{\"ServerType\":\"Local\",\"ServerAddress\":\"\"}},\"DesktopType\":\"local\"}}''' BrowserInstance=> Browser

极其重要的注意事项:
1. 只生成有把握正确的命令，不要猜测不确定的命令
2. 不要使用可能不存在的模块或操作，如不确定某个操作是否存在，请只使用基本命令
3. 对于复杂需求，应分解为多个基本步骤而不是使用猜测的高级命令
4. 确保所有用到的模块都是PAD中实际存在的
5. 如果一个需求只需一条命令就能完成，请只生成一条命令

用户需求: {requirement}

请仅输出Robin语言命令，不要包含任何解释、注释或markdown格式。只返回可以直接复制粘贴到PAD的命令，确保每条命令都能在PAD中正确执行。"""
            
            data = {
                "contents": [{
                    "parts": [{"text": prompt}]
                }]
            }
            
            self.root.after(0, self.log, "正在转换需求为PAD命令...")
            
            response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            
            result = response.json()
            pad_commands = result["candidates"][0]["content"]["parts"][0]["text"].strip()
            
            # 清理和验证命令
            pad_commands = self.clean_pad_commands(pad_commands)
            
            # 更新UI
            self.root.after(0, self.update_conversion_result, pad_commands)
            
            # 2. 启动PAD（如果未启动）
            self.root.after(0, self.log, "正在检查PAD是否已启动...")
            
            # 查找PAD窗口
            desktop = Desktop(backend="uia")
            pad_window = None
            try:
                pad_window = desktop.window(title_re=".*Power Automate.*")
                if pad_window.exists():
                    self.root.after(0, self.log, "找到PAD窗口，无需重新启动")
                else:
                    # 启动PAD
                    self.root.after(0, self.launch_pad)
                    # 等待PAD启动
                    time.sleep(5)
            except Exception:
                # 启动PAD
                self.root.after(0, self.launch_pad)
                # 等待PAD启动
                time.sleep(5)
            
            # 3. 执行粘贴
            self.root.after(0, self.log, "准备执行粘贴...")
            time.sleep(2)  # 给PAD一些启动时间
            
            # 执行粘贴
            self.root.after(0, self.execute_clipboard_input)
            
        except Exception as e:
            error_msg = f"一键执行过程中出错: {str(e)}"
            self.root.after(0, self.log, error_msg)
            self.root.after(0, self.status_var.set, "一键执行失败")
            self.root.after(0, messagebox.showerror, "错误", error_msg)
    
    def execute_clipboard_input(self):
        """执行复制粘贴操作"""
        # 从文本框获取命令
        command = self.command_text.get(1.0, tk.END).strip()
        if not command:
            messagebox.showwarning("警告", "请输入要粘贴的命令")
            return
        
        self.log(f"开始执行复制粘贴命令...")
        self.status_var.set("正在执行...")
        
        # 复制命令到剪贴板
        self.log("复制命令到剪贴板...")
        pyperclip.copy(command)
        time.sleep(0.5)
        
        # 查找 PAD 窗口
        try:
            self.log("查找 Power Automate Desktop 窗口...")
            desktop = Desktop(backend="uia")
            power_automate_window = desktop.window(title_re=".*Power Automate.*")
            
            if not power_automate_window.exists():
                self.log("错误: 未找到 Power Automate Desktop 窗口")
                self.status_var.set("错误: 未找到 PAD 窗口")
                messagebox.showerror("错误", "未找到 Power Automate Desktop 窗口，请确保应用已打开")
                return
            
            self.log(f"找到窗口: {power_automate_window.window_text()}")
            
            # 激活窗口
            power_automate_window.set_focus()
            time.sleep(0.5)
            
            # 根据选择的点击方式执行
            click_method = self.click_method_var.get()
            
            if click_method == "current":
                # 使用记录的鼠标位置
                if not self.recorded_position:
                    self.log("错误: 未记录鼠标位置")
                    self.status_var.set("错误: 未记录鼠标位置")
                    messagebox.showwarning("警告", "请先记录鼠标位置")
                    return
                
                x, y = self.recorded_position
                self.log(f"点击记录的位置: X={x}, Y={y}")
                pyautogui.click(x, y)
            else:
                # 自动查找编辑区
                self.log("尝试自动查找编辑区...")
                
                # 查找 Main 标签页
                tab_items = power_automate_window.descendants(control_type="TabItem")
                main_tab_found = False
                
                for tab_item in tab_items:
                    if tab_item.window_text() == "Main":
                        self.log("找到 Main 标签页，点击...")
                        tab_item.click_input()
                        main_tab_found = True
                        time.sleep(0.5)
                        break
                
                if not main_tab_found:
                    self.log("警告: 未找到 Main 标签页")
                
                # 获取窗口区域
                rect = power_automate_window.rectangle()
                
                # 计算可能的编辑区域位置（右侧中央区域）
                edit_x = rect.left + int(rect.width() * 0.7)  # 右侧 70% 位置
                edit_y = rect.top + int(rect.height() * 0.5)  # 中间高度
                
                self.log(f"点击可能的编辑区域: X={edit_x}, Y={edit_y}")
                pyautogui.click(edit_x, edit_y)
            
            # 点击后延迟
            delay = self.delay_var.get()
            time.sleep(delay)
            
            # 执行粘贴
            self.log("正在执行粘贴...")
            pyautogui.hotkey('ctrl', 'v')
            
            self.log("命令已粘贴到 PAD 中")
            self.status_var.set("完成: 命令已粘贴到 PAD 中")
            return True
            
        except Exception as e:
            error_msg = f"执行过程中出错: {str(e)}"
            self.log(error_msg)
            self.status_var.set("错误")
            messagebox.showerror("错误", error_msg)
            return False
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def main():
    app = PADClipboardInputTool()
    app.run()

if __name__ == "__main__":
    main() 