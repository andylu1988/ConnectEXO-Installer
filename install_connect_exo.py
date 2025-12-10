import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import subprocess
import json
import os
import sys
import threading
import requests
import time
import re

# 检查依赖库
try:
    from azure.identity import InteractiveBrowserCredential
except ImportError:
    # 如果缺少依赖，尝试自动安装
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "azure-identity", "requests"])
        from azure.identity import InteractiveBrowserCredential
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少依赖", f"无法自动安装依赖库，请手动运行：\npip install azure-identity requests\n\n错误: {e}")
        sys.exit(1)

class ConnectEXOInstallerApp:
    def __init__(self, root):
        # 启用 DPI 感知 (Windows)
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
            
        self.root = root
        self.root.title("Exchange Online PowerShell自动连接Module傻瓜化配置工具 v25")
        self.root.geometry("1000x900") # 增大窗口以适应 High DPI
        
        # 设置图标
        try:
            # 获取资源路径 (兼容 PyInstaller 打包)
            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_ico = os.path.join(base_path, "exchange.ico")
            icon_png = os.path.join(base_path, "exchange.png")

            # 优先尝试加载 .ico (用于窗口图标)
            if os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
            # 如果没有 .ico，尝试加载 .png
            elif os.path.exists(icon_png):
                icon_img = tk.PhotoImage(file=icon_png)
                self.root.iconphoto(True, icon_img)
        except Exception:
            pass

        # 使用 ttk 样式
        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except:
            self.style.theme_use('clam')
        
        # 定义一些样式 (优化字体和大小)
        self.style.configure("Title.TLabel", font=("Microsoft YaHei UI", 20, "bold"), foreground="#0078D7")
        self.style.configure("Header.TLabel", font=("Microsoft YaHei UI", 14, "bold"))
        self.style.configure("Info.TLabel", font=("Microsoft YaHei UI", 12))
        self.style.configure("Status.TLabel", font=("Microsoft YaHei UI", 12, "bold"), foreground="#333333")
        self.style.configure("Action.TButton", font=("Microsoft YaHei UI", 14, "bold"))
        
        # 自定义进度条样式
        self.style.configure("Thick.Horizontal.TProgressbar", thickness=25)

        # 初始化日志文件路径
        self.log_file_path = os.path.join(os.path.expanduser("~"), "Documents", "ConnectEXO_Install.log")
        self._init_log_file()

        # 主容器 (增加内边距)
        main_frame = ttk.Frame(root, padding="40 30 40 30")
        main_frame.pack(fill="both", expand=True)

        # 1. 标题区域
        title_label = ttk.Label(main_frame, text="Exchange Online PowerShell自动连接Module傻瓜化配置工具", style="Title.TLabel")
        title_label.pack(pady=(0, 25))

        # 2. 说明区域 (LabelFrame)
        info_frame = ttk.LabelFrame(main_frame, text="功能说明", padding="20")
        info_frame.pack(fill="x", pady=(0, 20))
        
        info_text = (
            "本工具将自动完成以下配置：\n\n"
            "• 登录 Azure AD (需全局管理员) 并获取租户信息\n"
            "• 本地生成自签名证书 & Azure AD 创建应用程序\n"
            "• 自动授予 Exchange.ManageAsApp 权限 & Exchange Administrator 角色\n"
            "• 检查并安装 ExchangeOnlineManagement 模块\n"
            "• 生成本地 PowerShell 连接脚本 (支持自定义模块名称)\n"
            "• 将脚本注册成本地 PowerShell 模块:\n"
            "   - 默认: ConnectEXO (国际版) / ConnectEXO21V (世纪互联)\n"
            "   - 自定义: 您指定的模块名称 (App名称将自动命名为 [模块名]-Automation-App)"
        )
        ttk.Label(info_frame, text=info_text, style="Info.TLabel", justify="left").pack(anchor="w")

        # 3. 配置区域 (环境选择)
        # 创建一个容器来容纳环境选择和操作选择，确保顺序正确
        self.config_container = ttk.Frame(main_frame)
        self.config_container.pack(fill="x", pady=(0, 10))

        config_frame = ttk.Frame(self.config_container)
        config_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(config_frame, text="1. 选择云环境:", style="Header.TLabel").pack(side="left", padx=(0, 15))
        
        self.env_var = tk.StringVar(value="") # 默认不选中
        ttk.Radiobutton(config_frame, text="Global (国际版)", variable=self.env_var, value="Global", command=self.on_env_selected).pack(side="left", padx=15)
        ttk.Radiobutton(config_frame, text="21Vianet (世纪互联)", variable=self.env_var, value="China", command=self.on_env_selected).pack(side="left", padx=15)

        # 3.2 自定义模块名称 (新增)
        self.name_frame = ttk.Frame(self.config_container)
        # self.name_frame.pack(fill="x", pady=(0, 10)) # 初始不显示

        ttk.Label(self.name_frame, text="2. 模块名称:", style="Header.TLabel").pack(side="left", padx=(0, 32))
        
        self.module_name_var = tk.StringVar()
        self.entry_module_name = ttk.Entry(self.name_frame, textvariable=self.module_name_var, width=30, font=("Microsoft YaHei UI", 12))
        self.entry_module_name.pack(side="left", padx=15)
        ttk.Label(self.name_frame, text="(字母数字组合, 例: MyConnectEXO)", foreground="gray").pack(side="left")

        # 3.5 配置区域 (操作选择) - 初始不显示
        self.action_frame = ttk.Frame(self.config_container)
        # self.action_frame.pack(fill="x", pady=(0, 20)) # 初始不 pack

        ttk.Label(self.action_frame, text="3. 选择操作:", style="Header.TLabel").pack(side="left", padx=(0, 32))

        self.action_var = tk.StringVar(value="") # 默认不选中
        ttk.Radiobutton(self.action_frame, text="安装 / 更新 (Install)", variable=self.action_var, value="Install").pack(side="left", padx=15)
        ttk.Radiobutton(self.action_frame, text="彻底卸载 (Uninstall)", variable=self.action_var, value="Uninstall").pack(side="left", padx=15)

        # 4. 进度条区域
        progress_frame = ttk.LabelFrame(main_frame, text="执行进度", padding="20")
        progress_frame.pack(fill="x", pady=(0, 20))

        self.status_var = tk.StringVar(value="准备就绪")
        self.lbl_status = ttk.Label(progress_frame, textvariable=self.status_var, style="Status.TLabel")
        self.lbl_status.pack(anchor="w", pady=(0, 10))

        self.progress_val = tk.DoubleVar(value=0)
        self.progressbar = ttk.Progressbar(progress_frame, variable=self.progress_val, maximum=100, style="Thick.Horizontal.TProgressbar")
        self.progressbar.pack(fill="x", ipady=2)

        # 5. 日志区域 (可折叠) - 调整顺序
        self.show_log = tk.BooleanVar(value=False)
        log_toggle_frame = ttk.Frame(main_frame)
        log_toggle_frame.pack(fill="x", pady=(0, 10))
        
        # 日志文件路径显示 (可点击) - 常驻显示
        self.lbl_log_path = ttk.Label(log_toggle_frame, text=f"日志文件: {self.log_file_path}", foreground="blue", cursor="hand2")
        self.lbl_log_path.pack(side="left")
        self.lbl_log_path.bind("<Button-1>", lambda e: self.open_log_file())

        self.btn_toggle_log = ttk.Button(log_toggle_frame, text="显示详细日志 ▼", command=self.toggle_log)
        self.btn_toggle_log.pack(side="right")

        # 6. 按钮区域 (固定在底部)
        self.btn_start = ttk.Button(main_frame, text="开始自动化配置", command=self.start_process, style="Action.TButton", width=30)
        self.btn_start.pack(side="bottom", pady=20, ipady=5)

        # 7. 日志内容框 (定义但不显示)
        self.log_frame = ttk.LabelFrame(main_frame, text="详细日志", padding="10")
        # 初始不显示 log_frame.pack(...)

        # 增加高度到 20 行，并添加垂直滚动条 (ScrolledText 自带)
        self.log_area = scrolledtext.ScrolledText(self.log_frame, height=20, state='disabled', font=("Consolas", 10))
        self.log_area.pack(fill="both", expand=True)
        
        self.log_area.tag_config("ERROR", foreground="red")
        self.log_area.tag_config("HEADER", foreground="blue", font=("Consolas", 10, "bold"))

    def open_log_file(self):
        try:
            os.startfile(self.log_file_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开日志文件: {e}")

    def on_env_selected(self):
        # 当用户选择环境后，显示模块名称输入框和操作选项
        env = self.env_var.get()
        if env == "China":
            default_name = "ConnectEXO21V"
        else:
            default_name = "ConnectEXO"
        
        # 只有当输入框为空或者等于另一个默认值时才自动填充，避免覆盖用户输入
        current_val = self.module_name_var.get()
        if not current_val or current_val in ["ConnectEXO", "ConnectEXO21V"]:
            self.module_name_var.set(default_name)

        if not self.name_frame.winfo_ismapped():
            self.name_frame.pack(fill="x", pady=(0, 10))
            
        if not self.action_frame.winfo_ismapped():
            self.action_frame.pack(fill="x", pady=(0, 20))

    def toggle_log(self):
        if self.show_log.get():
            self.log_frame.pack_forget()
            self.btn_toggle_log.config(text="显示详细日志 ▼")
            self.show_log.set(False)
            self.root.geometry("1000x900") # 恢复默认窗口大小
        else:
            # 展开时填充剩余空间 (位于 Toggle 和 底部按钮 之间)
            self.log_frame.pack(fill="both", expand=True, pady=(0, 10))
            self.btn_toggle_log.config(text="隐藏详细日志 ▲")
            self.show_log.set(True)
            self.root.geometry("1000x1100") # 增大窗口高度以显示更多日志

    def _init_log_file(self):
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*50}\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] 应用程序启动\n{'='*50}\n")
        except:
            pass

    def log(self, message, level="INFO"):
        self.log_area.config(state='normal')
        
        timestamp = time.strftime("%H:%M:%S")
        display_msg = f"[{timestamp}] {message}\n"
        
        # 简单的关键词自动着色逻辑
        tag = level
        if level == "INFO":
            if ">>>" in message: tag = "HEADER"
            elif "√" in message: tag = "SUCCESS"
            elif "!" in message: tag = "WARNING"
            elif "X" in message: tag = "ERROR"
        
        self.log_area.insert(tk.END, display_msg, tag)
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update()

        # 写入文件
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] [{level}] {message}\n")
        except:
            pass

    def update_progress(self, step, total_steps, message):
        progress = (step / total_steps) * 100
        self.progress_val.set(progress)
        self.status_var.set(f"步骤 {step}/{total_steps}: {message}")
        self.log(message, level="HEADER" if ">>>" in message else "INFO")

    def get_local_module_info(self, module_name):
        """尝试从本地已安装的模块文件中解析 AppID 和 Thumbprint"""
        base_module_path = self.get_best_module_path()
        psm1_path = os.path.join(base_module_path, module_name, f"{module_name}.psm1")
        
        info = {"AppID": None, "Thumbprint": None, "Path": psm1_path, "Exists": False}
        
        if os.path.exists(psm1_path):
            info["Exists"] = True
            try:
                with open(psm1_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    
                # 正则提取
                app_id_match = re.search(r'\$AppID\s*=\s*"([^"]+)"', content, re.IGNORECASE)
                thumb_match = re.search(r'\$Thumbprint\s*=\s*"([^"]+)"', content, re.IGNORECASE)
                
                if app_id_match: info["AppID"] = app_id_match.group(1)
                if thumb_match: info["Thumbprint"] = thumb_match.group(1)
            except Exception as e:
                print(f"Error parsing module file: {e}")
                
        return info

    def start_process(self):
        # 获取当前选择的环境和操作
        env = self.env_var.get()
        action = self.action_var.get()
        module_name = self.module_name_var.get().strip()

        if not env:
            messagebox.showwarning("提示", "请先选择云环境 (Global 或 21Vianet)")
            return
        
        if not module_name:
            messagebox.showwarning("提示", "请输入模块名称")
            return
            
        if not re.match(r'^[a-zA-Z0-9]+$', module_name):
            messagebox.showwarning("提示", "模块名称只能包含字母和数字")
            return

        if not action:
            messagebox.showwarning("提示", "请先选择操作 (安装 或 卸载)")
            return

        # 根据环境定义显示名称
        if env == "China":
            env_display = "21Vianet (世纪互联)"
        else:
            env_display = "Global (国际版)"
            
        # 统一命名规则
        app_display_name = f"{module_name}-Automation-App"
        cert_subject = f"{module_name}-Auto"

        # 清空日志区域
        self.log_area.config(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.config(state='disabled')

        # 获取本地模块信息 (用于卸载或更新检查)
        local_info = self.get_local_module_info(module_name)

        if action == "Uninstall":
            msg = f"确定要彻底卸载 [{env_display}] 环境的配置吗？\n\n模块名称: {module_name}\nApp名称: {app_display_name}\n\n这将执行以下操作：\n1. 删除本地 PowerShell 模块\n2. 清理 PowerShell Profile\n3. 删除本地证书\n4. 登录 Azure 并删除应用程序\n\n此操作不可撤销。"
            
            if not local_info["Exists"]:
                msg += "\n\n注意: 本地未找到该模块文件，将尝试根据名称规则清理云端资源。"
            
            if not messagebox.askyesno("确认卸载", msg):
                return
            
            self.btn_start.config(state='disabled', text="正在卸载...")
            self.progress_val.set(0)
            threading.Thread(target=self.run_uninstall, args=(env, module_name, app_display_name, cert_subject, local_info), daemon=True).start()
            return

        # Install 逻辑
        if local_info["Exists"]:
            # 询问用户是更新还是重装
            choice = messagebox.askyesno(
                "配置已存在", 
                f"检测到模块 ({module_name}) 已在本地安装。\n\n"
                "点击 '是 (Yes)' : 覆盖安装 (将删除旧配置，重新生成证书和 Azure App)。\n"
                "点击 '否 (No)' : 取消操作。"
            )
            if not choice:
                return

        self.btn_start.config(state='disabled', text="正在运行...")
        self.progress_val.set(0)
        # 将环境参数传递给 run_setup
        threading.Thread(target=self.run_setup, args=(env, module_name, app_display_name, cert_subject), daemon=True).start()

    def get_documents_dir(self):
        """获取真实的 'My Documents' 路径 (兼容 OneDrive)"""
        try:
            import ctypes.wintypes
            CSIDL_PERSONAL = 5       # My Documents
            SHGFP_TYPE_CURRENT = 0   # Get current, not default value
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
            return buf.value
        except:
            return os.path.join(os.path.expanduser("~"), "Documents")

    def get_best_module_path(self):
        """
        获取最佳的 PowerShell 模块安装路径。
        通过执行 PowerShell 命令来获取真实的 PSModulePath，而不是依赖 Python 进程的环境变量。
        """
        user_home = os.path.expanduser("~")
        
        # 尝试获取 PowerShell Core (pwsh) 和 Windows PowerShell (powershell) 的路径
        candidates = []
        
        for ps_cmd in ["pwsh", "powershell"]:
            try:
                # 获取当前用户的 PSModulePath
                cmd = [ps_cmd, "-NoProfile", "-Command", "[Environment]::GetEnvironmentVariable('PSModulePath', 'User')"]
                # creationflags=0x08000000 (CREATE_NO_WINDOW)
                creation_flags = 0x08000000 if sys.platform == 'win32' else 0
                process = subprocess.run(cmd, capture_output=True, text=True, creationflags=creation_flags)
                
                if process.returncode == 0:
                    paths = process.stdout.strip().split(';')
                    for p in paths:
                        p = p.strip()
                        # 只要路径不为空，就加入候选 (不再强制检查 startswith(user_home)，因为 OneDrive 路径可能不同)
                        if p:
                            candidates.append(p)
            except FileNotFoundError:
                continue # 如果没安装 pwsh，跳过
            except Exception:
                continue

        # 去重并保持顺序
        unique_candidates = []
        for p in candidates:
            if p not in unique_candidates:
                unique_candidates.append(p)

        # 1. 优先选择已存在的路径
        for path in unique_candidates:
            if os.path.exists(path):
                return path
                
        # 2. 如果都不存在，选择第一个看起来合理的路径 (优先选 Documents 下的)
        for path in unique_candidates:
            if "Documents" in path:
                return path
        
        # 3. 如果还是没有，Fallback 到默认路径 (使用真实的 Documents 路径)
        real_docs = self.get_documents_dir()
        
        # 优先尝试 PowerShell 7 (Core) 路径，如果用户装了的话
        ps7_path = os.path.join(real_docs, "PowerShell", "Modules")
        ps5_path = os.path.join(real_docs, "WindowsPowerShell", "Modules")
        
        # 如果检测到 pwsh 存在，优先用 ps7_path
        try:
            subprocess.run(["pwsh", "-v"], capture_output=True, creationflags=0x08000000 if sys.platform == 'win32' else 0)
            return ps7_path
        except:
            return ps5_path

    def get_profile_path(self):
        """
        根据模块路径推断 Profile 路径。
        """
        module_path = self.get_best_module_path()
        # module_path is usually .../Documents/PowerShell/Modules or .../Documents/WindowsPowerShell/Modules
        
        # Go up one level from Modules folder to get the config root
        # .../Documents/PowerShell/Modules -> .../Documents/PowerShell
        config_root = os.path.dirname(module_path)
        
        return os.path.join(config_root, "Microsoft.PowerShell_profile.ps1")

    def run_powershell_script(self, script):
        command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]
        process = subprocess.run(command, capture_output=True, text=True)
        if process.returncode != 0:
            raise Exception(f"PowerShell Error: {process.stderr}")
        return process.stdout.strip()

    def run_uninstall(self, env, module_name, app_display_name, cert_subject, local_info):
        total_steps = 5
        current_step = 0
        
        try:
            # 配置环境参数
            if env == "China":
                authority_host = "https://login.chinacloudapi.cn"
                graph_endpoint = "https://microsoftgraph.chinacloudapi.cn"
                scope = "https://microsoftgraph.chinacloudapi.cn/.default"
            else:
                authority_host = "https://login.microsoftonline.com"
                graph_endpoint = "https://graph.microsoft.com"
                scope = "https://graph.microsoft.com/.default"

            # --- Step 1: 清理本地模块 ---
            current_step += 1
            self.update_progress(current_step, total_steps, f">>> 正在清理本地 PowerShell 模块 ({module_name})...")
            
            base_module_path = self.get_best_module_path()
            module_path = os.path.join(base_module_path, module_name)
            
            if os.path.exists(module_path):
                try:
                    import shutil
                    shutil.rmtree(module_path)
                    self.log(f"√ 已删除本地模块文件夹: {module_path}", "SUCCESS")
                except Exception as e:
                    self.log(f"X 删除模块文件夹失败: {e}", "ERROR")
            else:
                self.log(f"! 本地模块文件夹不存在，跳过: {module_path}", "WARNING")

            # --- Step 2: 清理 Profile ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 正在清理 PowerShell Profile...")
            
            profile_path = self.get_profile_path()
            if os.path.exists(profile_path):
                try:
                    # 读取所有行
                    with open(profile_path, "r", encoding="utf-8-sig") as f:
                        lines = f.readlines()
                    
                    # 过滤掉包含模块名称的行 (精确匹配)
                    new_lines = []
                    removed_count = 0
                    for line in lines:
                        # 使用更严格的正则: 匹配 (行首或空白) + 模块名 + (空白或行尾)
                        pattern = rf"(?i)(^|\s){re.escape(module_name)}(\s|$)"
                        if re.search(pattern, line):
                            removed_count += 1
                            self.log(f"  - 移除 Profile 行: {line.strip()}")
                        else:
                            new_lines.append(line)
                    
                    if removed_count > 0:
                        with open(profile_path, "w", encoding="utf-8-sig") as f:
                            f.writelines(new_lines)
                        self.log(f"√ 已从 Profile 中移除 {removed_count} 行配置", "SUCCESS")
                    else:
                        self.log("! Profile 中未找到相关配置，无需修改", "WARNING")
                except Exception as e:
                    self.log(f"X 修改 Profile 失败: {e}", "ERROR")
            else:
                self.log("! 未找到 PowerShell Profile 文件，跳过", "WARNING")

            # --- Step 3: 清理本地证书 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 正在清理本地证书...")
            
            thumbprint_to_delete = local_info.get("Thumbprint")
            
            try:
                if thumbprint_to_delete:
                    self.log(f"正在根据指纹删除证书: {thumbprint_to_delete}")
                    ps_script = f"Get-ChildItem Cert:\\CurrentUser\\My | Where-Object {{ $_.Thumbprint -eq '{thumbprint_to_delete}' }} | Remove-Item"
                    self.run_powershell_script(ps_script)
                    self.log("√ 已尝试删除指定指纹的证书", "SUCCESS")
                else:
                    # 如果没有指纹，尝试按 Subject 删除
                    self.log(f"未找到本地指纹记录，尝试按 Subject 删除: {cert_subject}")
                    ps_script = f"Get-ChildItem Cert:\\CurrentUser\\My | Where-Object {{ $_.Subject -like '*{cert_subject}*' }} | Remove-Item"
                    self.run_powershell_script(ps_script)
                    self.log("√ 已尝试删除匹配 Subject 的证书", "SUCCESS")
            except Exception as e:
                self.log(f"X 删除证书失败: {e}", "WARNING")

            # --- Step 4: Azure 登录 ---
            current_step += 1
            self.update_progress(current_step, total_steps, f">>> 正在登录 Azure ({env}) 以清理应用程序...")
            
            try:
                credential = InteractiveBrowserCredential(authority=authority_host)
                token = credential.get_token(scope)
                access_token = token.token
                self.log("√ Azure 登录成功", "SUCCESS")
                
                # --- Step 5: 删除 Azure App ---
                current_step += 1
                
                headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
                app_id_to_delete = local_info.get("AppID")
                
                if app_id_to_delete:
                    self.update_progress(current_step, total_steps, f">>> 正在根据 AppID 删除 Azure App ({app_id_to_delete})...")
                    # 直接按 AppID (Client ID) 查询 Object ID
                    resp = requests.get(f"{graph_endpoint}/v1.0/applications?$filter=appId eq '{app_id_to_delete}'", headers=headers)
                    if resp.status_code == 200:
                        apps = resp.json().get('value', [])
                        if apps:
                            obj_id = apps[0]['id']
                            del_resp = requests.delete(f"{graph_endpoint}/v1.0/applications/{obj_id}", headers=headers)
                            if del_resp.status_code == 204:
                                self.log(f"√ 已删除 Azure App (AppID: {app_id_to_delete})", "SUCCESS")
                            else:
                                self.log(f"X 删除 Azure App 失败: {del_resp.text}", "ERROR")
                        else:
                            self.log(f"! 未找到 AppID 为 {app_id_to_delete} 的应用", "WARNING")
                    else:
                        self.log(f"X 查询 App 失败: {resp.text}", "ERROR")
                else:
                    # Fallback: 按名称删除
                    self.update_progress(current_step, total_steps, f">>> 正在根据名称查找并删除 Azure App ({app_display_name})...")
                    resp = requests.get(f"{graph_endpoint}/v1.0/applications?$filter=displayName eq '{app_display_name}'", headers=headers)
                    
                    if resp.status_code == 200:
                        apps = resp.json().get('value', [])
                        if apps:
                            for app in apps:
                                del_resp = requests.delete(f"{graph_endpoint}/v1.0/applications/{app['id']}", headers=headers)
                                if del_resp.status_code == 204:
                                    self.log(f"√ 已删除 Azure App: {app_display_name} (ID: {app['appId']})", "SUCCESS")
                                else:
                                    self.log(f"X 删除 Azure App 失败: {del_resp.text}", "ERROR")
                        else:
                            self.log(f"! 未找到名为 {app_display_name} 的 Azure App", "WARNING")
                    else:
                        self.log(f"X 查询 Azure App 失败: {resp.text}", "ERROR")

            except Exception as e:
                self.log(f"X Azure 操作失败 (可能是登录取消或网络问题): {e}", "ERROR")

            self.update_progress(total_steps, total_steps, ">>> 卸载操作完成")
            messagebox.showinfo("完成", f"[{env}] 环境的卸载清理已完成！")

        except Exception as e:
            self.log(f"发生未知错误: {e}", "ERROR")
            messagebox.showerror("错误", f"卸载过程中发生错误: {e}")
        finally:
            self.btn_start.config(state='normal', text="开始自动化配置")

    def run_setup(self, env, module_name, app_display_name, cert_subject):
        total_steps = 10
        current_step = 0
        
        try:
            # 配置环境参数
            if env == "China":
                authority_host = "https://login.chinacloudapi.cn"
                graph_endpoint = "https://microsoftgraph.chinacloudapi.cn"
                scope = "https://microsoftgraph.chinacloudapi.cn/.default"
                # 修正: 21Vianet 环境参数应为 O365China
                exo_env_param = "-ExchangeEnvironmentName O365China"
            else:
                authority_host = "https://login.microsoftonline.com"
                graph_endpoint = "https://graph.microsoft.com"
                scope = "https://graph.microsoft.com/.default"
                exo_env_param = "" 
            
            # 使用传入的参数
            function_name = module_name
            cert_dns_name = cert_subject

            # --- Step 1: Azure 登录 ---
            current_step += 1
            self.update_progress(current_step, total_steps, f">>> 正在启动 Azure ({env}) 浏览器登录...")
            
            credential = InteractiveBrowserCredential(authority=authority_host)
            token = credential.get_token(scope)
            access_token = token.token
            self.log("√ Azure 登录成功！", "SUCCESS")

            # --- Step 2: 获取租户信息 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 获取租户信息...")
            
            headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
            resp = requests.get(f"{graph_endpoint}/v1.0/organization", headers=headers)
            if resp.status_code != 200: raise Exception(f"无法获取租户信息: {resp.text}")
            
            org_info = resp.json()['value'][0]
            tenant_domain = org_info['verifiedDomains'][0]['name']
            for d in org_info['verifiedDomains']:
                if d['isDefault']:
                    tenant_domain = d['name']
                    break
            self.log(f"√ 检测到租户域名: {tenant_domain}", "SUCCESS")

            # --- Step 3: 清理旧应用 ---
            current_step += 1
            self.update_progress(current_step, total_steps, f">>> 检查并清理旧配置 ({app_display_name})...")
            
            resp = requests.get(f"{graph_endpoint}/v1.0/applications?$filter=displayName eq '{app_display_name}'", headers=headers)
            if resp.status_code == 200:
                existing_apps = resp.json().get('value', [])
                if existing_apps:
                    self.log(f"发现 {len(existing_apps)} 个同名旧应用，正在清理...", "WARNING")
                    for old_app in existing_apps:
                        del_resp = requests.delete(f"{graph_endpoint}/v1.0/applications/{old_app['id']}", headers=headers)
                        if del_resp.status_code == 204: self.log(f"√ 已删除旧应用: {old_app['appId']}", "SUCCESS")

            # --- Step 4: 本地生成证书 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 本地生成自签名证书...")
            
            ps_script = f"""
            $cert = New-SelfSignedCertificate -DnsName "{cert_dns_name}" -CertStoreLocation "cert:\\CurrentUser\\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddYears(2)
            $thumbprint = $cert.Thumbprint
            $certContent = [System.Convert]::ToBase64String($cert.GetRawCertData())
            $result = @{{Thumbprint=$thumbprint; Base64=$certContent}}
            $result | ConvertTo-Json -Compress
            """
            cert_json = self.run_powershell_script(ps_script)
            cert_data = json.loads(cert_json)
            thumbprint = cert_data['Thumbprint']
            cert_blob = cert_data['Base64']
            self.log(f"√ 证书已生成 (指纹: {thumbprint})", "SUCCESS")

            # --- Step 5: 创建 Azure AD 应用 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 创建 Azure AD 应用程序...")
            
            app_body = {
                "displayName": app_display_name,
                "signInAudience": "AzureADMyOrg",
                "keyCredentials": [{
                    "type": "AsymmetricX509Cert",
                    "usage": "Verify",
                    "key": cert_blob,
                    "displayName": f"{cert_dns_name}-Cert"
                }]
            }
            resp = requests.post(f"{graph_endpoint}/v1.0/applications", headers=headers, json=app_body)
            if resp.status_code != 201: raise Exception(f"创建应用失败: {resp.text}")
            
            app_info = resp.json()
            app_id = app_info['appId']
            self.log(f"√ 应用程序创建成功 (App ID: {app_id})", "SUCCESS")

            # --- Step 6: 创建服务主体 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 创建服务主体 (Service Principal)...")
            
            time.sleep(5) # 等待 App 传播
            sp_body = {"appId": app_id}
            resp = requests.post(f"{graph_endpoint}/v1.0/servicePrincipals", headers=headers, json=sp_body)
            if resp.status_code != 201:
                resp = requests.get(f"{graph_endpoint}/v1.0/servicePrincipals?$filter=appId eq '{app_id}'", headers=headers)
                sp_info = resp.json()['value'][0]
            else:
                sp_info = resp.json()
            sp_id = sp_info['id']
            self.log(f"√ 服务主体就绪 (Object ID: {sp_id})", "SUCCESS")

            # --- Step 7: 授予 API 权限 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 授予 Exchange.ManageAsApp 权限...")
            
            # Exchange Online AppId (Global & China 相同)
            exo_app_id = "00000002-0000-0ff1-ce00-000000000000"
            resp = requests.get(f"{graph_endpoint}/v1.0/servicePrincipals?$filter=appId eq '{exo_app_id}'", headers=headers)
            exo_sp_info = resp.json()['value'][0]
            exo_sp_id = exo_sp_info['id']
            
            manage_as_app_role_id = next((r['id'] for r in exo_sp_info['appRoles'] if r['value'] == "Exchange.ManageAsApp"), None)
            if not manage_as_app_role_id: raise Exception("未找到 Exchange.ManageAsApp 角色定义")

            assignment_body = {"principalId": sp_id, "resourceId": exo_sp_id, "appRoleId": manage_as_app_role_id}
            resp = requests.post(f"{graph_endpoint}/v1.0/servicePrincipals/{sp_id}/appRoleAssignments", headers=headers, json=assignment_body)
            if resp.status_code not in [201, 409]: raise Exception(f"权限授予失败: {resp.text}")
            self.log("√ 权限授予成功", "SUCCESS")

            # --- Step 8: 安装 PowerShell 模块 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 检查 ExchangeOnlineManagement 模块...")
            
            install_module_script = """
            $ErrorActionPreference = 'Stop'
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
            if (-not $nuget) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -Confirm:$false }
            if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
                Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -Confirm:$false
            }
            """
            try:
                self.run_powershell_script(install_module_script)
                self.log("√ 模块检查/安装完成", "SUCCESS")
            except Exception as e:
                self.log(f"X 模块安装遇到问题: {str(e)}", "WARNING")

            # --- Step 9: 分配管理员角色 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 分配 Exchange Administrator 角色...")
            
            role_template_id = "29232cdf-9323-42fd-ade2-1d097af3e4de"
            
            # 优先使用 Legacy API (directoryRoles) 以确保 Exchange Online 兼容性
            # 1. 激活角色 (如果尚未激活)
            requests.post(f"{graph_endpoint}/v1.0/directoryRoles", headers=headers, json={"roleTemplateId": role_template_id})
            
            # 2. 获取角色实例 ID
            resp = requests.get(f"{graph_endpoint}/v1.0/directoryRoles?$filter=roleTemplateId eq '{role_template_id}'", headers=headers)
            
            role_assigned = False
            if resp.status_code == 200 and resp.json().get('value'):
                role_id = resp.json()['value'][0]['id']
                member_body = {"@odata.id": f"{graph_endpoint}/v1.0/directoryObjects/{sp_id}"}
                
                # 3. 添加成员
                resp = requests.post(f"{graph_endpoint}/v1.0/directoryRoles/{role_id}/members/$ref", headers=headers, json=member_body)
                
                if resp.status_code == 204 or "already exist" in resp.text:
                    self.log("√ 角色分配成功 (Legacy API)", "SUCCESS")
                    role_assigned = True
                else:
                    self.log(f"! Legacy API 分配失败: {resp.text}，尝试 Unified API...", "WARNING")
            
            # 如果 Legacy API 失败，尝试 Unified API (roleManagement)
            if not role_assigned:
                role_assignment_body = {"principalId": sp_id, "roleDefinitionId": role_template_id, "directoryScopeId": "/"}
                resp = requests.post(f"{graph_endpoint}/v1.0/roleManagement/directory/roleAssignments", headers=headers, json=role_assignment_body)
                
                if resp.status_code == 201 or "already exists" in resp.text:
                    self.log("√ 角色分配成功 (Unified API)", "SUCCESS")
                else:
                    self.log(f"X 角色分配失败: {resp.text}", "ERROR")

            # 提示用户等待生效
            self.log("! 注意: 角色分配可能需要 5-15 分钟生效，如果连接报错请稍后重试。", "WARNING")

            # --- Step 10: 本地配置 ---
            current_step += 1
            self.update_progress(current_step, total_steps, f">>> 生成本地连接脚本 ({module_name})...")
            
            base_module_path = self.get_best_module_path()
            module_dir = os.path.join(base_module_path, module_name)
            if not os.path.exists(module_dir): os.makedirs(module_dir)
            
            psm1_content = f"""
function {function_name} {{
    param(
        [string]$Thumbprint = "{thumbprint}",
        [string]$AppID = "{app_id}",
        [string]$Organization = "{tenant_domain}"
    )
    Write-Host "Connecting to Exchange Online ({env}) for $Organization..." -ForegroundColor Cyan
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppID $AppID -Organization $Organization {exo_env_param}
}}
Export-ModuleMember -Function {function_name}
"""
            with open(os.path.join(module_dir, f"{module_name}.psm1"), "w", encoding="utf-8") as f:
                f.write(psm1_content)
            
            # Update Profile
            profile_path = self.get_profile_path()
            if not os.path.exists(os.path.dirname(profile_path)): os.makedirs(os.path.dirname(profile_path))
            
            current_content = ""
            if os.path.exists(profile_path):
                with open(profile_path, "r", encoding="utf-8") as f: current_content = f.read()
            
            import_cmd = f"Import-Module {module_name}"
            if import_cmd not in current_content:
                with open(profile_path, "a", encoding="utf-8") as f:
                    if current_content and not current_content.endswith("\n"): f.write("\n")
                    f.write(f"{import_cmd}\n")
            
            self.log("√ 本地模块配置完成", "SUCCESS")
            
            # 完成
            self.progress_val.set(100)
            self.status_var.set("配置全部完成！")
            self.log("-" * 30)
            self.log(f"全部完成！请重启 PowerShell 并运行 '{function_name}'。", "SUCCESS")
            messagebox.showinfo("成功", f"配置全部完成！\n\n请重启 PowerShell 并运行 '{function_name}' 进行测试。")

        except Exception as e:
            self.log(f"X 发生错误: {str(e)}", "ERROR")
            self.status_var.set("执行出错")
            messagebox.showerror("错误", f"执行过程中发生错误:\n{str(e)}")
        finally:
            self.btn_start.config(state='normal', text="开始自动化配置")

if __name__ == "__main__":
    root = tk.Tk()
    # 尝试设置高 DPI 感知 (Windows)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    app = ConnectEXOInstallerApp(root)
    root.mainloop()
