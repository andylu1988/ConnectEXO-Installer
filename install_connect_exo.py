import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import subprocess
import json
import os
import sys
import threading
import requests
import time

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
        self.root.title("ConnectEXO Module自动化配置工具")
        self.root.geometry("850x750") # 稍微调大默认窗口
        
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
        self.style.configure("Title.TLabel", font=("Microsoft YaHei UI", 26, "bold"), foreground="#0078D7")
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
        title_label = ttk.Label(main_frame, text="ConnectEXO自动连接模块自动化安装", style="Title.TLabel")
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
            "• 生成本地 PowerShell 连接脚本\n"
            "• 将脚本注册成本地 PowerShell 模块 (使用命令ConnectEXO自动运行)"
        )
        ttk.Label(info_frame, text=info_text, style="Info.TLabel", justify="left").pack(anchor="w")

        # 3. 配置区域 (环境选择)
        config_frame = ttk.Frame(main_frame)
        config_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(config_frame, text="选择云环境:", style="Header.TLabel").pack(side="left", padx=(0, 15))
        
        self.env_var = tk.StringVar(value="Global")
        ttk.Radiobutton(config_frame, text="Global (国际版)", variable=self.env_var, value="Global").pack(side="left", padx=15)
        ttk.Radiobutton(config_frame, text="21Vianet (世纪互联)", variable=self.env_var, value="China").pack(side="left", padx=15)

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
        
        # 配置日志颜色标签
        self.log_area.tag_config("INFO", foreground="black")
        self.log_area.tag_config("SUCCESS", foreground="green")
        self.log_area.tag_config("WARNING", foreground="#FF8C00") # DarkOrange
        self.log_area.tag_config("ERROR", foreground="red")
        self.log_area.tag_config("HEADER", foreground="blue", font=("Consolas", 10, "bold"))

    def toggle_log(self):
        if self.show_log.get():
            self.log_frame.pack_forget()
            self.btn_toggle_log.config(text="显示详细日志 ▼")
            self.show_log.set(False)
            self.root.geometry("850x750") # 恢复默认窗口大小
        else:
            # 展开时填充剩余空间 (位于 Toggle 和 底部按钮 之间)
            self.log_frame.pack(fill="both", expand=True, pady=(0, 10))
            self.btn_toggle_log.config(text="隐藏详细日志 ▲")
            self.show_log.set(True)
            self.root.geometry("850x950") # 增大窗口高度以显示更多日志

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

    def start_process(self):
        # 检查本地是否已安装
        user_profile = os.path.expanduser("~")
        ps_path = os.path.join(user_profile, "Documents", "WindowsPowerShell", "Modules", "ConnectEXO", "ConnectEXO.psm1")
        
        if os.path.exists(ps_path):
            if not messagebox.askyesno("配置已存在", "检测到 ConnectEXO 模块已在本地安装。\n\n是否要删除旧配置并重新添加？\n(这将清理 Azure 上的旧 App 并重新生成证书)"):
                return

        self.btn_start.config(state='disabled', text="正在运行...")
        self.progress_val.set(0)
        threading.Thread(target=self.run_setup, daemon=True).start()

    def run_powershell_script(self, script):
        command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]
        process = subprocess.run(command, capture_output=True, text=True)
        if process.returncode != 0:
            raise Exception(f"PowerShell Error: {process.stderr}")
        return process.stdout.strip()

    def run_setup(self):
        total_steps = 10
        current_step = 0
        
        try:
            # 获取环境配置
            env = self.env_var.get()
            if env == "China":
                authority_host = "https://login.chinacloudapi.cn"
                graph_endpoint = "https://microsoftgraph.chinacloudapi.cn"
                scope = "https://microsoftgraph.chinacloudapi.cn/.default"
                exo_env_param = "-ExchangeEnvironmentName AzureChinaCloud"
            else:
                authority_host = "https://login.microsoftonline.com"
                graph_endpoint = "https://graph.microsoft.com"
                scope = "https://graph.microsoft.com/.default"
                exo_env_param = "" 

            # --- Step 1: Azure 登录 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 正在启动 Azure 浏览器登录...")
            
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
            self.update_progress(current_step, total_steps, ">>> 检查并清理旧配置...")
            
            app_name = "ConnectEXO-Automation-App"
            resp = requests.get(f"{graph_endpoint}/v1.0/applications?$filter=displayName eq '{app_name}'", headers=headers)
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
            
            ps_script = """
            $cert = New-SelfSignedCertificate -DnsName "ConnectEXO-Auto" -CertStoreLocation "cert:\\CurrentUser\\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddYears(2)
            $thumbprint = $cert.Thumbprint
            $certContent = [System.Convert]::ToBase64String($cert.GetRawCertData())
            $result = @{Thumbprint=$thumbprint; Base64=$certContent}
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
                "displayName": app_name,
                "signInAudience": "AzureADMyOrg",
                "keyCredentials": [{
                    "type": "AsymmetricX509Cert",
                    "usage": "Verify",
                    "key": cert_blob,
                    "displayName": "ConnectEXO-Auto-Cert"
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
            role_assignment_body = {"principalId": sp_id, "roleDefinitionId": role_template_id, "directoryScopeId": "/"}
            resp = requests.post(f"{graph_endpoint}/v1.0/roleManagement/directory/roleAssignments", headers=headers, json=role_assignment_body)
            
            if resp.status_code == 201 or "already exists" in resp.text:
                self.log("√ 角色分配成功 (Unified API)", "SUCCESS")
            else:
                self.log("! 尝试使用 Legacy API 分配角色...", "WARNING")
                # Legacy fallback logic...
                requests.post(f"{graph_endpoint}/v1.0/directoryRoles", headers=headers, json={"roleTemplateId": role_template_id})
                resp = requests.get(f"{graph_endpoint}/v1.0/directoryRoles?$filter=roleTemplateId eq '{role_template_id}'", headers=headers)
                if resp.status_code == 200 and resp.json()['value']:
                    role_id = resp.json()['value'][0]['id']
                    member_body = {"@odata.id": f"{graph_endpoint}/v1.0/directoryObjects/{sp_id}"}
                    resp = requests.post(f"{graph_endpoint}/v1.0/directoryRoles/{role_id}/members/$ref", headers=headers, json=member_body)
                    if resp.status_code == 204 or "already exist" in resp.text:
                        self.log("√ 角色分配成功 (Legacy API)", "SUCCESS")
                    else:
                        self.log(f"X 角色分配失败: {resp.text}", "ERROR")

            # --- Step 10: 本地配置 ---
            current_step += 1
            self.update_progress(current_step, total_steps, ">>> 生成本地连接脚本...")
            
            user_profile = os.path.expanduser("~")
            module_dir = os.path.join(user_profile, "Documents", "WindowsPowerShell", "Modules", "ConnectEXO")
            if not os.path.exists(module_dir): os.makedirs(module_dir)
            
            psm1_content = f"""
function ConnectEXO {{
    param(
        [string]$Thumbprint = "{thumbprint}",
        [string]$AppID = "{app_id}",
        [string]$Organization = "{tenant_domain}"
    )
    Write-Host "Connecting to Exchange Online for $Organization..." -ForegroundColor Cyan
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppID $AppID -Organization $Organization {exo_env_param}
}}
Export-ModuleMember -Function ConnectEXO
"""
            with open(os.path.join(module_dir, "ConnectEXO.psm1"), "w", encoding="utf-8") as f:
                f.write(psm1_content)
            
            # Update Profile
            profile_path = os.path.join(user_profile, "Documents", "WindowsPowerShell", "Microsoft.PowerShell_profile.ps1")
            if not os.path.exists(os.path.dirname(profile_path)): os.makedirs(os.path.dirname(profile_path))
            
            current_content = ""
            if os.path.exists(profile_path):
                with open(profile_path, "r", encoding="utf-8") as f: current_content = f.read()
            
            if "Import-Module ConnectEXO" not in current_content:
                with open(profile_path, "a", encoding="utf-8") as f:
                    if current_content and not current_content.endswith("\n"): f.write("\n")
                    f.write("Import-Module ConnectEXO\n")
            
            self.log("√ 本地模块配置完成", "SUCCESS")
            
            # 完成
            self.progress_val.set(100)
            self.status_var.set("配置全部完成！")
            self.log("-" * 30)
            self.log("全部完成！请重启 PowerShell 并运行 'ConnectEXO'。", "SUCCESS")
            messagebox.showinfo("成功", "配置全部完成！\n\n请重启 PowerShell 并运行 'ConnectEXO' 进行测试。")

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
