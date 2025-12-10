# ConnectEXO Installer (Exchange Online PowerShell Automation)

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
## 🇬🇧 English

### Introduction
**ConnectEXO Installer** is a "lazy" configuration tool designed for Exchange Online administrators. It completely automates the complex process of setting up **Certificate-based Authentication (CBA)** for Exchange Online PowerShell.

Instead of manually creating Azure AD Apps, generating certificates, uploading public keys, and writing connection scripts, this tool does it all in one click. It generates a local PowerShell command (e.g., `Connect-EXO`) that lets you connect to Exchange Online instantly without entering credentials every time.

### Key Features
*   **Zero-Touch Configuration**:
    *   Logs into Azure AD (Interactive).
    *   Creates an Azure AD Application.
    *   Generates a Self-signed Certificate locally.
    *   Uploads the certificate to the Azure AD App.
    *   Grants `Exchange.ManageAsApp` API permissions.
    *   Assigns the `Exchange Administrator` role to the Service Principal.
*   **Local Integration**:
    *   Checks and installs the `ExchangeOnlineManagement` module (NuGet/PSGallery).
    *   Generates a custom PowerShell module wrapper.
    *   **Smart Path Detection**: Automatically installs to the correct user module path (supports PowerShell 5.1 & 7+).
    *   Updates `Microsoft.PowerShell_profile.ps1` for auto-loading.
*   **Multi-Cloud Support**: Fully supports **Global (International)** and **21Vianet (China)** environments.
*   **Clean Uninstallation**: A dedicated "Uninstall" mode that removes the local module, cleans the Profile, deletes the certificate, and removes the Azure AD App.

### Prerequisites
*   **OS**: Windows 10/11 or Windows Server.
*   **Permissions**: You must be a **Global Administrator** to register apps and assign roles.
*   **PowerShell**: PowerShell 5.1 or PowerShell 7 (Core) installed.

### Usage

#### 1. Installation
Download the latest executable `ConnectEXO_Setup_v23.exe` from the [Releases](../../releases) page.

#### 2. Setup (Install)
1.  Run the tool.
2.  **Cloud Environment**: Select "Global" or "21Vianet".
3.  **Module Name**: Enter a name for your command (Default: `ConnectEXO`).
4.  Click **"Start Automation Config"**.
5.  Log in with your Global Admin account when prompted.
6.  Wait for the process to complete (approx. 30-60 seconds).

#### 3. Connect
Once finished, open a **new** PowerShell window and type:
```powershell
ConnectEXO
```
*(Or whatever name you configured)*. You will be connected immediately.

#### 4. Uninstallation
1.  Run the tool.
2.  Select the **"Uninstall"** tab.
3.  Enter the Module Name you want to remove.
4.  Click **"Start Uninstall"**.
    *   *Note: This will permanently delete the Azure AD App and local certificate.*

### Build from Source
```bash
pip install azure-identity requests
# Build with PyInstaller
python -m PyInstaller --noconsole --onefile --name "ConnectEXO_Setup_v23" --add-data "exchange.png;." --hidden-import=azure.identity --hidden-import=requests --hidden-import=tkinter install_connect_exo.py
```

---

<a name="chinese"></a>
## 🇨🇳 中文 (Chinese)

### 简介
**ConnectEXO Installer** 是一个为 Exchange Online 管理员打造的"傻瓜化"配置工具。它旨在全自动完成 Exchange Online PowerShell 的 **证书认证 (Certificate-based Authentication)** 配置流程。

你不再需要手动去 Azure 门户创建应用、生成证书、上传公钥、写连接脚本。这个工具可以一键完成所有工作，并生成一个本地的 PowerShell 命令（如 `Connect-EXO`），让你以后无需输入密码即可秒连 Exchange Online。

### 主要功能
*   **全自动配置**:
    *   自动登录 Azure AD (交互式)。
    *   自动创建 Azure AD 应用程序 (App Registration)。
    *   自动在本地生成自签名证书。
    *   自动将证书公钥上传到 Azure AD 应用。
    *   自动授予 `Exchange.ManageAsApp` API 权限。
    *   自动为服务主体分配 `Exchange Administrator` (Exchange 管理员) 角色。
*   **本地集成**:
    *   自动检查并安装 `ExchangeOnlineManagement` 模块。
    *   生成自定义的 PowerShell 模块封装脚本。
    *   **智能路径检测**: 自动识别 PowerShell 模块安装路径 (支持 PS 5.1 和 PS 7+)。
    *   自动更新 `Microsoft.PowerShell_profile.ps1` 配置文件，实现启动即加载。
*   **多环境支持**: 完美支持 **Global (国际版)** 和 **21Vianet (世纪互联)** 环境。
*   **一键卸载**: 提供"卸载"模式，可自动清理本地模块、Profile 配置、本地证书，并删除云端的 Azure AD 应用。

### 前置条件
*   **操作系统**: Windows 10/11 或 Windows Server。
*   **权限**: 需要 **全局管理员 (Global Admin)** 权限以注册应用和分配角色。
*   **PowerShell**: 系统需安装 PowerShell 5.1 或 PowerShell 7 (Core)。

### 使用指南

#### 1. 下载
从 [Releases](../../releases) 页面下载最新的 `ConnectEXO_Setup_v23.exe`。

#### 2. 安装配置
1.  运行工具。
2.  **云环境**: 选择 "Global (国际版)" 或 "21Vianet (世纪互联)"。
3.  **模块名称**: 输入你想要的命令名称 (默认: `ConnectEXO`)。
4.  点击 **"开始自动化配置"**。
5.  在弹出的浏览器窗口中登录你的全局管理员账号。
6.  等待进度条走完 (约 30-60 秒)。

#### 3. 连接使用
配置完成后，打开一个新的 PowerShell 窗口，直接输入：
```powershell
ConnectEXO
```
*(或者你自定义的名称)*，即可立即连接，无需输入密码。

#### 4. 卸载清理
1.  运行工具。
2.  切换到 **"卸载 (Uninstall)"** 标签页。
3.  输入要卸载的模块名称。
4.  点击 **"开始卸载"**。
    *   *注意：这将永久删除云端的 Azure AD 应用和本地证书。*

### 源码构建
```bash
pip install azure-identity requests
# 使用 PyInstaller 打包
python -m PyInstaller --noconsole --onefile --name "ConnectEXO_Setup_v23" --add-data "exchange.png;." --hidden-import=azure.identity --hidden-import=requests --hidden-import=tkinter install_connect_exo.py
```
