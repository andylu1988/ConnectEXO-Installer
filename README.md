# ConnectEXO Module Automation Tool / 自动化配置工具

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
# ConnectEXO Module Automation Tool

A GUI tool to automate the setup and configuration of Exchange Online PowerShell connection using Certificate-based authentication (App-only).

## Features

*   **Multi-Environment Support**: Supports **Global** (International) and **21Vianet** (China) Office 365 environments.
*   **Automated Configuration**:
    *   **Azure AD Login**: Interactive login to your tenant.
    *   **Certificate Management**: Automatically generates a self-signed certificate locally.
    *   **App Registration**: Creates an Azure AD Application and configures the certificate.
    *   **Permissions**: Automatically grants `Exchange.ManageAsApp` API permission and assigns the `Exchange Administrator` role to the Service Principal.
*   **Local Setup**:
    *   Checks and installs the `ExchangeOnlineManagement` PowerShell module.
    *   Generates a local PowerShell module (wrapper) for easy connection.
    *   Updates your PowerShell Profile to auto-load the module.
*   **Customization**:
    *   **Custom Module Name**: You can specify your own name for the PowerShell command (e.g., `MyExo`).
*   **Uninstallation**:
    *   **Full Cleanup**: One-click uninstallation that removes the local module, cleans up the PowerShell Profile, deletes the local certificate, and removes the Azure AD Application.

## Usage

1.  Run `ConnectEXO_Setup_v22.exe`.
2.  **Select Cloud Environment**: Choose between "Global" or "21Vianet".
3.  **Module Name**: (Optional) Enter a custom name for your connection command. Defaults are `ConnectEXO` (Global) or `ConnectEXO21V` (China).
4.  **Select Action**:
    *   **Install / Update**: Sets up the environment.
    *   **Uninstall**: Removes all configurations and cleans up resources.
5.  Click **Start Automation Config** (开始自动化配置).
6.  Follow the prompts to log in to your Azure account.

### After Installation
Restart your PowerShell and run your command (e.g., `ConnectEXO`) to connect to Exchange Online.

## Build from Source

Requirements: Python 3.x, `azure-identity`, `requests`, `tkinter`.

```bash
pip install azure-identity requests
pyinstaller --noconsole --onefile --clean --name "ConnectEXO_Setup_v22" --add-data "exchange.png;." --hidden-import=azure.identity --hidden-import=requests --hidden-import=tkinter install_connect_exo.py
```

---

<a name="chinese"></a>
# ConnectEXO Module 自动化配置工具

一个用于自动化配置 Exchange Online PowerShell 基于证书认证（App-only）连接环境的 GUI 工具。

## 功能特点

*   **多环境支持**: 支持 **Global (国际版)** 和 **21Vianet (世纪互联)** Office 365 环境。
*   **全自动化配置**:
    *   **Azure AD 登录**: 交互式登录到您的租户。
    *   **证书管理**: 本地自动生成自签名证书。
    *   **应用注册**: 在 Azure AD 中创建应用程序并配置证书。
    *   **权限管理**: 自动授予 `Exchange.ManageAsApp` API 权限，并为服务主体分配 `Exchange Administrator` 角色。
*   **本地环境部署**:
    *   自动检查并安装 `ExchangeOnlineManagement` PowerShell 模块。
    *   生成本地 PowerShell 封装模块，方便快捷连接。
    *   自动更新 PowerShell Profile 配置文件以加载模块。
*   **自定义功能**:
    *   **自定义模块名称**: 您可以指定生成的 PowerShell 命令名称（例如：`MyExo`）。
*   **彻底卸载**:
    *   **一键清理**: 支持彻底卸载，包括删除本地模块文件、清理 PowerShell Profile 配置、删除本地证书以及移除 Azure AD 上的应用程序。

## 使用方法

1.  运行 `ConnectEXO_Setup_v22.exe`。
2.  **选择云环境**: 选择 "Global (国际版)" 或 "21Vianet (世纪互联)"。
3.  **模块名称**: (可选) 输入您想要的连接命令名称。默认为 `ConnectEXO` (国际版) 或 `ConnectEXO21V` (世纪互联)。
4.  **选择操作**:
    *   **安装 / 更新 (Install)**: 执行配置安装。
    *   **彻底卸载 (Uninstall)**: 删除所有配置并清理资源。
5.  点击 **开始自动化配置**。
6.  根据弹窗提示完成 Azure 账号登录。

### 安装完成后
重启 PowerShell 终端，直接运行您设置的命令（如 `ConnectEXO`）即可连接到 Exchange Online。

## 源码构建

依赖环境: Python 3.x, `azure-identity`, `requests`, `tkinter`.

```bash
pip install azure-identity requests
pyinstaller --noconsole --onefile --clean --name "ConnectEXO_Setup_v22" --add-data "exchange.png;." --hidden-import=azure.identity --hidden-import=requests --hidden-import=tkinter install_connect_exo.py
```
