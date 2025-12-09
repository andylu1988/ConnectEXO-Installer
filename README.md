# ConnectEXO Automatic Installer

这是一个用于自动化配置 Exchange Online PowerShell 连接环境的工具。

## 功能

*   **自动登录 Azure AD**: 支持 Global 和 21Vianet (世纪互联) 环境。
*   **自动配置**:
    *   获取租户信息。
    *   生成自签名证书。
    *   在 Azure AD 中创建应用程序并配置证书。
    *   创建服务主体 (Service Principal)。
    *   授予 `Exchange.ManageAsApp` 权限。
    *   分配 `Exchange Administrator` 角色。
*   **本地环境配置**:
    *   检查并安装 `ExchangeOnlineManagement` 模块。
    *   生成本地 PowerShell 模块 `ConnectEXO`。
    *   更新 PowerShell Profile 以自动加载模块。

## 使用方法

1.  运行 `ConnectEXO_Setup.exe` (或运行 Python 脚本)。
2.  选择云环境 (Global 或 21Vianet)。
3.  点击 "开始自动化配置"。
4.  根据提示完成 Azure 登录。
5.  等待配置完成。
6.  重启 PowerShell，运行 `ConnectEXO` 即可连接。

## 开发环境

*   Python 3.x
*   Tkinter (GUI)
*   Azure Identity SDK
*   Requests

## 构建

使用 PyInstaller 构建 EXE:

```bash
pyinstaller --noconsole --onefile --clean --name "ConnectEXO_Setup" --add-data "exchange.png;." --hidden-import=azure.identity --hidden-import=requests --hidden-import=tkinter install_connect_exo.py
```
