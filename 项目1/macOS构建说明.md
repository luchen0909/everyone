# macOS 构建说明

## 生成应用

1. 将整个构建包复制到苹果电脑。
2. 安装 Python 3.11 或更高版本。
3. 在终端进入 `项目1` 目录，执行：

```bash
chmod +x build_macos.command
./build_macos.command
```

输出应用位于：

```text
项目1/dist/供需协同工具V3.0.app
```

首次运行时，在 Finder 中右键应用并选择“打开”。

## 使用 GitHub Actions 自动打包

仓库已经包含 `.github/workflows/build-macos.yml`。

1. 将构建包内容推送到 GitHub 仓库，保持 `项目1` 和 `静态平衡表` 两个目录位于仓库根目录。
2. 打开仓库的 `Actions` 页面。
3. 选择 `Build macOS app`，点击 `Run workflow`。
4. 构建完成后，在该次运行底部下载 `供需协同工具V3.0-macOS` 制品。

工作流会自动执行源码编译、模板检查、PyInstaller 打包、临时签名、签名校验和 `Info.plist` 校验。普通 GitHub 临时签名不等于苹果开发者公证，首次打开仍可能需要在 Finder 中右键选择“打开”。

## 数据库模式

文件上传模式不需要额外配置。

如需直接连接 SQL Server，请先在 Mac 安装 Microsoft ODBC Driver 18 for SQL Server，并通过环境变量 `MRP_U9_CONN` 或以下文件保存连接串：

```text
~/.config/bom_tool/u9_db_conn.txt
```

未安装驱动时，取消勾选“PR/PO使用数据库”，使用文件上传即可。

## 架构说明

应用会按执行打包命令的 Mac 架构生成：Apple Silicon Mac 生成 arm64，Intel Mac 生成 x86_64。不同架构建议分别打包。
