# Visio 转 PDF 工具

# 下载源文件自己编译前请先将环境runtime.zip.001.....解压到脚本一起

版本：1.0.0
运行平台：Windows 10/11（64 位）

Windows 命令行小工具，支持 `.vsd`、`.vsdx`，把 PDF 输出到源文件所在位置。界面、日志和错误提示均为中文。

## 使用

1. 只需复制 `Visio转PDF.exe`，不需要复制源码或依赖目录。
2. 将一个或多个 VSD/VSDX 文件拖到 exe 图标上，松开鼠标。
3. 等待黑色命令行窗口显示“转换成功”和“输出位置”，按回车关闭窗口。
4. PDF 就在原 Visio 文件旁边。

适用系统：Windows 10/11，64 位。

不需要安装 Microsoft Visio、LibreOffice、Python 或其他运行库。转换在本机完成，不上传图纸，也不需要联网。

首次运行会释放内置组件到：

```text
%LOCALAPPDATA%\VisioToPDF\libreoffice-26.8.3-v1
```

首次准备组件约需几秒，同时需要约 600MB 本机磁盘空间；后续转换直接使用已释放的组件。exe 为 248543119 字节（约 237MB），这是内置转换引擎带来的体积，不是重复文件。

## 输出规则

- 原文件 `图纸.vsdx` 输出为同目录的 `图纸.pdf`。
- 原文件不修改、不删除。
- 已有同名 PDF 时不覆盖，自动使用 `图纸_1.pdf`、`图纸_2.pdf` 等名称。
- 一次可拖入多个文件；成功的文件会逐个列出完整 PDF 路径。
- 如果部分文件失败，其余文件仍继续处理，最后显示成功和失败数量。
- 程序检查生成的 PDF 是否能读取以及是否有页面，但不会替代人工核对图纸内容。

## 命令行用法

路径含中文或空格时加双引号：

```bat
Visio转PDF.exe "D:\图纸\示例.vsd"
Visio转PDF.exe "D:\图纸\示例.vsdx" "D:\图纸\另一张图.vsd"
Visio转PDF.exe --help
Visio转PDF.exe --version
```

也可以双击 exe，在窗口中粘贴带双引号的文件路径后按回车。

## 程序原理与具体实现

### 1. 整体处理流程

```text
拖入文件 / 命令行参数 / 双击后粘贴路径
                  ↓
        检查文件存在与扩展名
                  ↓
      计算源目录中的 PDF 输出名称
                  ↓
        优先尝试 Microsoft Visio
                  ↓ 不可用或导出失败
       内置 LibreOffice + libvisio 引擎
                  ↓
       检查 PDF 可读取且至少有一页
                  ↓
      中文日志输出成功状态及完整路径
```

每个文件单独处理。某个文件失败时记录失败原因，继续处理队列中的下一个文件，最后汇总成功与失败数量。

### 2. 为什么有两条转换路径

Microsoft Visio 原生导出更适合自己的文件格式，因此优先使用。没有安装 Visio 的电脑则使用内置引擎，避免还要安装办公软件。

两条路径都只在本机读取和写入文件，不调用网络转换服务。当前版本采用“Visio 失败则回退内置引擎”的策略：即使电脑装有 Visio，但文件无法被 Visio 自动化接口打开，也会再尝试内置引擎。

### 3. Visio 原生导出

`visio_export()` 通过 pywin32 调用 Visio COM 自动化接口：

```python
app = win32com.client.DispatchEx("Visio.Application")
app.Visible = False
app.AlertResponse = 1
document = app.Documents.OpenEx(str(source), 2 | 64 | 128 | 256)
document.ExportAsFixedFormat(1, str(target), 1, 0)
```

其中：

- `DispatchEx` 创建独立的 Visio 后台实例，避免复用正在编辑其他图纸的实例。
- 打开标志 `2 | 64 | 128 | 256` 分别对应只读、隐藏、禁用宏、不恢复工作区。
- `ExportAsFixedFormat` 的第一个参数 `1` 表示 PDF，随后依次是输出文件、打印质量和导出全部页面。
- 最后关闭文档并退出该后台实例，不保存源文件。

这部分参考了两个公开项目的实现思路，详见许可说明。

### 4. 内置引擎导出

`portable_export()` 调用随附的 LibreOffice 命令行程序 `soffice.com`。LibreOffice 通过 libvisio 读取 VSD/VSDX，导入为 Draw 图形文档，再用 `draw_pdf_Export` 导出 PDF。

实际命令的核心参数为：

```text
soffice.com --headless --invisible --nologo --nodefault
  --nolockcheck --norestore
  -env:UserInstallation=file:///独立配置目录
  --convert-to pdf:draw_pdf_Export
  --outdir 独立工作目录 输入文件
```

实现细节：

- 后台运行，不打开 LibreOffice 窗口。
- 每次转换使用独立用户配置，避免与电脑上已有 LibreOffice 的配置冲突。
- 子进程最多等待 5 分钟，超时后终止本次转换并报告失败。
- 只有返回码为 0、确实生成 PDF、PDF 通过结构校验时才报告成功。
- 转换结束后删除本次临时工作目录。

### 5. 中文路径和临时文件处理

源文件先复制到临时工作目录中的 `input.vsd` 或 `input.vsdx`，引擎只处理简单的英文临时文件名。成功后再将 PDF 移回原来的中文目录，避免旧版本办公软件对特殊路径的兼容问题。

拖放文件时 Windows 会把路径作为参数交给 exe，不要求手工加引号；在命令行手工输入含空格的路径时，需要加双引号。

控制台启动时设置 UTF-8 编码，保证中文日志正常显示。错误信息优先显示中文；底层组件返回的英文诊断也会保留，便于定位具体文件问题。

### 6. 输出保护与 PDF 校验

`output_path()` 从 `原名称.pdf` 开始检查。如果已存在，则递增为 `原名称_1.pdf`、`原名称_2.pdf`，不会覆盖旧 PDF。

`validate_pdf()` 使用 pypdf 严格模式读取成品，要求文件存在、大小有效且至少包含一页。这能识别未生成、空文件、损坏 PDF 等情况，但不能自动判断图形是否与 Visio 中完全一致，所以复杂图纸仍应人工抽查。

### 7. 单文件 exe 和运行组件

PyInstaller 将主程序、Python 运行环境、pypdf、pywin32 和 `runtime.zip` 一起封装为 exe。严格说它是“自包含的单文件程序”，不是把所有功能都写成一个原生静态链接库。

首次需要内置引擎时：

1. 从 exe 中读取组件压缩包。
2. 释放到 `%LOCALAPPDATA%\VisioToPDF\libreoffice-26.8.3-v1`。
3. 检查 `program/soffice.com` 是否存在。
4. 后续直接复用，不再重复释放。

这是文件释放，不是执行 Office 安装程序，也不修改系统文件关联。删除上述缓存目录只会导致下次重新释放，不影响源文件和已生成 PDF。

### 8. 组件裁剪与零安装

完整 LibreOffice 解包后包含大量与绘图转换无关的内容。随附的 `runtime` 是已经裁剪并验证的精简组件，保留 Draw、Visio 导入、PDF 导出、字体配置和中英文资源，裁去了其他语言包、扩展、帮助、模板、Python 开发组件等。

Microsoft Visual C++ 所需 DLL 直接放在转换程序的 `program` 目录，因此目标电脑无需另行安装 VC++ 运行库。当前成品在 Windows 11 64 位验证，其他 Windows 版本仍建议先试运行。

### 9. 关键源码函数

| 函数 | 作用 |
| --- | --- |
| `main()` | 解析参数、批量调度、日志及退出码 |
| `collect_inputs()` | 检查文件存在、扩展名与重复输入 |
| `output_path()` | 源目录输出及重名自动编号 |
| `visio_export()` | Visio COM 后台原生导出 |
| `prepare_runtime()` | 查找或首次释放内置组件 |
| `portable_export()` | 调用 LibreOffice 后台导出并管理临时文件 |
| `convert()` | 原生导出失败后切换内置引擎 |
| `validate_pdf()` | 校验成品可读且至少有一页 |
| `configure_console()` | 中文控制台编码设置 |

正常全部成功时退出码为 0；输入文件不存在或格式错误为 2；有转换失败为 1。

## 转换兼容性与使用注意事项

程序优先尝试调用本机已安装的 Microsoft Visio；不可用时自动改用随附的 LibreOffice 26.8.3 转换组件。两种引擎都不需要联网。

- 安装 Visio 的电脑上，Visio 原生导出通常具有更好的版式保真度。
- 内置引擎已经通过真实 VSD、VSDX 的转换和 PDF 画面检查，但不保证所有 Visio 图纸完全一致。
- 特殊模具、嵌入对象、字体缺失、复杂公式或个别文件结构可能导致错位、缺字或转换失败。
- 第一次使用前，建议用本单位的典型图纸试转几份，打开 PDF 对照检查。
- 程序未购买商业代码签名证书，Windows 可能提示“未知发布者”，可从“更多信息”选择运行。

## 交付目录

```text
Visio转PDF工具/
├─ Visio转PDF.exe             单文件成品，给领导只需复制这个
├─ README.md                  本说明
├─ 使用说明.txt               简短操作说明
├─ 许可说明/                  第三方许可与声明
├─ 测试样例/                  实际转换过的 VSD、VSDX 和输出 PDF
├─ 源码/
│  ├─ visio2pdf.py            主程序
│  ├─ tests.py                基础回归测试
│  ├─ build.py                测试、生成组件压缩包、打包 exe
│  ├─ 一键打包.bat            创建 Python 环境并调用 build.py
│  ├─ version_info.txt        exe 作者、产品名和版本元数据
│  ├─ requirements.txt        Python 依赖版本
└─ └─ runtime/                完整构建所需的精简转换组件
```

“源码”保留完整构建所需的最小集合：7 个源码/配置文件和一套 `runtime` 精简组件。`runtime.zip` 在构建时自动生成，构建完成后可删除；日常分发 exe 不需要源码或完整组件。

## 直接运行源码（开发用）

开发电脑安装 64 位 Python 3.12–3.14，并勾选加入 PATH。

```bat
cd 源码
py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python visio2pdf.py "D:\图纸\示例.vsdx"
```

源码版自动使用旁边的 `runtime` 精简组件。只有安装 Python 运行依赖即可，开发电脑同样不要求安装 Visio。当前打包验证使用 Python 3.14.9；日常维护推荐 Python 3.12 或 3.14 的 64 位版本。

## 重新打包

直接双击 `源码\一键打包.bat`，或在命令行执行：

```bat
cd 源码
.venv\Scripts\python build.py
```

脚本直接从 `runtime` 生成 `runtime.zip`，执行 6 项基础测试，再在交付目录根部生成 `Visio转PDF.exe`。不需要 `minimize.py`、离线 Python 安装包、预先存在的 `runtime.zip` 或外层完整组件。

构建产生的 `runtime.zip`、`build` 和 `__pycache__` 是临时产物，可在验证后删除。不要删除保留的 `runtime` 目录，它是完整构建的必要组成部分。

如需从官方 LibreOffice MSI 重新准备完整组件，可在 Windows 命令行使用管理员解包模式，将文件释放到指定目录，不执行普通安装：

```bat
msiexec /a "D:\下载\LibreOffice_26.8.3_Win_x86-64.msi" /qn TARGETDIR="D:\构建目录"
```

本包还随附 Microsoft Visual C++ 运行库所需的本地 DLL，避免目标电脑另外安装运行库。


## 常见问题排查

- 双击后没有转换：需要把文件拖到 exe 图标上，或在窗口中粘贴文件完整路径后按回车。
- 第一次启动较慢：正在释放内置组件；以后使用已释放的组件。请确认本机磁盘有足够空间。
- 找不到输出文件：查看窗口中以“输出位置：”开头的完整路径，PDF 不在 exe 所在目录，而是在各源文件所在目录。
- 出现 `_1`、`_2` 后缀：说明同名 PDF 已存在，这是防止覆盖旧文件的正常行为。
- 提示“源文件无法加载”：文件可能损坏、使用了内置引擎不兼容的结构，或不是有效的 Visio 图纸。可尝试在 Visio 中打开并另存为标准 VSDX 后重新转换。
- 提示 Visio 的 COM 错误：程序会继续尝试内置引擎，只有两条路径都失败才记为该文件失败。
- 安全软件提示未知发布者：此版本未进行商业代码签名；请只使用本交付目录复制出的程序，并核对下方的 SHA-256。
- 组件释放失败：确认 `%LOCALAPPDATA%` 可写，关闭正在运行的本工具后重试；不要删除正在使用的缓存目录。

## 成品

```text
文件：Visio转PDF.exe
大小：248543119 字节
SHA-256：9f7e7a40c68ed2c259783c6ad4ca23bf92f105d38c4b57bd90728823ab6e4660
```

Windows 校验命令：

```bat
certutil -hashfile "Visio转PDF.exe" SHA256
```

修改源码或重新打包后，哈希会变化，应重新生成 `校验信息.txt`。

## 参考与许可

参考了以下公开项目的 Visio 自动化思路：

- https://github.com/egalletta/visio2pdf
- https://github.com/hay-kot/Visio2PDF （MIT）

本程序没有复制 egalletta 项目的源码；其公开仓库未见许可证声明。完整转换使用 LibreOffice 及 libvisio，相关许可、源码获取地址和第三方声明见 `许可说明` 目录。
