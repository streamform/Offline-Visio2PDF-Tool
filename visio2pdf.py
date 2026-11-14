import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from pypdf import PdfReader

VERSION = "1.0.0"
AUTHOR = "乐乐"
PRODUCT_NAME = "Visio 转 PDF 工具"
RUNTIME_VERSION = "libreoffice-26.8.3-v1"
SUPPORTED_SUFFIXES = {".vsd", ".vsdx"}


def log(message):
    print(f"[{time.strftime('%H:%M:%S')}] {message}", flush=True)


def configure_console():
    if os.name == "nt":
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def collect_inputs(arguments):
    sources = []
    for argument in arguments:
        source = Path(argument.strip().strip('"')).expanduser().resolve()
        if not source.is_file():
            raise ValueError(f"文件不存在：{source}")
        if source.suffix.lower() not in SUPPORTED_SUFFIXES:
            raise ValueError(f"仅支持 VSD、VSDX 文件：{source}")
        if source not in sources:
            sources.append(source)
    return sources


def output_path(source):
    candidate = source.with_suffix(".pdf")
    number = 1
    while candidate.exists():
        candidate = source.with_name(f"{source.stem}_{number}.pdf")
        number += 1
    return candidate


def validate_pdf(target):
    if not target.is_file() or target.stat().st_size < 100:
        raise ValueError("PDF 未生成或文件为空")
    try:
        with target.open("rb") as stream:
            reader = PdfReader(stream, strict=True)
            if len(reader.pages) == 0:
                raise ValueError("PDF 没有页面")
    except ValueError:
        raise
    except Exception as error:
        raise ValueError(f"PDF 校验失败：{error}") from error


def visio_export(source, target):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        app = win32com.client.DispatchEx("Visio.Application")
        app.Visible = False
        app.AlertResponse = 1
        flags = 2 | 64 | 128 | 256
        document = app.Documents.OpenEx(str(source), flags)
        document.ExportAsFixedFormat(1, str(target), 1, 0)
    finally:
        if document is not None:
            document.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def runtime_root(base):
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return Path(local) / "VisioToPDF" / RUNTIME_VERSION
    return base / "runtime"


def prepare_runtime(base):
    if not getattr(sys, "frozen", False):
        local_runtime = base / "runtime" / "program" / "soffice.com"
        if local_runtime.is_file():
            return local_runtime
    destination = runtime_root(base)
    soffice = destination / "program" / "soffice.com"
    if soffice.is_file():
        return soffice
    package = Path(sys._MEIPASS) / "runtime.zip"
    log("首次运行，正在释放内置转换组件，请稍候……")
    destination.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix="prepare-", dir=destination.parent))
    try:
        with zipfile.ZipFile(package) as archive:
            archive.extractall(staging)
        extracted = staging / "runtime"
        if not (extracted / "program" / "soffice.com").is_file():
            raise RuntimeError("内置转换组件不完整，请重新获取程序")
        if destination.exists():
            shutil.rmtree(destination)
        extracted.replace(destination)
        shutil.rmtree(staging)
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise
    log("组件准备完成，后续转换无需再次释放")
    return destination / "program" / "soffice.com"


def portable_export(source, target, base):
    soffice = prepare_runtime(base)
    work = Path(tempfile.mkdtemp(prefix="visio2pdf-", dir=base))
    try:
        copied_source = work / ("input" + source.suffix.lower())
        shutil.copy2(source, copied_source)
        profile = work / "profile"
        command = [
            str(soffice),
            "--headless",
            "--invisible",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--norestore",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to",
            "pdf:draw_pdf_Export",
            "--outdir",
            str(work),
            str(copied_source),
        ]
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        try:
            output = process.communicate(timeout=300)[0]
        except subprocess.TimeoutExpired:
            process.kill()
            process.communicate()
            raise RuntimeError("内置引擎转换超时（5 分钟）")
        message = output.decode("utf-8", errors="replace").strip()
        if process.returncode != 0:
            raise RuntimeError(f"内置引擎无法读取此文件：{message}")
        generated = work / "input.pdf"
        if not generated.is_file():
            raise RuntimeError(f"内置引擎未生成 PDF：{message}")
        shutil.move(str(generated), str(target))
    finally:
        shutil.rmtree(work, ignore_errors=True)


def convert(source, target, base):
    visio_error = None
    try:
        log("尝试使用 Microsoft Visio 导出")
        visio_export(source, target)
        return "Microsoft Visio"
    except Exception as error:
        visio_error = error
        log("Visio 不可用，改用内置转换引擎")
    try:
        portable_export(source, target, base)
        return "内置转换引擎"
    except Exception as error:
        raise RuntimeError(
            f"Visio：{visio_error}\n        内置引擎：{error}"
        ) from error


def parse_line(line):
    match = re.match(r'^\s*(?:"([^"]+)"|(\S+))\s*(.*)$', line)
    if not match:
        return []
    path = match.group(1) or match.group(2)
    remainder = match.group(3)
    return [path] + (parse_line(remainder) if remainder else [])


def print_help():
    print(PRODUCT_NAME + "  v" + VERSION)
    print("作者：" + AUTHOR)
    print("支持 VSD、VSDX；Windows 10/11 64 位；无需安装 Visio、Python 或 LibreOffice。")
    print()
    print("使用方法：")
    print("  1. 把一个或多个 Visio 文件拖到本程序图标上。")
    print("  2. 命令行：Visio转PDF.exe \"D:\\图纸\\示例.vsdx\"")
    print("  3. 双击程序后，把文件路径粘贴到窗口，按回车。")
    print()
    print("PDF 输出到源文件所在文件夹。已有 PDF 不覆盖，自动命名为“名称_1.pdf”。")
    print("转换在本机完成，不上传文件。复杂图纸建议抽查版式。")


def pause_when_interactive():
    if sys.stdin.isatty():
        try:
            input("\n按回车键关闭窗口……")
        except EOFError:
            pass


def main(arguments=None):
    args = sys.argv[1:] if arguments is None else arguments
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("-h", "--help", action="store_true")
    parser.add_argument("--version", action="store_true")
    parser.add_argument("files", nargs="*")
    options = parser.parse_args(args)
    if options.help:
        print_help()
        return 0
    if options.version:
        print(VERSION)
        return 0
    interactive = not options.files
    if interactive:
        print_help()
        try:
            options.files = parse_line(input("\n请拖入文件或粘贴路径，然后按回车：\n"))
        except EOFError:
            return 1
    succeeded = 0
    failed = 0
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    try:
        sources = collect_inputs(options.files)
    except ValueError as error:
        log(f"错误：{error}")
        if interactive or len(sys.argv) > 1:
            pause_when_interactive()
        return 2
    log(f"共 {len(sources)} 个文件，开始转换")
    for index, source in enumerate(sources, 1):
        log(f"[{index}/{len(sources)}] {source}")
        target = output_path(source)
        try:
            engine = convert(source, target, base)
            validate_pdf(target)
            succeeded += 1
            log(f"转换成功（{engine}）")
            log(f"输出位置：{target}")
        except Exception as error:
            failed += 1
            if target.exists():
                target.unlink()
            log(f"转换失败：{error}")
    log(f"处理完成：成功 {succeeded} 个，失败 {failed} 个")
    if interactive or len(sys.argv) > 1:
        pause_when_interactive()
    return 0 if failed == 0 and succeeded > 0 else 1


if __name__ == "__main__":
    try:
        configure_console()
        sys.exit(main())
    except Exception as error:
        log(f"程序错误：{error}")
        pause_when_interactive()
        sys.exit(1)
