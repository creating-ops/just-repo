#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT转PPTX转换脚本：将.ppt格式无损转换为.pptx格式
支持多种转换方法
"""

import os
import subprocess
import shutil
from pathlib import Path

# ===================== 转换方法 =====================

def convert_with_libreoffice(ppt_file, output_dir=None):
    """
    使用LibreOffice进行转换（推荐）
    LibreOffice可以无损转换PPT到PPTX

    需要安装LibreOffice: https://www.libreoffice.org/
    """
    ppt_path = Path(ppt_file).resolve()
    if output_dir is None:
        output_dir = ppt_path.parent
    else:
        output_dir = Path(output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    # 查找LibreOffice路径
    libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]

    soffice = None
    for path in libreoffice_paths:
        if os.path.exists(path):
            soffice = path
            break

    if not soffice:
        # 尝试从环境变量或PATH查找
        soffice = shutil.which("soffice") or shutil.which("libreoffice")

    if not soffice:
        print("❌ 未找到LibreOffice，请先安装:")
        print("   下载地址: https://www.libreoffice.org/download/")
        return None

    print(f"使用LibreOffice: {soffice}")
    print(f"转换文件: {ppt_path}")

    # 执行转换命令
    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pptx",
        "--outdir", str(output_dir),
        str(ppt_path)
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        # 检查输出文件
        pptx_name = ppt_path.stem + ".pptx"
        pptx_path = output_dir / pptx_name

        if pptx_path.exists():
            print(f"✅ 转换成功: {pptx_path}")
            return str(pptx_path)
        else:
            print(f"❌ 转换失败，输出文件不存在")
            print(f"   stdout: {result.stdout}")
            print(f"   stderr: {result.stderr}")
            return None

    except subprocess.TimeoutExpired:
        print("❌ 转换超时")
        return None
    except Exception as e:
        print(f"❌ 转换出错: {e}")
        return None


def convert_with_com(ppt_file, output_dir=None):
    """
    使用Microsoft PowerPoint COM自动化转换
    需要安装Microsoft Office

    仅Windows系统可用
    """
    import sys
    if sys.platform != 'win32':
        print("❌ COM方法仅支持Windows系统")
        return None

    try:
        import win32com.client
    except ImportError:
        print("❌ 需要安装pywin32: pip install pywin32")
        return None

    ppt_path = Path(ppt_file).resolve()
    if output_dir is None:
        output_dir = ppt_path.parent
    else:
        output_dir = Path(output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    pptx_name = ppt_path.stem + ".pptx"
    pptx_path = output_dir / pptx_name

    print(f"使用Microsoft PowerPoint COM")
    print(f"转换文件: {ppt_path}")

    try:
        # 创建PowerPoint应用对象
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # 必须显示窗口

        # 打开PPT文件
        deck = powerpoint.Presentations.Open(str(ppt_path))

        # 保存为PPTX格式 (24 = pptx格式)
        deck.SaveAs(str(pptx_path), 24)
        deck.Close()
        powerpoint.Quit()

        print(f"✅ 转换成功: {pptx_path}")
        return str(pptx_path)

    except Exception as e:
        print(f"❌ COM转换出错: {e}")
        # 确保清理
        try:
            powerpoint.Quit()
        except:
            pass
        return None


def convert_ppt_to_pptx(ppt_file, output_dir=None, method='libreoffice'):
    """
    主转换函数

    参数:
        ppt_file: .ppt文件路径
        output_dir: 输出目录（默认与原文件同目录）
        method: 转换方法 ('libreoffice' 或 'com')

    返回:
        转换后的.pptx文件路径，失败返回None
    """
    ppt_path = Path(ppt_file)

    # 检查文件存在
    if not ppt_path.exists():
        print(f"❌ 文件不存在: {ppt_file}")
        return None

    # 检查文件扩展名
    if ppt_path.suffix.lower() != '.ppt':
        print(f"❌ 不是.ppt文件: {ppt_file}")
        return None

    print("=" * 50)
    print(f"PPT → PPTX 转换")
    print("=" * 50)

    if method == 'libreoffice':
        return convert_with_libreoffice(ppt_file, output_dir)
    elif method == 'com':
        return convert_with_com(ppt_file, output_dir)
    else:
        print(f"❌ 未知方法: {method}")
        print("   支持的方法: 'libreoffice', 'com'")
        return None


# ===================== 批量转换 =====================

def batch_convert(input_dir, output_dir=None, method='libreoffice'):
    """
    批量转换目录下所有.ppt文件
    """
    input_path = Path(input_dir)

    if output_dir is None:
        output_dir = input_path
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    # 查找所有.ppt文件
    ppt_files = list(input_path.glob("*.ppt"))

    if not ppt_files:
        print(f"未找到.ppt文件: {input_dir}")
        return []

    print(f"找到 {len(ppt_files)} 个.ppt文件")

    results = []
    for ppt_file in ppt_files:
        print(f"\n处理: {ppt_file.name}")
        pptx_path = convert_ppt_to_pptx(ppt_file, output_dir, method)
        if pptx_path:
            results.append((str(ppt_file), pptx_path))

    print(f"\n转换完成: {len(results)}/{len(ppt_files)}")
    return results


# ===================== 主函数 =====================

if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("用法:")
        print("  单文件转换: python convert_ppt_to_pptx.py <ppt文件路径> [方法]")
        print("  批量转换:   python convert_ppt_to_pptx.py <目录路径> --batch")
        print("")
        print("方法选项:")
        print("  libreoffice - 使用LibreOffice（需安装）")
        print("  com - 使用Microsoft Office COM（Windows自带）")
        print("")
        print("示例:")
        print("  python convert_ppt_to_pptx.py 2026货币金银重点工作思路4.2.ppt com")
        print("  python convert_ppt_to_pptx.py . --batch")
        sys.exit(1)

    target = sys.argv[1]
    is_batch = '--batch' in sys.argv

    # 确定方法
    method = 'com'  # Windows默认使用COM
    if len(sys.argv) > 2 and sys.argv[2] in ['libreoffice', 'com']:
        method = sys.argv[2]

    if is_batch:
        batch_convert(target, method=method)
    else:
        convert_ppt_to_pptx(target, method=method)