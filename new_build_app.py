#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EasyPsych 应用打包程序

此脚本使用 PyInstaller 将 EasyPsych_source_code.py 打包为可执行文件，并嵌入图标。
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# 设置路径
BASE_DIR = Path(__file__).parent
SCRIPT_PATH = BASE_DIR / "EasyPsych_source_code.py"
CONFIG_PATH = BASE_DIR / "config.py"
LANGUAGE_CONFIG_PATH = BASE_DIR / "language_config.py"
ICONS_DIR = BASE_DIR / "icons"
OUTPUT_DIR = BASE_DIR / "dist"
BUILD_DIR = BASE_DIR / "build"

# 图标文件路径
ICON_FILE = ICONS_DIR / "EasyPsych.ico"

# 检查必要的依赖
def check_dependencies():
    """检查打包所需的依赖"""
    print("正在检查打包依赖...")
    
    # 声明全局变量
    global ICON_FILE
    
    # 检查 PyInstaller
    try:
        import PyInstaller
        print("✓ PyInstaller 已安装")
    except ImportError:
        print("⚠️ PyInstaller 未安装，正在安装...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            print("✓ PyInstaller 安装成功")
        except subprocess.CalledProcessError:
            print("✗ 无法安装 PyInstaller，请手动安装后重试")
            return False
    
    # 检查脚本文件
    if not SCRIPT_PATH.exists():
        print(f"✗ 脚本文件不存在: {SCRIPT_PATH}")
        return False
    else:
        print(f"✓ 脚本文件存在: {SCRIPT_PATH}")
    
    # 检查配置文件
    if not CONFIG_PATH.exists():
        print(f"✗ 配置文件不存在: {CONFIG_PATH}")
        return False
    else:
        print(f"✓ 配置文件存在: {CONFIG_PATH}")
    
    # 检查多语言配置文件
    if not LANGUAGE_CONFIG_PATH.exists():
        print(f"✗ 多语言配置文件不存在: {LANGUAGE_CONFIG_PATH}")
        return False
    else:
        print(f"✓ 多语言配置文件存在: {LANGUAGE_CONFIG_PATH}")
    
    # 检查图标文件
    if not ICON_FILE.exists():
        print(f"⚠️ 图标文件不存在: {ICON_FILE}")
        # 尝试查找其他图标文件
        icon_files = list(ICONS_DIR.glob("*.ico"))
        if icon_files:
            ICON_FILE = icon_files[0]
            print(f"✓ 使用备用图标: {ICON_FILE}")
        else:
            print("⚠️ 未找到任何 .ico 格式图标，将使用默认图标")
    else:
        print(f"✓ 图标文件存在: {ICON_FILE}")
    
    return True

# 清理旧的构建文件
def clean_old_builds():
    """清理旧的构建文件"""
    print("正在清理旧的构建文件...")
    
    if OUTPUT_DIR.exists():
        try:
            shutil.rmtree(OUTPUT_DIR)
            print(f"✓ 清理旧的输出目录: {OUTPUT_DIR}")
        except Exception as e:
            print(f"⚠️ 清理输出目录时出错: {e}")
    
    if BUILD_DIR.exists():
        try:
            shutil.rmtree(BUILD_DIR)
            print(f"✓ 清理旧的构建目录: {BUILD_DIR}")
        except Exception as e:
            print(f"⚠️ 清理构建目录时出错: {e}")

# 打包应用
def build_app():
    """使用 PyInstaller 打包应用"""
    print("\n开始打包应用...")
    
    # 构建 PyInstaller 命令
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--onefile",  # 生成单个可执行文件
        "--windowed",  # 无控制台窗口
        "--name", "EasyPsych",  # 应用名称
        "--add-data", f"{CONFIG_PATH}{os.pathsep}.",  # 添加配置文件到根目录
        "--add-data", f"{LANGUAGE_CONFIG_PATH}{os.pathsep}.",  # 添加多语言配置文件到根目录
        "--add-data", f"{ICONS_DIR}{os.pathsep}icons",  # 添加图标文件夹
        "--distpath", str(OUTPUT_DIR),  # 输出目录
        "--workpath", str(BUILD_DIR),  # 工作目录
    ]
    
    # 添加图标（如果存在）
    if ICON_FILE.exists():
        cmd.extend(["--icon", str(ICON_FILE)])
    
    # 添加主脚本
    cmd.append(str(SCRIPT_PATH))
    
    # 打印命令
    print("执行命令:")
    print(' '.join(cmd))
    print()
    
    # 执行打包命令
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)
        
        # 打印输出
        if result.stdout:
            print("\n打包输出:")
            print(result.stdout)
        
        # 打印错误
        if result.stderr:
            print("\n打包错误:")
            print(result.stderr)
        
        # 检查结果
        if result.returncode == 0:
            print("\n✅ 应用打包成功！")
            
            # 查找生成的可执行文件
            exe_file = OUTPUT_DIR / "EasyPsych.exe"
            if exe_file.exists():
                print(f"\n可执行文件位置: {exe_file}")
                print("\n使用方法:")
                print(f"1. 双击运行: {exe_file}")
                print("2. 或在命令行中执行: {exe_file}")
            else:
                print("\n⚠️ 未找到生成的可执行文件，请检查输出目录")
        else:
            print(f"\n✗ 打包失败，返回代码: {result.returncode}")
            return False
            
    except Exception as e:
        print(f"\n✗ 打包过程中发生错误: {e}")
        return False
    
    return True

# 主函数
def main():
    """主函数"""
    print("=======================================")
    print("EasyPsych 应用打包程序")
    print("=======================================")
    
    # 检查依赖
    if not check_dependencies():
        print("\n✗ 依赖检查失败，无法继续")
        return 1
    
    # 清理旧的构建文件
    clean_old_builds()
    
    # 打包应用
    if not build_app():
        print("\n✗ 应用打包失败")
        return 1
    
    print("\n=======================================")
    print("打包过程完成！")
    print("=======================================")
    return 0

if __name__ == "__main__":
    sys.exit(main())
