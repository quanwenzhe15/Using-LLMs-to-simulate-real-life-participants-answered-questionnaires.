#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图标格式转换工具
将png格式转换为ico格式
"""
import os
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("❌ 错误: 缺少Pillow库，请先安装")
    print("运行: pip install pillow")
    input("按回车键退出...")
    exit()

def convert_image_to_ico(image_path, ico_path):
    """
    将图片文件转换为ico文件
    支持PNG、JPG、JPEG格式
    """
    try:
        # 打开图片
        img = Image.open(image_path)
        
        # 确保图片是RGB模式
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 调整图片大小（ICO文件通常使用16x16, 32x32, 48x48等尺寸）
        # 创建多个尺寸的图标
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128)]
        
        # 保存为ICO文件
        img.save(ico_path, format='ICO', sizes=sizes)
        
        print(f"✅ 成功将 {image_path} 转换为 {ico_path}")
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {str(e)}")
        return False

def main():
    print("=" * 60)
    print("图标格式转换工具")
    print("=" * 60)
    
    # 检查图标文件
    icons_dir = Path("icons")
    if not icons_dir.exists():
        print("❌ 错误: 找不到icons文件夹")
        input("按回车键退出...")
        return
    
    # 寻找jpg和png图标文件
    image_files = list(icons_dir.glob("*.png")) + list(icons_dir.glob("*.jpg")) + list(icons_dir.glob("*.jpeg"))
    if not image_files:
        print("❌ 错误: icons文件夹中没有图片文件")
        input("按回车键退出...")
        return
    
    print("找到的图片文件:")
    for i, image_file in enumerate(image_files):
        print(f"{i+1}. {image_file.name}")
    
    # 选择要转换的文件
    choice = input("请选择要转换的文件编号 (默认 1): ")
    if not choice:
        choice = "1"
    
    try:
        selected_idx = int(choice) - 1
        if 0 <= selected_idx < len(image_files):
            selected_image = image_files[selected_idx]
        else:
            print("❌ 无效的选择")
            input("按回车键退出...")
            return
    except ValueError:
        print("❌ 请输入数字")
        input("按回车键退出...")
        return
    
    # 创建目标ico文件路径
    ico_file = icons_dir / (selected_image.stem + ".ico")
    
    print(f"\n正在转换: {selected_image.name} -> {ico_file.name}")
    
    if convert_image_to_ico(selected_image, ico_file):
        print("\n🎉 转换完成！")
        print(f"您现在可以在打包时使用 {ico_file.name} 作为图标")
    else:
        print("\n❌ 转换失败，请检查错误信息")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()