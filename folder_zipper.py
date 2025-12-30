# Python env   : Python 3.6+
# -*- coding: utf-8 -*-
# @Time    : 2025/12/30 下午4:07
# @Author  : 李清水
# @File    : folder_zipper.py
# @Description : 批量压缩当前目录下的所有直接子文件夹，生成与文件夹同名的ZIP压缩包；
#               保留原文件夹的目录结构，包含完善的异常处理：跳过已存在的压缩包、
#               处理空文件夹、压缩失败时自动清理无效的空压缩包，兼容Windows/macOS/Linux系统。

import os
import zipfile

def batch_compress_folders():
    """
    批量压缩当前目录下的所有直接子文件夹，压缩包名称与文件夹名称一致
    """
    # 获取当前工作目录
    current_dir = os.getcwd()
    print(f"当前操作目录：{current_dir}")

    # 遍历当前目录下的所有条目
    for item in os.listdir(current_dir):
        # 拼接完整路径
        item_path = os.path.join(current_dir, item)

        # 仅处理直接子文件夹（排除文件、隐藏文件夹）
        if os.path.isdir(item_path) and not item.startswith('.'):
            # 压缩包名称：文件夹名 + .zip（与文件夹同名）
            zip_filename = f"{item}.zip"
            zip_filepath = os.path.join(current_dir, zip_filename)

            # 检查压缩包是否已存在，避免覆盖
            if os.path.exists(zip_filepath):
                print(f"⚠️  压缩包 {zip_filename} 已存在，跳过压缩")
                continue

            # 处理空文件夹
            if not os.listdir(item_path):
                print(f"⚠️  文件夹 {item} 为空，创建空压缩包")

            try:
                # 创建ZIP压缩包（'w'表示写入，压缩级别6为平衡压缩率和速度）
                with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
                    # 递归遍历文件夹内的所有文件和子文件夹
                    for root, dirs, files in os.walk(item_path):
                        # 对每个文件进行处理
                        for file in files:
                            # 文件完整路径
                            file_path = os.path.join(root, file)
                            # 压缩包内的相对路径（保持原文件夹结构）
                            arcname = os.path.relpath(file_path, current_dir)
                            # 将文件添加到压缩包
                            zipf.write(file_path, arcname=arcname)

                print(f"✅  成功压缩：{item} → {zip_filename}")

            except Exception as e:
                print(f"❌  压缩 {item} 失败：{str(e)}")
                # 若压缩失败，删除已创建的空压缩包
                if os.path.exists(zip_filepath):
                    os.remove(zip_filepath)

if __name__ == "__main__":
    batch_compress_folders()
    print("\n批量压缩操作完成！")