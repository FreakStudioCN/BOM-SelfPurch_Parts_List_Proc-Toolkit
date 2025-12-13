# Python env   : Python 3.8+（需支持 pathlib、f-string 及 pandas/openpyxl 最新 API）
# -*- coding: utf-8 -*-
# @Time    : 2025/12/13 下午6:40
# @Author  : 李清水
# @File    : bom_component_search.py
# @Description : 搜索同级文件夹下BOM_开头的CSV/Excel文件，筛选包含指定关键词的文件，返回匹配结果并汇总，同时处理文件读取异常

import os
import pandas as pd

def search_specific_content_in_bom_files(search_keyword):
    """
    搜索同级文件夹下BOM_开头的文件中包含指定关键词的文件
    :param search_keyword: 要搜索的关键词（如RES-ADJ-TH_3362P）
    :return: 符合条件的文件名列表
    """
    # 获取当前运行代码的文件夹路径
    current_dir = os.getcwd()
    # 存储符合条件的文件名
    target_files = []

    # 遍历当前文件夹下的所有同级子文件夹
    for folder_name in os.listdir(current_dir):
        folder_path = os.path.join(current_dir, folder_name)
        # 只处理文件夹，跳过文件和隐藏目录（可选，避免处理.开头的系统目录）
        if os.path.isdir(folder_path) and not folder_name.startswith('.'):
            # 遍历子文件夹内的文件
            for file_name in os.listdir(folder_path):
                # 筛选“BOM_开头 + .csv/.xlsx后缀”的文件，覆盖常见的BOM文件格式
                if file_name.startswith("BOM_") and (file_name.endswith(".csv") or file_name.endswith(".xlsx")):
                    file_full_path = os.path.join(folder_path, file_name)
                    try:
                        # 根据文件后缀选择对应的读取方法
                        if file_name.endswith(".csv"):
                            # 读取CSV文件，不指定表头，全量读取（避免表头干扰关键词搜索）
                            df = pd.read_csv(file_full_path, header=None, encoding='utf-8', errors='ignore')
                        else:  # .xlsx
                            df = pd.read_excel(file_full_path, header=None)

                        # 检查文件中是否存在包含指定关键词的单元格
                        # 先将所有单元格转为字符串，再检查是否包含关键词
                        has_keyword = df.apply(
                            lambda row: row.astype(str).str.contains(search_keyword, na=False).any(),
                            axis=1
                        ).any()

                        if has_keyword:
                            # 可以选择存储完整路径或仅文件名，这里保留文件名，也可改为file_full_path
                            target_files.append(file_name)
                            print(f"找到匹配文件：{file_name}（路径：{file_full_path}）")

                    except Exception as e:
                        print(f"读取文件{file_full_path}失败：{str(e)}")
                        continue

    return target_files

if __name__ == "__main__":
    # 定义要搜索的关键词：RES-ADJ-TH_3362P
    SEARCH_KEY = "RES-ADJ-TH_3362P"
    # 执行搜索
    result_files = search_specific_content_in_bom_files(SEARCH_KEY)

    # 输出最终结果汇总
    print("\n===== 搜索结果汇总 =====")
    if result_files:
        print(f"共找到{len(result_files)}个包含'{SEARCH_KEY}'的文件：")
        for idx, file in enumerate(result_files, 1):
            print(f"{idx}. {file}")
    else:
        print(f"未找到包含'{SEARCH_KEY}'的BOM文件")