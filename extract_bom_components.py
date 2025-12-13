# Python env   : Python 3.8+（需支持 pathlib、f-string 及 pandas/openpyxl 最新 API）
# -*- coding: utf-8 -*-
# @Time    : 2025/12/13 下午6:35
# @Author  : 李清水
# @File    : extract_bom_components.py
# @Description : 处理BOM文件（命名格式：BOM_模块名-v版本号.xlsx/xls），提取需自行采购的元器件数据
#                核心功能：筛选自采数据→按模块汇总并计算总价→生成带样式的Excel表（模块汇总表+去重类型表）→输出统计信息

import os
import re
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Side, Border

def extract_and_format_bom():
    root_dir = Path(os.getcwd())
    bom_file_pattern = re.compile(r'^BOM_.+-v\d+\.\d+(\.\d+)?\.(xlsx|xls)$', re.IGNORECASE)

    # 核心列（匹配你的BOM）
    core_columns = [
        'Manufacturer Part', 'Quantity', 'Designator', 'Supplier Part',
        'LCSC Price', 'Value', '淘宝链接', '下单配置', '最小起订量'
    ]
    unique_type_cols = ['Manufacturer Part', 'Supplier Part', 'LCSC Price', '淘宝链接']
    filter_columns = ['淘宝链接', '下单配置', '最小起订量']

    # 样式配置：新增细边框
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    colors = {
        'module_table': [
            PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid"),
            PatternFill(start_color="F0F8E6", end_color="F0F8E6", fill_type="solid"),
            PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")
        ],
        'type_table': [
            PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"),
            PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        ]
    }

    # 提取数据
    all_self_purchase = []
    print("🔍 搜索BOM文件...")
    for dir_path, _, file_names in os.walk(root_dir):
        for file_name in file_names:
            if bom_file_pattern.match(file_name):
                bom_path = Path(dir_path) / file_name
                print(f"📄 找到：{bom_path}")
                try:
                    df = pd.read_excel(bom_path, dtype=str, header=0)
                    df.columns = df.columns.str.strip()

                    missing_cols = [col for col in core_columns if col not in df.columns]
                    if missing_cols:
                        print(f"❌ 跳过{file_name}：缺少列{missing_cols}\n")
                        continue

                    # 筛选自行采购数据
                    df_filtered = df.copy()
                    mask = pd.Series(False, index=df_filtered.index)
                    for col in filter_columns:
                        col_mask = df_filtered[col].notna() & (df_filtered[col].str.strip() != '')
                        mask = mask | col_mask
                    df_filtered = df_filtered[mask]

                    if df_filtered.empty:
                        print(f"ℹ️ {file_name}无自采数据\n")
                        continue

                    # 处理数值列
                    df_calc = df_filtered[core_columns].copy()
                    for col in ['Quantity', 'LCSC Price', 'Value']:
                        df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').fillna(0)

                    # 提取模块名
                    module_name = re.sub(r'^BOM_(.+)-v\d+\.\d+(\.\d+)?\.(xlsx|xls)$', r'\1', file_name, re.IGNORECASE)
                    df_with_module = df_calc.copy()
                    df_with_module.insert(0, '模块名称', module_name)
                    all_self_purchase.append(df_with_module)
                    print(f"✅ 提取{file_name}：{len(df_with_module)}个器件\n")

                except Exception as e:
                    print(f"❌ 处理{file_name}失败：{str(e)}\n")
                    continue

    if not all_self_purchase:
        print("⚠️ 无自采数据")
        return
    df_with_module_all = pd.concat(all_self_purchase, ignore_index=True).sort_values(by='模块名称')

    # 计算“自采元器件总价”
    module_total = df_with_module_all.groupby('模块名称')['Value'].sum().reset_index()
    module_total.rename(columns={'Value': '自采元器件总价'}, inplace=True)  # 列名修改
    df_with_module_all = pd.merge(df_with_module_all, module_total, on='模块名称', how='left')
    cols = df_with_module_all.columns.tolist()
    cols.insert(1, cols.pop(cols.index('自采元器件总价')))
    df_with_module_all = df_with_module_all[cols]

    # 生成文件1：按模块汇总（带边框+新列名）
    file1_path = root_dir / "1_按模块汇总_自采元器件.xlsx"
    with pd.ExcelWriter(file1_path, engine='openpyxl') as writer:
        df_with_module_all.to_excel(writer, sheet_name='按模块汇总', index=False)

    wb1 = load_workbook(file1_path)
    ws1 = wb1['按模块汇总']
    max_row1, max_col1 = ws1.max_row, ws1.max_column

    # 合并单元格（模块名+自采元器件总价）
    print("📊 合并单元格...")
    module_ranges = []
    if max_row1 > 1:
        current_module = ws1['A2'].value
        start_row = 2
        for row in range(3, max_row1 + 1):
            if ws1[f'A{row}'].value != current_module:
                ws1.merge_cells(f'A{start_row}:A{row - 1}')
                ws1.merge_cells(f'B{start_row}:B{row - 1}')  # 第2列是新列名
                module_ranges.append((start_row, row - 1))
                current_module = ws1[f'A{row}'].value
                start_row = row
        ws1.merge_cells(f'A{start_row}:A{max_row1}')
        ws1.merge_cells(f'B{start_row}:B{max_row1}')
        module_ranges.append((start_row, max_row1))

    # 设置背景色+边框+居中
    print("🎨 设置样式...")
    color_idx = 0
    for (start_row, end_row) in module_ranges:
        current_color = colors['module_table'][color_idx % len(colors['module_table'])]
        for row in range(start_row, end_row + 1):
            for col in range(1, max_col1 + 1):
                ws1.cell(row=row, column=col).fill = current_color
                ws1.cell(row=row, column=col).border = thin_border  # 添加边框
                ws1.cell(row=row, column=col).alignment = Alignment(horizontal='center', vertical='center')
        color_idx += 1

    # 表头样式（补全边框+居中）
    for col in range(1, max_col1 + 1):
        ws1.cell(row=1, column=col).border = thin_border
        ws1.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')

    # 自适应列宽
    for col in range(1, max_col1 + 1):
        max_width = 0
        for row in range(1, max_row1 + 1):
            cell_val = str(ws1.cell(row=row, column=col).value or "")
            max_width = max(max_width, sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in cell_val))
        ws1.column_dimensions[ws1.cell(row=1, column=col).column_letter].width = max_width * 0.9

    wb1.save(file1_path)
    print(f"✅ 文件1生成：{file1_path}\n")

    # 生成文件2：去重类型（带边框）
    df_type_unique = df_with_module_all[core_columns].drop_duplicates(subset=unique_type_cols,
                                                                      keep='first').reset_index(drop=True)
    file2_path = root_dir / "2_去重_自采元器件类型.xlsx"
    with pd.ExcelWriter(file2_path, engine='openpyxl') as writer:
        df_type_unique.to_excel(writer, sheet_name='类型汇总', index=False)

    wb2 = load_workbook(file2_path)
    ws2 = wb2['类型汇总']
    max_row2, max_col2 = ws2.max_row, ws2.max_column

    # 奇偶行背景色+边框+居中
    for row in range(2, max_row2 + 1):
        color_idx = 0 if row % 2 == 0 else 1
        current_color = colors['type_table'][color_idx]
        for col in range(1, max_col2 + 1):
            ws2.cell(row=row, column=col).fill = current_color
            ws2.cell(row=row, column=col).border = thin_border
            ws2.cell(row=row, column=col).alignment = Alignment(horizontal='center', vertical='center')

    # 表头样式
    for col in range(1, max_col2 + 1):
        ws2.cell(row=1, column=col).border = thin_border
        ws2.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')

    # 自适应列宽
    for col in range(1, max_col2 + 1):
        max_width = 0
        for row in range(1, max_row2 + 1):
            cell_val = str(ws2.cell(row=row, column=col).value or "")
            max_width = max(max_width, sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in cell_val))
        ws2.column_dimensions[ws2.cell(row=1, column=col).column_letter].width = max_width * 0.9

    wb2.save(file2_path)
    print(f"✅ 文件2生成：{file2_path}\n")

    # 统计
    total = df_with_module_all.drop_duplicates(subset=['模块名称'])['自采元器件总价'].sum()
    print("📋 统计：")
    print(f"   - 自采器件数：{len(df_with_module_all)}个")
    print(f"   - 自采总金额：{total:.4f}元")
    print(f"   - 去重类型数：{len(df_type_unique)}种")


if __name__ == "__main__":
    extract_and_format_bom()