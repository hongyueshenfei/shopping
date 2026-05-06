#!/usr/bin/env python3
"""
Fix Enrich - 批量修复补全文件中问题行的脚本

扫描已有的 _补全.xlsx 文件，根据预定义规则自动识别有问题的行，
重新调用 enrich_product() 进行补全，覆盖原数据。

检查规则（每次运行会根据实际情况更新）:
  1. 铁线莲三级必须是「木本藤本」（不是草本藤本）
  2. 绣球四级必须是「绣球花科」（不是绣球科）
  3. 枫树四级必须是「槭树科」（不是槒树科）
  4. 百子莲三级必须是「鳞茎」（不是多年生草本）
  5. 品牌HY/HY-/Encore Azaleas → 统一为「安酷」
  6. 一级~四级分类/品牌完全缺失 → 重新补全

使用方法:
    python3 fix_enrich.py

说明:
    - 仅处理问题行，非问题行不会被重新调用
    - 每10行保存一次进度
    - 无问题行时直接退出
"""

import sys
import os
import json
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from product_enricher import load_categories, enrich_product

import openpyxl


# ===========================================================================
# 问题行检查规则
# ===========================================================================

def check_row(row_idx, name, yiji, erji, sanji, siji, brand, variety):
    """
    检查指定行的分类数据是否有已知问题。

    参数:
        row_idx ~ variety: Excel中各列的原始值

    返回:
        (is_problematic: bool, reason: str)
        - True + 说明文字 表示需要重新补全
        - False + '' 表示数据正常
    """
    name = name or ''

    # ---- 规则1: 铁线莲三级分类 ----
    # 铁线莲是木本藤本，不是草本藤本
    if '铁线莲' in name and sanji == '草本藤本':
        return True, '铁线莲三级分类错误（草本藤本→木本藤本）'

    # ---- 规则2: 绣球四级分类 ----
    # 绣球的科名是「绣球花科」，不是「绣球科」
    if siji in ('绣球科',):
        return True, f'绣球四级分类错误（{siji}→绣球花科）'

    # ---- 规则3: 枫树四级分类 ----
    # 枫树的科名是「槭树科」，不是「槒树科」
    if siji in ('槒树科',):
        return True, f'枫树四级分类错误（{siji}→槭树科）'

    # ---- 规则4: 百子莲三级分类 ----
    # 百子莲是鳞茎植物，不是多年生草本
    if '百子莲' in name and sanji == '多年生草本':
        return True, '百子莲三级分类错误（多年生草本→鳞茎）'

    # ---- 规则5: 安酷杜鹃品牌名称归一 ----
    # Encore Azaleas / HY / HY- 统一输出为「安酷」
    if brand and (brand.strip() in ('HY', 'Encore Azaleas（安酷杜鹃）', 'Encore Azaleas') or brand.strip().startswith('HY')):
        return True, f'品牌写法错误（{brand}→安酷）'

    # ---- 规则6: 数据完全缺失 ----
    # 一级分类为空说明该行从未成功补全
    if yiji is None:
        return True, '数据完全缺失'

    # 无问题
    return False, ''


# ===========================================================================
# 主流程
# ===========================================================================

def main():
    # ---- 1. 定位补全文件 ----
    excel_path = Path(__file__).parent / '1年内销售库存大于10育种+上海_补全.xlsx'
    if not excel_path.exists():
        print(f"文件不存在: {excel_path}")
        sys.exit(1)

    print(f"读取文件: {excel_path}")
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    total_rows = ws.max_row - 1
    print(f"总行数: {total_rows}\n")

    # ---- 2. 加载分类目录 ----
    categories = None
    for parent in [excel_path.parent, Path(__file__).parent]:
        xlsx = parent / '商品分类目录.xlsx'
        if xlsx.exists():
            print(f"加载分类目录: {xlsx}")
            categories = load_categories(str(xlsx))
            print(f"共 {len(categories)} 条分类记录\n")
            break

    if not categories:
        print("错误: 找不到商品分类目录.xlsx")
        sys.exit(1)

    # ---- 3. 全量扫描，收集问题行 ----
    print("=== 检查所有行 ===")
    problematic = []
    for row_idx in range(2, ws.max_row + 1):
        goods_name = ws.cell(row_idx, 2).value
        yiji  = ws.cell(row_idx, 8).value   # 一级分类
        erji  = ws.cell(row_idx, 9).value   # 二级分类
        sanji = ws.cell(row_idx, 10).value  # 三级分类
        siji  = ws.cell(row_idx, 11).value  # 四级分类
        brand = ws.cell(row_idx, 12).value  # 品牌
        variety = ws.cell(row_idx, 14).value  # 品种

        is_bad, reason = check_row(row_idx, goods_name, yiji, erji, sanji, siji, brand, variety)
        if is_bad:
            problematic.append({
                'row': row_idx,
                'name': goods_name,
                'reason': reason,
                '当前': f'三级={sanji}, 四级={siji}, 品牌={brand}'
            })

    print(f"发现问题行: {len(problematic)} 条\n")
    for p in problematic:
        print(f"  Row {p['row']} [{str(p['name'])[:30]}]")
        print(f"    原因: {p['reason']}")
        print(f"    当前: {p['当前']}\n")

    if not problematic:
        print("没有发现问题行，无需修复")
        sys.exit(0)

    # ---- 4. 逐行重新补全 ----
    print(f"开始重新补全 {len(problematic)} 条...\n")
    fixed_count = 0
    error_count = 0

    for i, item in enumerate(problematic):
        row_idx = item['row']
        goods_name = str(item['name'])

        # 进度条
        current = i + 1
        pct = current / len(problematic) * 100
        bar_len = 30
        filled = int(bar_len * current / len(problematic))
        bar = '█' * filled + '░' * (bar_len - filled)
        print(f"\r[{bar}] {pct:.1f}% ({current}/{len(problematic)}) - {goods_name[:25]}", end='', flush=True)

        result = enrich_product(goods_name, '', categories)

        if 'error' in result:
            error_count += 1
            print(f"\n  错误: {result['error']}")
        else:
            fixed_count += 1
            # 写入8个补全字段（第8列起）
            enrich_cols = ['一级分类', '二级分类', '三级分类', '四级分类', '品牌', '品类', '品种', '百科信息']
            for j, col_name in enumerate(enrich_cols):
                ws.cell(row_idx, 8 + j).value = result.get(col_name, '')

        # 每10行保存一次，防止意外中断丢失
        if (i + 1) % 10 == 0:
            wb.save(excel_path)

    print(f"\n\n修复完成！成功: {fixed_count}, 失败: {error_count}")
    wb.save(excel_path)
    print(f"已保存: {excel_path}")


if __name__ == '__main__':
    main()