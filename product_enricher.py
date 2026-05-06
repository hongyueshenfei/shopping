#!/usr/bin/env python3
"""
Product Enricher - 商品信息补全核心模块
根据商品名称调用大模型（MiniMax-M2.7）补全：4级分类、品牌、品类、品种、百科信息。

使用方法:
    单条: python product_enricher.py <商品名称> [规格信息]
    批量: 由 batch_enrich.py 调用 enrich_product() 函数

输出字段:
    一级分类 / 二级分类 / 三级分类 / 四级分类 / 品牌 / 品类 / 品种 / 百科信息
"""

import sys
import os
import json
import re
from pathlib import Path

# ===========================================================================
# 内嵌分类数据（来源：种苗苗木分类清单.xlsx，共142条）
# ===========================================================================

_EMBEDDED_CATEGORIES = [
    ["植物", "种苗苗木", "多年生草本", "松果菊"],
    ["植物", "种苗苗木", "多年生草本", "玉簪"],
    ["植物", "种苗苗木", "多年生草本", "蝴蝶兰"],
    ["植物", "种苗苗木", "多年生草本", "角堇"],
    ["植物", "种苗苗木", "多年生草本", "芍药"],
    ["植物", "种苗苗木", "多年生草本", "仙客来"],
    ["植物", "种苗苗木", "多年生草本", "鸢尾"],
    ["植物", "种苗苗木", "多年生草本", "天竺葵"],
    ["植物", "种苗苗木", "多年生草本", "其他多年生草本"],
    ["植物", "种苗苗木", "一年生草本", "矮牵牛"],
    ["植物", "种苗苗木", "一年生草本", "三色堇"],
    ["植物", "种苗苗木", "一年生草本", "金鱼草"],
    ["植物", "种苗苗木", "一年生草本", "百日草"],
    ["植物", "种苗苗木", "一年生草本", "天竺葵"],
    ["植物", "种苗苗木", "一年生草本", "飞燕草"],
    ["植物", "种苗苗木", "一年生草本", "凤仙花"],
    ["植物", "种苗苗木", "一年生草本", "美女樱"],
    ["植物", "种苗苗木", "一年生草本", "其他一年生草本"],
    ["植物", "种苗苗木", "灌木", "月季"],
    ["植物", "种苗苗木", "灌木", "绣球"],
    ["植物", "种苗苗木", "灌木", "三角梅"],
    ["植物", "种苗苗木", "灌木", "杜鹃"],
    ["植物", "种苗苗木", "灌木", "蓝莓"],
    ["植物", "种苗苗木", "灌木", "茶花"],
    ["植物", "种苗苗木", "灌木", "木槿"],
    ["植物", "种苗苗木", "灌木", "其他灌木"],
    ["植物", "种苗苗木", "藤本", "铁线莲"],
    ["植物", "种苗苗木", "藤本", "球兰"],
    ["植物", "种苗苗木", "藤本", "紫藤"],
    ["植物", "种苗苗木", "藤本", "葡萄"],
    ["植物", "种苗苗木", "藤本", "风车茉莉"],
    ["植物", "种苗苗木", "藤本", "黄木香"],
    ["植物", "种苗苗木", "藤本", "藤本月季"],
    ["植物", "种苗苗木", "藤本", "络石藤"],
    ["植物", "种苗苗木", "藤本", "其他藤本"],
    ["植物", "种苗苗木", "乔木", "枫树"],
    ["植物", "种苗苗木", "乔木", "油橄榄"],
    ["植物", "种苗苗木", "乔木", "发财树"],
    ["植物", "种苗苗木", "乔木", "琴叶榕"],
    ["植物", "种苗苗木", "乔木", "女贞"],
    ["植物", "种苗苗木", "乔木", "木绣球"],
    ["植物", "种苗苗木", "乔木", "冬青"],
    ["植物", "种苗苗木", "乔木", "紫薇"],
    ["植物", "种苗苗木", "乔木", "其他乔木"],
    ["植物", "种苗苗木", "造型苗", "多头"],
    ["植物", "种苗苗木", "造型苗", "高杆"],
    ["植物", "种苗苗木", "造型苗", "棒棒糖"],
    ["植物", "种苗苗木", "造型苗", "丛生"],
    ["植物", "种苗苗木", "造型苗", "其他造型苗"],
    ["植物", "种苗苗木", "水生植物", "碗莲"],
    ["植物", "种苗苗木", "水生植物", "荷花"],
    ["植物", "种苗苗木", "水生植物", "睡莲"],
    ["植物", "种苗苗木", "水生植物", "水仙"],
    ["植物", "种苗苗木", "水生植物", "其他水生植物"],
    ["植物", "种子", "观赏花卉种子", "矮牵牛"],
    ["植物", "种子", "观赏花卉种子", "百日草"],
    ["植物", "种子", "观赏花卉种子", "三色堇"],
    ["植物", "种子", "观赏花卉种子", "金鱼草"],
    ["植物", "种子", "观赏花卉种子", "波斯菊"],
    ["植物", "种子", "观赏花卉种子", "向日葵"],
    ["植物", "种子", "观赏花卉种子", "天竺葵"],
    ["植物", "种子", "观赏花卉种子", "飞燕草"],
    ["植物", "种子", "观赏花卉种子", "其他观赏花卉种子"],
    ["植物", "种子", "蔬果种子", "番茄"],
    ["植物", "种子", "蔬果种子", "草莓"],
    ["植物", "种子", "蔬果种子", "萝卜"],
    ["植物", "种子", "蔬果种子", "西红柿"],
    ["植物", "种子", "蔬果种子", "黄瓜"],
    ["植物", "种子", "蔬果种子", "其他果蔬种子"],
    ["植物", "种子", "草坪草种", "黑麦草"],
    ["植物", "种子", "草坪草种", "高羊茅"],
    ["植物", "种子", "草坪草种", "其他草坪草种"],
    ["植物", "球块茎", "鳞茎", "郁金香"],
    ["植物", "球块茎", "鳞茎", "朱顶红"],
    ["植物", "球块茎", "鳞茎", "百合"],
    ["植物", "球块茎", "鳞茎", "洋水仙"],
    ["植物", "球块茎", "鳞茎", "风信子"],
    ["植物", "球块茎", "鳞茎", "风雨兰"],
    ["植物", "球块茎", "鳞茎", "晚香玉"],
    ["植物", "球块茎", "鳞茎", "石蒜"],
    ["植物", "球块茎", "鳞茎", "其他鳞茎"],
    ["植物", "球块茎", "球茎", "唐菖蒲"],
    ["植物", "球块茎", "球茎", "香雪兰"],
    ["植物", "球块茎", "球茎", "彩叶芋"],
    ["植物", "球块茎", "球茎", "姜荷花"],
    ["植物", "球块茎", "球茎", "其他球茎"],
    ["植物", "球块茎", "块茎", "大丽花"],
    ["植物", "球块茎", "块茎", "仙客来"],
    ["植物", "球块茎", "块茎", "马蹄莲"],
    ["植物", "球块茎", "块茎", "海棠"],
    ["植物", "球块茎", "块茎", "姜荷花"],
    ["植物", "球块茎", "块茎", "花毛茛"],
    ["植物", "球块茎", "块茎", "其他块茎"],
    ["非植", "盆", "品牌盆", "帝罗马"],
    ["非植", "盆", "品牌盆", "爱好"],
    ["非植", "盆", "品牌盆", "ARTSTONE"],
    ["非植", "盆", "品牌盆", "蕾秀"],
    ["非植", "盆", "品牌盆", "爱丽思"],
    ["非植", "盆", "品牌盆", "华锦"],
    ["非植", "盆", "品牌盆", "嘉伯瑞"],
    ["非植", "盆", "品牌盆", "园艺家"],
    ["非植", "盆", "品牌盆", "其他品牌盆"],
    ["非植", "盆", "尺寸盆", "10寸"],
    ["非植", "盆", "尺寸盆", "12寸"],
    ["非植", "盆", "尺寸盆", "14寸"],
    ["非植", "盆", "尺寸盆", "16寸"],
    ["非植", "盆", "尺寸盆", "20寸"],
    ["非植", "盆", "尺寸盆", "其他尺寸"],
    ["非植", "土", "专用介质", "球根专用介质"],
    ["非植", "土", "专用介质", "花卉专用介质"],
    ["非植", "土", "专用介质", "其他专用介质"],
    ["非植", "土", "普通介质", "椰糠"],
    ["非植", "土", "普通介质", "营养土"],
    ["非植", "土", "普通介质", "赤玉土"],
    ["非植", "土", "普通介质", "泥炭土"],
    ["非植", "土", "普通介质", "其他普通介质"],
    ["非植", "土", "改良材料", "陶粒"],
    ["非植", "土", "改良材料", "轻石"],
    ["非植", "土", "改良材料", "兰卡椰壳"],
    ["非植", "土", "改良材料", "珍珠岩"],
    ["非植", "土", "改良材料", "其他改良材料"],
    ["非植", "肥", "特肥", "骨粉"],
    ["非植", "肥", "特肥", "其他特肥"],
    ["非植", "肥", "通用肥", "有机肥"],
    ["非植", "肥", "通用肥", "水溶肥"],
    ["非植", "肥", "通用肥", "复合肥"],
    ["非植", "肥", "通用肥", "缓释肥"],
    ["非植", "肥", "通用肥", "其他通用肥"],
    ["非植", "工具", "浇水工具", "浇水壶"],
    ["非植", "工具", "浇水工具", "滴水器"],
    ["非植", "工具", "浇水工具", "其他浇水工具"],
    ["非植", "工具", "种植工具", "铲子"],
    ["非植", "工具", "种植工具", "其他种植工具"],
    ["非植", "工具", "修剪工具", "园艺剪"],
    ["非植", "工具", "修剪工具", "其他修剪工具"],
    ["非植", "工具", "清洁工具", "清洁工具"],
    ["非植", "工具", "植保用品", "植保用品"],
    ["非植", "工具", "基础工具", "基础工具"],
    ["非植", "园艺周边", "装饰摆件", "装饰摆件"],
    ["非植", "园艺周边", "其他周边", "其他周边"],
    ["非植", "书籍", "园艺书籍", "图鉴类"],
    ["非植", "书籍", "园艺书籍", "百科类"],
    ["非植", "书籍", "园艺书籍", "设计类"],
]

# ---------------------------------------------------------------------------
# 依赖安装：首次运行时自动安装缺少的包
# ---------------------------------------------------------------------------
def install_and_import():
    """检查并安装所需依赖包（anthropic、openpyxl）"""
    modules = ['anthropic', 'openpyxl']
    for mod in modules:
        try:
            __import__(mod.replace('-', '_'))
        except ImportError:
            print(f"Installing {mod}...")
            os.system(f"python3 -m pip install {mod} -q 2>/dev/null")

install_and_import()

import anthropic


# ===========================================================================
# 第一部分：工具函数
# ===========================================================================

def load_categories(excel_path=None):
    """
    加载4级分类体系。

    优先使用内嵌数据（不依赖外部文件）。
    仅当 excel_path 有值且文件存在时，尝试从文件加载。

    参数:
        excel_path: 可选，种苗苗木分类清单.xlsx 文件路径
    返回:
        分类字典列表，每条: {'一级分类':..., '二级分类':..., '三级分类':..., '四级分类':...}
    """
    if excel_path is None or not Path(excel_path).exists():
        return [
            {'一级分类': r[0], '二级分类': r[1], '三级分类': r[2], '四级分类': r[3]}
            for r in _EMBEDDED_CATEGORIES
        ]

    try:
        import openpyxl
    except ImportError:
        os.system("python3 -m pip install openpyxl -q")
        import openpyxl

    wb = openpyxl.load_workbook(excel_path)
    ws = wb['Sheet1']

    categories = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            categories.append({
                '一级分类': row[0],
                '二级分类': row[1],
                '三级分类': row[2],
                '四级分类': row[3]
            })
    return categories


def build_category_tree(categories):
    """
    将扁平分类列表构建为层级树结构，便于统计和展示。

    返回结构:
        {一级分类: {二级分类: {三级分类: [四级分类列表]}}}
    """
    tree = {}
    for cat in categories:
        l1 = cat['一级分类']
        l2 = cat['二级分类']
        l3 = cat['三级分类']
        l4 = cat['四级分类']

        if l1 not in tree:
            tree[l1] = {}
        if l2 not in tree[l1]:
            tree[l1][l2] = {}
        if l3 not in tree[l1][l2]:
            tree[l1][l2][l3] = []
        tree[l1][l2][l3].append(l4)

    return tree


def get_api_key():
    """
    获取 Anthropic API 密钥。
    优先级: 硬编码key > 环境变量 > 配置文件 ~/.claude/settings.json
    """
    hardcoded_key = 'sk-cp-vGuw2bWAbA8aEdyb7erws4sEWts7xAPu6rOYSJX0fkr7r_j2EXgIQnNYvLnneoPO8eKfmDPxCFrkl0S2zOsUBLRwqqd7uV3728dTHIqlcw_daFdZhF6BLNQ'
    if hardcoded_key:
        return hardcoded_key

    for var in ['ANTHROPIC_API_KEY', 'CLAUDE_API_KEY', 'API_KEY', 'ANTHROPIC_AUTH_TOKEN']:
        key = os.environ.get(var)
        if key:
            return key

    config_paths = [
        Path.home() / '.claude' / 'settings.json',
        Path.home() / '.claude' / 'settings.local.json',
        Path.home() / '.claude' / 'config.json',
    ]

    for config_path in config_paths:
        if config_path.exists():
            try:
                import json as json_mod
                with open(config_path) as f:
                    data = json_mod.load(f)
                    if 'env' in data:
                        for var in ['ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_API_KEY', 'API_KEY']:
                            if var in data['env']:
                                return data['env'][var]
                    for key_path in ['api_key', 'ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN', 'key', 'anthropic_key']:
                        keys = data
                        for part in key_path.split('.'):
                            if isinstance(keys, dict) and part in keys:
                                keys = keys[part]
                            else:
                                keys = None
                                break
                        if keys:
                            return keys
            except Exception:
                pass

    return None


def get_base_url():
    """获取自定义API Base URL（用于配置代理或私有部署）"""
    for var in ['ANTHROPIC_BASE_URL']:
        url = os.environ.get(var)
        if url:
            return url

    config_paths = [
        Path.home() / '.claude' / 'settings.json',
        Path.home() / '.claude' / 'settings.local.json',
    ]

    for config_path in config_paths:
        if config_path.exists():
            try:
                import json as json_mod
                with open(config_path) as f:
                    data = json_mod.load(f)
                    if 'env' in data and 'ANTHROPIC_BASE_URL' in data['env']:
                        return data['env']['ANTHROPIC_BASE_URL']
            except Exception:
                pass

    return None


# ===========================================================================
# 第二部分：LLM 调用与解析
# ===========================================================================

def call_llm(prompt, model='MiniMax-M2.7'):
    """
    向大模型发送prompt并获取返回。

    参数:
        prompt: 给LLM的提示词（含商品信息和分类规则）
        model:  模型名称，默认 MiniMax-M2.7
    返回:
        {'result': text} 或 {'error': error_message}
    """
    api_key = get_api_key()
    base_url = get_base_url()
    if not api_key:
        return {'error': 'API key not found. Please set ANTHROPIC_API_KEY environment variable.'}

    client_kwargs = {'api_key': api_key}
    if base_url:
        client_kwargs['base_url'] = base_url

    client = anthropic.Anthropic(**client_kwargs)

    try:
        response = client.messages.create(
            model=model,
            max_tokens=1024,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        text_result = []
        for block in response.content:
            if hasattr(block, 'type') and block.type == 'text':
                text_result.append(block.text)
            elif hasattr(block, 'type') and block.type == 'thinking':
                pass
            elif hasattr(block, 'text'):
                text_result.append(block.text)
        return {'result': ''.join(text_result) if text_result else str(response.content)}
    except Exception as e:
        return {'error': str(e)}


def extract_json(text):
    """
    从LLM返回的文本中提取并解析JSON。
    """
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    for fence in ['```json', '```', '`']:
        if text.startswith(fence):
            text = text[len(fence):]
        if text.endswith(fence):
            text = text[:-len(fence)]
        text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    decoder = json.JSONDecoder()
    try:
        data, end_idx = decoder.raw_decode(text)
        return data
    except json.JSONDecodeError:
        pass

    start = text.find('{')
    if start == -1:
        return None

    depth = 0
    in_string = False
    escape = False
    for i, ch in enumerate(text[start:], start):
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '"' and not escape:
            in_string = not in_string
        if not in_string:
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i+1])
                    except json.JSONDecodeError:
                        pass
                    break
    return None


# ===========================================================================
# 第三部分：后处理与数据清洗
# ===========================================================================

def clean_variety(product_name, data):
    """
    品种/品类后处理：从商品名中提取引号内的品种名。
    """
    variety = data.get('品种', '')
    yiji = data.get('一级分类', '')

    if yiji == '非植':
        data['品类'] = ''
        data['品种'] = ''
        return data

    plant_type_map = {
        '铁线莲': '铁线莲', '蓝莓': '蓝莓', '葡萄': '葡萄',
        '杜鹃': '杜鹃', '枫树': '枫树', '冬青': '冬青',
        '绣球': '绣球', '月季': '月季', '玫瑰': '月季',
        '草莓': '草莓', '茶花': '茶花', '三角梅': '三角梅',
    }

    name_str = str(product_name)
    variety_extracted = None

    quoted = re.findall(r'[""“”]([^""“”]+)[""“”]', name_str)
    if quoted:
        parts = [q.strip() for q in quoted
                 if not re.search(r'\d年|\d球|\d号|\dL|\dkg|代工|混苗', q)]
        if parts:
            variety_extracted = min(parts, key=len)
        else:
            variety_extracted = quoted[-1].strip()
        if variety != variety_extracted:
            data['品种'] = variety_extracted
    else:
        if variety:
            pure_plant_types = ['铁线莲', '葡萄', '杜鹃', '枫树', '冬青',
                               '绣球', '月季', '玫瑰', '草莓', '茶花', '三角梅',
                               '草本', '灌木', '乔木', '藤本']
            if variety.strip() in pure_plant_types:
                data['品种'] = ''

    if variety_extracted is None:
        for keyword, cat in plant_type_map.items():
            if keyword in name_str:
                data['品类'] = cat
                break

    return data


# ===========================================================================
# 第四部分：核心补全逻辑
# ===========================================================================

def enrich_product(product_name, spec_info, categories):
    """
    商品信息补全核心函数。

    流程:
    1. 构建包含4级分类目录和层级约束规则的prompt
    2. 调用LLM获取JSON格式的分类结果
    3. 解析JSON
    4. 后处理：品种提取、品类推断
    5. 品牌归一化

    参数:
        product_name: 商品名称字符串
        spec_info:    规格信息（可为空字符串）
        categories:   load_categories() 返回的分类列表
    返回:
        包含8个字段的字典，或 {'error': ...} 表示失败
    """
    category_lines = []
    for cat in categories:
        line = f"- {cat['一级分类']} > {cat['二级分类']} > {cat['三级分类']} > {cat['四级分类']}"
        category_lines.append(line)

    tree = build_category_tree(categories)
    category_summary = []
    for l1, l2_dict in tree.items():
        for l2, l3_dict in l2_dict.items():
            l3_list = list(l3_dict.keys())
            l4_counts = {l3: len(l4s) for l3, l4s in l3_dict.items()}
            category_summary.append(f"  {l1} > {l2}: {l3_list} (四级分类数: {sum(l4_counts.values())})")

    prompt = f"""你是一个专业的园艺商品分类专家。

## 输入信息
商品名称：{product_name}
规格信息：{spec_info if spec_info else '未提供'}

## 4级分类目录（严格遵守）
一级分类 > 二级分类 > 三级分类 > 四级分类
{chr(10).join(category_lines)}

## 层级强约束规则（必须严格遵守）
1. **只能选择上述分类目录中存在的L1/L2/L3/L4**，禁止自行创造新分类
2. **L1 → L2 → L3 → L4 必须构成完整的父子链条**
   - L2必须在对应的L1分类下
   - L3必须在对应的L2分类下
   - L4必须在对应的L3分类下
3. **不允许跨级分类或跳跃分类**
4. **兜底规则**：无法明确归类时，选择对应的"其他"分类

## 分类判断标准
- **以植物成年形态为统一判断标准**
- 藤本植物（如铁线莲、葡萄、紫藤）→ 三级分类为「藤本」
- 灌木（如月季、绣球、杜鹃）→ 三级分类为「灌木」
- 乔木（如枫树、油橄榄）→ 三级分类为「乔木」
- 草本花卉（如松果菊、玉簪、鸢尾）→ 三级分类为「多年生草本」或「一年生草本」

## 品牌、品类、品种说明
- **品牌**：育种公司或苗圃品牌（如玫昂、美乐棵、花彩师等）。必须输出品牌
- **品类**：植物的生物学类型名称，如铁线莲、蓝莓、葡萄、杜鹃、月季等
- **品种**：具体植物品种名称（如龙沙宝石、阳光玫瑰等），从商品名引号内提取的短名称

## 输出格式（JSON）
{{
    "一级分类": "必须从分类目录中选择",
    "二级分类": "必须从分类目录中选择",
    "三级分类": "必须从分类目录中选择",
    "四级分类": "必须从分类目录中选择",
    "品牌": "育种公司或品牌名",
    "品类": "植物类型名（如有）",
    "品种": "具体品种名称（如有）",
    "百科信息": "商品简要描述、特点、用途、养护要点等（50-200字）"
}}

请严格按照JSON格式输出，不要包含其他文字。
"""

    result = call_llm(prompt)

    if 'error' in result:
        return result

    try:
        text = result['result']
        data = extract_json(text)
        if data is None:
            return {'error': '无法解析LLM返回结果', 'raw': text}
    except json.JSONDecodeError as e:
        return {'error': f'JSON解析失败: {str(e)}', 'raw': text}

    data = clean_variety(product_name, data)

    brand = data.get('品牌', '')
    if brand and brand.strip() in ('HY', 'Encore Azaleas（安酷杜鹃）', 'Encore Azaleas', 'HY-'):
        data['品牌'] = '安酷'

    variety = data.get('品种', '')
    if variety and ('科' in variety or '属' in variety or len(variety) > 20):
        data['品种'] = ''

    return data


# ===========================================================================
# 第五部分：单条测试入口
# ===========================================================================

def main():
    """
    命令行单条测试模式。
    用法: python product_enricher.py <商品名称> [规格信息]
    示例: python product_enricher.py 月季花苗 一年生
          python product_enricher.py 有机营养土 10L装
    """
    if len(sys.argv) < 2:
        print("使用方法: python product_enricher.py <商品名称> [规格信息]")
        print("示例: python product_enricher.py 月季花苗 一年生")
        print("示例: python product_enricher.py 有机营养土 10L装")
        sys.exit(1)

    product_name = sys.argv[1]
    spec_info = sys.argv[2] if len(sys.argv) > 2 else ''

    excel_paths = [
        Path(__file__).parent / '种苗苗木分类清单.xlsx',
        Path(__file__).parent / '商品分类目录.xlsx',
        Path.home() / 'Desktop' / '沈飞mac' / 'skills' / 'shopping' / '种苗苗木分类清单.xlsx',
        Path.cwd() / '种苗苗木分类清单.xlsx',
    ]

    excel_path = None
    for path in excel_paths:
        if path.exists():
            excel_path = path
            break

    if not excel_path:
        print("警告: 找不到分类文件，使用内嵌数据")

    categories = load_categories(excel_path)
    print(f"共加载 {len(categories)} 条分类记录")

    print(f"\n正在分析商品: {product_name} {spec_info}")
    result = enrich_product(product_name, spec_info, categories)

    if 'error' in result:
        print(f"\n错误: {result['error']}")
        if 'raw' in result:
            print(f"原始返回: {result['raw'][:500]}")
        sys.exit(1)

    print("\n=== 补全结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
