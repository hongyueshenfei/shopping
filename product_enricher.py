#!/usr/bin/env python3
"""
Product Enricher - 商品信息补全核心模块
根据商品名称调用大模型（MiniMax-M2.7）补全：4级分类、品牌、品类、品种、百科信息。

使用方法:
    单条: python product_enricher.py <商品名称> [规格信息]
    批量: 由 batch_enrich.py 调用 enrich_product() 函数
    修复: 由 fix_enrich.py 调用 enrich_product() 函数

输出字段:
    一级分类 / 二级分类 / 三级分类 / 四级分类 / 品牌 / 品类 / 品种 / 百科信息
"""

import sys
import os
import json
import re
from pathlib import Path

# ===========================================================================
# 内嵌分类数据（来源：商品分类目录.xlsx，共88条）
# ===========================================================================

_EMBEDDED_CATEGORIES = [
    ["植物", "种子", "观赏花卉种子", "一二年生观赏花卉种子"],
    ["植物", "种子", "观赏花卉种子", "多年生观赏花卉种子"],
    ["植物", "种子", "观赏花卉种子", "盆栽观赏花卉种子"],
    ["植物", "种子", "观赏花卉种子", "切花观赏花卉种子"],
    ["植物", "种子", "观赏花卉种子", "其他观赏花卉种子"],
    ["植物", "种子", "绿化林木种子", "景观乔木种子"],
    ["植物", "种子", "绿化林木种子", "景观灌木种子"],
    ["植物", "种子", "绿化林木种子", "生态修复林木种子"],
    ["植物", "种子", "绿化林木种子", "其他绿化林木种子"],
    ["植物", "种子", "草坪草种", "冷季型草坪草种"],
    ["植物", "种子", "草坪草种", "暖季型草坪草种"],
    ["植物", "种子", "草坪草种", "特殊用途草坪草种"],
    ["植物", "种子", "蔬果种子", "叶菜类种子"],
    ["植物", "种子", "蔬果种子", "果菜类种子"],
    ["植物", "种子", "蔬果种子", "根茎类种子"],
    ["植物", "球块茎", "鳞茎", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "球块茎", "球茎", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "球块茎", "块茎", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "球块茎", "根茎", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "一年生草本", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "二年生草本", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "多年生草本", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "乔木", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "灌木", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "木本藤本", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["植物", "种苗苗木", "草本藤本", "分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"],
    ["非植", "盆", "种植盆", "陶质种植盆"],
    ["非植", "盆", "种植盆", "塑料种植盆"],
    ["非植", "盆", "种植盆", "瓷质种植盆"],
    ["非植", "盆", "种植盆", "水泥种植盆"],
    ["非植", "盆", "种植盆", "木质种植盆"],
    ["非植", "盆", "种植盆", "其他种植盆"],
    ["非植", "盆", "套盆", "陶质套盆"],
    ["非植", "盆", "套盆", "塑料套盆"],
    ["非植", "盆", "套盆", "瓷质套盆"],
    ["非植", "盆", "套盆", "金属套盆"],
    ["非植", "盆", "套盆", "布艺套盆"],
    ["非植", "盆", "套盆", "其他套盆"],
    ["非植", "盆", "悬挂盆", "塑料悬挂盆"],
    ["非植", "盆", "悬挂盆", "陶质悬挂盆"],
    ["非植", "盆", "悬挂盆", "金属悬挂盆"],
    ["非植", "盆", "悬挂盆", "藤编悬挂盆"],
    ["非植", "盆", "悬挂盆", "其他悬挂盆"],
    ["非植", "盆", "特殊场景盆器", "水培盆（玻璃 / 塑料）"],
    ["非植", "盆", "特殊场景盆器", "壁挂盆（塑料 / 金属）"],
    ["非植", "盆", "特殊场景盆器", "浅口盆（陶 / 瓷）"],
    ["非植", "土", "单一介质（原料型）", "泥炭"],
    ["非植", "土", "单一介质（原料型）", "椰糠"],
    ["非植", "土", "单一介质（原料型）", "珍珠岩"],
    ["非植", "土", "单一介质（原料型）", "蛭石"],
    ["非植", "土", "单一介质（原料型）", "园土"],
    ["非植", "土", "单一介质（原料型）", "松鳞"],
    ["非植", "土", "单一介质（原料型）", "陶粒"],
    ["非植", "土", "单一介质（原料型）", "赤玉土"],
    ["非植", "土", "混合介质（成品功能型）", "通用型营养土"],
    ["非植", "土", "混合介质（成品功能型）", "透气控根型专用土"],
    ["非植", "土", "混合介质（成品功能型）", "喜湿保水型专用土"],
    ["非植", "土", "混合介质（成品功能型）", "开花结果型专用土"],
    ["非植", "土", "混合介质（成品功能型）", "无菌育苗型专用土"],
    ["非植", "肥", "植物营养类", "速效型营养肥"],
    ["非植", "肥", "植物营养类", "长效型营养肥"],
    ["非植", "肥", "植物营养类", "天然有机营养肥"],
    ["非植", "肥", "植物营养类", "专用型营养液"],
    ["非植", "肥", "病虫害防治类", "单一杀虫剂"],
    ["非植", "肥", "病虫害防治类", "病虫害综合套装"],
    ["非植", "肥", "病虫害防治类", "功能性驱虫精油"],
    ["非植", "工具", "浇水灌溉工具", "基础浇水工具"],
    ["非植", "工具", "浇水灌溉工具", "精细灌溉工具"],
    ["非植", "工具", "修剪造型工具", "通用修剪工具"],
    ["非植", "工具", "修剪造型工具", "轻便型修剪工具"],
    ["非植", "工具", "修剪造型工具", "安全型修剪工具"],
    ["非植", "工具", "辅助固定摆放工具", "植物支撑工具"],
    ["非植", "工具", "辅助固定摆放工具", "悬挂固定工具"],
    ["非植", "工具", "收纳整理工具", "工具收纳用品"],
    ["非植", "工具", "收纳整理工具", "耗材收纳用品"],
    ["非植", "工具", "防护装备", "手部防护"],
    ["非植", "工具", "防护装备", "其他防护"],
    ["非植", "工具", "种植辅助工具", "基础种植工具"],
    ["非植", "工具", "种植辅助工具", "精细操作工具"],
    ["非植", "工具", "通用配件及周边", "工具配件"],
    ["非植", "书籍", "园艺图书", "养护类"],
    ["非植", "书籍", "园艺图书", "设计类"],
    ["非植", "书籍", "园艺图书", "图鉴类"],
    ["非植", "书籍", "园艺百科", "植物百科"],
    ["非植", "书籍", "园艺百科", "资材百科"],
    ["非植", "园艺周边", "园艺周边", "分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分"],
    ["非植", "园艺周边", "园艺文创", "分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分"],
    ["非植", "园艺周边", "家居装饰", "分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分"],
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
    仅当 excel_path 有值且文件存在时，尝试从文件加载（用于未来更新分类数据）。

    参数:
        excel_path: 可选，商品分类目录.xlsx 文件路径
    返回:
        分类字典列表，每条: {'一级分类':..., '二级分类':..., '三级分类':..., '四级分类':...}
    """
    # 优先使用内嵌数据（88条，与商品分类目录.xlsx一致）
    if excel_path is None or not Path(excel_path).exists():
        return [
            {'一级分类': r[0], '二级分类': r[1], '三级分类': r[2], '四级分类': r[3]}
            for r in _EMBEDDED_CATEGORIES
        ]

    # 文件存在时从文件加载（允许外部覆盖）
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
    # 硬编码key（主要来源）
    hardcoded_key = 'sk-cp-vGuw2bWAbA8aEdyb7erws4sEWts7xAPu6rOYSJX0fkr7r_j2EXgIQnNYvLnneoPO8eKfmDPxCFrkl0S2zOsUBLRwqqd7uV3728dTHIqlcw_daFdZhF6BLNQ'
    if hardcoded_key:
        return hardcoded_key

    # 检查常见环境变量
    for var in ['ANTHROPIC_API_KEY', 'CLAUDE_API_KEY', 'API_KEY', 'ANTHROPIC_AUTH_TOKEN']:
        key = os.environ.get(var)
        if key:
            return key

    # 检查配置文件 ~/.claude/settings.json 等
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
                    # 优先检查 env 节
                    if 'env' in data:
                        for var in ['ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_API_KEY', 'API_KEY']:
                            if var in data['env']:
                                return data['env'][var]
                    # 检查常见的key路径
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
    """
    获取自定义API Base URL（用于配置代理或私有部署）。
    优先级: 环境变量 > 配置文件
    """
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
        # 兼容不同类型的content block（text / thinking）
        text_result = []
        for block in response.content:
            if hasattr(block, 'type') and block.type == 'text':
                text_result.append(block.text)
            elif hasattr(block, 'type') and block.type == 'thinking':
                pass  # 跳过思考过程
            elif hasattr(block, 'text'):
                text_result.append(block.text)
        return {'result': ''.join(text_result) if text_result else str(response.content)}
    except Exception as e:
        return {'error': str(e)}


def extract_json(text):
    """
    从LLM返回的文本中提取并解析JSON。
    处理多种异常情况:
    - 带markdown代码块（```json ... ```）
    - 带反引号（`...`）
    - 前后有多余空白
    - 嵌套花括号导致直接json.loads失败
    最后通过括号匹配找到最外层JSON对象。
    """
    # 1. 直接解析
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2. 去除markdown代码块围栏
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

    # 3. 使用raw_decode自动查找JSON边界
    decoder = json.JSONDecoder()
    try:
        data, end_idx = decoder.raw_decode(text)
        return data
    except json.JSONDecodeError:
        pass

    # 4. 兜底：逐字符括号匹配找最外层{}
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
    品种/品类后处理：从商品名中提取引号内的品种名，修正品类。

    处理逻辑:
    1. 非植物商品（一级分类=非植）：品类和品种强制清空
    2. 品类：根据商品名关键词映射表推断植物生物学类型
    3. 品种：
       - 有引号：从引号内提取（排除规格描述性质的片段）
       - 无引号：如品种名为纯植物类型名则清空（防止品类当品种）
    """
    import re

    variety = data.get('品种', '')
    yiji = data.get('一级分类', '')

    # 非植物：品类+品种必须为空
    if yiji == '非植':
        data['品类'] = ''
        data['品种'] = ''
        return data

    # ---- 品类推断 ----
    # 关键词→品类名 映射表
    plant_type_map = {
        '铁线莲': '铁线莲', '蓝莓': '蓝莓', '葡萄': '葡萄', '葡萄柚': '柚',
        '杜鹃': '杜鹃', '安酷杜鹃': '杜鹃', '枫树': '枫树', '冬青': '冬青',
        '绣球': '绣球', '枇杷': '枇杷', '月季': '月季', '玫瑰': '月季',
        '草莓': '草莓', '梨': '梨', '桃': '桃', '苹果': '苹果', '黑莓': '黑莓',
        '紫菀': '紫菀', '百子莲': '百子莲', '风雨兰': '风雨兰',
        '香檬': '香檬', '蜜桔': '蜜桔', '金桔': '蜜桔', '丁香': '丁香',
        '巴布豆': '巴布豆',
    }

    name_str = str(product_name)

    # ---- 品种提取（先于品类推断） ----
    # 从中文/英文引号内提取品种名
    quoted = re.findall(r'[""\u201c\u201d]([^""\u201c\u201d]+)[""\u201c\u201d]', name_str)
    variety_extracted = None
    if quoted:
        # 过滤掉明显是规格描述的部分（带年/球/号/L/kg/代工/混苗等）
        parts = [q.strip() for q in quoted
                 if not re.search(r'\d年|\d球|\d号|\dL|\dkg|代工|混苗', q)]
        if parts:
            # 取最短的有效品种名（最可能是真正的品种名）
            variety_extracted = min(parts, key=len)
        else:
            variety_extracted = quoted[-1].strip()
        if variety != variety_extracted:
            data['品种'] = variety_extracted
    else:
        # 无引号：品种若为纯植物类型名则清空
        if variety:
            pure_plant_types = ['铁线莲', '葡萄', '杜鹃', '枫树', '冬青',
                               '绣球', '月季', '玫瑰', '草莓', '梨', '桃',
                               '苹果', '黑莓', '柚', '草本', '灌木', '乔木', '藤本']
            if variety.strip() in pure_plant_types:
                data['品种'] = ''
            # 特殊：商品名含蓝莓时"蓝葡萄"是品种，非葡萄品种
            if variety.strip() == '蓝葡萄' and '蓝莓' in name_str:
                data['品种'] = ''

    # ---- 品类推断（品种提取成功时跳过，避免品种名含关键词被错误归类） ----
    # 只有当品种无法从引号提取时，才用关键词映射推断品类
    # 防止"葡萄太妃糖"→品类=葡萄 这种误判
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
    1. 构建包含分类目录和植物学规则的prompt
    2. 调用LLM获取JSON格式的分类结果
    3. 解析JSON（extract_json）
    4. 后处理：品种提取、品类推断（clean_variety）
    5. 品牌归一化：HY/Encore → 安酷
    6. 品种防污染：含"科/属"或超过20字的长文本强制清空

    参数:
        product_name: 商品名称字符串
        spec_info:    规格信息（可为空字符串）
        categories:   load_categories() 返回的分类列表
    返回:
        包含8个字段的字典，或 {'error': ...} 表示失败
    """
    # 构建分类目录行列表（用于prompt）
    category_lines = []
    for cat in categories:
        line = f"- {cat['一级分类']} > {cat['二级分类']} > {cat['三级分类']} > {cat['四级分类']}"
        category_lines.append(line)

    # 构建分类树（用于prompt里的统计摘要）
    tree = build_category_tree(categories)
    category_summary = []
    for l1, l2_dict in tree.items():
        for l2, l3_dict in l2_dict.items():
            l3_list = list(l3_dict.keys())
            l4_counts = {l3: len(l4s) for l3, l4s in l3_dict.items()}
            category_summary.append(f"  {l1} > {l2}: {l3_list} (四级分类数: {sum(l4_counts.values())})")

    # 组装prompt：商品信息 + 分类目录 + 植物学规则 + 品类品种品牌定义 + 输出格式
    prompt = f"""你是一个专业的园艺商品分类专家，精通植物分类学和园艺植物学。

## 输入信息
商品名称：{product_name}
规格信息：{spec_info if spec_info else '未提供'}

## 分类目录（商品分类目录.xlsx）
一级分类 > 二级分类 > 三级分类 > 四级分类
每行的分类层级结构如下：
{chr(10).join(category_lines)}

## 植物学分类规则
1. **以植物的成年形态为唯一判断标准**，不受苗龄或规格大小影响
   - 藤本植物（如月季、紫藤、葡萄）→ 三级分类为「木本藤本」
   - 灌木（如玫瑰、茶花）→ 三级分类为「灌木」
   - 乔木（如樱花、银杏）→ 三级分类为「乔木」

2. **如果四级分类为"分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上"**：
   - 必须根据植物的成年形态学特征确定三级分类
   - 然后进一步确定该植物所属的**植物学科属**（如：蔷薇科、茄科、豆科等）
   - 四级分类输出为具体的植物科名

3. **统一性原则**：同一植物品种必须获得一致的分类结果，不因规格描述不同而产生不同结果

**常见园艺植物的科属参照（必须严格遵守，禁止编造不存在的科名）**：
- **绣球花** → 四级分类为「绣球花科」（不是「绣球科」！），三级分类为「灌木」
- **铁线莲** → 四级分类为「毛茛科」，三级分类为「木本藤本」（不是「草本藤本」！）
- **枫树** → 四级分类为「槭树科」（不是「槒树科」！），三级分类为「乔木」
- **葡萄（寒香蜜、阳光玫瑰等）** → 四级分类为「葡萄科」，三级分类为「木本藤本」
- **百子莲** → 四级分类为「石蒜科」，三级分类为「鳞茎」（不是「多年生草本」！）
- **蓝莓** → 四级分类为「杜鹃花科」，三级分类为「灌木」
- **草莓** → 四级分类为「蔷薇科」，三级分类为「草本」
- **玫瑰/月季** → 四级分类为「蔷薇科」，三级分类为「灌木」
- **杜鹃/安酷杜鹃** → 四级分类为「杜鹃花科」，三级分类为「灌木」
- **枇杷/桃/梨/苹果/黑莓** → 四级分类为「蔷薇科」，三级分类为「乔木」
- **枫树** → 四级分类为「槭树科」，三级分类为「乔木」
- **冬青** → 四级分类为「冬青科」，三级分类为「灌木」
- **紫菀** → 四级分类为「菊科」，三级分类为「多年生草本」
- **风雨兰** → 四级分类为「石蒜科」，三级分类为「鳞茎」

**品牌、品类与品种说明：**
- **品牌**：指育种公司或苗圃品牌（如：玫昂、美乐棵、花彩师、德沃多等）。**必须输出品牌**，即使商品是通用类型无明确品牌，也应输出常见园艺品牌作为默认值。**安酷杜鹃（Encore Azaleas）品牌的商品，品牌统一输出为"安酷"**，不得输出为"HY"或"HY-"
- **品类**：指植物的生物学类型名称，如铁线莲、蓝莓、葡萄、杜鹃、枫树、冬青、绣球、枇杷、月季、草莓等。**所有植物商品必须输出品类**。对于植物，即使商品名没有明确写品类，也要根据植物学分类推断出正确的品类
- **品种**：指具体植物品种名称（如：龙沙宝石、追雪、乌托邦、阳光玫瑰等）。**品种名必须是从商品名引号内提取的短名称（2-10字）**，禁止输出包含"科""属""是..."等描述性文字。如果商品名没有引号包裹具体品种名，品种字段为空字符串""。**禁止**把品类名（如"蓝莓"、"铁线莲"、"葡萄"）当作品种输出！

**输出格式（JSON）：**
{{
    "一级分类": "...",
    "二级分类": "...",
    "三级分类": "...",
    "四级分类": "...",
    "品牌": "育种公司或品牌名",
    "品类": "植物类型名（如铁线莲、蓝莓等）",
    "品种": "具体品种名称（如有）",
    "百科信息": "商品的简要描述、特点、用途、养护要点等（50-200字）"
}}

请严格按照JSON格式输出，不要包含其他文字。
"""

    # ---- 调用LLM ----
    result = call_llm(prompt)

    if 'error' in result:
        return result

    # ---- 解析JSON ----
    try:
        text = result['result']
        data = extract_json(text)
        if data is None:
            return {'error': '无法解析LLM返回结果', 'raw': text}
    except json.JSONDecodeError as e:
        return {'error': f'JSON解析失败: {str(e)}', 'raw': text}

    # ---- 后处理 ----
    # 1. 品种提取+品类推断
    data = clean_variety(product_name, data)

    # 2. 品牌归一化：HY/Encore统一为"安酷"
    brand = data.get('品牌', '')
    if brand and brand.strip() in ('HY', 'Encore Azaleas（安酷杜鹃）', 'Encore Azaleas', 'HY-'):
        data['品牌'] = '安酷'

    # 3. 品种防污染：品种列只允许是真正的品种名（2-20字纯名称）
    #    禁止含"科""属"或超长的百科式描述
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
    示例: python product_enricher.py 虹安铁线莲 乌托邦
          python product_enricher.py 花彩师营养液 500ml
    """
    if len(sys.argv) < 2:
        print("使用方法: python product_enricher.py <商品名称> [规格信息]")
        print("示例: python product_enricher.py 月季花苗 一年生")
        print("示例: python product_enricher.py 有机营养土 10L装")
        sys.exit(1)

    product_name = sys.argv[1]
    spec_info = sys.argv[2] if len(sys.argv) > 2 else ''

    # 查找商品分类目录.xlsx（多个可能路径）
    excel_paths = [
        Path(__file__).parent / '商品分类目录.xlsx',
        Path.home() / 'Desktop' / '沈飞mac' / 'skills' / 'shaoping' / '商品分类目录.xlsx',
        Path.home() / 'Desktop' / '商品分类目录.xlsx',
        Path.cwd() / '商品分类目录.xlsx',
    ]

    excel_path = None
    for path in excel_paths:
        if path.exists():
            excel_path = path
            break

    if not excel_path:
        print("错误: 找不到商品分类目录.xlsx文件")
        sys.exit(1)

    print(f"加载分类目录: {excel_path}")
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