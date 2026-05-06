#!/usr/bin/env php
<?php
/**
 * Product Enricher - 商品信息补全核心模块（PHP版）
 *
 * 根据商品名称调用大模型（MiniMax-M2.7）补全：4级分类、品牌、品类、品种、百科信息。
 *
 * 使用方法:
 *   单条: php product_enricher.php <商品名称> [规格信息]
 *   批量: 由 batch_enrich.php 调用 enrichProduct() 函数
 *   修复: 由 fix_enrich.php 调用 enrichProduct() 函数
 *
 * 输出字段:
 *   一级分类 / 二级分类 / 三级分类 / 四级分类 / 品牌 / 品类 / 品种 / 百科信息
 *
 * 依赖:
 *   PHP 7.4+, curl, json, mbstring
 *   安装: composer require phpoffice/phpspreadsheet
 */

// ===========================================================================
// 第一部分：工具函数
// ===========================================================================

// 内嵌分类数据（来源：商品分类目录.xlsx，共88条）
$EMBEDDED_CATEGORIES = [
    ['植物', '种子', '观赏花卉种子', '一二年生观赏花卉种子'],
    ['植物', '种子', '观赏花卉种子', '多年生观赏花卉种子'],
    ['植物', '种子', '观赏花卉种子', '盆栽观赏花卉种子'],
    ['植物', '种子', '观赏花卉种子', '切花观赏花卉种子'],
    ['植物', '种子', '观赏花卉种子', '其他观赏花卉种子'],
    ['植物', '种子', '绿化林木种子', '景观乔木种子'],
    ['植物', '种子', '绿化林木种子', '景观灌木种子'],
    ['植物', '种子', '绿化林木种子', '生态修复林木种子'],
    ['植物', '种子', '绿化林木种子', '其他绿化林木种子'],
    ['植物', '种子', '草坪草种', '冷季型草坪草种'],
    ['植物', '种子', '草坪草种', '暖季型草坪草种'],
    ['植物', '种子', '草坪草种', '特殊用途草坪草种'],
    ['植物', '种子', '蔬果种子', '叶菜类种子'],
    ['植物', '种子', '蔬果种子', '果菜类种子'],
    ['植物', '种子', '蔬果种子', '根茎类种子'],
    ['植物', '球块茎', '鳞茎', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '球块茎', '球茎', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '球块茎', '块茎', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '球块茎', '根茎', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '一年生草本', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '二年生草本', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '多年生草本', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '乔木', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '灌木', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '木本藤本', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['植物', '种苗苗木', '草本藤本', '分析用户给出的植物名称、规格、单位信息和之前的分类将该植物准确归属到植物学中的科上'],
    ['非植', '盆', '种植盆', '陶质种植盆'],
    ['非植', '盆', '种植盆', '塑料种植盆'],
    ['非植', '盆', '种植盆', '瓷质种植盆'],
    ['非植', '盆', '种植盆', '水泥种植盆'],
    ['非植', '盆', '种植盆', '木质种植盆'],
    ['非植', '盆', '种植盆', '其他种植盆'],
    ['非植', '盆', '套盆', '陶质套盆'],
    ['非植', '盆', '套盆', '塑料套盆'],
    ['非植', '盆', '套盆', '瓷质套盆'],
    ['非植', '盆', '套盆', '金属套盆'],
    ['非植', '盆', '套盆', '布艺套盆'],
    ['非植', '盆', '套盆', '其他套盆'],
    ['非植', '盆', '悬挂盆', '塑料悬挂盆'],
    ['非植', '盆', '悬挂盆', '陶质悬挂盆'],
    ['非植', '盆', '悬挂盆', '金属悬挂盆'],
    ['非植', '盆', '悬挂盆', '藤编悬挂盆'],
    ['非植', '盆', '悬挂盆', '其他悬挂盆'],
    ['非植', '盆', '特殊场景盆器', '水培盆（玻璃 / 塑料）'],
    ['非植', '盆', '特殊场景盆器', '壁挂盆（塑料 / 金属）'],
    ['非植', '盆', '特殊场景盆器', '浅口盆（陶 / 瓷）'],
    ['非植', '土', '单一介质（原料型）', '泥炭'],
    ['非植', '土', '单一介质（原料型）', '椰糠'],
    ['非植', '土', '单一介质（原料型）', '珍珠岩'],
    ['非植', '土', '单一介质（原料型）', '蛭石'],
    ['非植', '土', '单一介质（原料型）', '园土'],
    ['非植', '土', '单一介质（原料型）', '松鳞'],
    ['非植', '土', '单一介质（原料型）', '陶粒'],
    ['非植', '土', '单一介质（原料型）', '赤玉土'],
    ['非植', '土', '混合介质（成品功能型）', '通用型营养土'],
    ['非植', '土', '混合介质（成品功能型）', '透气控根型专用土'],
    ['非植', '土', '混合介质（成品功能型）', '喜湿保水型专用土'],
    ['非植', '土', '混合介质（成品功能型）', '开花结果型专用土'],
    ['非植', '土', '混合介质（成品功能型）', '无菌育苗型专用土'],
    ['非植', '肥', '植物营养类', '速效型营养肥'],
    ['非植', '肥', '植物营养类', '长效型营养肥'],
    ['非植', '肥', '植物营养类', '天然有机营养肥'],
    ['非植', '肥', '植物营养类', '专用型营养液'],
    ['非植', '肥', '病虫害防治类', '单一杀虫剂'],
    ['非植', '肥', '病虫害防治类', '病虫害综合套装'],
    ['非植', '肥', '病虫害防治类', '功能性驱虫精油'],
    ['非植', '工具', '浇水灌溉工具', '基础浇水工具'],
    ['非植', '工具', '浇水灌溉工具', '精细灌溉工具'],
    ['非植', '工具', '修剪造型工具', '通用修剪工具'],
    ['非植', '工具', '修剪造型工具', '轻便型修剪工具'],
    ['非植', '工具', '修剪造型工具', '安全型修剪工具'],
    ['非植', '工具', '辅助固定摆放工具', '植物支撑工具'],
    ['非植', '工具', '辅助固定摆放工具', '悬挂固定工具'],
    ['非植', '工具', '收纳整理工具', '工具收纳用品'],
    ['非植', '工具', '收纳整理工具', '耗材收纳用品'],
    ['非植', '工具', '防护装备', '手部防护'],
    ['非植', '工具', '防护装备', '其他防护'],
    ['非植', '工具', '种植辅助工具', '基础种植工具'],
    ['非植', '工具', '种植辅助工具', '精细操作工具'],
    ['非植', '工具', '通用配件及周边', '工具配件'],
    ['非植', '书籍', '园艺图书', '养护类'],
    ['非植', '书籍', '园艺图书', '设计类'],
    ['非植', '书籍', '园艺图书', '图鉴类'],
    ['非植', '书籍', '园艺百科', '植物百科'],
    ['非植', '书籍', '园艺百科', '资材百科'],
    ['非植', '园艺周边', '园艺周边', '分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分'],
    ['非植', '园艺周边', '园艺文创', '分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分'],
    ['非植', '园艺周边', '家居装饰', '分析用户给出的产品名称、规格、单位信息和之前的分类将自行细分'],
];

/**
 * 加载4级分类体系。
 *
 * 优先使用内嵌数据（不依赖外部文件）。
 * 仅当 $excelPath 有值且文件存在时，尝试从文件加载（用于未来更新分类数据）。
 *
 * @param string|null $excelPath 可选，商品分类目录.xlsx 文件路径
 * @return array 分类字典数组，每条: ['一级分类'=>..., '二级分类'=>..., '三级分类'=>..., '四级分类'=>...]
 */
function loadCategories(?string $excelPath = null): array
{
    // 优先使用内嵌数据（88条，与商品分类目录.xlsx一致）
    if ($excelPath === null || !file_exists($excelPath)) {
        global $EMBEDDED_CATEGORIES;
        return array_map(function($r) {
            return ['一级分类'=>$r[0], '二级分类'=>$r[1], '三级分类'=>$r[2], '四级分类'=>$r[3]];
        }, $EMBEDDED_CATEGORIES);
    }

    // 文件存在时从文件加载（允许外部覆盖）
    if (!class_exists('\PhpOffice\PhpSpreadsheet\IOFactory')) {
        trigger_error("建议安装 phpoffice/phpspreadsheet 以获得更好的Excel支持", E_USER_WARNING);
        return [];
    }

    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excelPath);
    $sheet = $spreadsheet->getActiveSheet();
    $rows = $sheet->toArray();

    $categories = [];
    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];
        if (!empty($row[0])) {
            $categories[] = [
                '一级分类' => $row[0],
                '二级分类' => $row[1] ?? '',
                '三级分类' => $row[2] ?? '',
                '四级分类' => $row[3] ?? '',
            ];
        }
    }
    return $categories;
}

/**
 * 构建分类层级树
 *
 * @param array $categories loadCategories() 返回的数组
 * @return array 层级树: [一级=>[二级=>[三级=>[四级列表]]]]
 */
function buildCategoryTree(array $categories): array
{
    $tree = [];
    foreach ($categories as $cat) {
        $l1 = $cat['一级分类'];
        $l2 = $cat['二级分类'];
        $l3 = $cat['三级分类'];
        $l4 = $cat['四级分类'];

        if (!isset($tree[$l1])) $tree[$l1] = [];
        if (!isset($tree[$l1][$l2])) $tree[$l1][$l2] = [];
        if (!isset($tree[$l1][$l2][$l3])) $tree[$l1][$l2][$l3] = [];
        $tree[$l1][$l2][$l3][] = $l4;
    }
    return $tree;
}

/**
 * 获取 API Key
 * 优先级: 硬编码 > 环境变量 > 配置文件
 */
function getApiKey(): string
{
    // 硬编码key（主要来源）
    $hardcodedKey = 'sk-cp-vGuw2bWAbA8aEdyb7erws4sEWts7xAPu6rOYSJX0fkr7r_j2EXgIQnNYvLnneoPO8eKfmDPxCFrkl0S2zOsUBLRwqqd7uV3728dTHIqlcw_daFdZhF6BLNQ';
    if ($hardcodedKey) return $hardcodedKey;

    // 环境变量
    foreach (['ANTHROPIC_API_KEY', 'CLAUDE_API_KEY', 'API_KEY', 'ANTHROPIC_AUTH_TOKEN'] as $var) {
        $key = getenv($var);
        if ($key) return $key;
    }

    // 配置文件 ~/.claude/settings.json
    $home = getenv('HOME') ?: exec('echo $HOME');
    $configPaths = [
        $home . '/.claude/settings.json',
        $home . '/.claude/settings.local.json',
        $home . '/.claude/config.json',
    ];

    foreach ($configPaths as $path) {
        if (file_exists($path)) {
            $data = json_decode(file_get_contents($path), true);
            if (isset($data['env']['ANTHROPIC_API_KEY'])) {
                return $data['env']['ANTHROPIC_API_KEY'];
            }
            // 尝试常见key路径
            foreach (['api_key', 'ANTHROPIC_API_KEY', 'key'] as $kp) {
                $keys = $data;
                foreach (explode('.', $kp) as $part) {
                    if (is_array($keys) && isset($keys[$part])) {
                        $keys = $keys[$part];
                    } else {
                        $keys = null;
                        break;
                    }
                }
                if ($keys) return $keys;
            }
        }
    }

    return '';
}

/**
 * 获取自定义 Base URL
 */
function getBaseUrl(): ?string
{
    $url = getenv('ANTHROPIC_BASE_URL');
    if ($url) return $url;

    $home = getenv('HOME') ?: exec('echo $HOME');
    foreach ([$home . '/.claude/settings.json', $home . '/.claude/settings.local.json'] as $path) {
        if (file_exists($path)) {
            $data = json_decode(file_get_contents($path), true);
            if (isset($data['env']['ANTHROPIC_BASE_URL'])) {
                return $data['env']['ANTHROPIC_BASE_URL'];
            }
        }
    }
    return null;
}

// ===========================================================================
// 第二部分：LLM 调用与解析
// ===========================================================================

/**
 * 调用大模型 API
 *
 * @param string $prompt 给LLM的提示词
 * @param string $model  模型名，默认 MiniMax-M2.7
 * @return array ['result'=>text] 或 ['error'=>message]
 */
function callLlm(string $prompt, string $model = 'MiniMax-M2.7'): array
{
    $apiKey = getApiKey();
    if (!$apiKey) {
        return ['error' => 'API key not found'];
    }

    $baseUrl = getBaseUrl() ?? 'https://api.minimaxi.chat/v1';

    $url = rtrim($baseUrl, '/') . '/text/chatcompletion_v2';

    $payload = json_encode([
        'model' => $model,
        'max_tokens' => 1024,
        'messages' => [
            ['role' => 'user', 'content' => $prompt]
        ]
    ], JSON_UNESCAPED_UNICODE);

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_POSTFIELDS => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER => [
            'Authorization: Bearer ' . $apiKey,
            'Content-Type: application/json',
        ],
        CURLOPT_TIMEOUT => 60,
    ]);

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $curlError = curl_error($ch);
    curl_close($ch);

    if ($curlError) {
        return ['error' => "cURL error: $curlError"];
    }

    if ($httpCode !== 200) {
        return ['error' => "HTTP $httpCode: $response"];
    }

    $data = json_decode($response, true);

    // 解析 MiniMax API 响应格式
    // 实际格式需要根据具体API响应结构调整
    if (isset($data['choices'][0]['message']['content'])) {
        return ['result' => $data['choices'][0]['message']['content']];
    }

    // 兼容其他格式
    if (isset($data['output']['text'])) {
        return ['result' => $data['output']['text']];
    }

    if (isset($data['error'])) {
        return ['error' => is_array($data['error']) ? json_encode($data['error'], JSON_UNESCAPED_UNICODE) : $data['error']];
    }

    return ['result' => $response];
}

/**
 * 从 LLM 返回文本中提取 JSON
 *
 * 处理: markdown代码块、前后空白、嵌套括号等导致json_decode失败的情况
 *
 * @param string $text LLM原始返回文本
 * @return array|null 解析后的数组，或null表示失败
 */
function extractJson(string $text): ?array
{
    // 1. 直接解析
    $data = json_decode($text, true);
    if ($data !== null) return $data;

    // 2. 去除markdown围栏
    $text = preg_replace('/^```json\s*/', '', $text);
    $text = preg_replace('/^```\s*/', '', $text);
    $text = preg_replace('/\`+$/', '', $text);
    $text = trim($text);
    $text = trim($text, '`');

    $data = json_decode($text, true);
    if ($data !== null) return $data;

    // 3. raw_decode 自动查找JSON边界
    try {
        $data = json_decode($text, true, 512, JSON_THROW_ON_ERROR);
        return $data;
    } catch (\JsonException $e) {
        // 继续尝试括号匹配
    }

    // 4. 括号匹配找最外层{}
    $start = strpos($text, '{');
    if ($start === false) return null;

    $len = strlen($text);
    $depth = 0;
    $inString = false;
    $escape = false;

    for ($i = $start; $i < $len; $i++) {
        $ch = $text[$i];

        if ($escape) {
            $escape = false;
            continue;
        }
        if ($ch === '\\') {
            $escape = true;
            continue;
        }
        if ($ch === '"') {
            $inString = !$inString;
            continue;
        }

        if (!$inString) {
            if ($ch === '{') {
                $depth++;
            } elseif ($ch === '}') {
                $depth--;
                if ($depth === 0) {
                    $candidate = substr($text, $start, $i - $start + 1);
                    $data = json_decode($candidate, true);
                    if ($data !== null) return $data;
                    break;
                }
            }
        }
    }

    return null;
}

// ===========================================================================
// 第三部分：后处理与数据清洗
// ===========================================================================

/**
 * 品种/品类后处理
 *
 * 处理逻辑:
 * 1. 非植物商品（一级分类=非植）：品类和品种强制清空
 * 2. 品类：根据商品名关键词映射表推断植物生物学类型
 * 3. 品种：
 *    - 有引号：从引号内提取（排除规格描述性质的片段）
 *    - 无引号：如品种名为纯植物类型名则清空（防止品类当品种）
 *
 * @param string $productName 商品名称
 * @param array  $data        LLM返回的原始数据数组（引用传参，会被直接修改）
 * @return array 修改后的$data
 */
function cleanVariety(string $productName, array &$data): array
{
    $variety = $data['品种'] ?? '';
    $yiji    = $data['一级分类'] ?? '';

    // ---- 非植物：品类+品种必须为空 ----
    if ($yiji === '非植') {
        $data['品类'] = '';
        $data['品种'] = '';
        return $data;
    }

    // ---- 品类推断 ----
    $plantTypeMap = [
        '铁线莲' => '铁线莲', '蓝莓' => '蓝莓', '葡萄' => '葡萄', '杜鹃' => '杜鹃',
        '安酷杜鹃' => '杜鹃', '枫树' => '枫树', '冬青' => '冬青', '绣球' => '绣球',
        '葡萄柚' => '柚', '枇杷' => '枇杷', '月季' => '月季', '玫瑰' => '月季', '草莓' => '草莓',
        '梨' => '梨', '桃' => '桃', '苹果' => '苹果', '黑莓' => '黑莓',
        '紫菀' => '紫菀', '百子莲' => '百子莲', '风雨兰' => '风雨兰',
        '香檬' => '香檬', '蜜桔' => '蜜桔', '金桔' => '金桔', '丁香' => '丁香',
        '巴布豆' => '巴布豆',
    ];

    // ---- 品种提取（先于品类推断） ----
    // 从中文/英文引号内提取品种名
    preg_match_all('/[""\u201c\u201d]([^""\u201c\u201d]+)[""\u201c\u201d]/u', $productName, $matches);
    $quoted = $matches[1] ?? [];
    $varietyExtracted = null;

    if (!empty($quoted)) {
        // 过滤掉规格描述性片段（含数字年/球/号/L/kg/代工/混苗等）
        $parts = [];
        foreach ($quoted as $q) {
            $q = trim($q);
            if (!preg_match('/\d年|\d球|\d号|\dL|\dkg|代工|混苗/', $q)) {
                $parts[] = $q;
            }
        }
        if (!empty($parts)) {
            usort($parts, function($a, $b) { return mb_strlen($a) - mb_strlen($b); });
            $varietyExtracted = $parts[0];
        } else {
            $varietyExtracted = trim(end($quoted));
        }
        if ($variety !== $varietyExtracted) {
            $data['品种'] = $varietyExtracted;
        }
    } else {
        // 无引号：品种若为纯植物类型名则清空
        if ($variety) {
            $purePlantTypes = ['铁线莲', '葡萄', '杜鹃', '枫树', '冬青',
                             '绣球', '月季', '玫瑰', '草莓', '梨', '桃',
                             '苹果', '黑莓', '柚', '草本', '灌木', '乔木', '藤本'];
            if (in_array(trim($variety), $purePlantTypes, true)) {
                $data['品种'] = '';
            }
            if (trim($variety) === '蓝葡萄' && mb_strpos($productName, '蓝莓') !== false) {
                $data['品种'] = '';
            }
        }
    }

    // ---- 品类推断（品种提取成功时跳过，避免品种名含关键词被错误归类） ----
    // 只有当品种无法从引号提取时，才用关键词映射推断品类
    // 防止"葡萄太妃糖"→品类=葡萄 这种误判
    if ($varietyExtracted === null) {
        foreach ($plantTypeMap as $keyword => $cat) {
            if (mb_strpos($productName, $keyword) !== false) {
                $data['品类'] = $cat;
                break;
            }
        }
    }

    return $data;
}

// ===========================================================================
// 第四部分：核心补全逻辑
// ===========================================================================

/**
 * 商品信息补全核心函数
 *
 * 流程:
 * 1. 构建包含分类目录和植物学规则的prompt
 * 2. 调用LLM获取JSON格式的分类结果
 * 3. 解析JSON（extractJson）
 * 4. 后处理：品种提取、品类推断、品牌归一化、品种防污染
 *
 * @param string $productName 商品名称
 * @param string $specInfo    规格信息（可为空字符串）
 * @param array  $categories  loadCategories() 返回的分类列表
 * @return array  包含8个字段的数组，或 ['error'=>...] 表示失败
 */
function enrichProduct(string $productName, string $specInfo, array $categories): array
{
    // 构建分类目录行列表
    $categoryLines = [];
    foreach ($categories as $cat) {
        $categoryLines[] = sprintf(
            '%s > %s > %s > %s',
            $cat['一级分类'], $cat['二级分类'], $cat['三级分类'], $cat['四级分类']
        );
    }

    // 构建分类树摘要
    $tree = buildCategoryTree($categories);
    $categorySummary = [];
    foreach ($tree as $l1 => $l2Dict) {
        foreach ($l2Dict as $l2 => $l3Dict) {
            $l3List = array_keys($l3Dict);
            $l4Counts = array_sum(array_map('count', $l3Dict));
            $categorySummary[] = "  $l1 > $l2: " . json_encode($l3List, JSON_UNESCAPED_UNICODE) . " (四级分类数: $l4Counts)";
        }
    }

    // 组装 prompt
    $prompt = <<<PROMPT
你是一个专业的园艺商品分类专家，精通植物分类学和园艺植物学。

## 输入信息
商品名称：{$productName}
规格信息：{$specInfo}

## 分类目录（商品分类目录.xlsx）
一级分类 > 二级分类 > 三级分类 > 四级分类
每行的分类层级结构如下：
PROMPT;

    foreach ($categoryLines as $line) {
        $prompt .= "\n- " . $line;
    }

    $prompt .= <<<'PROMPT'

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
- **冬青** → 四级分类为「冬青科」，三级分类为「灌木」
- **紫菀** → 四级分类为「菊科」，三级分类为「多年生草本」
- **风雨兰** → 四级分类为「石蒜科」，三级分类为「鳞茎」

**品牌、品类与品种说明：**
- **品牌**：指育种公司或苗圃品牌（如：玫昂、美乐棵、花彩师、德沃多等）。**必须输出品牌**，即使商品是通用类型无明确品牌，也应输出常见园艺品牌作为默认值。**安酷杜鹃（Encore Azaleas）品牌的商品，品牌统一输出为"安酷"**，不得输出为"HY"或"HY-"
- **品类**：指植物的生物学类型名称，如铁线莲、蓝莓、葡萄、杜鹃、枫树、冬青、绣球、枇杷、月季、草莓等。**所有植物商品必须输出品类**。对于植物，即使商品名没有明确写品类，也要根据植物学分类推断出正确的品类
- **品种**：指具体植物品种名称（如：龙沙宝石、追雪、乌托邦、阳光玫瑰等）。**品种名必须是从商品名引号内提取的短名称（2-10字）**，禁止输出包含"科""属""是..."等描述性文字。如果商品名没有引号包裹具体品种名，品种字段为空字符串""。**禁止**把品类名（如"蓝莓"、"铁线莲"、"葡萄"）当作品种输出！

**输出格式（JSON）：**
{
    "一级分类": "...",
    "二级分类": "...",
    "三级分类": "...",
    "四级分类": "...",
    "品牌": "育种公司或品牌名",
    "品类": "植物类型名（如铁线莲、蓝莓等）",
    "品种": "具体品种名称（如有）",
    "百科信息": "商品的简要描述、特点、用途、养护要点等（50-200字）"
}

请严格按照JSON格式输出，不要包含其他文字。
PROMPT;

    // ---- 调用LLM ----
    $result = callLlm($prompt);

    if (isset($result['error'])) {
        return $result;
    }

    // ---- 解析JSON ----
    $text = $result['result'] ?? '';
    $data = extractJson($text);

    if ($data === null) {
        return ['error' => '无法解析LLM返回结果', 'raw' => mb_substr($text, 0, 500)];
    }

    // ---- 后处理 ----
    // 1. 品种提取+品类推断
    $data = cleanVariety($productName, $data);

    // 2. 品牌归一化：HY/Encore统一为"安酷"
    $brand = $data['品牌'] ?? '';
    if ($brand && in_array(trim($brand), ['HY', 'Encore Azaleas（安酷杜鹃）', 'Encore Azaleas', 'HY-'], true)) {
        $data['品牌'] = '安酷';
    }

    // 3. 品种防污染：品种列只允许2-20字的纯品种名
    //    禁止含"科""属"或超长的百科式描述
    $variety = $data['品种'] ?? '';
    if ($variety && (mb_strlen($variety) > 20 || mb_strpos($variety, '科') !== false || mb_strpos($variety, '属') !== false)) {
        $data['品种'] = '';
    }

    return $data;
}

// ===========================================================================
// 第五部分：单条测试入口
// ===========================================================================

/**
 * 命令行单条测试模式
 */
function main(): void
{
    array_shift($GLOBALS['argv']); // 去掉脚本名

    if (empty($GLOBALS['argv'])) {
        echo "使用方法: php product_enricher.php <商品名称> [规格信息]\n";
        echo "示例: php product_enricher.php 月季花苗 一年生\n";
        echo "示例: php product_enricher.php 有机营养土 10L装\n";
        exit(1);
    }

    $productName = $GLOBALS['argv'][0];
    $specInfo = $GLOBALS['argv'][1] ?? '';

    // 查找商品分类目录.xlsx
    $scriptDir = __DIR__;
    $home = getenv('HOME') ?: exec('echo $HOME');
    $searchPaths = [
        $scriptDir . '/商品分类目录.xlsx',
        $home . '/Desktop/沈飞mac/skills/shopping/商品分类目录.xlsx',
        $home . '/Desktop/商品分类目录.xlsx',
        getcwd() . '/商品分类目录.xlsx',
    ];

    $excelPath = null;
    foreach ($searchPaths as $path) {
        if (file_exists($path)) {
            $excelPath = $path;
            break;
        }
    }

    if (!$excelPath) {
        echo "错误: 找不到商品分类目录.xlsx文件\n";
        exit(1);
    }

    echo "加载分类目录: $excelPath\n";
    $categories = loadCategories($excelPath);
    echo "共加载 " . count($categories) . " 条分类记录\n\n";

    echo "正在分析商品: $productName $specInfo\n";
    $result = enrichProduct($productName, $specInfo, $categories);

    if (isset($result['error'])) {
        echo "\n错误: {$result['error']}\n";
        if (isset($result['raw'])) {
            echo "原始返回: {$result['raw']}\n";
        }
        exit(1);
    }

    echo "\n=== 补全结果 ===\n";
    echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE) . "\n";
}

// ---------------------------------------------------------------------------
// 自动加载 Composer 依赖（如果存在）
// ---------------------------------------------------------------------------
$autoloadPaths = [
    __DIR__ . '/../vendor/autoload.php',
    __DIR__ . '/vendor/autoload.php',
    __DIR__ . '/composer.json',  // 触发 composer autoload
];
foreach ($autoloadPaths as $path) {
    if (file_exists($path)) {
        if (strpos($path, 'autoload.php') !== false) {
            require_once $path;
        }
        break;
    }
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------
if (php_sapi_name() === 'cli' && realpath($GLOBALS['argv'][0] ?? '') === realpath(__FILE__)) {
    main();
}
