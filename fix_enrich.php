#!/usr/bin/env php
<?php
/**
 * Fix Enrich - 批量修复补全文件中问题行的脚本（PHP版）
 *
 * 扫描已有的 _补全.xlsx 文件，根据预定义规则自动识别有问题的行，
 * 重新调用 enrichProduct() 进行补全，覆盖原数据。
 *
 * 检查规则:
 *   1. 铁线莲三级必须是「木本藤本」（不是草本藤本）
 *   2. 绣球四级必须是「绣球花科」（不是绣球科）
 *   3. 枫树四级必须是「槭树科」（不是槒树科）
 *   4. 百子莲三级必须是「鳞茎」（不是多年生草本）
 *   5. 品牌HY/HY-/Encore Azaleas → 统一为「安酷」
 *   6. 一级~四级分类/品牌完全缺失 → 重新补全
 *
 * 使用方法:
 *   php fix_enrich.php
 */

define('SOURCE_FILE', '1年内销售库存大于10育种+上海_补全.xlsx');
define('CATEGORY_FILE', '商品分类目录.xlsx');

// ===========================================================================
// 工具函数（与 batch_enrich.php 共享）
// ===========================================================================

function loadCategories(string $path): array
{
    if (!class_exists('\PhpOffice\PhpSpreadsheet\IOFactory')) {
        die("错误: 需要安装 phpoffice/phpspreadsheet\n");
    }
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
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

function loadExcel(string $path): array
{
    if (!class_exists('\PhpOffice\PhpSpreadsheet\IOFactory')) {
        die("错误: 需要安装 phpoffice/phpspreadsheet\n");
    }
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
    return $spreadsheet->getActiveSheet()->toArray();
}

function saveExcelWithPhpSpreadsheet(array $data, string $path): void
{
    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    foreach ($data as $rowIdx => $row) {
        foreach ($row as $colIdx => $value) {
            $sheet->setCellValueByColumnAndRow($colIdx + 1, $rowIdx + 1, $value ?? '');
        }
    }

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($path);
}

function findFile(string $filename): ?string
{
    $scriptDir = __DIR__;
    $home = getenv('HOME') ?: exec('echo $HOME');

    $candidates = [
        $scriptDir . '/' . $filename,
        $home . '/Desktop/沈飞mac/skills/shopping/' . $filename,
    ];

    foreach ($candidates as $p) {
        if (file_exists($p)) return $p;
    }
    return null;
}

function autoload(): void
{
    $paths = [
        __DIR__ . '/vendor/autoload.php',
        dirname(__DIR__) . '/vendor/autoload.php',
    ];
    foreach ($paths as $p) {
        if (file_exists($p)) {
            require_once $p;
            return;
        }
    }
}

// ===========================================================================
// 问题行检查规则
// ===========================================================================

/**
 * 检查指定行是否有已知问题
 *
 * @param int    $rowIdx  Excel行号（从2开始，数据行）
 * @param string $name    商品名
 * @param string $yiji    一级分类
 * @param string $erji    二级分类
 * @param string $sanji   三级分类
 * @param string $siji    四级分类
 * @param string $brand   品牌
 * @param string $variety 品种
 *
 * @return array [isProblematic: bool, reason: string]
 */
function checkRow(int $rowIdx, string $name, string $yiji, string $erji, string $sanji, string $siji, string $brand, string $variety): array
{
    // ---- 规则1: 铁线莲三级分类 ----
    // 铁线莲是木本藤本，不是草本藤本
    if (mb_strpos($name, '铁线莲') !== false && $sanji === '草本藤本') {
        return [true, '铁线莲三级分类错误（草本藤本→木本藤本）'];
    }

    // ---- 规则2: 绣球四级分类 ----
    // 绣球的科名是「绣球花科」，不是「绣球科」
    if ($siji === '绣球科') {
        return [true, "绣球四级分类错误（{$siji}→绣球花科）"];
    }

    // ---- 规则3: 枫树四级分类 ----
    // 枫树的科名是「槭树科」，不是「槒树科」
    if ($siji === '槒树科') {
        return [true, "枫树四级分类错误（{$siji}→槭树科）"];
    }

    // ---- 规则4: 百子莲三级分类 ----
    // 百子莲是鳞茎植物，不是多年生草本
    if (mb_strpos($name, '百子莲') !== false && $sanji === '多年生草本') {
        return [true, '百子莲三级分类错误（多年生草本→鳞茎）'];
    }

    // ---- 规则5: 安酷杜鹃品牌名称归一 ----
    // Encore Azaleas / HY / HY- 统一输出为「安酷」
    $brandTrim = trim($brand);
    if ($brandTrim && (
        in_array($brandTrim, ['HY', 'Encore Azaleas（安酷杜鹃）', 'Encore Azaleas'], true) ||
        strpos($brandTrim, 'HY') === 0
    )) {
        return [true, "品牌写法错误（{$brand}→安酷）"];
    }

    // ---- 规则6: 数据完全缺失 ----
    // 一级分类为空说明该行从未成功补全
    if ($yiji === null || $yiji === '') {
        return [true, '数据完全缺失'];
    }

    // 无问题
    return [false, ''];
}

// ===========================================================================
// 主流程
// ===========================================================================

function main(): void
{
    autoload();

    // ---- 1. 定位补全文件 ----
    $excelPath = findFile(SOURCE_FILE);
    if (!$excelPath) {
        echo "文件不存在: " . SOURCE_FILE . "\n";
        exit(1);
    }

    echo "读取文件: $excelPath\n";
    $allRows = loadExcel($excelPath);
    $totalRows = count($allRows) - 1;
    echo "总行数: $totalRows\n\n";

    // ---- 2. 加载分类目录 ----
    $catPath = findFile(CATEGORY_FILE);
    if (!$catPath) {
        echo "错误: 找不到 " . CATEGORY_FILE . "\n";
        exit(1);
    }

    echo "加载分类目录: $catPath\n";
    $categories = loadCategories($catPath);
    echo "共 " . count($categories) . " 条分类记录\n\n";

    // ---- 3. 全量扫描，收集问题行 ----
    echo "=== 检查所有行 ===\n";
    $problematic = [];

    for ($rowIdx = 2; $rowIdx <= count($allRows); $rowIdx++) {
        $row = $allRows[$rowIdx - 1] ?? [];
        $goodsName = $row[1] ?? '';   // B列: goods_name
        $yiji     = $row[7] ?? '';   // H列: 一级分类
        $erji     = $row[8] ?? '';   // I列: 二级分类
        $sanji    = $row[9] ?? '';   // J列: 三级分类
        $siji     = $row[10] ?? '';  // K列: 四级分类
        $brand    = $row[11] ?? '';  // L列: 品牌
        $variety  = $row[13] ?? '';  // N列: 品种

        [$isBad, $reason] = checkRow($rowIdx, $goodsName, $yiji, $erji, $sanji, $siji, $brand, $variety);

        if ($isBad) {
            $problematic[] = [
                'row' => $rowIdx,
                'name' => $goodsName,
                'reason' => $reason,
                '当前' => "三级={$sanji}, 四级={$siji}, 品牌={$brand}",
            ];
        }
    }

    echo "发现问题行: " . count($problematic) . " 条\n\n";
    foreach ($problematic as $p) {
        echo "  Row {$p['row']} [" . mb_substr($p['name'], 0, 30) . "]\n";
        echo "    原因: {$p['reason']}\n";
        echo "    当前: {$p['当前']}\n\n";
    }

    if (empty($problematic)) {
        echo "没有发现问题行，无需修复\n";
        exit(0);
    }

    // ---- 4. 逐行重新补全 ----
    echo "开始重新补全 " . count($problematic) . " 条...\n\n";
    $fixedCount = 0;
    $errorCount = 0;
    $total = count($problematic);

    foreach ($problematic as $idx => $item) {
        $rowIdx = $item['row'];
        $goodsName = $item['name'];

        // 进度条
        $current = $idx + 1;
        $pct = $current / $total * 100;
        $barLen = 30;
        $filled = (int)($barLen * $current / $total);
        $bar = str_repeat('█', $filled) . str_repeat('░', $barLen - $filled);
        $extra = mb_substr($goodsName, 0, 25);
        echo "\r[{$bar}] " . sprintf("%.1f%% (%d/%d) - %s", $pct, $current, $total, $extra);
        echo str_repeat(' ', 20);

        $enriched = enrichProduct($goodsName, '', $categories);

        if (isset($enriched['error'])) {
            $errorCount++;
            echo "\n  错误: {$enriched['error']}\n";
        } else {
            $fixedCount++;
            // 更新数据行中的8个补全列（第8列N~第15列O）
            $enrichCols = ['一级分类', '二级分类', '三级分类', '四级分类', '品牌', '品类', '品种', '百科信息'];
            $rowNum = $rowIdx; // 行号从1开始，$rowIdx就是数据行号

            foreach ($enrichCols as $colIdx => $colName) {
                // 列号: 一级分类=H(8), 二级分类=I(9), ... 百科信息=O(15)
                $colLetter = chr(ord('H') + $colIdx); // H=8, I=9, ...
                $allRows[$rowNum - 1][7 + $colIdx] = $enriched[$colName] ?? '';
            }
        }

        // 每10行保存一次
        if ($current % 10 === 0) {
            saveExcelWithPhpSpreadsheet($allRows, $excelPath);
        }
    }

    echo "\n\n修复完成！成功: $fixedCount, 失败: $errorCount\n";
    saveExcelWithPhpSpreadsheet($allRows, $excelPath);
    echo "已保存: $excelPath\n";
}

// ===========================================================================
// 入口
// ===========================================================================
main();
