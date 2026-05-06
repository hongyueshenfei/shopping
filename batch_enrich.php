#!/usr/bin/env php
<?php
/**
 * Batch Enrich - 批量商品信息补全脚本（PHP版）
 *
 * 从 Excel 读取商品名称列表，调用 enrichProduct() 逐条补全分类/品牌/品类/品种/百科信息，
 * 将结果追加写入原文件的右侧新列（每次运行会覆盖之前的补全列）。
 *
 * 输出文件: <原文件名>_补全.xlsx（如 1年内销售库存大于10育种+上海_补全.xlsx）
 *
 * 使用方法:
 *   php batch_enrich.php
 *
 * 依赖:
 *   PHP 7.4+, curl, json, mbstring
 *   composer require phpoffice/phpspreadsheet
 */

// ===========================================================================
// 配置
// ===========================================================================

// 源文件名（相对于脚本目录）
define('SOURCE_FILE', '1年内销售库存大于10育种+上海.xlsx');
define('CATEGORY_FILE', '商品分类目录.xlsx');

// ===========================================================================
// 工具函数
// ===========================================================================

/**
 * 加载 Excel 文件（使用 PhpSpreadsheet）
 *
 * @param string $path
 * @return array 二维数组，第一行是表头
 */
function loadExcel(string $path): array
{
    if (!class_exists('\PhpOffice\PhpSpreadsheet\IOFactory')) {
        die("错误: 需要安装 phpoffice/phpspreadsheet\n运行: composer require phpoffice/phpspreadsheet\n");
    }
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
    return $spreadsheet->getActiveSheet()->toArray();
}

/**
 * 保存 Excel 文件（使用 PhpSpreadsheet）
 *
 * @param array  $data     二维数组
 * @param string $path     输出路径
 * @param array  $headers  表头数组
 */
function saveExcel(array $data, string $path, array $headers): void
{
    if (!class_exists('\PhpOffice\PhpSpreadsheet\IOFactory')) {
        die("错误: 需要安装 phpoffice/phpspreadsheet\n");
    }

    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // 写入表头
    foreach ($headers as $col => $header) {
        $sheet->setCellValueByColumnAndRow($col + 1, 1, $header);
    }

    // 写入数据
    foreach ($data as $rowIdx => $row) {
        foreach ($row as $colIdx => $value) {
            $sheet->setCellValueByColumnAndRow($colIdx + 1, $rowIdx + 2, $value ?? '');
        }
    }

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($path);
}

/**
 * 查找源 Excel 文件路径
 */
function findSourceFile(): ?string
{
    $scriptDir = __DIR__;
    $home = getenv('HOME') ?: exec('echo $HOME');

    $candidates = [
        $scriptDir . '/' . SOURCE_FILE,
        $home . '/Desktop/沈飞mac/skills/shopping/' . SOURCE_FILE,
    ];

    foreach ($candidates as $p) {
        if (file_exists($p)) return $p;
    }
    return null;
}

/**
 * 查找分类目录文件
 */
function findCategoryFile(): ?string
{
    $scriptDir = __DIR__;
    $home = getenv('HOME') ?: exec('echo $HOME');

    $candidates = [
        $scriptDir . '/' . CATEGORY_FILE,
        $home . '/Desktop/沈飞mac/skills/shopping/' . CATEGORY_FILE,
    ];

    foreach ($candidates as $p) {
        if (file_exists($p)) return $p;
    }
    return null;
}

/**
 * 生成输出文件路径
 */
function getOutputPath(string $inputPath): string
{
    $info = pathinfo($inputPath);
    $dir = $info['dirname'] ?? '.';
    $name = $info['filename'];
    $ext = $info['extension'] ?? 'xlsx';
    return $dir . '/' . $name . '_补全.' . $ext;
}

/**
 * 打印进度条
 */
function progressBar(float $pct, int $current, int $total, string $extra = ''): void
{
    $barLen = 30;
    $filled = (int)($barLen * $pct / 100);
    $bar = str_repeat('█', $filled) . str_repeat('░', $barLen - $filled);
    $msg = sprintf("[%s] %.1f%% (%d/%d) %s", $bar, $pct, $current, $total, $extra);
    echo "\r" . str_pad($msg, 120) . str_repeat(' ', 20);
}

/**
 * 自动加载依赖
 */
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
// 主流程
// ===========================================================================

function main(): void
{
    autoload();

    // ---- 1. 定位源文件 ----
    $excelPath = findSourceFile();
    if (!$excelPath) {
        echo "错误: 找不到 " . SOURCE_FILE . "\n";
        exit(1);
    }

    echo "读取文件: $excelPath\n";
    $allRows = loadExcel($excelPath);
    $headers = $allRows[0] ?? [];
    $totalRows = count($allRows) - 1; // 减去表头
    echo "共 $totalRows 条记录待处理\n\n";

    // ---- 2. 加载分类目录 ----
    $catPath = findCategoryFile();
    if (!$catPath) {
        echo "错误: 找不到 " . CATEGORY_FILE . "\n";
        exit(1);
    }

    echo "加载分类目录: $catPath\n";
    $categories = loadCategories($catPath);
    echo "共 " . count($categories) . " 条分类记录\n\n";

    // ---- 3. 定义输出列 ----
    $headerCols = [
        'goods_no', 'goods_name', 'department_name', 'department_code',
        'all_qty', 'all_amount', '结存数量'
    ];
    $enrichCols = [
        '一级分类', '二级分类', '三级分类', '四级分类',
        '品牌', '品类', '品种', '百科信息'
    ];
    $allHeaders = array_merge($headerCols, $enrichCols);

    // ---- 4. 逐行处理 ----
    $successCount = 0;
    $errorCount = 0;

    // 结果数组：第一行是表头
    $results = [$allHeaders];

    for ($i = 1; $i <= $totalRows; $i++) {
        $row = $allRows[$i] ?? [];
        $goodsName = $row[1] ?? '';  // 第2列（B列）

        if (empty($goodsName)) {
            $results[] = array_fill(0, count($headerCols) + count($enrichCols), '');
            continue;
        }

        // 进度条
        $current = $i;
        $pct = $current / $totalRows * 100;
        progressBar($pct, $current, $totalRows, mb_substr($goodsName, 0, 20));

        // 调用核心补全函数
        $enriched = enrichProduct($goodsName, '', $categories);

        if (isset($enriched['error'])) {
            $errorCount++;
            $enrichValues = array_fill(0, count($enrichCols), '');
        } else {
            $successCount++;
            $enrichValues = [];
            foreach ($enrichCols as $col) {
                $enrichValues[] = $enriched[$col] ?? '';
            }
        }

        // 拼接原始列 + 补全列
        $fullRow = array_merge(
            array_slice($row, 0, count($headerCols)),
            $enrichValues
        );
        $results[] = $fullRow;

        // 每50行保存一次checkpoint
        if ($i % 50 === 0) {
            $outPath = getOutputPath($excelPath);
            saveExcel($results, $outPath, $allHeaders);
        }
    }

    echo "\n\n处理完成！成功: $successCount, 失败: $errorCount\n";

    // ---- 5. 最终保存 ----
    $outPath = getOutputPath($excelPath);
    saveExcel($results, $outPath, $allHeaders);
    echo "结果已保存: $outPath\n";
}

// ===========================================================================
// 入口
// ===========================================================================
main();
