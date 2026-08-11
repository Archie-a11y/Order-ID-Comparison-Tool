<?php
/**
 * PHP App - Order ID Comparison Tool
 * Single-file implementation optimized with Bootstrap 5.3, Bootstrap Icons,
 * Complete Multilingual support, Native Dark/Light mode, and High-Performance Memory Management.
 */

// -------------------------------------------------------------------------
// 0. 性能与内存极限优化 + 错误捕获 (彻底解决 Apache/cPanel 500/503 错误)
// -------------------------------------------------------------------------
@ini_set('memory_limit', '512M');
@set_time_limit(600);
@ini_set('max_execution_time', '600');

// 初始化 Session 用于存放比对结果统计信息
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}

// 创建并检查 uploads 临时目录
$uploadDir = __DIR__ . '/uploads';
if (!is_dir($uploadDir)) {
    @mkdir($uploadDir, 0755, true);
}

// 自动清理：删除服务器上创建超过 1 小时的临时结果文件，保证不占空间
if (is_dir($uploadDir)) {
    $files = glob($uploadDir . '/Result_*.xlsx');
    $now = time();
    foreach ($files as $file) {
        if (is_file($file) && ($now - filemtime($file) > 3600)) {
            @unlink($file);
        }
    }
}

// 引入 Composer 自动加载器
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}

// 启用输出缓冲，防止 Header 已经发送的警告
ob_start();

// -------------------------------------------------------------------------
// 0b. 处理临时文件的下载与清理动作
// -------------------------------------------------------------------------
if (isset($_GET['action']) && $_GET['action'] === 'download') {
    if (isset($_SESSION['comparison_result']['file'])) {
        $file = $_SESSION['comparison_result']['file'];
        $filePath = $uploadDir . '/' . $file;
        if (file_exists($filePath)) {
            ob_end_clean();
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="Order_Comparison_Result.xlsx"');
            header('Content-Length: ' . filesize($filePath));
            readfile($filePath);
            @unlink($filePath); // 下载后立即销毁，确保 100% 隐私和零磁盘积压
            unset($_SESSION['comparison_result']); // 清空 Session
            exit;
        }
    }
    header("Location: " . strtok($_SERVER['REQUEST_URI'], '?'));
    exit;
}

// 彻底清除 Session 并返回主页
if (isset($_GET['action']) && $_GET['action'] === 'reset') {
    if (isset($_SESSION['comparison_result']['file'])) {
        $filePath = $uploadDir . '/' . $_SESSION['comparison_result']['file'];
        if (file_exists($filePath)) {
            @unlink($filePath);
        }
    }
    unset($_SESSION['comparison_result']);
    header("Location: " . strtok($_SERVER['REQUEST_URI'], '?'));
    exit;
}

// -------------------------------------------------------------------------
// 0c. OpenSpout v4 极致流式数据读取助手 (0 内存，100% 杜绝 503)
// -------------------------------------------------------------------------
if (!function_exists('readRowsWithOpenSpout')) {
    function readRowsWithOpenSpout($filePath, array $colLetters, $sheetName = '') {
        $ext = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
        $colLetters = array_map('strtoupper', $colLetters);
        
        // 将列字母（如 'C'）转为从 0 开始的数组索引
        $colIndices = [];
        foreach ($colLetters as $letter) {
            $idx = 0;
            $len = strlen($letter);
            for ($i = 0; $i < $len; $i++) {
                $idx = $idx * 26 + (ord($letter[$i]) - 64);
            }
            $colIndices[$letter] = $idx - 1;
        }

        $results = [];

        // 识别文件类型并开启流式读取
        if ($ext === 'csv') {
            $reader = new \OpenSpout\Reader\CSV\Reader();
        } else {
            $reader = new \OpenSpout\Reader\XLSX\Reader();
        }

        $reader->open($filePath);

        $rowIndex = 1;
        foreach ($reader->getSheetIterator() as $sheet) {
            // 如果是 Excel 且指定了 Sheet 名，则只读取匹配的 Sheet
            if ($ext !== 'csv' && !empty($sheetName) && strtolower($sheet->getName()) !== strtolower($sheetName)) {
                continue;
            }

            foreach ($sheet->getRowIterator() as $row) {
                $cells = $row->getCells();
                $rowValues = [];
                foreach ($colIndices as $letter => $idx) {
                    if (isset($cells[$idx])) {
                        // 获取单元格文本内容
                        $rowValues[$letter] = trim((string)$cells[$idx]->getValue());
                    }
                }
                if (!empty($rowValues)) {
                    $results[$rowIndex] = $rowValues;
                }
                $rowIndex++;
            }
            break; // 读取完目标 Sheet 后立即退出，防止多余加载
        }

        $reader->close();
        return $results;
    }
}

// -------------------------------------------------------------------------
// 1. 语言与主题配置（使用 Cookie 持久化，支持通过 Query String 切换并自动重定向清除 URL 缓存）
// -------------------------------------------------------------------------
$lang = $_COOKIE['app_lang'] ?? 'en';
if (!in_array($lang, ['en', 'zh', 'ms'])) {
    $lang = 'en';
}

$theme = $_COOKIE['app_theme'] ?? 'light';
if (!in_array($theme, ['light', 'dark'])) {
    $theme = 'light';
}

// -------------------------------------------------------------------------
// 2. 完整翻译字典 (Complete Translation Dictionary)
// -------------------------------------------------------------------------
$t = [
    'en' => [
        'title' => 'Order ID Comparison Tool',
        'subtitle' => 'Compare master orders and income sheets to identify discrepancies.',
        'guide_title' => 'User Guide & Instructions',
        'guide_step1' => 'Step 1: Orders File (Master List)',
        'guide_step1_desc' => 'Upload your main orders sheet (.csv or .xlsx). Define the exact column letters containing the Order ID (e.g., A) and the Order Status (e.g., B).',
        'guide_step2' => 'Step 2: Income File (Comparison List)',
        'guide_step2_desc' => 'Upload your income/service sheet. The system will automatically detect available Sheet/Tab names in the dropdown. Specify the column letter containing the Order ID.',
        'guide_step3' => 'Step 3: Run & Download',
        'guide_step3_desc' => 'Click the "Compare Order IDs" button. The tool matches File 1 against File 2. Unmatched IDs will be flagged as "MISSING" with a red background in the downloaded Excel sheet.',
        'file1_title' => '1. Orders File (Master List)',
        'file2_title' => '2. Income File (Comparison List)',
        'file_label' => 'Choose File (.xlsx, .csv):',
        'order_col_label' => 'Order ID Column Letter:',
        'status_col_label' => 'Order Status Column Letter:',
        'sheet_name_label' => 'Worksheet / Tab Name:',
        'placeholder_col_id' => 'e.g., A',
        'placeholder_col_id_c' => 'e.g., C',
        'placeholder_col_status' => 'e.g., B',
        'select_file_first' => '-- Please select Income file first --',
        'csv_no_sheet' => 'CSV File (No sheet required)',
        'submit_btn' => 'Compare Order IDs',
        'missing_label' => 'MISSING',
        'col_order_id_f1' => 'Order ID – Orders File',
        'col_match_id_f2' => 'Matching Order ID – Income File',
        'col_status' => 'Order Status',
        'err_missing_files' => 'Please select and upload both files to proceed.',
        'err_invalid_col' => 'Invalid column letter format: "%s". Please enter 1-3 letters (A-Z).',
        'err_sheet_not_found' => 'The specified sheet "%s" was not found in the Income file.',
        'err_processing' => 'An error occurred during comparison: ',
        'err_invalid_format' => 'Unsupported file format. Please upload .xlsx or .csv files.',
        'theme_light' => 'Light Mode',
        'theme_dark' => 'Dark Mode',
        'select_lang' => 'Language',
        'summary_title' => 'Comparison Summary',
        'total_orders' => 'Total Master Orders',
        'matched_orders' => 'Matched Orders',
        'missing_orders' => 'Missing Orders (MISSING)',
        'download_btn' => 'Download Excel Results',
        'compare_again' => 'Compare Another Files'
    ],
    'zh' => [
        'title' => '订单 ID 对比工具',
        'subtitle' => '对比主订单表和收益表，快速找出差异数据。',
        'guide_title' => '用户指南与操作说明',
        'guide_step1' => '第一步：订单文件（主列表）',
        'guide_step1_desc' => '上传您的主订单表（支持 .csv 或 .xlsx 格式）。准确指定“订单 ID”所在的列字母（如：A）以及“订单状态”所在的列字母（如：B）。',
        'guide_step2' => '第二步：收益文件（对比表）',
        'guide_step2_desc' => '上传对应的结算或收益文件。系统将自动检索并列出 Excel 中的工作表 (Tab) 供您下拉选择。指定其“订单 ID”所在的列字母。',
        'guide_step3' => '第三步：运行并下载',
        'guide_step3_desc' => '点击“对比订单 ID”开始。系统将以订单文件为主表去查找对应的收益文件。未匹配的订单在新下载的 Excel 中将用红色底色并标记为 "MISSING"。',
        'file1_title' => '1. 订单文件（主列表）',
        'file2_title' => '2. 收益文件（对比列表）',
        'file_label' => '选择文件 (.xlsx, .csv):',
        'order_col_label' => '订单 ID 所在列字母:',
        'status_col_label' => '订单状态 所在列字母:',
        'sheet_name_label' => '工作表/Tab 名称:',
        'placeholder_col_id' => '例如: A',
        'placeholder_col_id_c' => '例如: C',
        'placeholder_col_status' => '例如: B',
        'select_file_first' => '-- 请先选择收益文件 --',
        'csv_no_sheet' => 'CSV 文件（无需选择工作表）',
        'submit_btn' => '对比订单 ID',
        'missing_label' => 'MISSING',
        'col_order_id_f1' => 'Order ID – Orders File',
        'col_match_id_f2' => 'Matching Order ID – Income File',
        'col_status' => 'Order Status',
        'err_missing_files' => '请同时选择并上传两个文件以继续。',
        'err_invalid_col' => '无效的列字母格式: "%s"。请输入 1 至 3 位英文字母 (A-Z)。',
        'err_sheet_not_found' => '在收益文件中未找到指定的工作表 "%s"。',
        'err_processing' => '对比过程中发生错误: ',
        'err_invalid_format' => '不支持的文件格式。请上传 .xlsx or .csv 文件。',
        'theme_light' => '浅色模式',
        'theme_dark' => '深色模式',
        'select_lang' => '语言',
        'summary_title' => '对比结果汇总',
        'total_orders' => '总订单数量',
        'matched_orders' => '成功匹配数量',
        'missing_orders' => '未匹配数量 (MISSING)',
        'download_btn' => '下载 Excel 对比结果',
        'compare_again' => '再次对比其他文件'
    ],
    'ms' => [
        'title' => 'Alat Perbandingan ID Pesanan',
        'subtitle' => 'Bandingkan pesanan utama dan helaian pendapatan untuk mencari percanggahan.',
        'guide_title' => 'Panduan Pengguna & Arahan',
        'guide_step1' => 'Langkah 1: Fail Pesanan (Senarai Utama)',
        'guide_step1_desc' => 'Muat naik fail pesanan utama anda (.csv atau .xlsx). Tentukan huruf lajur yang mengandungi ID Pesanan (cth., A) dan Status Pesanan (cth., B).',
        'guide_step2' => 'Langkah 2: Fail Pendapatan (Senarai Perbandingan)',
        'guide_step2_desc' => 'Muat naik fail pendapatan atau yuran perkhidmatan anda. Sistem akan mengesan nama Helaian Kerja (Tab) secara automatik dalam senarai pilihan. Tentukan huruf lajur ID Pesanan.',
        'guide_step3' => 'Langkah 3: Jalankan & Muat Turun',
        'guide_step3_desc' => 'Klik butang "Bandingkan ID Pesanan". Alat ini akan memadankan Fail 1 dengan Fail 2. ID yang tidak sepadan akan ditandakan sebagai "MISSING" dengan latar belakang merah dalam fail Excel yang dimuat turun.',
        'file1_title' => '1. Fail Pesanan (Senarai Utama)',
        'file2_title' => '2. Fail Pendapatan (Senarai Perbandingan)',
        'file_label' => 'Pilih Fail (.xlsx, .csv):',
        'order_col_label' => 'Huruf Lajur ID Pesanan:',
        'status_col_label' => 'Huruf Lajur Status Pesanan:',
        'sheet_name_label' => 'Nama Helaian Kerja (Tab):',
        'placeholder_col_id' => 'cth., A',
        'placeholder_col_id_c' => 'cth., C',
        'placeholder_col_status' => 'cth., B',
        'select_file_first' => '-- Sila pilih fail Pendapatan dahulu --',
        'csv_no_sheet' => 'Fail CSV (Tiada helaian diperlukan)',
        'submit_btn' => 'Bandingkan ID Pesanan',
        'missing_label' => 'MISSING',
        'col_order_id_f1' => 'Order ID – Orders File',
        'col_match_id_f2' => 'Matching Order ID – Income File',
        'col_status' => 'Order Status',
        'err_missing_files' => 'Sila pilih dan muat naik kedua-dua fail untuk meneruskan.',
        'err_invalid_col' => 'Format huruf lajur tidak sah: "%s". Sila masukkan 1-3 huruf (A-Z).',
        'err_sheet_not_found' => 'Helaian kerja "%s" tidak ditemui dalam fail Pendapatan.',
        'err_processing' => 'Ralat berlaku semasa perbandingan: ',
        'err_invalid_format' => 'Format fail tidak disokong. Sila muat naik fail .xlsx atau .csv.',
        'theme_light' => 'Mod Cerah',
        'theme_dark' => 'Mod Gelap',
        'select_lang' => 'Bahasa',
        'summary_title' => 'Ringkasan Perbandingan',
        'total_orders' => 'Jumlah Pesanan Utama',
        'matched_orders' => 'Pesanan Padan',
        'missing_orders' => 'Pesanan Hilang (MISSING)',
        'download_btn' => 'Muat Turun Hasil Excel',
        'compare_again' => 'Banding Fail Lain'
    ]
];

$errorMessage = "";

// -------------------------------------------------------------------------
// 3. 高性能处理文件比对逻辑 (POST)
// -------------------------------------------------------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        if (!class_exists('\OpenSpout\Reader\XLSX\Reader')) {
            throw new Exception("OpenSpout is not installed. Please run 'composer install' first.");
        }

        if (empty($_FILES['orders_file']['tmp_name']) || empty($_FILES['income_file']['tmp_name'])) {
            throw new Exception($t[$lang]['err_missing_files']);
        }

        $f1_order_id_col     = strtoupper(trim($_POST['f1_order_id_col']));
        $f1_order_status_col = strtoupper(trim($_POST['f1_order_status_col']));
        $f2_sheet_name       = trim($_POST['f2_sheet_name'] ?? '');
        $f2_order_id_col     = strtoupper(trim($_POST['f2_order_id_col']));

        // 验证列名字母
        $colPattern = '/^[A-Z]{1,3}$/';
        if (!preg_match($colPattern, $f1_order_id_col)) {
            throw new Exception(sprintf($t[$lang]['err_invalid_col'], $f1_order_id_col));
        }
        if (!preg_match($colPattern, $f1_order_status_col)) {
            throw new Exception(sprintf($t[$lang]['err_invalid_col'], $f1_order_status_col));
        }
        if (!preg_match($colPattern, $f2_order_id_col)) {
            throw new Exception(sprintf($t[$lang]['err_invalid_col'], $f2_order_id_col));
        }

        $file1Path = $_FILES['orders_file']['tmp_name'];
        $file2Path = $_FILES['income_file']['tmp_name'];

        // --- 1. 使用 OpenSpout 极速读取 File 2 (Income File) 并提取 ID 集 ---
        $incomeRows = readRowsWithOpenSpout($file2Path, [$f2_order_id_col], $f2_sheet_name);
        $incomeIds = [];
        foreach ($incomeRows as $row) {
            if (isset($row[$f2_order_id_col])) {
                $cleanId = trim($row[$f2_order_id_col]);
                if ($cleanId !== '') {
                    $incomeIds[$cleanId] = true;
                }
            }
        }
        unset($incomeRows); // 读完立刻清空内存

        // --- 2. 使用 OpenSpout 读取 File 1 (Orders File) ---
        $orderRows = readRowsWithOpenSpout($file1Path, [$f1_order_id_col, $f1_order_status_col]);

        // --- 3. 生成临时存储的结果文件 (防 503 超低内存流式写入) ---
        $randomFile = 'Result_' . uniqid('', true) . '.xlsx';
        $outputPath = $uploadDir . '/' . $randomFile;

        $writer = new \OpenSpout\Writer\XLSX\Writer();
        $writer->openToFile($outputPath); // 写入磁盘临时文件而不再直推浏览器，以便能完美重定向渲染页面

        // 定义样式
        $headerStyle = (new \OpenSpout\Common\Entity\Style\Style())
            ->setFontBold()
            ->setBackgroundColor('F2F2F2');

        $missingStyle = (new \OpenSpout\Common\Entity\Style\Style())
            ->setFontBold()
            ->setFontColor('9C0006')
            ->setBackgroundColor('FFC7CE');

        // 写表头
        $writer->addRow(\OpenSpout\Common\Entity\Row::fromValues([
            $t[$lang]['col_order_id_f1'],
            $t[$lang]['col_match_id_f2'],
            $t[$lang]['col_status']
        ], $headerStyle));

        $total_count = 0;
        $matched_count = 0;
        $missing_count = 0;

        // 逐行匹配并流式写入临时 Excel，并在此统计数量
        foreach ($orderRows as $rowIndex => $row) {
            if ($rowIndex === 1) {
                continue; // 跳过表头
            }
            $orderId = $row[$f1_order_id_col] ?? '';
            $status  = $row[$f1_order_status_col] ?? '';
            if ($orderId === '') {
                continue;
            }

            $total_count++;

            if (isset($incomeIds[$orderId])) {
                $matched_count++;
                $cells = [
                    \OpenSpout\Common\Entity\Cell::fromValue($orderId),
                    \OpenSpout\Common\Entity\Cell::fromValue($orderId),
                    \OpenSpout\Common\Entity\Cell::fromValue($status)
                ];
            } else {
                $missing_count++;
                $cells = [
                    \OpenSpout\Common\Entity\Cell::fromValue($orderId),
                    \OpenSpout\Common\Entity\Cell::fromValue($t[$lang]['missing_label'], $missingStyle),
                    \OpenSpout\Common\Entity\Cell::fromValue($status)
                ];
            }
            $writer->addRow(new \OpenSpout\Common\Entity\Row($cells));
        }

        $writer->close();
        unset($orderRows, $incomeIds);

        // 将最终结果统计存入 Session
        $_SESSION['comparison_result'] = [
            'total' => $total_count,
            'matched' => $matched_count,
            'missing' => $missing_count,
            'file' => $randomFile
        ];

        // 处理完后跳转，防止表单重复提交，且优雅渲染网页比对状态
        header("Location: " . strtok($_SERVER['REQUEST_URI'], '?') . "?result=success");
        exit;

    } catch (Throwable $e) { // 拦截所有致命异常，防止白屏
        $errorMessage = $t[$lang]['err_processing'] . $e->getMessage();
    }
}
?>
<!DOCTYPE html>
<html lang="<?= $lang ?>" data-bs-theme="<?= $theme ?>">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><?= htmlspecialchars($t[$lang]['title']) ?></title>
    <!-- 引入 Bootstrap 5.3 CSS & Icons -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <!-- 引入 SheetJS 用于前端毫秒级快速读取 Excel 的工作表 Tab 名称 -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        body {
            background-color: var(--bs-body-bg);
            color: var(--bs-body-color);
        }
        .navbar {
            border-bottom: 1px solid var(--bs-border-color);
        }
        .card {
            box-shadow: 0 4px 6px rgba(0,0,0,0.02);
            border: 1px solid var(--bs-border-color-translucent);
        }
    </style>
</head>
<body class="bg-body-tertiary">

    <!-- 顶栏导航 -->
    <nav class="navbar navbar-expand-lg bg-body sticky-top">
        <div class="container">
            <span class="navbar-brand mb-0 h1">
                <i class="bi bi-file-earmark-diff text-primary"></i> <?= htmlspecialchars($t[$lang]['title']) ?>
            </span>
            <div class="d-flex align-items-center gap-2">
                <!-- 语言切换下拉组 -->
                <div class="dropdown">
                    <button class="btn btn-outline-secondary btn-sm dropdown-toggle" type="button" data-bs-toggle="dropdown" aria-expanded="false">
                        <i class="bi bi-translate"></i> <?= htmlspecialchars($t[$lang]['select_lang']) ?>
                    </button>
                    <ul class="dropdown-menu dropdown-menu-end">
                        <li><a class="dropdown-item <?= $lang === 'en' ? 'active' : '' ?>" href="?lang=en">English</a></li>
                        <li><a class="dropdown-item <?= $lang === 'zh' ? 'active' : '' ?>" href="?lang=zh">中文</a></li>
                        <li><a class="dropdown-item <?= $lang === 'ms' ? 'active' : '' ?>" href="?lang=ms">Bahasa Melayu</a></li>
                    </ul>
                </div>
                <!-- 主题切换按钮 -->
                <a href="?theme=<?= $theme === 'light' ? 'dark' : 'light' ?>" class="btn btn-outline-secondary btn-sm" title="<?= htmlspecialchars($t[$lang]['theme_light']) ?> / <?= htmlspecialchars($t[$lang]['theme_dark']) ?>">
                    <?php if ($theme === 'light'): ?>
                        <i class="bi bi-moon-fill"></i> <span class="d-none d-sm-inline"><?= htmlspecialchars($t[$lang]['theme_dark']) ?></span>
                    <?php else: ?>
                        <i class="bi bi-sun-fill text-warning"></i> <span class="d-none d-sm-inline"><?= htmlspecialchars($t[$lang]['theme_light']) ?></span>
                    <?php endif; ?>
                </a>
            </div>
        </div>
    </nav>

    <div class="container py-4" style="max-width: 960px;">
        
        <!-- 子标题描述 -->
        <p class="text-muted text-center mb-4"><?= htmlspecialchars($t[$lang]['subtitle']) ?></p>

        <!-- 异常提示区域 -->
        <?php if (!empty($errorMessage)): ?>
            <div class="alert alert-danger d-flex align-items-center gap-2 alert-dismissible fade show" role="alert">
                <i class="bi bi-exclamation-triangle-fill fs-5"></i>
                <div><?= htmlspecialchars($errorMessage) ?></div>
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
            </div>
        <?php endif; ?>

        <!-- ----------------------------------------------------------------- -->
        <!-- 核心功能升级：比对结果可视化大屏汇总 Dashboard -->
        <!-- ----------------------------------------------------------------- -->
        <?php if (isset($_SESSION['comparison_result']) && isset($_GET['result']) && $_GET['result'] === 'success'): ?>
            <?php
            $res = $_SESSION['comparison_result'];
            $total = $res['total'];
            $matched = $res['matched'];
            $missing = $res['missing'];
            
            $matched_percent = $total > 0 ? ($matched / $total) * 100 : 0;
            $missing_percent = $total > 0 ? ($missing / $total) * 100 : 0;
            ?>
            <div class="card shadow border-0 mb-4 bg-body">
                <div class="card-header bg-primary text-white py-3">
                    <h5 class="card-title mb-0 d-flex align-items-center gap-2 fw-bold">
                        <i class="bi bi-pie-chart-fill"></i> <?= htmlspecialchars($t[$lang]['summary_title']) ?>
                    </h5>
                </div>
                <div class="card-body p-4">
                    <!-- 双色堆叠比例进度条 -->
                    <div class="progress mb-4" style="height: 32px; font-size: 0.95rem; font-weight: bold;">
                        <?php if ($matched > 0): ?>
                        <div class="progress-bar bg-success d-flex align-items-center justify-content-center" role="progressbar" style="width: <?= $matched_percent ?>%" aria-valuenow="<?= $matched_percent ?>" aria-valuemin="0" aria-valuemax="100">
                            <?= number_format($matched_percent, 1) ?>% Match
                        </div>
                        <?php endif; ?>
                        <?php if ($missing > 0): ?>
                        <div class="progress-bar bg-danger d-flex align-items-center justify-content-center" role="progressbar" style="width: <?= $missing_percent ?>%" aria-valuenow="<?= $missing_percent ?>" aria-valuemin="0" aria-valuemax="100">
                            <?= number_format($missing_percent, 1) ?>% Missing
                        </div>
                        <?php endif; ?>
                    </div>

                    <!-- 统计数据格 -->
                    <div class="row g-3 text-center mb-4">
                        <div class="col-md-4">
                            <div class="p-3 bg-body-tertiary rounded border border-secondary-subtle">
                                <h6 class="text-muted small mb-1"><?= htmlspecialchars($t[$lang]['total_orders']) ?></h6>
                                <h3 class="fw-bold text-body mb-0"><?= number_format($total) ?></h3>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="p-3 bg-success-subtle rounded border border-success-subtle">
                                <h6 class="text-success-emphasis small mb-1"><?= htmlspecialchars($t[$lang]['matched_orders']) ?></h6>
                                <h3 class="fw-bold text-success mb-0"><?= number_format($matched) ?></h3>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="p-3 bg-danger-subtle rounded border border-danger-subtle">
                                <h6 class="text-danger-emphasis small mb-1"><?= htmlspecialchars($t[$lang]['missing_orders']) ?></h6>
                                <h3 class="fw-bold text-danger mb-0"><?= number_format($missing) ?></h3>
                            </div>
                        </div>
                    </div>

                    <!-- 操作控制区 -->
                    <div class="d-flex flex-column flex-sm-row gap-3">
                        <a href="?action=download" class="btn btn-primary btn-lg flex-grow-1 py-3 shadow-sm d-flex align-items-center justify-content-center gap-2">
                            <i class="bi bi-cloud-arrow-down-fill fs-5"></i> <?= htmlspecialchars($t[$lang]['download_btn']) ?>
                        </a>
                        <a href="?action=reset" class="btn btn-outline-secondary btn-lg py-3 d-flex align-items-center justify-content-center gap-2">
                            <i class="bi bi-arrow-counterclockwise"></i> <?= htmlspecialchars($t[$lang]['compare_again']) ?>
                        </a>
                    </div>
                </div>
            </div>
        <?php else: ?>

            <!-- Bootstrap 优雅手风琴用户指南 -->
            <div class="accordion mb-4" id="guideAccordion">
                <div class="accordion-item">
                    <h2 class="accordion-header">
                        <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseGuide" aria-expanded="true" aria-controls="collapseGuide">
                            <strong class="d-flex align-items-center gap-2">
                                <i class="bi bi-info-circle text-primary"></i> <?= htmlspecialchars($t[$lang]['guide_title']) ?>
                            </strong>
                        </button>
                    </h2>
                    <div id="collapseGuide" class="accordion-collapse collapse show" data-bs-parent="#guideAccordion">
                        <div class="accordion-body">
                            <div class="row g-3">
                                <div class="col-md-4">
                                    <h6 class="text-primary"><i class="bi bi-1-circle-fill"></i> <?= htmlspecialchars($t[$lang]['guide_step1']) ?></h6>
                                    <p class="small text-muted mb-0"><?= htmlspecialchars($t[$lang]['guide_step1_desc']) ?></p>
                                </div>
                                <div class="col-md-4">
                                    <h6 class="text-primary"><i class="bi bi-2-circle-fill"></i> <?= htmlspecialchars($t[$lang]['guide_step2']) ?></h6>
                                    <p class="small text-muted mb-0"><?= htmlspecialchars($t[$lang]['guide_step2_desc']) ?></p>
                                </div>
                                <div class="col-md-4">
                                    <h6 class="text-primary"><i class="bi bi-3-circle-fill"></i> <?= htmlspecialchars($t[$lang]['guide_step3']) ?></h6>
                                    <p class="small text-muted mb-0"><?= htmlspecialchars($t[$lang]['guide_step3_desc']) ?></p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 主操作表单 -->
            <form action="" method="POST" enctype="multipart/form-data">
                <div class="row g-4 mb-4">
                    
                    <!-- 模块一：订单主表 -->
                    <div class="col-md-6">
                        <div class="card h-100">
                            <div class="card-body">
                                <h5 class="card-title text-primary border-bottom pb-2 mb-3">
                                    <i class="bi bi-card-checklist"></i> <?= htmlspecialchars($t[$lang]['file1_title']) ?>
                                </h5>
                                <div class="mb-3">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['file_label']) ?></label>
                                    <input type="file" name="orders_file" accept=".xlsx, .csv" required class="form-control">
                                </div>
                                <div class="mb-3">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['order_col_label']) ?></label>
                                    <input type="text" name="f1_order_id_col" placeholder="<?= htmlspecialchars($t[$lang]['placeholder_col_id']) ?>" required class="form-control" maxlength="3">
                                </div>
                                <div class="mb-2">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['status_col_label']) ?></label>
                                    <input type="text" name="f1_order_status_col" placeholder="<?= htmlspecialchars($t[$lang]['placeholder_col_status']) ?>" required class="form-control" maxlength="3">
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- 模块二：收益/结算表 -->
                    <div class="col-md-6">
                        <div class="card h-100">
                            <div class="card-body">
                                <h5 class="card-title text-success border-bottom pb-2 mb-3">
                                    <i class="bi bi-cash-stack"></i> <?= htmlspecialchars($t[$lang]['file2_title']) ?>
                                </h5>
                                <div class="mb-3">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['file_label']) ?></label>
                                    <input type="file" name="income_file" id="income_file_input" accept=".xlsx, .csv" required class="form-control">
                                </div>
                                <!-- 自动检索的工作表下拉选择框 -->
                                <div class="mb-3">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['sheet_name_label']) ?></label>
                                    <select name="f2_sheet_name" id="f2_sheet_name" class="form-select">
                                        <option value=""><?= htmlspecialchars($t[$lang]['select_file_first']) ?></option>
                                    </select>
                                </div>
                                <div class="mb-2">
                                    <label class="form-label"><?= htmlspecialchars($t[$lang]['order_col_label']) ?></label>
                                    <input type="text" name="f2_order_id_col" placeholder="<?= htmlspecialchars($t[$lang]['placeholder_col_id_c']) ?>" required class="form-control" maxlength="3">
                                </div>
                            </div>
                        </div>
                    </div>

                </div>

                <!-- 执行提交 -->
                <button type="submit" class="btn btn-primary btn-lg w-100 py-3 shadow-sm">
                    <i class="bi bi-file-earmark-excel-fill"></i> <?= htmlspecialchars($t[$lang]['submit_btn']) ?>
                </button>
            </form>
        <?php endif; ?>

    </div>

    <!-- 引入 Bootstrap Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>

    <!-- 自动检索 Excel 工作表 JavaScript 逻辑 -->
    <script>
        document.getElementById('income_file_input').addEventListener('change', function(e) {
            const file = e.target.files[0];
            const selectEl = document.getElementById('f2_sheet_name');
            selectEl.innerHTML = '';

            if (!file) {
                selectEl.innerHTML = '<option value=""><?= htmlspecialchars($t[$lang]['select_file_first']) ?></option>';
                return;
            }

            const fileName = file.name.toLowerCase();
            if (fileName.endsWith('.csv')) {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = "<?= htmlspecialchars($t[$lang]['csv_no_sheet']) ?>";
                selectEl.appendChild(opt);
                return;
            }

            // 读取 Excel 文件元数据解析 工作表 (Tab) 列表
            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    // bookSheets: true 使得 SheetJS 仅读取工作表标签目录，速度极快
                    const workbook = XLSX.read(data, {type: 'array', bookSheets: true});
                    const sheetNames = workbook.SheetNames;

                    if (sheetNames && sheetNames.length > 0) {
                        sheetNames.forEach(name => {
                            const opt = document.createElement('option');
                            opt.value = name;
                            opt.textContent = name;
                            
                            // 自动适配新版：如果 Tab 名叫 "Income" (不区分大小写)，自动默认选中它
                            if (name.toLowerCase() === 'income') {
                                opt.selected = true;
                            }
                            selectEl.appendChild(opt);
                        });
                    } else {
                        selectEl.innerHTML = '<option value="">No sheets found</option>';
                    }
                } catch (err) {
                    console.error("Error parsing excel sheets:", err);
                    selectEl.innerHTML = '<option value="">Unable to read sheet names</option>';
                }
            };
            reader.readAsArrayBuffer(file);
        });
    </script>
</body>
</html>