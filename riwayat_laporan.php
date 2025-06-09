<?php
require_once 'config.php';
requireUnit();

// Get unit info
$stmt = $pdo->prepare("SELECT name FROM units WHERE id = ?");
$stmt->execute([$_SESSION['unit_id']]);
$unit = $stmt->fetch();

// Get filter parameters
$filter_date_from = $_GET['date_from'] ?? date('Y-m-01'); // First day of current month
$filter_date_to = $_GET['date_to'] ?? date('Y-m-d'); // Today
$filter_status = $_GET['status'] ?? '';
$filter_shift = $_GET['shift'] ?? '';
$page = (int)($_GET['page'] ?? 1);
$per_page = 10;
$offset = ($page - 1) * $per_page;

// Build query
$sql = "
    SELECT 
        rs.*,
        COUNT(r.id) as answered_questions,
        COUNT(pr.id) as patient_referrals,
        (SELECT COUNT(*) FROM questions q WHERE q.unit_id = rs.unit_id AND q.is_active = 1 AND q.deleted_at IS NULL) as total_questions
    FROM report_sessions rs
    LEFT JOIN reports r ON rs.id = r.session_id
    LEFT JOIN patient_referrals pr ON rs.id = pr.session_id
    WHERE rs.unit_id = ? AND rs.report_date BETWEEN ? AND ?
";
$params = [$_SESSION['unit_id'], $filter_date_from, $filter_date_to];

if ($filter_status) {
    $sql .= " AND rs.status = ?";
    $params[] = $filter_status;
}

if ($filter_shift) {
    $sql .= " AND rs.shift = ?";
    $params[] = $filter_shift;
}

$sql .= " GROUP BY rs.id ORDER BY rs.report_date DESC, rs.created_at DESC";

// Get total count for pagination (separate simpler query)
$count_sql = "
    SELECT COUNT(DISTINCT rs.id) as total
    FROM report_sessions rs
    WHERE rs.unit_id = ? AND rs.report_date BETWEEN ? AND ?
";
$count_params = [$_SESSION['unit_id'], $filter_date_from, $filter_date_to];

if ($filter_status) {
    $count_sql .= " AND rs.status = ?";
    $count_params[] = $filter_status;
}

if ($filter_shift) {
    $count_sql .= " AND rs.shift = ?";
    $count_params[] = $filter_shift;
}

$count_stmt = $pdo->prepare($count_sql);
$count_stmt->execute($count_params);
$count_result = $count_stmt->fetch();
$total_records = $count_result ? $count_result['total'] : 0;
$total_pages = ceil($total_records / $per_page);

// Add pagination to main query
$sql .= " LIMIT $per_page OFFSET $offset";

$stmt = $pdo->prepare($sql);
$stmt->execute($params);
$reports = $stmt->fetchAll() ?: [];

// Get statistics for current filters
$stats_sql = "
    SELECT 
        COUNT(rs.id) as total_reports,
        COUNT(CASE WHEN rs.status = 'completed' THEN 1 END) as completed_reports,
        COUNT(CASE WHEN rs.status = 'pending' THEN 1 END) as pending_reports,
        COUNT(DISTINCT rs.report_date) as unique_dates
    FROM report_sessions rs
    WHERE rs.unit_id = ? AND rs.report_date BETWEEN ? AND ?
";
$stats_params = [$_SESSION['unit_id'], $filter_date_from, $filter_date_to];

if ($filter_status) {
    $stats_sql .= " AND rs.status = ?";
    $stats_params[] = $filter_status;
}

if ($filter_shift) {
    $stats_sql .= " AND rs.shift = ?";
    $stats_params[] = $filter_shift;
}

$stats_stmt = $pdo->prepare($stats_sql);
$stats_stmt->execute($stats_params);
$stats_result = $stats_stmt->fetch();
$stats = $stats_result ?: [
    'total_reports' => 0,
    'completed_reports' => 0,
    'pending_reports' => 0,
    'unique_dates' => 0
];
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Riwayat Laporan - <?= $unit['name'] ?></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        .sidebar {
            min-height: 100vh;
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        }
        .sidebar .nav-link {
            color: rgba(255, 255, 255, 0.8);
            border-radius: 10px;
            margin: 2px 0;
            transition: all 0.3s ease;
        }
        .sidebar .nav-link:hover {
            background: rgba(255, 255, 255, 0.2);
            color: white;
            transform: translateX(5px);
        }
        .sidebar .nav-link.active {
            background: rgba(255, 255, 255, 0.2);
            color: white;
        }
        .content-wrapper {
            background: #f8f9fa;
            min-height: 100vh;
        }
        .stats-card {
            border-radius: 15px;
            border: none;
            transition: transform 0.3s ease;
        }
        .stats-card:hover {
            transform: translateY(-5px);
        }
        .report-card {
            border-radius: 15px;
            border: none;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            transition: all 0.3s ease;
        }
        .report-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }
        .report-card.completed {
            border-left: 5px solid #28a745;
        }
        .report-card.pending {
            border-left: 5px solid #ffc107;
        }
        .filter-card {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            border-radius: 15px;
        }
        .progress-circle {
            width: 60px;
            height: 60px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 14px;
        }
        .progress-100 { background: #28a745; color: white; }
        .progress-75 { background: #17a2b8; color: white; }
        .progress-50 { background: #ffc107; color: black; }
        .progress-25 { background: #dc3545; color: white; }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="row">
            <!-- Sidebar -->
            <div class="col-md-3 col-lg-2 px-0">
                <div class="sidebar p-3">
                    <div class="text-center mb-4">
                        <i class="bi bi-hospital fs-1 text-white"></i>
                        <h5 class="text-white mt-2"><?= $unit['name'] ?></h5>
                        <small class="text-white-50">Halo, <?= $_SESSION['name'] ?></small>
                    </div>
                    
                    <nav class="nav flex-column">
                        <a href="unit_dashboard.php" class="nav-link">
                            <i class="bi bi-speedometer2 me-2"></i>Dashboard
                        </a>
                        <a href="fill_report.php" class="nav-link">
                            <i class="bi bi-file-text me-2"></i>Isi Laporan
                        </a>
                        <a href="patient_referral.php" class="nav-link">
                            <i class="bi bi-person-plus me-2"></i>Data Pasien Rujuk
                        </a>
                        <a href="#" class="nav-link active">
                            <i class="bi bi-clock-history me-2"></i>Riwayat Laporan
                        </a>
                        <hr class="text-white-50">
                        <a href="logout.php" class="nav-link">
                            <i class="bi bi-box-arrow-right me-2"></i>Logout
                        </a>
                    </nav>
                </div>
            </div>
            
            <!-- Main Content -->
            <div class="col-md-9 col-lg-10 px-0">
                <div class="content-wrapper p-4">
                    <!-- Header -->
                    <div class="d-flex justify-content-between align-items-center mb-4">
                        <div>
                            <h2 class="mb-0">Riwayat Laporan</h2>
                            <p class="text-muted mb-0">Unit: <?= $unit['name'] ?></p>
                        </div>
                        <a href="fill_report.php" class="btn btn-primary">
                            <i class="bi bi-plus-circle me-2"></i>Buat Laporan Baru
                        </a>
                    </div>
                    
                    <!-- Filter -->
                    <div class="filter-card p-4 mb-4">
                        <h5 class="mb-3">
                            <i class="bi bi-funnel me-2"></i>
                            Filter Laporan
                        </h5>
                        <form method="GET" class="row g-3">
                            <div class="col-md-3">
                                <label class="form-label text-white">Tanggal Dari</label>
                                <input type="date" class="form-control" name="date_from" value="<?= $filter_date_from ?>">
                            </div>
                            
                            <div class="col-md-3">
                                <label class="form-label text-white">Tanggal Sampai</label>
                                <input type="date" class="form-control" name="date_to" value="<?= $filter_date_to ?>">
                            </div>
                            
                            <div class="col-md-2">
                                <label class="form-label text-white">Status</label>
                                <select class="form-select" name="status">
                                    <option value="">Semua Status</option>
                                    <option value="completed" <?= $filter_status === 'completed' ? 'selected' : '' ?>>Selesai</option>
                                    <option value="pending" <?= $filter_status === 'pending' ? 'selected' : '' ?>>Pending</option>
                                </select>
                            </div>
                            
                            <div class="col-md-2">
                                <label class="form-label text-white">Shift</label>
                                <select class="form-select" name="shift">
                                    <option value="">Semua Shift</option>
                                    <option value="pagi" <?= $filter_shift === 'pagi' ? 'selected' : '' ?>>Pagi</option>
                                    <option value="siang" <?= $filter_shift === 'siang' ? 'selected' : '' ?>>Siang</option>
                                    <option value="malam" <?= $filter_shift === 'malam' ? 'selected' : '' ?>>Malam</option>
                                </select>
                            </div>
                            
                            <div class="col-md-2 d-flex align-items-end">
                                <button type="submit" class="btn btn-light me-2">
                                    <i class="bi bi-search me-2"></i>Filter
                                </button>
                                <a href="my_reports.php" class="btn btn-outline-light">
                                    <i class="bi bi-arrow-clockwise"></i>
                                </a>
                            </div>
                        </form>
                    </div>
                    
                    <!-- Statistics -->
                    <div class="row mb-4">
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-primary text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['total_reports'] ?></h4>
                                    <p class="card-text mb-0">Total Laporan</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-success text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['completed_reports'] ?></h4>
                                    <p class="card-text mb-0">Selesai</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-warning text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['pending_reports'] ?></h4>
                                    <p class="card-text mb-0">Pending</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-info text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['unique_dates'] ?></h4>
                                    <p class="card-text mb-0">Hari Aktif</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Reports List -->
                    <div class="card border-0 shadow-sm">
                        <div class="card-header bg-white">
                            <h5 class="mb-0">
                                <i class="bi bi-list-ul me-2"></i>
                                Daftar Laporan
                                <?php if ($total_records > 0): ?>
                                    <span class="badge bg-primary ms-2"><?= $total_records ?> laporan</span>
                                <?php endif; ?>
                            </h5>
                        </div>
                        <div class="card-body">
                            <?php if (empty($reports)): ?>
                                <div class="text-center py-5">
                                    <i class="bi bi-inbox fs-1 text-muted"></i>
                                    <h5 class="mt-3 text-muted">Tidak Ada Laporan</h5>
                                    <p class="text-muted">Tidak ada laporan ditemukan untuk filter yang dipilih.</p>
                                    <a href="fill_report.php" class="btn btn-primary">
                                        <i class="bi bi-plus-circle me-2"></i>
                                        Buat Laporan Pertama
                                    </a>
                                </div>
                            <?php else: ?>
                                <?php foreach ($reports as $report): ?>
                                    <?php 
                                    $progress_percentage = $report['total_questions'] > 0 ? round(($report['answered_questions'] / $report['total_questions']) * 100) : 0;
                                    $progress_class = '';
                                    if ($progress_percentage >= 100) $progress_class = 'progress-100';
                                    elseif ($progress_percentage >= 75) $progress_class = 'progress-75';
                                    elseif ($progress_percentage >= 50) $progress_class = 'progress-50';
                                    else $progress_class = 'progress-25';
                                    ?>
                                    
                                    <div class="report-card <?= $report['status'] ?>">
                                        <div class="card-body">
                                            <div class="row align-items-center">
                                                <div class="col-md-1">
                                                    <div class="progress-circle <?= $progress_class ?>">
                                                        <?= $progress_percentage ?>%
                                                    </div>
                                                </div>
                                                
                                                <div class="col-md-3">
                                                    <h6 class="mb-1">
                                                        <i class="bi bi-calendar3 me-2"></i>
                                                        <?= formatDate($report['report_date']) ?>
                                                    </h6>
                                                    <small class="text-muted">
                                                        <i class="bi bi-clock me-1"></i>
                                                        Shift <?= ucfirst($report['shift']) ?>
                                                    </small>
                                                </div>
                                                
                                                <div class="col-md-2">
                                                    <h6 class="mb-1">Pelapor</h6>
                                                    <p class="mb-0"><?= htmlspecialchars($report['reporter_name']) ?></p>
                                                </div>
                                                
                                                <div class="col-md-2">
                                                    <h6 class="mb-1">Status</h6>
                                                    <?php if ($report['status'] === 'completed'): ?>
                                                        <span class="badge bg-success">
                                                            <i class="bi bi-check-circle me-1"></i>Selesai
                                                        </span>
                                                    <?php else: ?>
                                                        <span class="badge bg-warning">
                                                            <i class="bi bi-clock me-1"></i>Pending
                                                        </span>
                                                    <?php endif; ?>
                                                </div>
                                                
                                                <div class="col-md-2">
                                                    <small class="text-muted d-block">Jawaban</small>
                                                    <strong><?= $report['answered_questions'] ?>/<?= $report['total_questions'] ?></strong>
                                                    
                                                    <?php if ($report['patient_referrals'] > 0): ?>
                                                        <br><small class="text-muted">Rujuk</small>
                                                        <span class="badge bg-danger ms-1"><?= $report['patient_referrals'] ?></span>
                                                    <?php endif; ?>
                                                </div>
                                                
                                                <div class="col-md-2 text-end">
                                                    <div class="btn-group" role="group">
                                                        <a href="view_report.php?id=<?= $report['id'] ?>" 
                                                           class="btn btn-sm btn-outline-primary">
                                                            <i class="bi bi-eye"></i> Lihat
                                                        </a>
                                                        
                                                        <?php if ($report['status'] === 'pending'): ?>
                                                            <a href="fill_report.php?continue=<?= $report['id'] ?>" 
                                                               class="btn btn-sm btn-outline-warning">
                                                                <i class="bi bi-pencil"></i> Lanjut
                                                            </a>
                                                        <?php endif; ?>
                                                    </div>
                                                    
                                                    <br><small class="text-muted">
                                                        <?= formatDateTime($report['created_at']) ?>
                                                    </small>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                <?php endforeach; ?>
                                
                                <!-- Pagination -->
                                <?php if ($total_pages > 1): ?>
                                    <nav aria-label="Page navigation" class="mt-4">
                                        <ul class="pagination justify-content-center">
                                            <?php if ($page > 1): ?>
                                                <li class="page-item">
                                                    <a class="page-link" href="?<?= http_build_query(array_merge($_GET, ['page' => $page - 1])) ?>">
                                                        <i class="bi bi-chevron-left"></i>
                                                    </a>
                                                </li>
                                            <?php endif; ?>
                                            
                                            <?php 
                                            $start_page = max(1, $page - 2);
                                            $end_page = min($total_pages, $page + 2);
                                            ?>
                                            
                                            <?php for ($i = $start_page; $i <= $end_page; $i++): ?>
                                                <li class="page-item <?= $i === $page ? 'active' : '' ?>">
                                                    <a class="page-link" href="?<?= http_build_query(array_merge($_GET, ['page' => $i])) ?>">
                                                        <?= $i ?>
                                                    </a>
                                                </li>
                                            <?php endfor; ?>
                                            
                                            <?php if ($page < $total_pages): ?>
                                                <li class="page-item">
                                                    <a class="page-link" href="?<?= http_build_query(array_merge($_GET, ['page' => $page + 1])) ?>">
                                                        <i class="bi bi-chevron-right"></i>
                                                    </a>
                                                </li>
                                            <?php endif; ?>
                                        </ul>
                                    </nav>
                                    
                                    <div class="text-center text-muted">
                                        Halaman <?= $page ?> dari <?= $total_pages ?> 
                                        (Total: <?= $total_records ?> laporan)
                                    </div>
                                <?php endif; ?>
                            <?php endif; ?>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Auto-submit form when date changes
        document.querySelectorAll('input[type="date"], select').forEach(element => {
            element.addEventListener('change', function() {
                // Auto submit after short delay to allow user to select both dates
                setTimeout(() => {
                    this.form.submit();
                }, 100);
            });
        });
        
        // Highlight current report if coming from dashboard
        const urlParams = new URLSearchParams(window.location.search);
        const highlightId = urlParams.get('highlight');
        if (highlightId) {
            const reportCard = document.querySelector(`a[href*="id=${highlightId}"]`)?.closest('.report-card');
            if (reportCard) {
                reportCard.style.border = '2px solid #007bff';
                reportCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        }
        
        // Tooltip for progress circles
        document.querySelectorAll('.progress-circle').forEach(circle => {
            const percentage = circle.textContent.trim();
            let status = '';
            
            if (percentage === '100%') status = 'Laporan Lengkap';
            else if (parseInt(percentage) >= 75) status = 'Hampir Selesai';
            else if (parseInt(percentage) >= 50) status = 'Setengah Jalan';
            else status = 'Baru Dimulai';
            
            circle.setAttribute('title', status);
            circle.setAttribute('data-bs-toggle', 'tooltip');
        });
        
        // Initialize tooltips
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    </script>
</body>
</html>