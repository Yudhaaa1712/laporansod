<?php
require_once 'config.php';
requireAdmin();

// Get filter parameters
$filter_date = $_GET['date'] ?? date('Y-m-d');
$filter_unit = $_GET['unit'] ?? '';
$filter_shift = $_GET['shift'] ?? '';

// Get all units for filter
$stmt = $pdo->query("SELECT id, name FROM units WHERE deleted_at IS NULL ORDER BY name");
$units = $stmt->fetchAll();

// Build query with filters
$sql = "
    SELECT rs.*, u.name as unit_name, 
           COUNT(r.id) as answered_questions,
           COUNT(pr.id) as patient_referrals,
           (SELECT COUNT(*) FROM questions q WHERE q.unit_id = rs.unit_id AND q.is_active = 1 AND q.deleted_at IS NULL) as total_questions
    FROM report_sessions rs
    LEFT JOIN units u ON rs.unit_id = u.id
    LEFT JOIN reports r ON rs.id = r.session_id
    LEFT JOIN patient_referrals pr ON rs.id = pr.session_id
    WHERE 1=1
";

$params = [];

if ($filter_date) {
    $sql .= " AND rs.report_date = ?";
    $params[] = $filter_date;
}

if ($filter_unit) {
    $sql .= " AND rs.unit_id = ?";
    $params[] = $filter_unit;
}

if ($filter_shift) {
    $sql .= " AND rs.shift = ?";
    $params[] = $filter_shift;
}

$sql .= " GROUP BY rs.id ORDER BY rs.report_date DESC, rs.created_at DESC";

$stmt = $pdo->prepare($sql);
$stmt->execute($params);
$report_sessions = $stmt->fetchAll();

// Get summary statistics
$stats_sql = "
    SELECT 
        COUNT(DISTINCT rs.id) as total_reports,
        COUNT(DISTINCT rs.unit_id) as units_reported,
        COUNT(DISTINCT pr.id) as total_referrals,
        (SELECT COUNT(*) FROM units WHERE deleted_at IS NULL) as total_units
    FROM report_sessions rs
    LEFT JOIN patient_referrals pr ON rs.id = pr.session_id
    WHERE rs.report_date = ?
";

$stmt = $pdo->prepare($stats_sql);
$stmt->execute([$filter_date]);
$stats = $stmt->fetch();

// Get units that haven't reported today
$unreported_sql = "
    SELECT u.name 
    FROM units u 
    WHERE u.deleted_at IS NULL 
    AND u.id NOT IN (
        SELECT DISTINCT rs.unit_id 
        FROM report_sessions rs 
        WHERE rs.report_date = ?
    )
    ORDER BY u.name
";

$stmt = $pdo->prepare($unreported_sql);
$stmt->execute([$filter_date]);
$unreported_units = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitoring Laporan - Sistem Laporan RS</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        .sidebar {
            min-height: 100vh;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
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
        .completion-bar {
            height: 8px;
            border-radius: 4px;
            overflow: hidden;
        }
        .report-card {
            border-radius: 12px;
            border: none;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
        }
        .report-card:hover {
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }
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
                        <h5 class="text-white mt-2">Admin Panel</h5>
                        <small class="text-white-50">Halo, <?= $_SESSION['name'] ?></small>
                    </div>
                    
                    <nav class="nav flex-column">
                        <a href="admin_dashboard.php" class="nav-link">
                            <i class="bi bi-speedometer2 me-2"></i>Dashboard
                        </a>
                        <a href="master_accounts.php" class="nav-link">
                            <i class="bi bi-people me-2"></i>Master Akun
                        </a>
                        <a href="master_units.php" class="nav-link">
                            <i class="bi bi-building me-2"></i>Master Unit
                        </a>
                        <a href="manage_questions.php" class="nav-link">
                            <i class="bi bi-question-circle me-2"></i>Kelola Target Laporan
                        </a>
                        <a href="#" class="nav-link active">
                            <i class="bi bi-graph-up me-2"></i>Monitoring Laporan
                        </a>
                        <a href="download_excel.php" class="nav-link">
                            <i class="bi bi-download me-2"></i>Download Excel
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
                            <h2 class="mb-0">Monitoring Laporan</h2>
                            <p class="text-muted mb-0">Pantau status pengisian laporan unit</p>
                        </div>
                        <a href="download_excel.php?date=<?= $filter_date ?>" class="btn btn-success">
                            <i class="bi bi-download me-2"></i>Download Excel
                        </a>
                    </div>
                    
                    <!-- Filter Section -->
                    <div class="card border-0 shadow-sm mb-4">
                        <div class="card-body">
                            <form method="GET" class="row g-3">
                                <div class="col-md-3">
                                    <label class="form-label">Tanggal</label>
                                    <input type="date" class="form-control" name="date" value="<?= $filter_date ?>">
                                </div>
                                <div class="col-md-3">
                                    <label class="form-label">Unit</label>
                                    <select class="form-select" name="unit">
                                        <option value="">Semua Unit</option>
                                        <?php foreach ($units as $unit): ?>
                                            <option value="<?= $unit['id'] ?>" <?= $filter_unit == $unit['id'] ? 'selected' : '' ?>>
                                                <?= $unit['name'] ?>
                                            </option>
                                        <?php endforeach; ?>
                                    </select>
                                </div>
                                <div class="col-md-3">
                                    <label class="form-label">Shift</label>
                                    <select class="form-select" name="shift">
                                        <option value="">Semua Shift</option>
                                        <option value="pagi" <?= $filter_shift === 'pagi' ? 'selected' : '' ?>>Pagi</option>
                                        <option value="siang" <?= $filter_shift === 'siang' ? 'selected' : '' ?>>Siang</option>
                                        <option value="malam" <?= $filter_shift === 'malam' ? 'selected' : '' ?>>Malam</option>
                                    </select>
                                </div>
                                <div class="col-md-3 d-flex align-items-end">
                                    <button type="submit" class="btn btn-primary me-2">
                                        <i class="bi bi-funnel me-2"></i>Filter
                                    </button>
                                    <a href="?" class="btn btn-outline-secondary">
                                        <i class="bi bi-arrow-clockwise"></i>
                                    </a>
                                </div>
                            </form>
                        </div>
                    </div>
                    
                    <!-- Statistics Cards -->
                    <div class="row mb-4">
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-primary text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $stats['total_reports'] ?></h4>
                                            <p class="card-text">Total Laporan</p>
                                        </div>
                                        <i class="bi bi-file-text fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-success text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $stats['units_reported'] ?>/<?= $stats['total_units'] ?></h4>
                                            <p class="card-text">Unit Melaporkan</p>
                                        </div>
                                        <i class="bi bi-building fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-info text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $stats['total_referrals'] ?></h4>
                                            <p class="card-text">Pasien Rujuk</p>
                                        </div>
                                        <i class="bi bi-person-plus fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-warning text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= round(($stats['units_reported']/$stats['total_units'])*100) ?>%</h4>
                                            <p class="card-text">Tingkat Kepatuhan</p>
                                        </div>
                                        <i class="bi bi-graph-up fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Unreported Units Alert -->
                    <?php if (!empty($unreported_units)): ?>
                        <div class="alert alert-warning alert-dismissible fade show" role="alert">
                            <i class="bi bi-exclamation-triangle me-2"></i>
                            <strong>Unit belum melaporkan (<?= count($unreported_units) ?>):</strong>
                            <?php foreach ($unreported_units as $index => $unit): ?>
                                <span class="badge bg-warning text-dark me-1"><?= $unit['name'] ?></span>
                            <?php endforeach; ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Reports List -->
                    <div class="card border-0 shadow-sm">
                        <div class="card-header bg-white">
                            <h5 class="mb-0">
                                <i class="bi bi-list-ul me-2"></i>
                                Daftar Laporan (<?= count($report_sessions) ?> laporan)
                            </h5>
                        </div>
                        <div class="card-body">
                            <?php if (empty($report_sessions)): ?>
                                <div class="text-center py-5">
                                    <i class="bi bi-inbox fs-1 text-muted"></i>
                                    <h5 class="mt-3 text-muted">Tidak Ada Laporan</h5>
                                    <p class="text-muted">Belum ada laporan untuk filter yang dipilih.</p>
                                </div>
                            <?php else: ?>
                                <div class="row">
                                    <?php foreach ($report_sessions as $session): ?>
                                        <?php 
                                        $completion_percentage = $session['total_questions'] > 0 ? 
                                            round(($session['answered_questions'] / $session['total_questions']) * 100) : 0;
                                        $is_complete = $completion_percentage >= 100;
                                        ?>
                                        <div class="col-md-6 col-lg-4 mb-4">
                                            <div class="card report-card h-100">
                                                <div class="card-body">
                                                    <div class="d-flex justify-content-between align-items-start mb-3">
                                                        <h6 class="card-title mb-0">
                                                            <i class="bi bi-building text-primary me-2"></i>
                                                            <?= $session['unit_name'] ?>
                                                        </h6>
                                                        <span class="badge <?= $is_complete ? 'bg-success' : 'bg-warning' ?>">
                                                            <?= $is_complete ? 'Lengkap' : 'Belum Lengkap' ?>
                                                        </span>
                                                    </div>
                                                    
                                                    <div class="mb-3">
                                                        <small class="text-muted d-block">Pelapor:</small>
                                                        <strong><?= $session['reporter_name'] ?></strong>
                                                    </div>
                                                    
                                                    <div class="row mb-3">
                                                        <div class="col-6">
                                                            <small class="text-muted d-block">Tanggal:</small>
                                                            <span class="fw-bold"><?= formatDate($session['report_date']) ?></span>
                                                        </div>
                                                        <div class="col-6">
                                                            <small class="text-muted d-block">Shift:</small>
                                                            <span class="badge bg-secondary"><?= ucfirst($session['shift']) ?></span>
                                                        </div>
                                                    </div>
                                                    
                                                    <div class="mb-3">
                                                        <div class="d-flex justify-content-between mb-1">
                                                            <small>Progress Jawaban</small>
                                                            <small><?= $session['answered_questions'] ?>/<?= $session['total_questions'] ?></small>
                                                        </div>
                                                        <div class="completion-bar bg-light">
                                                            <div class="bg-<?= $is_complete ? 'success' : 'warning' ?>" 
                                                                 style="width: <?= $completion_percentage ?>%; height: 100%;"></div>
                                                        </div>
                                                        <small class="text-muted"><?= $completion_percentage ?>% selesai</small>
                                                    </div>
                                                    
                                                    <?php if ($session['patient_referrals'] > 0): ?>
                                                        <div class="mb-3">
                                                            <span class="badge bg-info">
                                                                <i class="bi bi-person-plus me-1"></i>
                                                                <?= $session['patient_referrals'] ?> Pasien Rujuk
                                                            </span>
                                                        </div>
                                                    <?php endif; ?>
                                                    
                                                    <div class="d-flex gap-2">
                                                        <button class="btn btn-sm btn-outline-primary flex-fill" 
                                                                onclick="viewReport(<?= $session['id'] ?>)">
                                                            <i class="bi bi-eye me-1"></i>Lihat
                                                        </button>
                                                        <button class="btn btn-sm btn-outline-success" 
                                                                onclick="downloadReport(<?= $session['id'] ?>)">
                                                            <i class="bi bi-download"></i>
                                                        </button>
                                                    </div>
                                                    
                                                    <small class="text-muted d-block mt-2">
                                                        <i class="bi bi-clock me-1"></i>
                                                        <?= formatDateTime($session['created_at']) ?>
                                                    </small>
                                                </div>
                                            </div>
                                        </div>
                                    <?php endforeach; ?>
                                </div>
                            <?php endif; ?>
                        </div>
                    </div>
                    
                    <!-- Quick Summary -->
                    <div class="row mt-4">
                        <div class="col-md-6">
                            <div class="card border-0 shadow-sm">
                                <div class="card-header bg-white">
                                    <h6 class="mb-0">
                                        <i class="bi bi-pie-chart me-2"></i>
                                        Ringkasan Shift
                                    </h6>
                                </div>
                                <div class="card-body">
                                    <?php
                                    $shift_stats = [];
                                    foreach ($report_sessions as $session) {
                                        $shift = ucfirst($session['shift']);
                                        $shift_stats[$shift] = ($shift_stats[$shift] ?? 0) + 1;
                                    }
                                    ?>
                                    <?php if (empty($shift_stats)): ?>
                                        <p class="text-muted mb-0">Tidak ada data</p>
                                    <?php else: ?>
                                        <?php foreach ($shift_stats as $shift => $count): ?>
                                            <div class="d-flex justify-content-between align-items-center mb-2">
                                                <span>Shift <?= $shift ?></span>
                                                <span class="badge bg-primary"><?= $count ?> laporan</span>
                                            </div>
                                        <?php endforeach; ?>
                                    <?php endif; ?>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-6">
                            <div class="card border-0 shadow-sm">
                                <div class="card-header bg-white">
                                    <h6 class="mb-0">
                                        <i class="bi bi-clock-history me-2"></i>
                                        Waktu Laporan Terakhir
                                    </h6>
                                </div>
                                <div class="card-body">
                                    <?php if (!empty($report_sessions)): ?>
                                        <?php $latest = $report_sessions[0]; ?>
                                        <div class="d-flex align-items-center">
                                            <div class="flex-grow-1">
                                                <strong><?= $latest['unit_name'] ?></strong>
                                                <br>
                                                <small class="text-muted">
                                                    <?= $latest['reporter_name'] ?> - 
                                                    Shift <?= ucfirst($latest['shift']) ?>
                                                </small>
                                            </div>
                                            <div class="text-end">
                                                <div class="fw-bold"><?= formatDateTime($latest['created_at']) ?></div>
                                                <small class="text-muted">Laporan terakhir</small>
                                            </div>
                                        </div>
                                    <?php else: ?>
                                        <p class="text-muted mb-0">Belum ada laporan</p>
                                    <?php endif; ?>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- View Report Modal -->
    <div class="modal fade" id="viewReportModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="bi bi-file-text me-2"></i>
                        Detail Laporan
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div id="reportContent">
                        <div class="text-center py-3">
                            <div class="spinner-border text-primary" role="status">
                                <span class="visually-hidden">Loading...</span>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Tutup</button>
                    <button type="button" class="btn btn-primary" onclick="printReport()">
                        <i class="bi bi-printer me-2"></i>Print
                    </button>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function viewReport(sessionId) {
            // Show modal first
            var modal = new bootstrap.Modal(document.getElementById('viewReportModal'));
            modal.show();
            
            // Load report content via AJAX (you'll need to create view_report_ajax.php)
            fetch('view_report_ajax.php?id=' + sessionId)
                .then(response => response.text())
                .then(data => {
                    document.getElementById('reportContent').innerHTML = data;
                })
                .catch(error => {
                    document.getElementById('reportContent').innerHTML = 
                        '<div class="alert alert-danger">Gagal memuat laporan</div>';
                });
        }
        
        function downloadReport(sessionId) {
            window.open('download_excel.php?session_id=' + sessionId, '_blank');
        }
        
        function printReport() {
            var content = document.getElementById('reportContent').innerHTML;
            var printWindow = window.open('', '_blank');
            printWindow.document.write(`
                <html>
                    <head>
                        <title>Laporan</title>
                        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
                        <style>
                            @media print { 
                                .no-print { display: none; }
                                body { padding: 20px; }
                            }
                        </style>
                    </head>
                    <body onload="window.print(); window.close();">
                        ${content}
                    </body>
                </html>
            `);
        }
        
        // Auto refresh every 5 minutes
        setInterval(function() {
            if (!document.querySelector('.modal.show')) {
                location.reload();
            }
        }, 300000);
        
        // Real-time status indicator
        function updateStatus() {
            const now = new Date();
            const statusElements = document.querySelectorAll('.status-indicator');
            
            statusElements.forEach(element => {
                const lastUpdate = new Date(element.dataset.lastUpdate);
                const diffMinutes = Math.floor((now - lastUpdate) / 60000);
                
                if (diffMinutes < 5) {
                    element.className = 'status-indicator bg-success';
                    element.title = 'Baru saja';
                } else if (diffMinutes < 30) {
                    element.className = 'status-indicator bg-warning';
                    element.title = diffMinutes + ' menit yang lalu';
                } else {
                    element.className = 'status-indicator bg-danger';
                    element.title = diffMinutes + ' menit yang lalu';
                }
            });
        }
        
        // Update status every minute
        setInterval(updateStatus, 60000);
        updateStatus();
        
        // Filter form auto-submit on change
        document.querySelectorAll('select[name="unit"], select[name="shift"]').forEach(select => {
            select.addEventListener('change', function() {
                this.form.submit();
            });
        });
    </script>
</body>
</html>