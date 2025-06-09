<?php
require_once 'config.php';
requireUnit();

// Get unit info
$stmt = $pdo->prepare("SELECT name FROM units WHERE id = ?");
$stmt->execute([$_SESSION['unit_id']]);
$unit = $stmt->fetch();

// Get today's report status
$stmt = $pdo->prepare("
    SELECT COUNT(*) as report_count 
    FROM report_sessions 
    WHERE unit_id = ? AND DATE(created_at) = CURDATE()
");
$stmt->execute([$_SESSION['unit_id']]);
$today_report = $stmt->fetch()['report_count'];

// Get total questions for this unit
$stmt = $pdo->prepare("
    SELECT COUNT(*) as total_questions 
    FROM questions 
    WHERE unit_id = ? AND is_active = 1 AND deleted_at IS NULL
");
$stmt->execute([$_SESSION['unit_id']]);
$total_questions = $stmt->fetch()['total_questions'];

// Get recent reports
$stmt = $pdo->prepare("
    SELECT rs.*, COUNT(r.id) as answered_questions
    FROM report_sessions rs
    LEFT JOIN reports r ON rs.id = r.session_id
    WHERE rs.unit_id = ?
    GROUP BY rs.id
    ORDER BY rs.created_at DESC
    LIMIT 5
");
$stmt->execute([$_SESSION['unit_id']]);
$recent_reports = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Unit - <?= $unit['name'] ?></title>
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
        .sidebar .nav-link:hover,
        .sidebar .nav-link.active {
            background: rgba(255, 255, 255, 0.2);
            color: white;
            transform: translateX(5px);
        }
        .stats-card {
            border-radius: 15px;
            border: none;
            transition: transform 0.3s ease;
        }
        .stats-card:hover {
            transform: translateY(-5px);
        }
        .content-wrapper {
            background: #f8f9fa;
            min-height: 100vh;
        }
        .action-card {
            border-radius: 15px;
            border: none;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .action-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
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
                        <h5 class="text-white mt-2"><?= $unit['name'] ?></h5>
                        <small class="text-white-50">Halo, <?= $_SESSION['name'] ?></small>
                    </div>
                    
                    <nav class="nav flex-column">
                        <a href="#" class="nav-link active">
                            <i class="bi bi-speedometer2 me-2"></i>Dashboard
                        </a>
                        <a href="fill_report.php" class="nav-link">
                            <i class="bi bi-file-text me-2"></i>Isi Laporan
                        </a>
                        <a href="patient_referral.php" class="nav-link">
                            <i class="bi bi-person-plus me-2"></i>Data Pasien Rujuk
                        </a>
                         <a href="riwayat_laporan.php" class="nav-link">
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
                            <h2 class="mb-0">Dashboard Unit</h2>
                            <p class="text-muted mb-0">Unit: <?= $unit['name'] ?></p>
                        </div>
                        <div class="d-flex align-items-center">
                            <span class="badge bg-success me-2">Online</span>
                            <span class="text-muted"><?= date('d F Y') ?></span>
                        </div>
                    </div>
                    
                   <!-- Statistics Cards -->
                    <div class="row mb-4">
                        <div class="col-md-4 mb-3">
                            <div class="card stats-card bg-primary text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $total_questions ?></h4>
                                            <p class="card-text">Total Pertanyaan</p>
                                        </div>
                                        <i class="bi bi-question-circle fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-4 mb-3">
                            <div class="card stats-card <?= $today_report > 0 ? 'bg-success' : 'bg-warning' ?> text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $today_report ?></h4>
                                            <p class="card-text">Laporan Hari Ini</p>
                                        </div>
                                        <i class="bi bi-file-text fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-4 mb-3">
                            <div class="card stats-card bg-info text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= count($recent_reports) ?></h4>
                                            <p class="card-text">Total Laporan</p>
                                        </div>
                                        <i class="bi bi-graph-up fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Quick Actions -->
                    <div class="row mb-4">
                        <div class="col-12">
                            <h5 class="mb-3">
                                <i class="bi bi-lightning me-2"></i>
                                Aksi Cepat
                            </h5>
                        </div>
                        
                        <div class="col-md-6 mb-3">
                            <div class="card action-card h-100" onclick="window.location.href='fill_report.php'">
                                <div class="card-body text-center p-4">
                                    <i class="bi bi-file-text fs-1 text-primary mb-3"></i>
                                    <h5 class="card-title">Isi Laporan Harian</h5>
                                    <p class="card-text text-muted">Isi laporan harian untuk unit Anda</p>
                                    <span class="badge bg-primary">
                                        <?= $today_report > 0 ? 'Sudah Diisi' : 'Belum Diisi' ?>
                                    </span>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-6 mb-3">
                            <div class="card action-card h-100" onclick="window.location.href='patient_referral.php'">
                                <div class="card-body text-center p-4">
                                    <i class="bi bi-person-plus fs-1 text-success mb-3"></i>
                                    <h5 class="card-title">Data Pasien Rujuk</h5>
                                    <p class="card-text text-muted">Input data pasien yang dirujuk ke RS lain</p>
                                    <span class="badge bg-success">
                                        <i class="bi bi-plus-circle me-1"></i>Tambah Data
                                    </span>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                 
                    
                    <!-- Alert for today's report -->
                    <?php if ($today_report == 0): ?>
                        <div class="position-fixed bottom-0 end-0 p-3" style="z-index: 11">
                            <div class="alert alert-warning alert-dismissible fade show" role="alert">
                                <i class="bi bi-exclamation-triangle me-2"></i>
                                <strong>Perhatian!</strong> Anda belum mengisi laporan untuk hari ini.
                                <a href="fill_report.php" class="alert-link">Isi sekarang</a>
                                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                            </div>
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Auto-hide alert after 10 seconds
        setTimeout(function() {
            var alert = document.querySelector('.alert');
            if (alert) {
                var bsAlert = new bootstrap.Alert(alert);
                bsAlert.close();
            }
        }, 10000);
    </script>
</body>
</html>