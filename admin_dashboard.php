<?php
require_once 'config.php';
requireAdmin();

// Get statistics
$stmt = $pdo->query("SELECT COUNT(*) as total_units FROM units WHERE deleted_at IS NULL");
$total_units = $stmt->fetch()['total_units'];

$stmt = $pdo->query("SELECT COUNT(*) as total_users FROM users WHERE role = 'unit' AND deleted_at IS NULL");
$total_users = $stmt->fetch()['total_users'];

$stmt = $pdo->query("SELECT COUNT(*) as total_questions FROM questions WHERE deleted_at IS NULL");
$total_questions = $stmt->fetch()['total_questions'];

$stmt = $pdo->query("SELECT COUNT(*) as total_reports FROM report_sessions WHERE DATE(created_at) = CURDATE()");
$today_reports = $stmt->fetch()['total_reports'];

// Get recent activities
$stmt = $pdo->query("
    SELECT rs.*, u.name as unit_name, rs.created_at
    FROM report_sessions rs 
    JOIN units u ON rs.unit_id = u.id 
    ORDER BY rs.created_at DESC 
    LIMIT 5
");
$recent_activities = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Admin - Sistem Laporan RS</title>
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
                        <a href="#" class="nav-link active">
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
                        <a href="monitoring_reports.php" class="nav-link">
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
                        <h2 class="mb-0">Dashboard Admin</h2>
                        <div class="d-flex align-items-center">
                            <span class="badge bg-success me-2">Online</span>
                            <span class="text-muted"><?= date('d F Y') ?></span>
                        </div>
                    </div>
                    
                    <!-- Statistics Cards -->
                    <div class="row mb-4">
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-primary text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $total_units ?></h4>
                                            <p class="card-text">Total Unit</p>
                                        </div>
                                        <i class="bi bi-building fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-success text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $total_users ?></h4>
                                            <p class="card-text">User Unit</p>
                                        </div>
                                        <i class="bi bi-people fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-warning text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $total_questions ?></h4>
                                            <p class="card-text">Pertanyaan</p>
                                        </div>
                                        <i class="bi bi-question-circle fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-3 mb-3">
                            <div class="card stats-card bg-info text-white">
                                <div class="card-body">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <h4 class="card-title"><?= $today_reports ?></h4>
                                            <p class="card-text">Laporan Hari Ini</p>
                                        </div>
                                        <i class="bi bi-file-text fs-1 opacity-50"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Recent Activities -->
                    <div class="row">
                        <div class="col-12">
                            <div class="card border-0 shadow-sm">
                                <div class="card-header bg-white">
                                    <h5 class="mb-0">
                                        <i class="bi bi-clock-history me-2"></i>
                                        Aktivitas Terbaru
                                    </h5>
                                </div>
                                <div class="card-body">
                                    <?php if (empty($recent_activities)): ?>
                                        <div class="text-center text-muted py-4">
                                            <i class="bi bi-inbox fs-1"></i>
                                            <p class="mt-2">Belum ada aktivitas hari ini</p>
                                        </div>
                                    <?php else: ?>
                                        <div class="table-responsive">
                                            <table class="table table-hover">
                                                <thead>
                                                    <tr>
                                                        <th>Unit</th>
                                                        <th>Pelapor</th>
                                                        <th>Tanggal</th>
                                                        <th>Shift</th>
                                                        <th>Waktu</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <?php foreach ($recent_activities as $activity): ?>
                                                        <tr>
                                                            <td>
                                                                <span class="badge bg-primary"><?= $activity['unit_name'] ?></span>
                                                            </td>
                                                            <td><?= $activity['reporter_name'] ?></td>
                                                            <td><?= formatDate($activity['report_date']) ?></td>
                                                            <td>
                                                                <span class="badge bg-secondary"><?= ucfirst($activity['shift']) ?></span>
                                                            </td>
                                                            <td><?= formatDateTime($activity['created_at']) ?></td>
                                                        </tr>
                                                    <?php endforeach; ?>
                                                </tbody>
                                            </table>
                                        </div>
                                    <?php endif; ?>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>