<?php
require_once 'config.php';
requireAdmin();

// Handle Excel download
if (isset($_GET['download']) && $_GET['download'] === 'excel') {
    $date = $_GET['date'] ?? date('Y-m-d');
    $unit_filter = $_GET['unit'] ?? '';
    
    // Generate Excel
    generateExcel($date, $unit_filter);
    exit;
}

// Get filter values
$filter_date = $_GET['date'] ?? date('Y-m-d');
$filter_unit = $_GET['unit'] ?? '';

// Get all units for dropdown
$stmt = $pdo->query("SELECT id, name FROM units WHERE deleted_at IS NULL ORDER BY name");
$units = $stmt->fetchAll();

// Get statistics for selected date
$sql = "
    SELECT 
        COUNT(DISTINCT rs.id) as total_reports,
        COUNT(DISTINCT CASE WHEN rs.status = 'completed' THEN rs.id END) as completed_reports,
        COUNT(DISTINCT CASE WHEN rs.status = 'pending' THEN rs.id END) as pending_reports,
        COUNT(DISTINCT rs.unit_id) as units_reported,
        COUNT(pr.id) as total_referrals
    FROM report_sessions rs
    LEFT JOIN patient_referrals pr ON rs.id = pr.session_id
    WHERE DATE(rs.created_at) = ?
";
$params = [$filter_date];

if ($filter_unit) {
    $sql .= " AND rs.unit_id = ?";
    $params[] = $filter_unit;
}

$stmt = $pdo->prepare($sql);
$stmt->execute($params);
$stats = $stmt->fetch();

// Get recent reports for preview
$sql = "
    SELECT 
        rs.id, rs.reporter_name, rs.report_date, rs.shift, rs.status, rs.created_at,
        u.name as unit_name,
        COUNT(r.id) as answered_questions,
        COUNT(pr.id) as patient_referrals
    FROM report_sessions rs
    LEFT JOIN units u ON rs.unit_id = u.id
    LEFT JOIN reports r ON rs.id = r.session_id
    LEFT JOIN patient_referrals pr ON rs.id = pr.session_id
    WHERE DATE(rs.created_at) = ?
";
$params = [$filter_date];

if ($filter_unit) {
    $sql .= " AND rs.unit_id = ?";
    $params[] = $filter_unit;
}

$sql .= " GROUP BY rs.id ORDER BY rs.created_at DESC LIMIT 10";

$stmt = $pdo->prepare($sql);
$stmt->execute($params);
$recent_reports = $stmt->fetchAll();

function generateExcel($date, $unit_filter = '') {
    global $pdo;
    
    // Create filename
    $filename = 'Laporan_Harian_' . date('Y-m-d', strtotime($date));
    if ($unit_filter) {
        $stmt = $pdo->prepare("SELECT name FROM units WHERE id = ?");
        $stmt->execute([$unit_filter]);
        $unit = $stmt->fetch();
        if ($unit) {
            $filename .= '_' . str_replace(' ', '_', $unit['name']);
        }
    }
    $filename .= '.xls';
    
    // Set headers for Excel download
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Cache-Control: max-age=0');
    
    // Start output
    echo '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">';
    echo '<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>';
    echo '<body>';
    
    // Title
    echo '<h1>LAPORAN HARIAN RUMAH SAKIT</h1>';
    echo '<h2>Tanggal: ' . date('d F Y', strtotime($date)) . '</h2>';
    if ($unit_filter) {
        echo '<h3>Unit: ' . $unit['name'] . '</h3>';
    }
    echo '<hr>';
    
    // Get all report sessions for the date
    $sql = "
        SELECT rs.*, u.name as unit_name
        FROM report_sessions rs
        LEFT JOIN units u ON rs.unit_id = u.id
        WHERE DATE(rs.created_at) = ?
    ";
    $params = [$date];
    
    if ($unit_filter) {
        $sql .= " AND rs.unit_id = ?";
        $params[] = $unit_filter;
    }
    
    $sql .= " ORDER BY u.name, rs.created_at";
    
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $sessions = $stmt->fetchAll();
    
    if (empty($sessions)) {
        echo '<p><strong>Tidak ada laporan untuk tanggal ' . date('d F Y', strtotime($date)) . '</strong></p>';
        echo '</body></html>';
        return;
    }
    
    foreach ($sessions as $session) {
        echo '<h2>UNIT: ' . strtoupper($session['unit_name']) . '</h2>';
        echo '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
        
        // Session info header
        echo '<tr style="background-color: #4472C4; color: white;">';
        echo '<td colspan="2"><strong>INFORMASI PELAPOR</strong></td>';
        echo '</tr>';
        echo '<tr><td><strong>Nama Pelapor</strong></td><td>' . htmlspecialchars($session['reporter_name']) . '</td></tr>';
        echo '<tr><td><strong>Tanggal Laporan</strong></td><td>' . date('d F Y', strtotime($session['report_date'])) . '</td></tr>';
        echo '<tr><td><strong>Shift</strong></td><td>' . strtoupper($session['shift']) . '</td></tr>';
        echo '<tr><td><strong>Status</strong></td><td>' . strtoupper($session['status']) . '</td></tr>';
        echo '<tr><td><strong>Waktu Input</strong></td><td>' . date('d F Y H:i', strtotime($session['created_at'])) . '</td></tr>';
        
        // Questions and answers
        echo '<tr style="background-color: #70AD47; color: white;">';
        echo '<td colspan="2"><strong>PERTANYAAN DAN JAWABAN</strong></td>';
        echo '</tr>';
        
        // Get questions and answers for this session
        $stmt = $pdo->prepare("
            SELECT q.question_text, q.question_order, r.answer
            FROM questions q
            LEFT JOIN reports r ON q.id = r.question_id AND r.session_id = ?
            WHERE q.unit_id = ? AND q.is_active = 1 AND q.deleted_at IS NULL
            ORDER BY q.question_order, q.id
        ");
        $stmt->execute([$session['id'], $session['unit_id']]);
        $questions = $stmt->fetchAll();
        
        if (empty($questions)) {
            echo '<tr><td colspan="2"><em>Tidak ada pertanyaan untuk unit ini</em></td></tr>';
        } else {
            $no = 1;
            foreach ($questions as $question) {
                $answer = $question['answer'] ?: '<em>Belum dijawab</em>';
                echo '<tr>';
                echo '<td style="width: 60%;"><strong>' . $no . '. ' . htmlspecialchars($question['question_text']) . '</strong></td>';
                echo '<td style="width: 40%;">' . htmlspecialchars($answer) . '</td>';
                echo '</tr>';
                $no++;
            }
        }
        
        // Patient referrals
        echo '<tr style="background-color: #E7E6E6;">';
        echo '<td colspan="2"><strong>DATA PASIEN RUJUK</strong></td>';
        echo '</tr>';
        
        $stmt = $pdo->prepare("
            SELECT * FROM patient_referrals 
            WHERE session_id = ? 
            ORDER BY created_at
        ");
        $stmt->execute([$session['id']]);
        $referrals = $stmt->fetchAll();
        
        if (empty($referrals)) {
            echo '<tr><td colspan="2"><em>Tidak ada data pasien rujuk</em></td></tr>';
        } else {
            echo '<tr style="background-color: #FFC000;">';
            echo '<td colspan="2"><strong>DETAIL PASIEN RUJUK (' . count($referrals) . ' pasien)</strong></td>';
            echo '</tr>';
            
            $no_rujuk = 1;
            foreach ($referrals as $referral) {
                echo '<tr><td colspan="2"><strong>PASIEN RUJUK #' . $no_rujuk . '</strong></td></tr>';
                echo '<tr><td>Nama Pasien</td><td>' . htmlspecialchars($referral['patient_name']) . '</td></tr>';
                echo '<tr><td>Umur</td><td>' . htmlspecialchars($referral['patient_age']) . '</td></tr>';
                echo '<tr><td>No. Rekam Medis</td><td>' . htmlspecialchars($referral['medical_record']) . '</td></tr>';
                echo '<tr><td>Asal Ruangan</td><td>' . htmlspecialchars($referral['origin_room']) . '</td></tr>';
                echo '<tr><td>Jenis Jaminan</td><td>' . htmlspecialchars($referral['insurance_type']) . '</td></tr>';
                echo '<tr><td>Diagnosa</td><td>' . htmlspecialchars($referral['diagnosis']) . '</td></tr>';
                echo '<tr><td>RS Tujuan</td><td>' . htmlspecialchars($referral['destination_hospital']) . '</td></tr>';
                echo '<tr><td>Alasan Rujukan</td><td>' . htmlspecialchars($referral['referral_reason'] ?: 'Tidak ada keterangan') . '</td></tr>';
                echo '<tr><td>Waktu Input</td><td>' . date('d F Y H:i', strtotime($referral['created_at'])) . '</td></tr>';
                
                if ($no_rujuk < count($referrals)) {
                    echo '<tr><td colspan="2" style="border-top: 2px solid #000;"></td></tr>';
                }
                $no_rujuk++;
            }
        }
        
        echo '</table>';
        echo '<br><br>';
    }
    
    // Summary table
    echo '<h2>RINGKASAN LAPORAN</h2>';
    echo '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">';
    echo '<tr style="background-color: #4472C4; color: white;">';
    echo '<th>Unit</th><th>Status</th><th>Pelapor</th><th>Shift</th><th>Jumlah Jawaban</th><th>Pasien Rujuk</th>';
    echo '</tr>';
    
    foreach ($sessions as $session) {
        // Count answers for this session
        $stmt = $pdo->prepare("SELECT COUNT(*) as count FROM reports WHERE session_id = ?");
        $stmt->execute([$session['id']]);
        $answer_count = $stmt->fetch()['count'];
        
        // Count referrals for this session
        $stmt = $pdo->prepare("SELECT COUNT(*) as count FROM patient_referrals WHERE session_id = ?");
        $stmt->execute([$session['id']]);
        $referral_count = $stmt->fetch()['count'];
        
        $status_color = $session['status'] === 'completed' ? '#70AD47' : '#FFC000';
        
        echo '<tr>';
        echo '<td>' . htmlspecialchars($session['unit_name']) . '</td>';
        echo '<td style="background-color: ' . $status_color . '; color: white;"><strong>' . strtoupper($session['status']) . '</strong></td>';
        echo '<td>' . htmlspecialchars($session['reporter_name']) . '</td>';
        echo '<td>' . strtoupper($session['shift']) . '</td>';
        echo '<td>' . $answer_count . '</td>';
        echo '<td>' . $referral_count . '</td>';
        echo '</tr>';
    }
    echo '</table>';
    
    // Footer
    echo '<br><hr>';
    echo '<p><em>Laporan digenerate pada: ' . date('d F Y H:i:s') . '</em></p>';
    echo '<p><em>Sistem Laporan Harian Rumah Sakit</em></p>';
    
    echo '</body></html>';
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Download Excel - Sistem Laporan RS</title>
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
        .download-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 20px;
            border: none;
        }
        .preview-table {
            font-size: 0.9em;
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
                        <a href="monitoring_reports.php" class="nav-link">
                            <i class="bi bi-graph-up me-2"></i>Monitoring Laporan
                        </a>
                        <a href="#" class="nav-link active">
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
                        <h2 class="mb-0">Download Laporan Excel</h2>
                        <div class="d-flex align-items-center">
                            <span class="badge bg-success me-2">Online</span>
                            <span class="text-muted"><?= date('d F Y') ?></span>
                        </div>
                    </div>
                    
                    <!-- Filter Form -->
                    <div class="download-card p-4 mb-4">
                        <h4 class="mb-3">
                            <i class="bi bi-funnel me-2"></i>
                            Filter & Download
                        </h4>
                        <form method="GET" class="row g-3">
                            <div class="col-md-4">
                                <label class="form-label text-white">Tanggal Laporan</label>
                                <input type="date" class="form-control" name="date" value="<?= $filter_date ?>" required>
                            </div>
                            
                            <div class="col-md-4">
                                <label class="form-label text-white">Unit (Opsional)</label>
                                <select class="form-select" name="unit">
                                    <option value="">Semua Unit</option>
                                    <?php foreach ($units as $unit): ?>
                                        <option value="<?= $unit['id'] ?>" <?= $filter_unit == $unit['id'] ? 'selected' : '' ?>>
                                            <?= $unit['name'] ?>
                                        </option>
                                    <?php endforeach; ?>
                                </select>
                            </div>
                            
                            <div class="col-md-4 d-flex align-items-end">
                                <button type="submit" class="btn btn-light me-2">
                                    <i class="bi bi-search me-2"></i>Preview
                                </button>
                                <a href="?download=excel&date=<?= $filter_date ?>&unit=<?= $filter_unit ?>" 
                                   class="btn btn-warning">
                                    <i class="bi bi-download me-2"></i>Download Excel
                                </a>
                            </div>
                        </form>
                    </div>
                    
                    <!-- Statistics -->
                    <div class="row mb-4">
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-primary text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['total_reports'] ?></h4>
                                    <p class="card-text mb-0">Total Laporan</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-success text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['completed_reports'] ?></h4>
                                    <p class="card-text mb-0">Selesai</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-warning text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['pending_reports'] ?></h4>
                                    <p class="card-text mb-0">Pending</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-info text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['units_reported'] ?></h4>
                                    <p class="card-text mb-0">Unit Lapor</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-danger text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= $stats['total_referrals'] ?></h4>
                                    <p class="card-text mb-0">Pasien Rujuk</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-2 mb-3">
                            <div class="card stats-card bg-secondary text-white">
                                <div class="card-body text-center">
                                    <h4 class="card-title"><?= count($units) ?></h4>
                                    <p class="card-text mb-0">Total Unit</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Preview Data -->
                    <div class="card border-0 shadow-sm">
                        <div class="card-header bg-white">
                            <h5 class="mb-0">
                                <i class="bi bi-eye me-2"></i>
                                Preview Laporan: <?= date('d F Y', strtotime($filter_date)) ?>
                                <?php if ($filter_unit): ?>
                                    <?php 
                                    $unit_name = '';
                                    foreach ($units as $unit) {
                                        if ($unit['id'] == $filter_unit) {
                                            $unit_name = $unit['name'];
                                            break;
                                        }
                                    }
                                    ?>
                                    - <?= $unit_name ?>
                                <?php endif; ?>
                            </h5>
                        </div>
                        <div class="card-body">
                            <?php if (empty($recent_reports)): ?>
                                <div class="text-center py-5">
                                    <i class="bi bi-inbox fs-1 text-muted"></i>
                                    <h5 class="mt-3 text-muted">Tidak Ada Laporan</h5>
                                    <p class="text-muted">Tidak ada laporan untuk tanggal dan filter yang dipilih.</p>
                                </div>
                            <?php else: ?>
                                <div class="table-responsive">
                                    <table class="table table-hover preview-table">
                                        <thead class="table-light">
                                            <tr>
                                                <th>#</th>
                                                <th>Unit</th>
                                                <th>Pelapor</th>
                                                <th>Tanggal</th>
                                                <th>Shift</th>
                                                <th>Status</th>
                                                <th>Jawaban</th>
                                                <th>Rujukan</th>
                                                <th>Waktu Input</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <?php foreach ($recent_reports as $index => $report): ?>
                                                <tr>
                                                    <td><?= $index + 1 ?></td>
                                                    <td>
                                                        <span class="badge bg-primary"><?= $report['unit_name'] ?></span>
                                                    </td>
                                                    <td><?= $report['reporter_name'] ?></td>
                                                    <td><?= formatDate($report['report_date']) ?></td>
                                                    <td>
                                                        <span class="badge bg-secondary"><?= ucfirst($report['shift']) ?></span>
                                                    </td>
                                                    <td>
                                                        <?php if ($report['status'] === 'completed'): ?>
                                                            <span class="badge bg-success">Selesai</span>
                                                        <?php else: ?>
                                                            <span class="badge bg-warning">Pending</span>
                                                        <?php endif; ?>
                                                    </td>
                                                    <td>
                                                        <span class="badge bg-info"><?= $report['answered_questions'] ?> jawaban</span>
                                                    </td>
                                                    <td>
                                                        <?php if ($report['patient_referrals'] > 0): ?>
                                                            <span class="badge bg-danger"><?= $report['patient_referrals'] ?> rujuk</span>
                                                        <?php else: ?>
                                                            <span class="text-muted">-</span>
                                                        <?php endif; ?>
                                                    </td>
                                                    <td><?= formatDateTime($report['created_at']) ?></td>
                                                </tr>
                                            <?php endforeach; ?>
                                        </tbody>
                                    </table>
                                </div>
                                
                                <div class="text-center mt-3">
                                    <a href="?download=excel&date=<?= $filter_date ?>&unit=<?= $filter_unit ?>" 
                                       class="btn btn-success btn-lg">
                                        <i class="bi bi-file-earmark-excel me-2"></i>
                                        Download Excel Lengkap
                                    </a>
                                </div>
                            <?php endif; ?>
                        </div>
                    </div>
                    
                    <!-- Info Card -->
                    <div class="card border-0 shadow-sm mt-4">
                        <div class="card-body">
                            <h6 class="card-title">
                                <i class="bi bi-info-circle me-2"></i>
                                Informasi Download Excel
                            </h6>
                            <div class="row">
                                <div class="col-md-6">
                                    <h6 class="text-primary">Data yang Disertakan:</h6>
                                    <ul class="list-unstyled">
                                        <li><i class="bi bi-check-circle-fill text-success me-2"></i>Identitas pelapor (nama, tanggal, shift)</li>
                                        <li><i class="bi bi-check-circle-fill text-success me-2"></i>Semua pertanyaan dan jawaban per unit</li>
                                        <li><i class="bi bi-check-circle-fill text-success me-2"></i>Detail data pasien rujuk lengkap</li>
                                        <li><i class="bi bi-check-circle-fill text-success me-2"></i>Ringkasan laporan per unit</li>
                                    </ul>
                                </div>
                                <div class="col-md-6">
                                    <h6 class="text-warning">Format Excel:</h6>
                                    <ul class="list-unstyled">
                                        <li><i class="bi bi-file-earmark-excel text-success me-2"></i>Format .xls (kompatibel semua versi Excel)</li>
                                        <li><i class="bi bi-table text-info me-2"></i>Tabel terstruktur dengan border</li>
                                        <li><i class="bi bi-palette text-primary me-2"></i>Header berwarna untuk mudah dibaca</li>
                                        <li><i class="bi bi-clock text-secondary me-2"></i>Timestamp generate otomatis</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Auto-refresh statistics every 30 seconds
        setInterval(function() {
            if (!document.hidden) {
                location.reload();
            }
        }, 30000);
        
        // Download confirmation
        document.querySelector('a[href*="download=excel"]').addEventListener('click', function(e) {
            const date = '<?= $filter_date ?>';
            const unit = '<?= $filter_unit ?>';
            const reportCount = <?= $stats['total_reports'] ?>;
            
            if (reportCount === 0) {
                e.preventDefault();
                alert('Tidak ada data laporan untuk di-download!');
                return;
            }
            
            const message = `Akan mendownload laporan Excel untuk:\n\nTanggal: ${date}\n${unit ? 'Unit: ' + unit + '\n' : 'Semua Unit\n'}Total: ${reportCount} laporan\n\nLanjutkan?`;
            
            if (!confirm(message)) {
                e.preventDefault();
            }
        });
        
        // Form validation
        document.querySelector('form').addEventListener('submit', function(e) {
            const dateInput = this.querySelector('input[name="date"]');
            if (!dateInput.value) {
                e.preventDefault();
                alert('Tanggal laporan wajib diisi!');
                dateInput.focus();
            }
        });
    </script>
</body>
</html>