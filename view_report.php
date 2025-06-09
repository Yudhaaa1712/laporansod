<?php
require_once 'config.php';
requireLogin();

$success = $_GET['success'] ?? null;
$session_id = $_GET['id'] ?? null;

if (!$session_id) {
    header('Location: ' . (isAdmin() ? 'admin_dashboard.php' : 'unit_dashboard.php'));
    exit;
}

// Get report session details
$stmt = $pdo->prepare("
    SELECT rs.*, u.name as unit_name 
    FROM report_sessions rs 
    JOIN units u ON rs.unit_id = u.id 
    WHERE rs.id = ?
");
$stmt->execute([$session_id]);
$session = $stmt->fetch();

if (!$session) {
    header('Location: ' . (isAdmin() ? 'admin_dashboard.php' : 'unit_dashboard.php'));
    exit;
}

// Check access rights for unit users
if (isUnit() && $session['unit_id'] != $_SESSION['unit_id']) {
    header('Location: unauthorized.php');
    exit;
}

// Get questions and answers
$stmt = $pdo->prepare("
    SELECT q.question_text, q.question_order, r.answer, r.created_at as answered_at
    FROM questions q 
    LEFT JOIN reports r ON q.id = r.question_id AND r.session_id = ?
    WHERE q.unit_id = ? AND q.is_active = 1 AND q.deleted_at IS NULL
    ORDER BY q.question_order, q.id
");
$stmt->execute([$session_id, $session['unit_id']]);
$questions = $stmt->fetchAll();

// Get patient referrals
$stmt = $pdo->prepare("
    SELECT * FROM patient_referrals 
    WHERE session_id = ?
    ORDER BY created_at
");
$stmt->execute([$session_id]);
$referrals = $stmt->fetchAll();

// Calculate completion stats
$total_questions = count($questions);
$answered_questions = count(array_filter($questions, fn($q) => !empty($q['answer'])));
$completion_percentage = $total_questions > 0 ? round(($answered_questions / $total_questions) * 100) : 0;
$is_complete = $completion_percentage >= 100;
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detail Laporan - <?= $session['unit_name'] ?></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        .sidebar {
            min-height: 100vh;
            background: linear-gradient(135deg, <?= isAdmin() ? '#667eea 0%, #764ba2' : '#4facfe 0%, #00f2fe' ?> 100%);
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
        .question-card {
            border-radius: 15px;
            border: none;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            transition: all 0.3s ease;
        }
        .question-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }
        .question-card.answered {
            border-left: 5px solid #28a745;
        }
        .question-card.unanswered {
            border-left: 5px solid #dc3545;
        }
        .completion-bar {
            height: 10px;
            border-radius: 5px;
            overflow: hidden;
        }
        .report-header {
            background: linear-gradient(135deg, <?= isAdmin() ? '#667eea 0%, #764ba2' : '#4facfe 0%, #00f2fe' ?> 100%);
            color: white;
            border-radius: 15px;
            padding: 25px;
            margin-bottom: 25px;
        }
        .referral-card {
            border-radius: 12px;
            border: none;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 15px;
        }
        @media print {
            .sidebar, .no-print { display: none !important; }
            .content-wrapper { margin-left: 0 !important; }
            .report-header { background: #6c757d !important; }
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="row">
            <!-- Sidebar -->
            <div class="col-md-3 col-lg-2 px-0 no-print">
                <div class="sidebar p-3">
                    <div class="text-center mb-4">
                        <i class="bi bi-hospital fs-1 text-white"></i>
                        <h5 class="text-white mt-2"><?= isAdmin() ? 'Admin Panel' : $session['unit_name'] ?></h5>
                        <small class="text-white-50">Halo, <?= $_SESSION['name'] ?></small>
                    </div>
                    
                    <nav class="nav flex-column">
                        <?php if (isAdmin()): ?>
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
                            <a href="monitoring_reports.php" class="nav-link active">
                                <i class="bi bi-graph-up me-2"></i>Monitoring Laporan
                            </a>
                            <a href="download_excel.php" class="nav-link">
                                <i class="bi bi-download me-2"></i>Download Excel
                            </a>
                        <?php else: ?>
                            <a href="unit_dashboard.php" class="nav-link">
                                <i class="bi bi-speedometer2 me-2"></i>Dashboard
                            </a>
                            <a href="fill_report.php" class="nav-link">
                                <i class="bi bi-file-text me-2"></i>Isi Laporan
                            </a>
                            <a href="patient_referral.php" class="nav-link">
                                <i class="bi bi-person-plus me-2"></i>Data Pasien Rujuk
                            </a>
                           
                        <?php endif; ?>
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
                    <!-- Success Message -->
                    <?php if ($success): ?>
                        <div class="alert alert-success alert-dismissible fade show" role="alert">
                            <i class="bi bi-check-circle me-2"></i>
                            <strong>Laporan berhasil disimpan!</strong> Terima kasih telah melengkapi laporan harian.
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Report Header -->
                    <div class="report-header">
                        <div class="row align-items-center">
                            <div class="col-md-8">
                                <h2 class="mb-2">
                                    <i class="bi bi-file-text me-2"></i>
                                    Detail Laporan Harian
                                </h2>
                                <h4 class="mb-3"><?= $session['unit_name'] ?></h4>
                                <div class="row">
                                    <div class="col-md-6">
                                        <p class="mb-1">
                                            <i class="bi bi-person me-2"></i>
                                            <strong>Pelapor:</strong> <?= $session['reporter_name'] ?>
                                        </p>
                                        <p class="mb-1">
                                            <i class="bi bi-calendar me-2"></i>
                                            <strong>Tanggal:</strong> <?= formatDate($session['report_date']) ?>
                                        </p>
                                    </div>
                                    <div class="col-md-6">
                                        <p class="mb-1">
                                            <i class="bi bi-clock me-2"></i>
                                            <strong>Shift:</strong> <?= ucfirst($session['shift']) ?>
                                        </p>
                                        <p class="mb-1">
                                            <i class="bi bi-clock-history me-2"></i>
                                            <strong>Waktu Input:</strong> <?= formatDateTime($session['created_at']) ?>
                                        </p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-4 text-end">
                                <div class="d-flex flex-column align-items-end no-print">
                                    <button class="btn btn-light btn-lg mb-2" onclick="window.print()">
                                        <i class="bi bi-printer me-2"></i>Print
                                    </button>
                                    <a href="download_excel.php?session_id=<?= $session_id ?>" class="btn btn-success btn-lg">
                                        <i class="bi bi-download me-2"></i>Download Excel
                                    </a>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Progress Overview -->
                    <div class="card border-0 shadow-sm mb-4">
                        <div class="card-body">
                            <div class="row align-items-center">
                                <div class="col-md-8">
                                    <h6 class="mb-2">Progress Pengisian Laporan</h6>
                                    <div class="completion-bar bg-light">
                                        <div class="bg-<?= $is_complete ? 'success' : 'warning' ?>" 
                                             style="width: <?= $completion_percentage ?>%; height: 100%;"></div>
                                    </div>
                                    <small class="text-muted mt-1 d-block">
                                        <?= $answered_questions ?> dari <?= $total_questions ?> pertanyaan terjawab (<?= $completion_percentage ?>%)
                                    </small>
                                </div>
                                <div class="col-md-4 text-end">
                                    <span class="badge <?= $is_complete ? 'bg-success' : 'bg-warning' ?> fs-6 p-3">
                                        <i class="bi bi-<?= $is_complete ? 'check-circle' : 'exclamation-triangle' ?> me-2"></i>
                                        <?= $is_complete ? 'Lengkap' : 'Belum Lengkap' ?>
                                    </span>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Questions and Answers -->
                    <div class="card border-0 shadow-sm mb-4">
                        <div class="card-header bg-white">
                            <h5 class="mb-0">
                                <i class="bi bi-question-circle me-2"></i>
                                Pertanyaan dan Jawaban (<?= count($questions) ?> pertanyaan)
                            </h5>
                        </div>
                        <div class="card-body">
                            <?php if (empty($questions)): ?>
                                <div class="text-center py-4">
                                    <i class="bi bi-question-circle fs-1 text-muted"></i>
                                    <h6 class="mt-3 text-muted">Tidak Ada Pertanyaan</h6>
                                    <p class="text-muted">Belum ada pertanyaan yang dibuat untuk unit ini.</p>
                                </div>
                            <?php else: ?>
                                <?php foreach ($questions as $index => $question): ?>
                                    <div class="question-card card <?= !empty($question['answer']) ? 'answered' : 'unanswered' ?>">
                                        <div class="card-body">
                                            <div class="d-flex align-items-start">
                                                <span class="badge bg-primary me-3 mt-1"><?= $index + 1 ?></span>
                                                <div class="flex-grow-1">
                                                    <h6 class="question-text mb-3"><?= nl2br(htmlspecialchars($question['question_text'])) ?></h6>
                                                    
                                                    <?php if (!empty($question['answer'])): ?>
                                                        <div class="answer-section p-3 bg-light rounded">
                                                            <div class="d-flex align-items-center mb-2">
                                                                <i class="bi bi-check-circle text-success me-2"></i>
                                                                <strong class="text-success">Jawaban:</strong>
                                                                <?php if ($question['answered_at']): ?>
                                                                    <small class="text-muted ms-auto">
                                                                        <?= formatDateTime($question['answered_at']) ?>
                                                                    </small>
                                                                <?php endif; ?>
                                                            </div>
                                                            <p class="mb-0"><?= nl2br(htmlspecialchars($question['answer'])) ?></p>
                                                        </div>
                                                    <?php else: ?>
                                                        <div class="answer-section p-3 bg-light rounded border border-danger">
                                                            <div class="d-flex align-items-center">
                                                                <i class="bi bi-x-circle text-danger me-2"></i>
                                                                <span class="text-danger"><strong>Belum dijawab</strong></span>
                                                            </div>
                                                        </div>
                                                    <?php endif; ?>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                <?php endforeach; ?>
                            <?php endif; ?>
                        </div>
                    </div>
                    
                    <!-- Patient Referrals -->
                    <?php if (!empty($referrals)): ?>
                        <div class="card border-0 shadow-sm mb-4">
                            <div class="card-header bg-white">
                                <h5 class="mb-0">
                                    <i class="bi bi-person-plus me-2"></i>
                                    Data Pasien Rujuk (<?= count($referrals) ?> pasien)
                                </h5>
                            </div>
                            <div class="card-body">
                                <?php foreach ($referrals as $index => $referral): ?>
                                    <div class="referral-card card">
                                        <div class="card-body">
                                            <div class="row">
                                                <div class="col-md-6">
                                                    <h6 class="text-primary mb-3">
                                                        <i class="bi bi-person-heart me-2"></i>
                                                        Data Pasien #<?= $index + 1 ?>
                                                    </h6>
                                                    <table class="table table-sm table-borderless">
                                                        <tr>
                                                            <td width="40%"><strong>Nama</strong></td>
                                                            <td><?= $referral['patient_name'] ?></td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>Umur</strong></td>
                                                            <td><?= $referral['patient_age'] ?></td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>No. RM</strong></td>
                                                            <td><code><?= $referral['medical_record'] ?></code></td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>Asal Ruangan</strong></td>
                                                            <td><?= $referral['origin_room'] ?></td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>Jaminan</strong></td>
                                                            <td><span class="badge bg-info"><?= $referral['insurance_type'] ?></span></td>
                                                        </tr>
                                                    </table>
                                                </div>
                                                <div class="col-md-6">
                                                    <h6 class="text-success mb-3">
                                                        <i class="bi bi-hospital me-2"></i>
                                                        Informasi Rujukan
                                                    </h6>
                                                    <table class="table table-sm table-borderless">
                                                        <tr>
                                                            <td width="40%"><strong>RS Tujuan</strong></td>
                                                            <td><?= $referral['destination_hospital'] ?></td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>Waktu Input</strong></td>
                                                            <td><?= formatDateTime($referral['created_at']) ?></td>
                                                        </tr>
                                                    </table>
                                                </div>
                                            </div>
                                            
                                            <div class="row mt-3">
                                                <div class="col-12">
                                                    <h6 class="text-warning mb-2">
                                                        <i class="bi bi-clipboard-heart me-2"></i>
                                                        Diagnosa
                                                    </h6>
                                                    <div class="p-3 bg-light rounded">
                                                        <?= nl2br(htmlspecialchars($referral['diagnosis'])) ?>
                                                    </div>
                                                </div>
                                            </div>
                                            
                                            <?php if ($referral['referral_reason']): ?>
                                                <div class="row mt-3">
                                                    <div class="col-12">
                                                        <h6 class="text-danger mb-2">
                                                            <i class="bi bi-info-circle me-2"></i>
                                                            Alasan Rujukan
                                                        </h6>
                                                        <div class="p-3 bg-light rounded">
                                                            <?= nl2br(htmlspecialchars($referral['referral_reason'])) ?>
                                                        </div>
                                                    </div>
                                                </div>
                                            <?php endif; ?>
                                        </div>
                                    </div>
                                <?php endforeach; ?>
                            </div>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Action Buttons -->
                    <div class="d-flex justify-content-between align-items-center no-print">
                        <a href="<?= isAdmin() ? 'monitoring_reports.php' : 'unit_dashboard.php' ?>" class="btn btn-outline-secondary">
                            <i class="bi bi-arrow-left me-2"></i>
                            Kembali
                        </a>
                        
                        <div class="btn-group">
                            <button class="btn btn-outline-primary" onclick="window.print()">
                                <i class="bi bi-printer me-2"></i>Print Laporan
                            </button>
                            <a href="download_excel.php?session_id=<?= $session_id ?>" class="btn btn-success">
                                <i class="bi bi-download me-2"></i>Download Excel
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Auto-hide success message after 5 seconds
        setTimeout(function() {
            var alert = document.querySelector('.alert-success');
            if (alert) {
                var bsAlert = new bootstrap.Alert(alert);
                bsAlert.close();
            }
        }, 5000);
        
        // Print function with custom styling
        function printReport() {
            window.print();
        }
        
        // Smooth scroll to incomplete questions
        function scrollToIncomplete() {
            const incompleteCard = document.querySelector('.question-card.unanswered');
            if (incompleteCard) {
                incompleteCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
                incompleteCard.style.boxShadow = '0 0 20px rgba(220, 53, 69, 0.5)';
                setTimeout(() => {
                    incompleteCard.style.boxShadow = '';
                }, 2000);
            }
        }
        
        // Keyboard shortcuts
        document.addEventListener('keydown', function(e) {
            // Ctrl+P for print
            if (e.ctrlKey && e.key === 'p') {
                e.preventDefault();
                window.print();
            }
            
            // Ctrl+D for download
            if (e.ctrlKey && e.key === 'd') {
                e.preventDefault();
                window.location.href = 'download_excel.php?session_id=<?= $session_id ?>';
            }
            
            // Escape to go back
            if (e.key === 'Escape') {
                window.location.href = '<?= isAdmin() ? 'monitoring_reports.php' : 'unit_dashboard.php' ?>';
            }
        });
    </script>
</body>
</html>