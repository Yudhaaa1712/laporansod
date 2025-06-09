<?php
require_once 'config.php';
requireUnit();

$message = '';
$message_type = '';

// Check if user has existing report today (completed or pending)
$stmt = $pdo->prepare("
    SELECT id, status FROM report_sessions 
    WHERE unit_id = ? AND DATE(created_at) = CURDATE()
");
$stmt->execute([$_SESSION['unit_id']]);
$existing_session = $stmt->fetch();

// If completed report exists, redirect to view
if ($existing_session && $existing_session['status'] === 'completed') {
    header('Location: view_report.php?id=' . $existing_session['id']);
    exit;
}

// Get unit info
$stmt = $pdo->prepare("SELECT name FROM units WHERE id = ?");
$stmt->execute([$_SESSION['unit_id']]);
$unit = $stmt->fetch();

// Get questions for this unit
$stmt = $pdo->prepare("
    SELECT * FROM questions 
    WHERE unit_id = ? AND is_active = 1 AND deleted_at IS NULL 
    ORDER BY question_order, id
");
$stmt->execute([$_SESSION['unit_id']]);
$questions = $stmt->fetchAll();

// If continuing existing pending report
$continue_session = false;
$existing_answers = [];
if ($existing_session && $existing_session['status'] === 'pending') {
    $continue_session = true;
    $_SESSION['current_session_id'] = $existing_session['id'];
    
    // Get existing answers
    $stmt = $pdo->prepare("
        SELECT question_id, answer FROM reports 
        WHERE session_id = ?
    ");
    $stmt->execute([$existing_session['id']]);
    $answers = $stmt->fetchAll();
    
    foreach ($answers as $answer) {
        $existing_answers[$answer['question_id']] = $answer['answer'];
    }
    
    // Get session details
    $stmt = $pdo->prepare("
        SELECT reporter_name, report_date, shift 
        FROM report_sessions WHERE id = ?
    ");
    $stmt->execute([$existing_session['id']]);
    $session_details = $stmt->fetch();
}

// Handle form submission
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_POST['step']) && $_POST['step'] === '1') {
        // Step 1: Save reporter info
        $reporter_name = sanitize($_POST['reporter_name']);
        $report_date = $_POST['report_date'];
        $shift = $_POST['shift'];
        
        if (!empty($reporter_name) && !empty($report_date) && !empty($shift)) {
            // Check if report for this date already exists (only completed ones)
            $stmt = $pdo->prepare("
                SELECT id FROM report_sessions 
                WHERE unit_id = ? AND report_date = ? AND status = 'completed'
            ");
            $stmt->execute([$_SESSION['unit_id'], $report_date]);
            
            if ($stmt->rowCount() == 0) {
                $stmt = $pdo->prepare("
                    INSERT INTO report_sessions (unit_id, reporter_name, report_date, shift, status) 
                    VALUES (?, ?, ?, ?, 'pending')
                ");
                
                if ($stmt->execute([$_SESSION['unit_id'], $reporter_name, $report_date, $shift])) {
                    $session_id = $pdo->lastInsertId();
                    $_SESSION['current_session_id'] = $session_id;
                    $show_questions = true;
                } else {
                    $message = 'Gagal menyimpan data pelapor!';
                    $message_type = 'danger';
                }
            } else {
                $message = 'Laporan lengkap untuk tanggal ini sudah ada!';
                $message_type = 'warning';
            }
        } else {
            $message = 'Harap isi semua field!';
            $message_type = 'danger';
        }
    }
    
    if (isset($_POST['step']) && $_POST['step'] === '2') {
        // Step 2: Save answers (can be partial)
        $session_id = $_SESSION['current_session_id'];
        $answers = $_POST['answers'];
        $action = $_POST['action'] ?? 'save'; // 'save' or 'complete'
        
        // Delete existing answers first to update
        $stmt = $pdo->prepare("DELETE FROM reports WHERE session_id = ?");
        $stmt->execute([$session_id]);
        
        $success_count = 0;
        $total_questions = count($questions);
        $answered_questions = 0;
        
        foreach ($answers as $question_id => $answer) {
            if (!empty(trim($answer))) {
                $stmt = $pdo->prepare("
                    INSERT INTO reports (session_id, question_id, answer) 
                    VALUES (?, ?, ?)
                ");
                
                if ($stmt->execute([$session_id, $question_id, trim($answer)])) {
                    $success_count++;
                    $answered_questions++;
                }
            }
        }
        
        // Determine status based on answers and action
        $new_status = 'pending';
        if ($action === 'complete') {
            if ($answered_questions >= $total_questions) {
                $new_status = 'completed';
            } else {
                $message = 'Tidak bisa menyelesaikan laporan! Masih ada ' . ($total_questions - $answered_questions) . ' pertanyaan yang belum dijawab.';
                $message_type = 'warning';
            }
        }
        
        // Update session status
        $stmt = $pdo->prepare("UPDATE report_sessions SET status = ? WHERE id = ?");
        $stmt->execute([$new_status, $session_id]);
        
        if ($success_count > 0 || $answered_questions === 0) {
            if ($new_status === 'completed') {
                unset($_SESSION['current_session_id']);
                header('Location: view_report.php?id=' . $session_id . '&success=1');
                exit;
            } else {
                $message = 'Jawaban berhasil disimpan! Status: Pending (' . $answered_questions . '/' . $total_questions . ' pertanyaan dijawab)';
                $message_type = 'info';
                // Keep the form open for further editing
                $show_questions = true;
                
                // Refresh answers
                $stmt = $pdo->prepare("
                    SELECT question_id, answer FROM reports 
                    WHERE session_id = ?
                ");
                $stmt->execute([$session_id]);
                $saved_answers = $stmt->fetchAll();
                
                $existing_answers = [];
                foreach ($saved_answers as $answer) {
                    $existing_answers[$answer['question_id']] = $answer['answer'];
                }
            }
        } else {
            $message = 'Tidak ada jawaban yang tersimpan!';
            $message_type = 'warning';
        }
    }
}

$show_questions = isset($show_questions) && $show_questions;
if ($continue_session) {
    $show_questions = true;
}

// Calculate progress
$answered_count = count($existing_answers);
$total_count = count($questions);
$progress_percentage = $total_count > 0 ? round(($answered_count / $total_count) * 100) : 0;
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Isi Laporan - <?= $unit['name'] ?></title>
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
            border-left: 4px solid #28a745;
        }
        .question-card.unanswered {
            border-left: 4px solid #ffc107;
        }
        .step-indicator {
            display: flex;
            justify-content: center;
            margin-bottom: 30px;
        }
        .step {
            display: flex;
            align-items: center;
            padding: 10px 20px;
            background: #e9ecef;
            border-radius: 25px;
            margin: 0 10px;
            transition: all 0.3s ease;
        }
        .step.active {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
        }
        .step.completed {
            background: #28a745;
            color: white;
        }
        .progress-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 15px;
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
                        <a href="unit_dashboard.php" class="nav-link">
                            <i class="bi bi-speedometer2 me-2"></i>Dashboard
                        </a>
                        <a href="#" class="nav-link active">
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
                            <h2 class="mb-0">
                                <?= $continue_session ? 'Lanjutkan Laporan Harian' : 'Isi Laporan Harian' ?>
                                <?php if ($continue_session): ?>
                                    <span class="badge bg-warning ms-2">Pending</span>
                                <?php endif; ?>
                            </h2>
                            <p class="text-muted mb-0">Unit: <?= $unit['name'] ?></p>
                        </div>
                        <a href="unit_dashboard.php" class="btn btn-outline-secondary">
                            <i class="bi bi-arrow-left me-2"></i>Kembali
                        </a>
                    </div>
                    
                    <!-- Progress Bar for continuing session -->
                    <?php if ($continue_session): ?>
                        <div class="progress-card p-4 mb-4">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <h6 class="mb-0">Progress Laporan</h6>
                                <span><?= $answered_count ?>/<?= $total_count ?> pertanyaan</span>
                            </div>
                            <div class="progress" style="height: 8px;">
                                <div class="progress-bar bg-light" role="progressbar" 
                                     style="width: <?= $progress_percentage ?>%" 
                                     aria-valuenow="<?= $progress_percentage ?>" 
                                     aria-valuemin="0" aria-valuemax="100"></div>
                            </div>
                            <small class="opacity-75 mt-2 d-block">
                                <?= $progress_percentage ?>% selesai | 
                                Tanggal: <?= formatDate($session_details['report_date']) ?> | 
                                Shift: <?= ucfirst($session_details['shift']) ?>
                            </small>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Step Indicator -->
                    <?php if (!$continue_session): ?>
                    <div class="step-indicator">
                        <div class="step <?= !$show_questions ? 'active' : 'completed' ?>">
                            <i class="bi bi-person-circle me-2"></i>
                            <span>Data Pelapor</span>
                        </div>
                        <div class="step <?= $show_questions ? 'active' : '' ?>">
                            <i class="bi bi-file-text me-2"></i>
                            <span>Isi Pertanyaan</span>
                        </div>
                    </div>
                    <?php endif; ?>
                    
                    <?php if ($message): ?>
                        <div class="alert alert-<?= $message_type ?> alert-dismissible fade show" role="alert">
                            <?= $message ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <?php if (!$show_questions && !$continue_session): ?>
                        <!-- Step 1: Reporter Information -->
                        <div class="card border-0 shadow-sm">
                            <div class="card-header bg-white">
                                <h5 class="mb-0">
                                    <i class="bi bi-person-circle me-2"></i>
                                    Informasi Pelapor
                                </h5>
                            </div>
                            <div class="card-body">
                                <form method="POST">
                                    <input type="hidden" name="step" value="1">
                                    
                                    <div class="row">
                                        <div class="col-md-6">
                                            <div class="mb-3">
                                                <label class="form-label">Nama Pelapor *</label>
                                                <input type="text" class="form-control" name="reporter_name" 
                                                       placeholder="Masukkan nama pelapor" required>
                                            </div>
                                        </div>
                                        
                                        <div class="col-md-6">
                                            <div class="mb-3">
                                                <label class="form-label">Tanggal Laporan *</label>
                                                <input type="date" class="form-control" name="report_date" 
                                                       value="<?= date('Y-m-d') ?>" required>
                                            </div>
                                        </div>
                                    </div>
                                    
                                    <div class="row">
                                        <div class="col-md-6">
                                            <div class="mb-3">
                                                <label class="form-label">Shift *</label>
                                                <select class="form-select" name="shift" required>
                                                    <option value="">Pilih Shift</option>
                                                    <option value="pagi">Pagi</option>
                                                    <option value="siang">Siang</option>
                                                    <option value="malam">Malam</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                    
                                    <div class="d-flex justify-content-end">
                                        <button type="submit" class="btn btn-primary">
                                            <i class="bi bi-arrow-right me-2"></i>
                                            Lanjut ke Pertanyaan
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                        
                    <?php else: ?>
                        <!-- Step 2: Questions -->
                        <div class="card border-0 shadow-sm">
                            <div class="card-header bg-white d-flex justify-content-between align-items-center">
                                <h5 class="mb-0">
                                    <i class="bi bi-file-text me-2"></i>
                                    Pertanyaan Laporan (<?= count($questions) ?> pertanyaan)
                                </h5>
                                <?php if ($continue_session): ?>
                                    <span class="badge bg-info">
                                        <?= $answered_count ?>/<?= $total_count ?> dijawab
                                    </span>
                                <?php endif; ?>
                            </div>
                            <div class="card-body">
                                <?php if (empty($questions)): ?>
                                    <div class="text-center py-5">
                                        <i class="bi bi-question-circle fs-1 text-muted"></i>
                                        <h5 class="mt-3 text-muted">Belum Ada Pertanyaan</h5>
                                        <p class="text-muted">Admin belum membuat pertanyaan untuk unit ini.</p>
                                        <a href="unit_dashboard.php" class="btn btn-primary">Kembali ke Dashboard</a>
                                    </div>
                                <?php else: ?>
                                    <form method="POST" id="reportForm">
                                        <input type="hidden" name="step" value="2">
                                        <input type="hidden" name="action" id="formAction" value="save">
                                        
                                        <?php foreach ($questions as $index => $question): ?>
                                            <?php 
                                            $is_answered = isset($existing_answers[$question['id']]) && !empty(trim($existing_answers[$question['id']]));
                                            ?>
                                            <div class="question-card <?= $is_answered ? 'answered' : 'unanswered' ?>">
                                                <div class="card-body">
                                                    <div class="d-flex align-items-start">
                                                        <div class="me-3">
                                                            <span class="badge bg-primary mt-1"><?= $index + 1 ?></span>
                                                            <?php if ($is_answered): ?>
                                                                <i class="bi bi-check-circle-fill text-success ms-1"></i>
                                                            <?php else: ?>
                                                                <i class="bi bi-exclamation-circle-fill text-warning ms-1"></i>
                                                            <?php endif; ?>
                                                        </div>
                                                        <div class="flex-grow-1">
                                                            <h6 class="question-text mb-3"><?= nl2br(htmlspecialchars($question['question_text'])) ?></h6>
                                                            <textarea class="form-control answer-textarea" 
                                                                      name="answers[<?= $question['id'] ?>]" 
                                                                      rows="3" 
                                                                      placeholder="Jumlah Target..."
                                                                      data-question-id="<?= $question['id'] ?>"><?= htmlspecialchars($existing_answers[$question['id']] ?? '') ?></textarea>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        <?php endforeach; ?>
                                        
                                        <div class="d-flex justify-content-between mt-4">
                                            <button type="button" class="btn btn-outline-secondary" onclick="history.back()">
                                                <i class="bi bi-arrow-left me-2"></i>
                                                Kembali
                                            </button>
                                            <div>
                                                <button type="button" class="btn btn-info me-2" onclick="saveAsDraft()">
                                                    <i class="bi bi-save me-2"></i>
                                                    Simpan Draft
                                                </button>
                                                <button type="button" class="btn btn-success" onclick="completeReport()">
                                                    <i class="bi bi-check-circle me-2"></i>
                                                    Selesaikan Laporan
                                                </button>
                                            </div>
                                        </div>
                                    </form>
                                <?php endif; ?>
                            </div>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Progress Info -->
                    <?php if (!empty($questions) && $show_questions): ?>
                        <div class="card border-0 shadow-sm mt-4">
                            <div class="card-body">
                                <div class="row text-center">
                                    <div class="col-md-3">
                                        <i class="bi bi-file-text fs-2 text-primary"></i>
                                        <h6 class="mt-2">Total Pertanyaan</h6>
                                        <p class="mb-0 fw-bold"><?= count($questions) ?></p>
                                    </div>
                                    <div class="col-md-3">
                                        <i class="bi bi-check-circle fs-2 text-success"></i>
                                        <h6 class="mt-2">Sudah Dijawab</h6>
                                        <p class="mb-0 fw-bold"><?= $answered_count ?></p>
                                    </div>
                                    <div class="col-md-3">
                                        <i class="bi bi-clock fs-2 text-warning"></i>
                                        <h6 class="mt-2">Estimasi Waktu</h6>
                                        <p class="mb-0 fw-bold"><?= ceil(count($questions) * 0.5) ?> menit</p>
                                    </div>
                                    <div class="col-md-3">
                                        <i class="bi bi-info-circle fs-2 text-info"></i>
                                        <h6 class="mt-2">Progress</h6>
                                        <p class="mb-0 fw-bold"><?= $progress_percentage ?>%</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function saveAsDraft() {
            document.getElementById('formAction').value = 'save';
            document.getElementById('reportForm').submit();
        }
        
        function completeReport() {
            const textareas = document.querySelectorAll('.answer-textarea');
            let unanswered = 0;
            let unansweredQuestions = [];
            
            textareas.forEach((textarea, index) => {
                if (!textarea.value.trim()) {
                    unanswered++;
                    unansweredQuestions.push(index + 1);
                }
            });
            
            if (unanswered > 0) {
                if (confirm(`Masih ada ${unanswered} pertanyaan yang belum dijawab (No: ${unansweredQuestions.join(', ')}). Apakah Anda yakin ingin menyelesaikan laporan? Pertanyaan yang belum dijawab tidak akan disimpan.`)) {
                    document.getElementById('formAction').value = 'complete';
                    document.getElementById('reportForm').submit();
                }
            } else {
                document.getElementById('formAction').value = 'complete';
                document.getElementById('reportForm').submit();
            }
        }
        
        // Real-time visual feedback
        document.querySelectorAll('.answer-textarea').forEach(textarea => {
            textarea.addEventListener('input', function() {
                const questionCard = this.closest('.question-card');
                const hasContent = this.value.trim().length > 0;
                
                if (hasContent) {
                    questionCard.classList.remove('unanswered');
                    questionCard.classList.add('answered');
                    questionCard.querySelector('.bi-exclamation-circle-fill')?.classList.replace('bi-exclamation-circle-fill', 'bi-check-circle-fill');
                    questionCard.querySelector('.text-warning')?.classList.replace('text-warning', 'text-success');
                } else {
                    questionCard.classList.remove('answered');
                    questionCard.classList.add('unanswered');
                    questionCard.querySelector('.bi-check-circle-fill')?.classList.replace('bi-check-circle-fill', 'bi-exclamation-circle-fill');
                    questionCard.querySelector('.text-success')?.classList.replace('text-success', 'text-warning');
                }
                
                // Update progress
                updateProgress();
            });
        });
        
        function updateProgress() {
            const textareas = document.querySelectorAll('.answer-textarea');
            let answered = 0;
            
            textareas.forEach(textarea => {
                if (textarea.value.trim()) answered++;
            });
            
            const total = textareas.length;
            const percentage = Math.round((answered / total) * 100);
            
            // Update progress info if exists
            const progressCard = document.querySelector('.progress-card');
            if (progressCard) {
                const progressBar = progressCard.querySelector('.progress-bar');
                const progressText = progressCard.querySelector('span');
                
                if (progressBar) {
                    progressBar.style.width = percentage + '%';
                    progressBar.setAttribute('aria-valuenow', percentage);
                }
                
                if (progressText) {
                    progressText.textContent = answered + '/' + total + ' pertanyaan';
                }
            }
        }
        
        // Auto-save draft every 2 minutes
        setInterval(function() {
            if (document.getElementById('reportForm')) {
                saveAsDraft();
            }
        }, 120000);
        
        // Confirm before leaving page if there are unsaved changes
        window.addEventListener('beforeunload', function(e) {
            const textareas = document.querySelectorAll('.answer-textarea');
            let hasChanges = false;
            
            textareas.forEach(textarea => {
                if (textarea.value.trim()) {
                    hasChanges = true;
                }
            });
            
            if (hasChanges) {
                e.preventDefault();
                e.returnValue = '';
                return '';
            }
        });
        
        // Remove beforeunload when form is submitted
        document.querySelector('#reportForm')?.addEventListener('submit', function() {
            window.removeEventListener('beforeunload', function(){});
        });
    </script>
</body>
</html>