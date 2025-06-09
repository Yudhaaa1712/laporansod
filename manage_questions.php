<?php
require_once 'config.php';
requireAdmin();

$message = '';
$message_type = '';

// Handle form submissions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_POST['action'])) {
        if ($_POST['action'] === 'create') {
            $unit_id = $_POST['unit_id'];
            $question_text = sanitize($_POST['question_text']);
            $question_order = $_POST['question_order'] ?: 0;
            
            if (!empty($unit_id) && !empty($question_text)) {
                $stmt = $pdo->prepare("INSERT INTO questions (unit_id, question_text, question_order, created_by) VALUES (?, ?, ?, ?)");
                
                if ($stmt->execute([$unit_id, $question_text, $question_order, $_SESSION['user_id']])) {
                    $message = 'Pertanyaan berhasil ditambahkan!';
                    $message_type = 'success';
                } else {
                    $message = 'Gagal menambahkan pertanyaan!';
                    $message_type = 'danger';
                }
            } else {
                $message = 'Harap isi semua field!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'edit') {
            $question_id = $_POST['question_id'];
            $unit_id = $_POST['unit_id'];
            $question_text = sanitize($_POST['question_text']);
            $question_order = $_POST['question_order'] ?: 0;
            $is_active = isset($_POST['is_active']) ? 1 : 0;
            
            $stmt = $pdo->prepare("UPDATE questions SET unit_id = ?, question_text = ?, question_order = ?, is_active = ? WHERE id = ?");
            
            if ($stmt->execute([$unit_id, $question_text, $question_order, $is_active, $question_id])) {
                $message = 'Pertanyaan berhasil diupdate!';
                $message_type = 'success';
            } else {
                $message = 'Gagal mengupdate pertanyaan!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'delete') {
            $question_id = $_POST['question_id'];
            $stmt = $pdo->prepare("UPDATE questions SET deleted_at = NOW() WHERE id = ?");
            
            if ($stmt->execute([$question_id])) {
                $message = 'Pertanyaan berhasil dihapus!';
                $message_type = 'success';
            } else {
                $message = 'Gagal menghapus pertanyaan!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'toggle_status') {
            $question_id = $_POST['question_id'];
            $stmt = $pdo->prepare("UPDATE questions SET is_active = NOT is_active WHERE id = ?");
            
            if ($stmt->execute([$question_id])) {
                $message = 'Status pertanyaan berhasil diubah!';
                $message_type = 'success';
            } else {
                $message = 'Gagal mengubah status pertanyaan!';
                $message_type = 'danger';
            }
        }
    }
}

// Get filter
$filter_unit = $_GET['unit'] ?? '';

// Get all units for dropdown
$stmt = $pdo->query("SELECT id, name FROM units WHERE deleted_at IS NULL ORDER BY name");
$units = $stmt->fetchAll();

// Get questions with filter
$sql = "
    SELECT q.*, u.name as unit_name, creator.name as creator_name
    FROM questions q 
    LEFT JOIN units u ON q.unit_id = u.id 
    LEFT JOIN users creator ON q.created_by = creator.id
    WHERE q.deleted_at IS NULL
";
$params = [];

if ($filter_unit) {
    $sql .= " AND q.unit_id = ?";
    $params[] = $filter_unit;
}

$sql .= " ORDER BY u.name, q.question_order, q.id";

$stmt = $pdo->prepare($sql);
$stmt->execute($params);
$questions = $stmt->fetchAll();

// Group questions by unit
$questions_by_unit = [];
foreach ($questions as $question) {
    $unit_name = $question['unit_name'] ?: 'Unit Tidak Diketahui';
    $questions_by_unit[$unit_name][] = $question;
}
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kelola Pertanyaan - Sistem Laporan RS</title>
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
        .question-card {
            border-left: 4px solid #007bff;
            transition: all 0.3s ease;
        }
        .question-card:hover {
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .question-card.inactive {
            border-left-color: #6c757d;
            opacity: 0.7;
        }
        .unit-section {
            background: white;
            border-radius: 15px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            margin-bottom: 25px;
        }
        .unit-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 15px 15px 0 0;
            padding: 15px 20px;
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
                        <a href="#" class="nav-link active">
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
                        <h2 class="mb-0">Kelola Target Laporan</h2>
                        <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#createModal">
                            <i class="bi bi-plus-circle me-2"></i>Tambah Pertanyaan
                        </button>
                    </div>
                    
                    <?php if ($message): ?>
                        <div class="alert alert-<?= $message_type ?> alert-dismissible fade show" role="alert">
                            <?= $message ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Filter -->
                    <div class="card border-0 shadow-sm mb-4">
                        <div class="card-body">
                            <form method="GET" class="row g-3">
                                <div class="col-md-4">
                                    <label class="form-label">Filter Unit</label>
                                    <select class="form-select" name="unit" onchange="this.form.submit()">
                                        <option value="">Semua Unit</option>
                                        <?php foreach ($units as $unit): ?>
                                            <option value="<?= $unit['id'] ?>" <?= $filter_unit == $unit['id'] ? 'selected' : '' ?>>
                                                <?= $unit['name'] ?>
                                            </option>
                                        <?php endforeach; ?>
                                    </select>
                                </div>
                                <div class="col-md-8 d-flex align-items-end">
                                    <a href="?" class="btn btn-outline-secondary">
                                        <i class="bi bi-arrow-clockwise me-2"></i>Reset Filter
                                    </a>
                                </div>
                            </form>
                        </div>
                    </div>
                    
                    <!-- Questions by Unit -->
                    <?php if (empty($questions_by_unit)): ?>
                        <div class="text-center py-5">
                            <i class="bi bi-question-circle fs-1 text-muted"></i>
                            <h5 class="mt-3 text-muted">Belum Ada Pertanyaan</h5>
                            <p class="text-muted">Silakan tambahkan pertanyaan pertama untuk unit.</p>
                            <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#createModal">
                                <i class="bi bi-plus-circle me-2"></i>Tambah Pertanyaan
                            </button>
                        </div>
                    <?php else: ?>
                        <?php foreach ($questions_by_unit as $unit_name => $unit_questions): ?>
                            <div class="unit-section">
                                <div class="unit-header">
                                    <div class="d-flex justify-content-between align-items-center">
                                        <h5 class="mb-0">
                                            <i class="bi bi-building me-2"></i>
                                            <?= $unit_name ?>
                                        </h5>
                                        <span class="badge bg-light text-dark"><?= count($unit_questions) ?> pertanyaan</span>
                                    </div>
                                </div>
                                <div class="p-3">
                                    <?php foreach ($unit_questions as $index => $question): ?>
                                        <div class="question-card card mb-3 <?= !$question['is_active'] ? 'inactive' : '' ?>">
                                            <div class="card-body">
                                                <div class="d-flex justify-content-between align-items-start">
                                                    <div class="flex-grow-1">
                                                        <div class="d-flex align-items-center mb-2">
                                                            <span class="badge bg-primary me-2"><?= $question['question_order'] ?: ($index + 1) ?></span>
                                                            <?php if (!$question['is_active']): ?>
                                                                <span class="badge bg-secondary me-2">Nonaktif</span>
                                                            <?php endif; ?>
                                                        </div>
                                                        <p class="card-text mb-2"><?= nl2br(htmlspecialchars($question['question_text'])) ?></p>
                                                        <small class="text-muted">
                                                            Dibuat oleh: <?= $question['creator_name'] ?> | 
                                                            <?= formatDateTime($question['created_at']) ?>
                                                        </small>
                                                    </div>
                                                    <div class="btn-group" role="group">
                                                        <button class="btn btn-sm btn-outline-primary" 
                                                                onclick="editQuestion(<?= htmlspecialchars(json_encode($question)) ?>)">
                                                            <i class="bi bi-pencil"></i>
                                                        </button>
                                                        <button class="btn btn-sm btn-outline-<?= $question['is_active'] ? 'warning' : 'success' ?>" 
                                                                onclick="toggleStatus(<?= $question['id'] ?>, '<?= $question['is_active'] ? 'nonaktifkan' : 'aktifkan' ?>')">
                                                            <i class="bi bi-<?= $question['is_active'] ? 'pause' : 'play' ?>"></i>
                                                        </button>
                                                        <button class="btn btn-sm btn-outline-danger" 
                                                                onclick="deleteQuestion(<?= $question['id'] ?>, '<?= addslashes(substr($question['question_text'], 0, 50)) ?>...')">
                                                            <i class="bi bi-trash"></i>
                                                        </button>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    <?php endforeach; ?>
                                </div>
                            </div>
                        <?php endforeach; ?>
                   <?php endif; ?>
                </div>
            </div>
        </div>
    </div>

    <!-- Create Modal -->
    <div class="modal fade" id="createModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Tambah Pertanyaan Baru</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST">
                    <div class="modal-body">
                        <input type="hidden" name="action" value="create">
                        
                        <div class="mb-3">
                            <label class="form-label">Unit *</label>
                            <select class="form-select" name="unit_id" required>
                                <option value="">Pilih Unit</option>
                                <?php foreach ($units as $unit): ?>
                                    <option value="<?= $unit['id'] ?>"><?= $unit['name'] ?></option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Pertanyaan *</label>
                            <textarea class="form-control" name="question_text" rows="4" 
                                      placeholder="Masukkan teks pertanyaan..." required></textarea>
                            <div class="form-text">Gunakan Enter untuk membuat baris baru jika diperlukan.</div>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Urutan Pertanyaan</label>
                            <input type="number" class="form-control" name="question_order" 
                                   min="0" placeholder="0 = otomatis di akhir">
                            <div class="form-text">Nomor urutan untuk mengurutkan pertanyaan dalam unit.</div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                        <button type="submit" class="btn btn-primary">Simpan Pertanyaan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Edit Modal -->
    <div class="modal fade" id="editModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Edit Pertanyaan</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST">
                    <div class="modal-body">
                        <input type="hidden" name="action" value="edit">
                        <input type="hidden" name="question_id" id="edit_question_id">
                        
                        <div class="mb-3">
                            <label class="form-label">Unit *</label>
                            <select class="form-select" name="unit_id" id="edit_unit_id" required>
                                <option value="">Pilih Unit</option>
                                <?php foreach ($units as $unit): ?>
                                    <option value="<?= $unit['id'] ?>"><?= $unit['name'] ?></option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Pertanyaan *</label>
                            <textarea class="form-control" name="question_text" id="edit_question_text" 
                                      rows="4" required></textarea>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Urutan Pertanyaan</label>
                            <input type="number" class="form-control" name="question_order" 
                                   id="edit_question_order" min="0">
                        </div>
                        
                        <div class="mb-3">
                            <div class="form-check">
                                <input class="form-check-input" type="checkbox" name="is_active" 
                                       id="edit_is_active" checked>
                                <label class="form-check-label" for="edit_is_active">
                                    Aktif (Tampil dalam laporan)
                                </label>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                        <button type="submit" class="btn btn-primary">Update Pertanyaan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Delete Form -->
    <form method="POST" id="deleteForm" style="display: none;">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="question_id" id="delete_question_id">
    </form>

    <!-- Toggle Status Form -->
    <form method="POST" id="toggleForm" style="display: none;">
        <input type="hidden" name="action" value="toggle_status">
        <input type="hidden" name="question_id" id="toggle_question_id">
    </form>

    <!-- Statistics Card -->
    <div class="position-fixed bottom-0 end-0 p-3" style="z-index: 1050;">
        <div class="card border-0 shadow-lg" style="width: 250px;">
            <div class="card-body text-center">
                <h6 class="card-title">
                    <i class="bi bi-bar-chart me-2"></i>
                    Statistik Pertanyaan
                </h6>
                <div class="row text-center">
                    <div class="col-6">
                        <h5 class="text-primary mb-0"><?= count($questions) ?></h5>
                        <small class="text-muted">Total</small>
                    </div>
                    <div class="col-6">
                        <h5 class="text-success mb-0"><?= count(array_filter($questions, fn($q) => $q['is_active'])) ?></h5>
                        <small class="text-muted">Aktif</small>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function editQuestion(question) {
            document.getElementById('edit_question_id').value = question.id;
            document.getElementById('edit_unit_id').value = question.unit_id;
            document.getElementById('edit_question_text').value = question.question_text;
            document.getElementById('edit_question_order').value = question.question_order;
            document.getElementById('edit_is_active').checked = question.is_active == 1;
            
            var editModal = new bootstrap.Modal(document.getElementById('editModal'));
            editModal.show();
        }
        
        function deleteQuestion(id, text) {
            if (confirm('Apakah Anda yakin ingin menghapus pertanyaan:\n\n"' + text + '"?\n\nTindakan ini tidak dapat dibatalkan.')) {
                document.getElementById('delete_question_id').value = id;
                document.getElementById('deleteForm').submit();
            }
        }
        
        function toggleStatus(id, action) {
            if (confirm('Apakah Anda yakin ingin ' + action + ' pertanyaan ini?')) {
                document.getElementById('toggle_question_id').value = id;
                document.getElementById('toggleForm').submit();
            }
        }
        
        // Auto-resize textarea
        document.querySelectorAll('textarea').forEach(textarea => {
            textarea.addEventListener('input', function() {
                this.style.height = 'auto';
                this.style.height = this.scrollHeight + 'px';
            });
        });
        
        // Character counter for question text
        function addCharacterCounter(textareaId, counterId) {
            const textarea = document.getElementById(textareaId);
            const counter = document.getElementById(counterId);
            
            if (textarea && counter) {
                textarea.addEventListener('input', function() {
                    const length = this.value.length;
                    counter.textContent = length + ' karakter';
                    
                    if (length > 500) {
                        counter.classList.add('text-warning');
                    } else {
                        counter.classList.remove('text-warning');
                    }
                });
            }
        }
        
        // Drag and drop reordering (future enhancement)
        // You can implement drag-and-drop functionality here
        
        // Bulk operations
        function selectAll() {
            const checkboxes = document.querySelectorAll('.question-checkbox');
            checkboxes.forEach(cb => cb.checked = true);
            updateBulkActions();
        }
        
        function selectNone() {
            const checkboxes = document.querySelectorAll('.question-checkbox');
            checkboxes.forEach(cb => cb.checked = false);
            updateBulkActions();
        }
        
        function updateBulkActions() {
            const checked = document.querySelectorAll('.question-checkbox:checked');
            const bulkActions = document.getElementById('bulkActions');
            
            if (bulkActions) {
                bulkActions.style.display = checked.length > 0 ? 'block' : 'none';
            }
        }
        
        // Search functionality
        function searchQuestions() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const questionCards = document.querySelectorAll('.question-card');
            
            questionCards.forEach(card => {
                const text = card.textContent.toLowerCase();
                const unitSection = card.closest('.unit-section');
                
                if (text.includes(searchTerm)) {
                    card.style.display = 'block';
                } else {
                    card.style.display = 'none';
                }
            });
            
            // Hide empty unit sections
            document.querySelectorAll('.unit-section').forEach(section => {
                const visibleCards = section.querySelectorAll('.question-card[style*="block"], .question-card:not([style*="none"])');
                section.style.display = visibleCards.length > 0 ? 'block' : 'none';
            });
        }
        
        // Preview question formatting
        function previewQuestion() {
            const text = document.querySelector('#createModal textarea[name="question_text"]').value;
            const preview = document.getElementById('questionPreview');
            
            if (preview) {
                preview.innerHTML = text.replace(/\n/g, '<br>');
            }
        }
        
        // Auto-save draft
        let draftTimer;
        function saveDraft() {
            clearTimeout(draftTimer);
            draftTimer = setTimeout(() => {
                const formData = new FormData(document.querySelector('#createModal form'));
                const draft = {};
                
                for (let [key, value] of formData.entries()) {
                    draft[key] = value;
                }
                
                localStorage.setItem('question_draft', JSON.stringify(draft));
            }, 1000);
        }
        
        // Load draft
        function loadDraft() {
            const draft = localStorage.getItem('question_draft');
            if (draft) {
                const data = JSON.parse(draft);
                
                Object.keys(data).forEach(key => {
                    const field = document.querySelector(`#createModal [name="${key}"]`);
                    if (field) field.value = data[key];
                });
            }
        }
        
        // Initialize
        document.addEventListener('DOMContentLoaded', function() {
            // Load draft when create modal is opened
            document.getElementById('createModal').addEventListener('show.bs.modal', loadDraft);
            
            // Clear draft when form is submitted successfully
            if (document.querySelector('.alert-success')) {
                localStorage.removeItem('question_draft');
            }
            
            // Add input listeners for draft saving
            document.querySelectorAll('#createModal input, #createModal textarea, #createModal select').forEach(field => {
                field.addEventListener('input', saveDraft);
            });
        });
    </script>
</body>
</html>