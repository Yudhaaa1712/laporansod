<?php
require_once 'config.php';
requireAdmin();

$message = '';
$message_type = '';

// Handle form submissions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_POST['action'])) {
        if ($_POST['action'] === 'create') {
            $name = sanitize($_POST['name']);
            $description = sanitize($_POST['description']);
            
            if (!empty($name)) {
                // Check if unit name already exists
                $stmt = $pdo->prepare("SELECT id FROM units WHERE name = ? AND deleted_at IS NULL");
                $stmt->execute([$name]);
                
                if ($stmt->rowCount() == 0) {
                    $stmt = $pdo->prepare("INSERT INTO units (name, description) VALUES (?, ?)");
                    
                    if ($stmt->execute([$name, $description])) {
                        $message = 'Unit berhasil ditambahkan!';
                        $message_type = 'success';
                    } else {
                        $message = 'Gagal menambahkan unit!';
                        $message_type = 'danger';
                    }
                } else {
                    $message = 'Nama unit sudah ada!';
                    $message_type = 'danger';
                }
            } else {
                $message = 'Nama unit wajib diisi!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'edit') {
            $unit_id = $_POST['unit_id'];
            $name = sanitize($_POST['name']);
            $description = sanitize($_POST['description']);
            
            if (!empty($name)) {
                // Check if name already exists (except current unit)
                $stmt = $pdo->prepare("SELECT id FROM units WHERE name = ? AND id != ? AND deleted_at IS NULL");
                $stmt->execute([$name, $unit_id]);
                
                if ($stmt->rowCount() == 0) {
                    $stmt = $pdo->prepare("UPDATE units SET name = ?, description = ? WHERE id = ?");
                    
                    if ($stmt->execute([$name, $description, $unit_id])) {
                        $message = 'Unit berhasil diupdate!';
                        $message_type = 'success';
                    } else {
                        $message = 'Gagal mengupdate unit!';
                        $message_type = 'danger';
                    }
                } else {
                    $message = 'Nama unit sudah ada!';
                    $message_type = 'danger';
                }
            } else {
                $message = 'Nama unit wajib diisi!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'delete') {
            $unit_id = $_POST['unit_id'];
            
            // Check if unit has users
            $stmt = $pdo->prepare("SELECT COUNT(*) as user_count FROM users WHERE unit_id = ? AND deleted_at IS NULL");
            $stmt->execute([$unit_id]);
            $user_count = $stmt->fetch()['user_count'];
            
            if ($user_count > 0) {
                $message = "Tidak dapat menghapus unit karena masih ada $user_count user yang terkait!";
                $message_type = 'warning';
            } else {
                $stmt = $pdo->prepare("UPDATE units SET deleted_at = NOW() WHERE id = ?");
                
                if ($stmt->execute([$unit_id])) {
                    $message = 'Unit berhasil dihapus!';
                    $message_type = 'success';
                } else {
                    $message = 'Gagal menghapus unit!';
                    $message_type = 'danger';
                }
            }
        }
    }
}

// Get all units with statistics
$stmt = $pdo->query("
    SELECT u.*, 
           COUNT(DISTINCT users.id) as user_count,
           COUNT(DISTINCT q.id) as question_count,
           COUNT(DISTINCT rs.id) as report_count
    FROM units u
    LEFT JOIN users ON u.id = users.unit_id AND users.deleted_at IS NULL
    LEFT JOIN questions q ON u.id = q.unit_id AND q.deleted_at IS NULL
    LEFT JOIN report_sessions rs ON u.id = rs.unit_id
    WHERE u.deleted_at IS NULL
    GROUP BY u.id
    ORDER BY u.created_at DESC
");
$units = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Master Unit - Sistem Laporan RS</title>
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
        .unit-card {
            border-radius: 15px;
            border: none;
            transition: all 0.3s ease;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .unit-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }
        .stats-badge {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            border-radius: 20px;
            padding: 5px 15px;
            font-size: 0.85em;
            font-weight: 600;
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
                        <a href="#" class="nav-link active">
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
                        <h2 class="mb-0">Master Unit</h2>
                        <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#createModal">
                            <i class="bi bi-plus-circle me-2"></i>Tambah Unit
                        </button>
                    </div>
                    
                    <?php if ($message): ?>
                        <div class="alert alert-<?= $message_type ?> alert-dismissible fade show" role="alert">
                            <?= $message ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Units Grid -->
                    <?php if (empty($units)): ?>
                        <div class="text-center py-5">
                            <i class="bi bi-building fs-1 text-muted"></i>
                            <h5 class="mt-3 text-muted">Belum Ada Unit</h5>
                            <p class="text-muted">Silakan tambahkan unit pertama untuk rumah sakit.</p>
                            <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#createModal">
                                <i class="bi bi-plus-circle me-2"></i>Tambah Unit Pertama
                            </button>
                        </div>
                    <?php else: ?>
                        <div class="row">
                            <?php foreach ($units as $unit): ?>
                                <div class="col-md-6 col-lg-4 mb-4">
                                    <div class="card unit-card h-100">
                                        <div class="card-body">
                                            <div class="d-flex justify-content-between align-items-start mb-3">
                                                <h5 class="card-title mb-0">
                                                    <i class="bi bi-building text-primary me-2"></i>
                                                    <?= $unit['name'] ?>
                                                </h5>
                                                <div class="dropdown">
                                                    <button class="btn btn-sm btn-outline-secondary" data-bs-toggle="dropdown">
                                                        <i class="bi bi-three-dots-vertical"></i>
                                                    </button>
                                                    <ul class="dropdown-menu">
                                                        <li>
                                                            <a class="dropdown-item" href="#" 
                                                               onclick="editUnit(<?= htmlspecialchars(json_encode($unit)) ?>)">
                                                                <i class="bi bi-pencil me-2"></i>Edit
                                                            </a>
                                                        </li>
                                                        <li><hr class="dropdown-divider"></li>
                                                        <li>
                                                            <a class="dropdown-item text-danger" href="#" 
                                                               onclick="deleteUnit(<?= $unit['id'] ?>, '<?= $unit['name'] ?>', <?= $unit['user_count'] ?>)">
                                                                <i class="bi bi-trash me-2"></i>Hapus
                                                            </a>
                                                        </li>
                                                    </ul>
                                                </div>
                                            </div>
                                            
                                            <?php if ($unit['description']): ?>
                                                <p class="card-text text-muted mb-3"><?= nl2br(htmlspecialchars($unit['description'])) ?></p>
                                            <?php endif; ?>
                                            
                                            <div class="row text-center mb-3">
                                                <div class="col-4">
                                                    <div class="stats-badge">
                                                        <div class="fw-bold"><?= $unit['user_count'] ?></div>
                                                        <small>User</small>
                                                    </div>
                                                </div>
                                                <div class="col-4">
                                                    <div class="stats-badge">
                                                        <div class="fw-bold"><?= $unit['question_count'] ?></div>
                                                        <small>Pertanyaan</small>
                                                    </div>
                                                </div>
                                                <div class="col-4">
                                                    <div class="stats-badge">
                                                        <div class="fw-bold"><?= $unit['report_count'] ?></div>
                                                        <small>Laporan</small>
                                                    </div>
                                                </div>
                                            </div>
                                            
                                            <small class="text-muted">
                                                <i class="bi bi-calendar me-1"></i>
                                                Dibuat: <?= formatDateTime($unit['created_at']) ?>
                                            </small>
                                        </div>
                                    </div>
                                </div>
                            <?php endforeach; ?>
                        </div>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </div>

  

    <!-- Delete Form -->
    <form method="POST" id="deleteForm" style="display: none;">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="unit_id" id="delete_unit_id">
    </form>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function editUnit(unit) {
            document.getElementById('edit_unit_id').value = unit.id;
            document.getElementById('edit_name').value = unit.name;
            document.getElementById('edit_description').value = unit.description || '';
            
            var editModal = new bootstrap.Modal(document.getElementById('editModal'));
            editModal.show();
        }
        
        function deleteUnit(id, name, userCount) {
            if (userCount > 0) {
                alert('Tidak dapat menghapus unit "' + name + '" karena masih ada ' + userCount + ' user yang terkait!');
                return;
            }
            
            if (confirm('Apakah Anda yakin ingin menghapus unit "' + name + '"?\n\nTindakan ini tidak dapat dibatalkan.')) {
                document.getElementById('delete_unit_id').value = id;
                document.getElementById('deleteForm').submit();
            }
        }
    </script>
</body>
</html>