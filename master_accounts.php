<?php
require_once 'config.php';
requireAdmin();

$message = '';
$message_type = '';

// Handle form submissions
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_POST['action'])) {
        if ($_POST['action'] === 'create') {
            $username = sanitize($_POST['username']);
            $password = $_POST['password'];
            $name = sanitize($_POST['name']);
            $email = sanitize($_POST['email']);
            $unit_id = $_POST['unit_id'];
            
            if (!empty($username) && !empty($password) && !empty($name) && !empty($email) && !empty($unit_id)) {
                // Check if username or email already exists
                $stmt = $pdo->prepare("SELECT id FROM users WHERE (username = ? OR email = ?) AND deleted_at IS NULL");
                $stmt->execute([$username, $email]);
                
                if ($stmt->rowCount() == 0) {
                    $hashed_password = password_hash($password, PASSWORD_DEFAULT);
                    $stmt = $pdo->prepare("INSERT INTO users (username, password, name, email, role, unit_id) VALUES (?, ?, ?, ?, 'unit', ?)");
                    
                    if ($stmt->execute([$username, $hashed_password, $name, $email, $unit_id])) {
                        $message = 'Akun berhasil dibuat!';
                        $message_type = 'success';
                    } else {
                        $message = 'Gagal membuat akun!';
                        $message_type = 'danger';
                    }
                } else {
                    $message = 'Username atau email sudah digunakan!';
                    $message_type = 'danger';
                }
            } else {
                $message = 'Harap isi semua field!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'delete') {
            $user_id = $_POST['user_id'];
            $stmt = $pdo->prepare("UPDATE users SET deleted_at = NOW() WHERE id = ? AND role = 'unit'");
            
            if ($stmt->execute([$user_id])) {
                $message = 'Akun berhasil dihapus!';
                $message_type = 'success';
            } else {
                $message = 'Gagal menghapus akun!';
                $message_type = 'danger';
            }
        }
        
        if ($_POST['action'] === 'edit') {
            $user_id = $_POST['user_id'];
            $username = sanitize($_POST['username']);
            $name = sanitize($_POST['name']);
            $email = sanitize($_POST['email']);
            $unit_id = $_POST['unit_id'];
            
            $sql = "UPDATE users SET username = ?, name = ?, email = ?, unit_id = ?";
            $params = [$username, $name, $email, $unit_id];
            
            // If password is provided, update it too
            if (!empty($_POST['password'])) {
                $hashed_password = password_hash($_POST['password'], PASSWORD_DEFAULT);
                $sql .= ", password = ?";
                $params[] = $hashed_password;
            }
            
            $sql .= " WHERE id = ? AND role = 'unit'";
            $params[] = $user_id;
            
            $stmt = $pdo->prepare($sql);
            
            if ($stmt->execute($params)) {
                $message = 'Akun berhasil diupdate!';
                $message_type = 'success';
            } else {
                $message = 'Gagal mengupdate akun!';
                $message_type = 'danger';
            }
        }
    }
}

// Get all units for dropdown
$stmt = $pdo->query("SELECT id, name FROM units WHERE deleted_at IS NULL ORDER BY name");
$units = $stmt->fetchAll();

// Get all unit users
$stmt = $pdo->query("
    SELECT u.*, un.name as unit_name 
    FROM users u 
    LEFT JOIN units un ON u.unit_id = un.id 
    WHERE u.role = 'unit' AND u.deleted_at IS NULL 
    ORDER BY u.created_at DESC
");
$users = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Master Akun - Sistem Laporan RS</title>
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
                        <a href="#" class="nav-link active">
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
                        <h2 class="mb-0">Master Akun Unit</h2>
                        <button class="btn btn-primary" data-bs-toggle="modal" data-bs-target="#createModal">
                            <i class="bi bi-plus-circle me-2"></i>Tambah Akun
                        </button>
                    </div>
                    
                    <?php if ($message): ?>
                        <div class="alert alert-<?= $message_type ?> alert-dismissible fade show" role="alert">
                            <?= $message ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Users Table -->
                    <div class="card border-0 shadow-sm">
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="table table-hover">
                                    <thead class="table-light">
                                        <tr>
                                            <th>#</th>
                                            <th>Username</th>
                                            <th>Nama</th>
                                            <th>Email</th>
                                            <th>Unit</th>
                                            <th>Dibuat</th>
                                            <th>Aksi</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <?php if (empty($users)): ?>
                                            <tr>
                                                <td colspan="7" class="text-center py-4">
                                                    <i class="bi bi-inbox fs-1 text-muted"></i>
                                                    <p class="text-muted mt-2">Belum ada akun unit</p>
                                                </td>
                                            </tr>
                                        <?php else: ?>
                                            <?php foreach ($users as $index => $user): ?>
                                                <tr>
                                                    <td><?= $index + 1 ?></td>
                                                    <td>
                                                        <strong><?= $user['username'] ?></strong>
                                                    </td>
                                                    <td><?= $user['name'] ?></td>
                                                    <td><?= $user['email'] ?></td>
                                                    <td>
                                                        <span class="badge bg-primary"><?= $user['unit_name'] ?: 'Tidak ada unit' ?></span>
                                                    </td>
                                                    <td><?= formatDateTime($user['created_at']) ?></td>
                                                    <td>
                                                        <button class="btn btn-sm btn-outline-primary me-1" 
                                                                onclick="editUser(<?= htmlspecialchars(json_encode($user)) ?>)">
                                                            <i class="bi bi-pencil"></i>
                                                        </button>
                                                        <button class="btn btn-sm btn-outline-danger" 
                                                                onclick="deleteUser(<?= $user['id'] ?>, '<?= $user['name'] ?>')">
                                                            <i class="bi bi-trash"></i>
                                                        </button>
                                                    </td>
                                                </tr>
                                            <?php endforeach; ?>
                                        <?php endif; ?>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Create Modal -->
    <div class="modal fade" id="createModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Tambah Akun Unit Baru</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST">
                    <div class="modal-body">
                        <input type="hidden" name="action" value="create">
                        
                        <div class="mb-3">
                            <label class="form-label">Username</label>
                            <input type="text" class="form-control" name="username" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Password</label>
                            <input type="password" class="form-control" name="password" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Nama Lengkap</label>
                            <input type="text" class="form-control" name="name" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Email</label>
                            <input type="email" class="form-control" name="email" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Unit</label>
                            <select class="form-select" name="unit_id" required>
                                <option value="">Pilih Unit</option>
                                <?php foreach ($units as $unit): ?>
                                    <option value="<?= $unit['id'] ?>"><?= $unit['name'] ?></option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                        <button type="submit" class="btn btn-primary">Simpan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Edit Modal -->
    <div class="modal fade" id="editModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Edit Akun Unit</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <form method="POST">
                    <div class="modal-body">
                        <input type="hidden" name="action" value="edit">
                        <input type="hidden" name="user_id" id="edit_user_id">
                        
                        <div class="mb-3">
                            <label class="form-label">Username</label>
                            <input type="text" class="form-control" name="username" id="edit_username" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Password Baru (kosongkan jika tidak ingin mengubah)</label>
                            <input type="password" class="form-control" name="password">
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Nama Lengkap</label>
                            <input type="text" class="form-control" name="name" id="edit_name" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Email</label>
                            <input type="email" class="form-control" name="email" id="edit_email" required>
                        </div>
                        
                        <div class="mb-3">
                            <label class="form-label">Unit</label>
                            <select class="form-select" name="unit_id" id="edit_unit_id" required>
                                <option value="">Pilih Unit</option>
                                <?php foreach ($units as $unit): ?>
                                    <option value="<?= $unit['id'] ?>"><?= $unit['name'] ?></option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                        <button type="submit" class="btn btn-primary">Update</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Delete Form -->
    <form method="POST" id="deleteForm" style="display: none;">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="user_id" id="delete_user_id">
    </form>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function editUser(user) {
            document.getElementById('edit_user_id').value = user.id;
            document.getElementById('edit_username').value = user.username;
            document.getElementById('edit_name').value = user.name;
            document.getElementById('edit_email').value = user.email;
            document.getElementById('edit_unit_id').value = user.unit_id;
            
            var editModal = new bootstrap.Modal(document.getElementById('editModal'));
            editModal.show();
        }
        
        function deleteUser(id, name) {
            if (confirm('Apakah Anda yakin ingin menghapus akun "' + name + '"?')) {
                document.getElementById('delete_user_id').value = id;
                document.getElementById('deleteForm').submit();
            }
        }
    </script>
</body>
</html>