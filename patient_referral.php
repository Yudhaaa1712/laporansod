<?php
require_once 'config.php';
requireUnit();

$message = '';
$message_type = '';

// Get unit info
$stmt = $pdo->prepare("SELECT name FROM units WHERE id = ?");
$stmt->execute([$_SESSION['unit_id']]);
$unit = $stmt->fetch();

// Handle form submission
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_POST['action']) && $_POST['action'] === 'add_referral') {
        $reporter_name = sanitize($_POST['reporter_name']);
        $report_date = $_POST['report_date'];
        $shift = $_POST['shift'];
        $patient_name = sanitize($_POST['patient_name']);
        $patient_age = sanitize($_POST['patient_age']);
        $medical_record = sanitize($_POST['medical_record']);
        $origin_room = sanitize($_POST['origin_room']);
        $insurance_type = sanitize($_POST['insurance_type']);
        $diagnosis = sanitize($_POST['diagnosis']);
        $destination_hospital = sanitize($_POST['destination_hospital']);
        $referral_reason = sanitize($_POST['referral_reason']);
        
        if (!empty($reporter_name) && !empty($report_date) && !empty($shift) && 
            !empty($patient_name) && !empty($patient_age) && !empty($medical_record) && 
            !empty($origin_room) && !empty($insurance_type) && !empty($diagnosis) && 
            !empty($destination_hospital)) {
            
            try {
                $pdo->beginTransaction();
                
                // Create or get report session
                $stmt = $pdo->prepare("
                    SELECT id FROM report_sessions 
                    WHERE unit_id = ? AND report_date = ? AND reporter_name = ? AND shift = ?
                ");
                $stmt->execute([$_SESSION['unit_id'], $report_date, $reporter_name, $shift]);
                $session = $stmt->fetch();
                
                if (!$session) {
                    // Create new session
                    $stmt = $pdo->prepare("
                        INSERT INTO report_sessions (unit_id, reporter_name, report_date, shift) 
                        VALUES (?, ?, ?, ?)
                    ");
                    $stmt->execute([$_SESSION['unit_id'], $reporter_name, $report_date, $shift]);
                    $session_id = $pdo->lastInsertId();
                } else {
                    $session_id = $session['id'];
                }
                
                // Insert patient referral
                $stmt = $pdo->prepare("
                    INSERT INTO patient_referrals 
                    (session_id, patient_name, patient_age, medical_record, origin_room, 
                     insurance_type, diagnosis, destination_hospital, referral_reason) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ");
                
                $stmt->execute([
                    $session_id, $patient_name, $patient_age, $medical_record, 
                    $origin_room, $insurance_type, $diagnosis, $destination_hospital, $referral_reason
                ]);
                
                $pdo->commit();
                $message = 'Data pasien rujuk berhasil disimpan!';
                $message_type = 'success';
                
                // Clear form data
                $_POST = array();
                
            } catch (Exception $e) {
                $pdo->rollBack();
                $message = 'Gagal menyimpan data: ' . $e->getMessage();
                $message_type = 'danger';
            }
        } else {
            $message = 'Harap isi semua field yang wajib!';
            $message_type = 'danger';
        }
    }
}

// Get recent referrals for this unit
$stmt = $pdo->prepare("
    SELECT pr.*, rs.reporter_name, rs.report_date, rs.shift 
    FROM patient_referrals pr 
    JOIN report_sessions rs ON pr.session_id = rs.id 
    WHERE rs.unit_id = ? 
    ORDER BY pr.created_at DESC 
    LIMIT 10
");
$stmt->execute([$_SESSION['unit_id']]);
$recent_referrals = $stmt->fetchAll();
?>

<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Data Pasien Rujuk - <?= $unit['name'] ?></title>
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
        .form-section {
            background: white;
            border-radius: 15px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .section-header {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
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
                        <a href="#" class="nav-link active">
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
                            <h2 class="mb-0">Data Pasien Rujuk</h2>
                            <p class="text-muted mb-0">Unit: <?= $unit['name'] ?></p>
                        </div>
                        <a href="unit_dashboard.php" class="btn btn-outline-secondary">
                            <i class="bi bi-arrow-left me-2"></i>Kembali
                        </a>
                    </div>
                    
                    <?php if ($message): ?>
                        <div class="alert alert-<?= $message_type ?> alert-dismissible fade show" role="alert">
                            <?= $message ?>
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    <?php endif; ?>
                    
                    <!-- Form Input -->
                    <div class="form-section">
                        <div class="section-header">
                            <h5 class="mb-0">
                                <i class="bi bi-person-plus me-2"></i>
                                Tambah Data Pasien Rujuk
                            </h5>
                        </div>
                        <div class="p-4">
                            <form method="POST">
                                <input type="hidden" name="action" value="add_referral">
                                
                                <!-- Reporter Information -->
                                <h6 class="mb-3 text-primary">
                                    <i class="bi bi-person-circle me-2"></i>
                                    Informasi Pelapor
                                </h6>
                                <div class="row mb-4">
                                    <div class="col-md-4">
                                        <label class="form-label">Nama Pelapor *</label>
                                        <input type="text" class="form-control" name="reporter_name" 
                                               value="<?= $_POST['reporter_name'] ?? '' ?>" required>
                                    </div>
                                    <div class="col-md-4">
                                        <label class="form-label">Tanggal *</label>
                                        <input type="date" class="form-control" name="report_date" 
                                               value="<?= $_POST['report_date'] ?? date('Y-m-d') ?>" required>
                                    </div>
                                    <div class="col-md-4">
                                        <label class="form-label">Shift *</label>
                                        <select class="form-select" name="shift" required>
                                            <option value="">Pilih Shift</option>
                                            <option value="pagi" <?= ($_POST['shift'] ?? '') === 'pagi' ? 'selected' : '' ?>>Pagi</option>
                                            <option value="siang" <?= ($_POST['shift'] ?? '') === 'siang' ? 'selected' : '' ?>>Siang</option>
                                            <option value="malam" <?= ($_POST['shift'] ?? '') === 'malam' ? 'selected' : '' ?>>Malam</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <!-- Patient Information -->
                                <h6 class="mb-3 text-success">
                                    <i class="bi bi-person-heart me-2"></i>
                                    Data Pasien
                                </h6>
                                <div class="row mb-4">
                                    <div class="col-md-6">
                                        <label class="form-label">Nama Pasien *</label>
                                        <input type="text" class="form-control" name="patient_name" 
                                               value="<?= $_POST['patient_name'] ?? '' ?>" 
                                               placeholder="Masukkan nama lengkap pasien" required>
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">Umur *</label>
                                        <input type="text" class="form-control" name="patient_age" 
                                               value="<?= $_POST['patient_age'] ?? '' ?>" 
                                               placeholder="contoh: 45 tahun" required>
                                    </div>
                                    <div class="col-md-3">
                                        <label class="form-label">No. Rekam Medis *</label>
                                        <input type="text" class="form-control" name="medical_record" 
                                               value="<?= $_POST['medical_record'] ?? '' ?>" 
                                               placeholder="No. RM" required>
                                    </div>
                                </div>
                                
                                <div class="row mb-4">
                                    <div class="col-md-6">
                                        <label class="form-label">Asal Ruangan *</label>
                                        <input type="text" class="form-control" name="origin_room" 
                                               value="<?= $_POST['origin_room'] ?? '' ?>" 
                                               placeholder="contoh: ICU, Rawat Inap Lantai 3" required>
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Jenis Jaminan *</label>
                                        <select class="form-select" name="insurance_type" required>
                                            <option value="">Pilih Jaminan</option>
                                            <option value="BPJS" <?= ($_POST['insurance_type'] ?? '') === 'BPJS' ? 'selected' : '' ?>>BPJS</option>
                                            <option value="Umum" <?= ($_POST['insurance_type'] ?? '') === 'Umum' ? 'selected' : '' ?>>Umum</option>
                                            <option value="Asuransi" <?= ($_POST['insurance_type'] ?? '') === 'Asuransi' ? 'selected' : '' ?>>Asuransi</option>
                                            <option value="Perusahaan" <?= ($_POST['insurance_type'] ?? '') === 'Perusahaan' ? 'selected' : '' ?>>Perusahaan</option>
                                            <option value="InHealth" <?= ($_POST['insurance_type'] ?? '') === 'InHealth' ? 'selected' : '' ?>>InHealth</option>
                                            <option value="BPJS TK" <?= ($_POST['insurance_type'] ?? '') === 'BPJS TK' ? 'selected' : '' ?>>BPJS TK</option>
                                            <option value="BPJS COB" <?= ($_POST['insurance_type'] ?? '') === 'BPJS COB' ? 'selected' : '' ?>>BPJS COB</option>
                                            <option value="Beban Direksi" <?= ($_POST['insurance_type'] ?? '') === 'Beban Direksi' ? 'selected' : '' ?>>Beban Direksi</option>
                                        </select>
                                    </div>
                                </div>
                                
                                <!-- Medical Information -->
                                <h6 class="mb-3 text-warning">
                                    <i class="bi bi-clipboard-heart me-2"></i>
                                    Informasi Medis
                                </h6>
                                <div class="row mb-4">
                                    <div class="col-12">
                                        <label class="form-label">Diagnosa *</label>
                                        <textarea class="form-control" name="diagnosis" rows="3" 
                                                  placeholder="Masukkan diagnosa lengkap pasien" required><?= $_POST['diagnosis'] ?? '' ?></textarea>
                                    </div>
                                </div>
                                
                                <!-- Referral Information -->
                                <h6 class="mb-3 text-danger">
                                    <i class="bi bi-hospital me-2"></i>
                                    Informasi Rujukan
                                </h6>
                                <div class="row mb-4">
                                    <div class="col-md-6">
                                        <label class="form-label">Rumah Sakit Tujuan *</label>
                                        <input type="text" class="form-control" name="destination_hospital" 
                                               value="<?= $_POST['destination_hospital'] ?? '' ?>" 
                                               placeholder="contoh: RS. Siloam, RSUP Dr. Mohammad Hoesin" required>
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Alasan Rujukan</label>
                                        <textarea class="form-control" name="referral_reason" rows="3" 
                                                  placeholder="Alasan pasien dirujuk (opsional)"><?= $_POST['referral_reason'] ?? '' ?></textarea>
                                    </div>
                                </div>
                                
                                <div class="d-flex justify-content-end">
                                    <button type="reset" class="btn btn-outline-secondary me-2">
                                        <i class="bi bi-arrow-clockwise me-2"></i>Reset
                                    </button>
                                    <button type="submit" class="btn btn-primary">
                                        <i class="bi bi-plus-circle me-2"></i>Simpan Data Rujuk
                                    </button>
                                </div>
                            </form>
                        </div>
                    </div>
                    
                    <!-- Recent Referrals -->
                    <div class="form-section">
                        <div class="section-header">
                            <h5 class="mb-0">
                                <i class="bi bi-clock-history me-2"></i>
                                Data Rujukan Terbaru
                            </h5>
                        </div>
                        <div class="p-4">
                            <?php if (empty($recent_referrals)): ?>
                                <div class="text-center py-5">
                                    <i class="bi bi-inbox fs-1 text-muted"></i>
                                    <h5 class="mt-3 text-muted">Belum Ada Data Rujukan</h5>
                                    <p class="text-muted">Silakan tambahkan data pasien rujuk pertama.</p>
                                </div>
                            <?php else: ?>
                                <div class="table-responsive">
                                    <table class="table table-hover">
                                        <thead class="table-light">
                                            <tr>
                                                <th>#</th>
                                                <th>Tanggal</th>
                                                <th>Nama Pasien</th>
                                                <th>Umur</th>
                                                <th>No. RM</th>
                                                <th>Jaminan</th>
                                                <th>RS Tujuan</th>
                                                <th>Pelapor</th>
                                                <th>Aksi</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <?php foreach ($recent_referrals as $index => $referral): ?>
                                                <tr>
                                                    <td><?= $index + 1 ?></td>
                                                    <td><?= formatDate($referral['report_date']) ?></td>
                                                    <td>
                                                        <strong><?= $referral['patient_name'] ?></strong>
                                                        <br><small class="text-muted"><?= $referral['origin_room'] ?></small>
                                                    </td>
                                                    <td><?= $referral['patient_age'] ?></td>
                                                    <td><code><?= $referral['medical_record'] ?></code></td>
                                                    <td>
                                                        <span class="badge bg-info"><?= $referral['insurance_type'] ?></span>
                                                    </td>
                                                    <td><?= $referral['destination_hospital'] ?></td>
                                                    <td>
                                                        <?= $referral['reporter_name'] ?>
                                                        <br><small class="text-muted">Shift: <?= ucfirst($referral['shift']) ?></small>
                                                    </td>
                                                    <td>
                                                        <button class="btn btn-sm btn-outline-primary" 
                                                                onclick="viewReferralDetail(<?= htmlspecialchars(json_encode($referral)) ?>)">
                                                            <i class="bi bi-eye"></i>
                                                        </button>
                                                    </td>
                                                </tr>
                                            <?php endforeach; ?>
                                        </tbody>
                                    </table>
                                </div>
                                
                                <div class="text-center mt-3">
                                    <a href="all_referrals.php" class="btn btn-outline-primary">
                                        <i class="bi bi-arrow-right me-2"></i>
                                        Lihat Semua Data Rujukan
                                    </a>
                                </div>
                            <?php endif; ?>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Detail Modal -->
    <div class="modal fade" id="detailModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="bi bi-person-heart me-2"></i>
                        Detail Pasien Rujuk
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <div class="col-md-6">
                            <h6 class="text-primary mb-3">Data Pasien</h6>
                            <table class="table table-sm">
                                <tr>
                                    <td width="40%"><strong>Nama</strong></td>
                                    <td id="detail_patient_name">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Umur</strong></td>
                                    <td id="detail_patient_age">-</td>
                                </tr>
                                <tr>
                                    <td><strong>No. RM</strong></td>
                                    <td id="detail_medical_record">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Asal Ruangan</strong></td>
                                    <td id="detail_origin_room">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Jaminan</strong></td>
                                    <td id="detail_insurance_type">-</td>
                                </tr>
                            </table>
                        </div>
                        <div class="col-md-6">
                            <h6 class="text-success mb-3">Informasi Rujukan</h6>
                            <table class="table table-sm">
                                <tr>
                                    <td width="40%"><strong>RS Tujuan</strong></td>
                                    <td id="detail_destination_hospital">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Tanggal</strong></td>
                                    <td id="detail_report_date">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Shift</strong></td>
                                    <td id="detail_shift">-</td>
                                </tr>
                                <tr>
                                    <td><strong>Pelapor</strong></td>
                                    <td id="detail_reporter_name">-</td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    <div class="row mt-3">
                        <div class="col-12">
                            <h6 class="text-warning mb-3">Diagnosa</h6>
                            <div class="p-3 bg-light rounded">
                                <p id="detail_diagnosis" class="mb-0">-</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="row mt-3">
                        <div class="col-12">
                            <h6 class="text-danger mb-3">Alasan Rujukan</h6>
                            <div class="p-3 bg-light rounded">
                                <p id="detail_referral_reason" class="mb-0">-</p>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Tutup</button>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        function viewReferralDetail(referral) {
            // Fill modal with referral data
            document.getElementById('detail_patient_name').textContent = referral.patient_name;
            document.getElementById('detail_patient_age').textContent = referral.patient_age;
            document.getElementById('detail_medical_record').textContent = referral.medical_record;
            document.getElementById('detail_origin_room').textContent = referral.origin_room;
            document.getElementById('detail_insurance_type').textContent = referral.insurance_type;
            document.getElementById('detail_destination_hospital').textContent = referral.destination_hospital;
            document.getElementById('detail_report_date').textContent = new Date(referral.report_date).toLocaleDateString('id-ID');
            document.getElementById('detail_shift').textContent = referral.shift.charAt(0).toUpperCase() + referral.shift.slice(1);
            document.getElementById('detail_reporter_name').textContent = referral.reporter_name;
            document.getElementById('detail_diagnosis').textContent = referral.diagnosis || '-';
            document.getElementById('detail_referral_reason').textContent = referral.referral_reason || 'Tidak ada alasan khusus';
            
            // Show modal
            var detailModal = new bootstrap.Modal(document.getElementById('detailModal'));
            detailModal.show();
        }
        
        // Auto-fill reporter name from session storage
        const reporterNameInput = document.querySelector('input[name="reporter_name"]');
        if (reporterNameInput && !reporterNameInput.value) {
            const savedReporter = localStorage.getItem('reporter_name');
            if (savedReporter) {
                reporterNameInput.value = savedReporter;
            }
        }
        
        // Save reporter name to session storage
        if (reporterNameInput) {
            reporterNameInput.addEventListener('blur', function() {
                if (this.value.trim()) {
                    localStorage.setItem('reporter_name', this.value.trim());
                }
            });
        }
        
        // Form validation
        document.querySelector('form').addEventListener('submit', function(e) {
            const requiredFields = this.querySelectorAll('input[required], select[required], textarea[required]');
            let isValid = true;
            
            requiredFields.forEach(field => {
                if (!field.value.trim()) {
                    field.classList.add('is-invalid');
                    isValid = false;
                } else {
                    field.classList.remove('is-invalid');
                }
            });
            
            if (!isValid) {
                e.preventDefault();
                alert('Harap isi semua field yang wajib diisi!');
            }
        });
        
        // Real-time validation feedback
        document.querySelectorAll('input[required], select[required], textarea[required]').forEach(field => {
            field.addEventListener('blur', function() {
                if (!this.value.trim()) {
                    this.classList.add('is-invalid');
                } else {
                    this.classList.remove('is-invalid');
                }
            });
        });
    </script>
</body>
</html>