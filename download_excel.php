<?php
require_once 'config.php';
requireAdmin();

// Get parameters
$date = $_GET['date'] ?? date('Y-m-d');
$session_id = $_GET['session_id'] ?? null;

// If specific session requested
if ($session_id) {
    // Get specific session data
    $stmt = $pdo->prepare("
        SELECT rs.*, u.name as unit_name 
        FROM report_sessions rs 
        JOIN units u ON rs.unit_id = u.id 
        WHERE rs.id = ?
    ");
    $stmt->execute([$session_id]);
    $session = $stmt->fetch();
    
    if (!$session) {
        die('Session tidak ditemukan');
    }
    
    $sessions = [$session];
    $filename = "Laporan_{$session['unit_name']}_" . date('Y-m-d', strtotime($session['report_date'])) . ".xlsx";
} else {
    // Get all sessions for the date
    $stmt = $pdo->prepare("
        SELECT rs.*, u.name as unit_name 
        FROM report_sessions rs 
        JOIN units u ON rs.unit_id = u.id 
        WHERE rs.report_date = ? 
        ORDER BY u.name, rs.shift
    ");
    $stmt->execute([$date]);
    $sessions = $stmt->fetchAll();
    
    $filename = "Laporan_Semua_Unit_" . date('Y-m-d', strtotime($date)) . ".xlsx";
}

if (empty($sessions)) {
    die('Tidak ada data laporan untuk tanggal tersebut');
}

// Start output
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Cache-Control: max-age=0');

// Create simple Excel format using HTML table
echo '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
echo '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' . "\n";

// Styles
echo '<Styles>
    <Style ss:ID="Header">
        <Font ss:Bold="1"/>
        <Interior ss:Color="#4472C4" ss:Pattern="Solid"/>
        <Font ss:Color="#FFFFFF"/>
    </Style>
    <Style ss:ID="SubHeader">
        <Font ss:Bold="1"/>
        <Interior ss:Color="#D9E2F3" ss:Pattern="Solid"/>
    </Style>
</Styles>' . "\n";

echo '<Worksheet ss:Name="Laporan Harian">' . "\n";
echo '<Table>' . "\n";

// Header
echo '<Row>';
echo '<Cell ss:StyleID="Header"><Data ss:Type="String">LAPORAN HARIAN RUMAH SAKIT</Data></Cell>';
echo '</Row>' . "\n";

echo '<Row>';
echo '<Cell><Data ss:Type="String">Tanggal: ' . formatDate($date) . '</Data></Cell>';
echo '</Row>' . "\n";

echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n"; // Empty row

foreach ($sessions as $session) {
    // Unit Header
    echo '<Row>';
    echo '<Cell ss:StyleID="SubHeader"><Data ss:Type="String">UNIT: ' . strtoupper($session['unit_name']) . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    // Session Info
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Pelapor: ' . $session['reporter_name'] . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Shift: ' . ucfirst($session['shift']) . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Waktu Lapor: ' . formatDateTime($session['created_at']) . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n"; // Empty row
    
    // Get questions and answers
    $stmt = $pdo->prepare("
        SELECT q.question_text, r.answer 
        FROM questions q 
        LEFT JOIN reports r ON q.id = r.question_id AND r.session_id = ?
        WHERE q.unit_id = ? AND q.is_active = 1 AND q.deleted_at IS NULL
        ORDER BY q.question_order, q.id
    ");
    $stmt->execute([$session['id'], $session['unit_id']]);
    $questions = $stmt->fetchAll();
    
    // Questions Header
    echo '<Row>';
    echo '<Cell ss:StyleID="Header"><Data ss:Type="String">No</Data></Cell>';
    echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Pertanyaan</Data></Cell>';
    echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Jawaban</Data></Cell>';
    echo '</Row>' . "\n";
    
    // Questions and Answers
    foreach ($questions as $index => $question) {
        echo '<Row>';
        echo '<Cell><Data ss:Type="Number">' . ($index + 1) . '</Data></Cell>';
        echo '<Cell><Data ss:Type="String">' . htmlspecialchars($question['question_text']) . '</Data></Cell>';
        echo '<Cell><Data ss:Type="String">' . htmlspecialchars($question['answer'] ?: '-') . '</Data></Cell>';
        echo '</Row>' . "\n";
    }
    
    echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n"; // Empty row
    
    // Get patient referrals
    $stmt = $pdo->prepare("
        SELECT * FROM patient_referrals 
        WHERE session_id = ?
        ORDER BY created_at
    ");
    $stmt->execute([$session['id']]);
    $referrals = $stmt->fetchAll();
    
    if (!empty($referrals)) {
        echo '<Row>';
        echo '<Cell ss:StyleID="SubHeader"><Data ss:Type="String">DATA PASIEN RUJUK</Data></Cell>';
        echo '</Row>' . "\n";
        
        // Referral Header
        echo '<Row>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">No</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Nama Pasien</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Umur</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">No. RM</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Asal Ruangan</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Jaminan</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Diagnosa</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">RS Tujuan</Data></Cell>';
        echo '<Cell ss:StyleID="Header"><Data ss:Type="String">Alasan Rujuk</Data></Cell>';
        echo '</Row>' . "\n";
        
        // Referral Data
        foreach ($referrals as $index => $referral) {
            echo '<Row>';
            echo '<Cell><Data ss:Type="Number">' . ($index + 1) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['patient_name']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['patient_age']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['medical_record']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['origin_room']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['insurance_type']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['diagnosis']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['destination_hospital']) . '</Data></Cell>';
            echo '<Cell><Data ss:Type="String">' . htmlspecialchars($referral['referral_reason'] ?: '-') . '</Data></Cell>';
            echo '</Row>' . "\n";
        }
    }
    
    echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n"; // Empty row
    echo '<Row><Cell><Data ss:Type="String">==========================================</Data></Cell></Row>' . "\n";
    echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n"; // Empty row
}

// Summary Section
if (count($sessions) > 1) {
    echo '<Row>';
    echo '<Cell ss:StyleID="Header"><Data ss:Type="String">RINGKASAN</Data></Cell>';
    echo '</Row>' . "\n";
    
    $total_questions = 0;
    $total_answered = 0;
    $total_referrals = 0;
    
    foreach ($sessions as $session) {
        // Count questions and answers for this session
        $stmt = $pdo->prepare("
            SELECT 
                COUNT(q.id) as total_q,
                COUNT(r.id) as answered_q
            FROM questions q 
            LEFT JOIN reports r ON q.id = r.question_id AND r.session_id = ?
            WHERE q.unit_id = ? AND q.is_active = 1 AND q.deleted_at IS NULL
        ");
        $stmt->execute([$session['id'], $session['unit_id']]);
        $counts = $stmt->fetch();
        
        $total_questions += $counts['total_q'];
        $total_answered += $counts['answered_q'];
        
        // Count referrals
        $stmt = $pdo->prepare("SELECT COUNT(*) as ref_count FROM patient_referrals WHERE session_id = ?");
        $stmt->execute([$session['id']]);
        $total_referrals += $stmt->fetch()['ref_count'];
    }
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Total Unit Melaporkan: ' . count($sessions) . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Total Pertanyaan: ' . $total_questions . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Total Jawaban: ' . $total_answered . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Total Pasien Rujuk: ' . $total_referrals . '</Data></Cell>';
    echo '</Row>' . "\n";
    
    $completion_rate = $total_questions > 0 ? round(($total_answered / $total_questions) * 100, 2) : 0;
    echo '<Row>';
    echo '<Cell><Data ss:Type="String">Tingkat Kelengkapan: ' . $completion_rate . '%</Data></Cell>';
    echo '</Row>' . "\n";
}

// Footer
echo '<Row><Cell><Data ss:Type="String"></Data></Cell></Row>' . "\n";
echo '<Row>';
echo '<Cell><Data ss:Type="String">Digenerate pada: ' . date('d/m/Y H:i:s') . '</Data></Cell>';
echo '</Row>' . "\n";

echo '<Row>';
echo '<Cell><Data ss:Type="String">Oleh: ' . $_SESSION['name'] . '</Data></Cell>';
echo '</Row>' . "\n";

echo '</Table>' . "\n";
echo '</Worksheet>' . "\n";
echo '</Workbook>' . "\n";

exit;
?>