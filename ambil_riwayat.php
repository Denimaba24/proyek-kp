<?php
header("Access-Control-Allow-Origin: *");
header("Content-Type: application/json; charset=UTF-8");

include 'koneksi.php';
$sql = "SELECT id, kode_export, tanggal, jumlah_data, detail_data FROM riwayat_export ORDER BY tanggal DESC";
$result = $conn->query($sql);

$riwayat = array();

if ($result->num_rows > 0) {
    while($row = $result->fetch_assoc()) {
        $row['detail_data'] = json_decode($row['detail_data']);
        $riwayat[] = $row;
    }
}

echo json_encode(array("status" => "success", "data" => $riwayat));

$conn->close();
?>