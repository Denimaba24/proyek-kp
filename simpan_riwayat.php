<?php
header("Access-Control-Allow-Origin: *");
header("Content-Type: application/json; charset=UTF-8");
header("Access-Control-Allow-Methods: POST");
header("Access-Control-Allow-Headers: Content-Type");

include 'koneksi.php';

$data = json_decode(file_get_contents("php://input"));

if (!empty($data->kode_export) && isset($data->jumlah_data)) {
    
    $kode_export = $conn->real_escape_string($data->kode_export);
    $jumlah_data = (int)$data->jumlah_data;
    
    $detail_data = $conn->real_escape_string(json_encode($data->detail_data));

    $sql = "INSERT INTO riwayat_export (kode_export, jumlah_data, detail_data) 
            VALUES ('$kode_export', $jumlah_data, '$detail_data')";

    if ($conn->query($sql) === TRUE) {
        http_response_code(201);
        echo json_encode(array("status" => "success", "message" => "Riwayat berhasil di-backup ke Server."));
    } else {
        http_response_code(503);
        echo json_encode(array("status" => "error", "message" => "Gagal menyimpan ke Server: " . $conn->error));
    }
} else {
    http_response_code(400);
    echo json_encode(array("status" => "error", "message" => "Data payload tidak lengkap."));
}

$conn->close();
?>