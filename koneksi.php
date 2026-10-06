<?php
$host = "sql300.infinityfree.com";
$user = "if0_42695977";
$pass = "ZiAklhm2NK07"; 
$db   = "if0_42695977_barcode";

$conn = new mysqli($host, $user, $pass, $db);

if ($conn->connect_error) {
    die("Koneksi gagal: " . $conn->connect_error);
}
?>