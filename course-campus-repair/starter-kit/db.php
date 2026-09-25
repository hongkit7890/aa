<?php
// 資料庫連線（教師預先設定好，學生 require 即可）
// XAMPP 預設：帳號 root、密碼空白

$host = "localhost";
$user = "root";
$pass = "";
$dbname = "campus_system";

$conn = new mysqli($host, $user, $pass, $dbname);

if ($conn->connect_error) {
    die("連線失敗：" . $conn->connect_error);
}

// 設定編碼，避免中文亂碼
$conn->set_charset("utf8mb4");
