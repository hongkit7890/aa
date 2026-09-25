<?php
// 第3堂：接收表單並寫入資料庫（prepared statement）
require "db.php";

$name = $_POST["name"] ?? "";
$location = $_POST["location"] ?? "";
$description = $_POST["description"] ?? "";

// ★加分：空欄位檢查
if ($name === "" || $location === "" || $description === "") {
    die("請填寫完整資料。<a href='form.html'>返回</a>");
}

$sql = "INSERT INTO repairs (name, location, description, status) VALUES (?, ?, ?, 'pending')";
$stmt = $conn->prepare($sql);
$stmt->bind_param("sss", $name, $location, $description);

if ($stmt->execute()) {
    // ★★加分：寫入後跳到列表
    header("Location: list.php");
    exit;
}

echo "新增失敗：" . $conn->error;
$stmt->close();
$conn->close();
