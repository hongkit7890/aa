<?php
// 第5堂：更新狀態
require "db.php";

$id = isset($_GET["id"]) ? (int)$_GET["id"] : 0;
$status = $_GET["status"] ?? "";

if ($id > 0 && ($status === "pending" || $status === "done")) {
    $sql = "UPDATE repairs SET status = ? WHERE id = ?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("si", $status, $id);
    $stmt->execute();
    $stmt->close();
}

header("Location: list.php");
exit;
