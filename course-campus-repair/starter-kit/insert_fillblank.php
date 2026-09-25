<?php
// 第3堂填空版：把 ____ 填上正確關鍵字
require "____";  // 提示：db.php

$name = $_POST["____"] ?? "";
$location = $_POST["____"] ?? "";
$description = $_POST["____"] ?? "";

$sql = "INSERT INTO repairs (name, location, description, status) VALUES (?, ?, ?, 'pending')";
$stmt = $conn->prepare($sql);
$stmt->bind_param("sss", $name, $location, $description);

if ($stmt->____()) {  // 提示：execute
    echo "新增成功！<a href='list.php'>查看列表</a>";
} else {
    echo "新增失敗";
}
