<?php
// 第2堂：只接收並顯示，不寫入資料庫
$name = $_POST["name"] ?? "";
$location = $_POST["location"] ?? "";
$description = $_POST["description"] ?? "";
?>
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <title>收到報修</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <main class="wrap">
    <h1>收到報修！</h1>
    <p class="big">報修人：<?php echo htmlspecialchars($name); ?></p>
    <p class="big">地點：<?php echo htmlspecialchars($location); ?></p>
    <p class="big">問題：<?php echo htmlspecialchars($description); ?></p>
    <p><a href="form.html">再送一筆</a></p>
  </main>
</body>
</html>
