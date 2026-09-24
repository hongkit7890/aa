<?php
// 第4堂：查詢並顯示＋簡單狀態篩選
require "db.php";

$status = $_GET["status"] ?? "all";

if ($status === "pending" || $status === "done") {
    $sql = "SELECT * FROM repairs WHERE status = ? ORDER BY created_at DESC";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("s", $status);
    $stmt->execute();
    $result = $stmt->get_result();
} else {
    $result = $conn->query("SELECT * FROM repairs ORDER BY created_at DESC");
}
?>
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <title>報修列表</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <main class="wrap wide">
    <h1>報修管理列表</h1>
    <p><a href="form.html">新增報修</a></p>

    <form method="get" class="filter">
      <label>狀態篩選
        <select name="status" onchange="this.form.submit()">
          <option value="all" <?php if ($status === "all") echo "selected"; ?>>全部</option>
          <option value="pending" <?php if ($status === "pending") echo "selected"; ?>>待處理</option>
          <option value="done" <?php if ($status === "done") echo "selected"; ?>>已完成</option>
        </select>
      </label>
    </form>

    <table>
      <thead>
        <tr>
          <th>編號</th>
          <th>報修人</th>
          <th>地點</th>
          <th>問題描述</th>
          <th>狀態</th>
          <th>時間</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
      <?php while ($row = $result->fetch_assoc()): ?>
        <tr>
          <td><?php echo (int)$row["id"]; ?></td>
          <td><?php echo htmlspecialchars($row["name"]); ?></td>
          <td><?php echo htmlspecialchars($row["location"]); ?></td>
          <td><?php echo htmlspecialchars($row["description"]); ?></td>
          <td><?php echo htmlspecialchars($row["status"]); ?></td>
          <td><?php echo htmlspecialchars($row["created_at"]); ?></td>
          <td>
            <?php if ($row["status"] !== "done"): ?>
              <a href="update.php?id=<?php echo (int)$row["id"]; ?>&status=done">標示完成</a>
            <?php else: ?>
              <a href="update.php?id=<?php echo (int)$row["id"]; ?>&status=pending">改回待處理</a>
            <?php endif; ?>
          </td>
        </tr>
      <?php endwhile; ?>
      </tbody>
    </table>
  </main>
</body>
</html>
