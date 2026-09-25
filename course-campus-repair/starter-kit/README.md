# 校園報修 Starter Kit

給中三學生的填空／範本專案。請複製到 XAMPP 的 `htdocs/campus_repair/`。

## 檔案說明

| 檔案 | 堂次 | 用途 |
|---|---|---|
| `sql/schema.sql` | 1 | 建立資料庫與範例資料 |
| `db.php` | 3+ | 連線（教師預設） |
| `form.html` | 2–5 | 報修表單 |
| `show.php` | 2 | 只顯示、不寫入 |
| `insert.php` | 3 | 寫入資料庫（完整版） |
| `insert_fillblank.php` | 3 | 填空練習版 |
| `list.php` | 4–5 | 列表＋篩選 |
| `update.php` | 5 | 更新狀態 |
| `style.css` | 5 | 現成美化（可改 `--brand`） |

## 建議學習順序

1. 匯入 `schema.sql` 或手動建表  
2. 第2堂把 `form.html` 的 action 暫時改成 `show.php`  
3. 第3堂改回 `insert.php`  
4. 第4–5堂使用 `list.php`／`update.php`

## 安全提醒

本套件僅供課堂本機練習，不要公開部署到網際網路。
