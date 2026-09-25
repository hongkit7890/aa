-- 校園報修系統：資料庫與資料表
-- 在 phpMyAdmin 可直接匯入，或手動執行

CREATE DATABASE IF NOT EXISTS campus_system
  DEFAULT CHARACTER SET utf8mb4
  COLLATE utf8mb4_general_ci;

USE campus_system;

CREATE TABLE IF NOT EXISTS repairs (
  id INT AUTO_INCREMENT PRIMARY KEY COMMENT '編號（主鍵）',
  name VARCHAR(50) NOT NULL COMMENT '報修人',
  location VARCHAR(100) NOT NULL COMMENT '地點',
  description TEXT NOT NULL COMMENT '問題描述',
  status VARCHAR(20) NOT NULL DEFAULT 'pending' COMMENT '狀態 pending/done',
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '建立時間'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 範例資料（可選）
INSERT INTO repairs (name, location, description, status) VALUES
('陳同學', '3樓電腦室', '投影機無法開機', 'pending'),
('李老師', '圖書館', '冷氣漏水', 'pending'),
('王同學', '操場旁洗手間', '水龍頭損壞', 'done');
