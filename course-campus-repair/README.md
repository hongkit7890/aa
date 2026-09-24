# 校園小系統實作：數據庫與 PHP 入門（中三教材）

## 設計原則

- 全班統一做「校園報修系統」
- 每堂 1 個新概念 + 1 個可見成果
- 筆記精簡；提供 Starter Kit 與填空程式
- 教案含 VASK（價值觀／態度／技能／知識）

## 產出檔案（`output/`）

1. `01_完整教案_含VASK.docx`
2. `02_五節教學進度表.docx`
3. `03_導學案_五堂合一.docx`
4. `04_*.pptx`（合訂＋分堂簡報）
5. `05_精簡教師教案_五堂.docx`
6. `06_學生半頁筆記與任務單_*.docx`
7. `07_評分規準_作弊條_錯誤對照.docx`
8. `08_PPT大綱與逐頁要點.md`／`.docx`
9. `09_成功畫面卡.docx`

## Starter Kit（`starter-kit/`）

SQL、db.php、表單、insert／list／update、填空版、style.css

## 重新產生文件

```bash
pip install -r requirements.txt
python3 scripts/generate_materials.py
python3 scripts/generate_plan_pack.py
```

