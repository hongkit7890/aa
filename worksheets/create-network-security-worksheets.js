/**
 * DSE ICT 中五「網上威脅及保安」工作紙
 * Aligned with EDB ICT C&A Guide (Internet and its Applications)
 */
const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  LevelFormat,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageNumber,
  PageBreak,
} = require("docx");

const OUT_DIR = __dirname;
const A4_WIDTH = 11906;
const A4_HEIGHT = 16838;
const MARGIN = 720; // 0.5 inch — more writing space for worksheets
const CONTENT_WIDTH = A4_WIDTH - MARGIN * 2; // 10466

const thin = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: thin, bottom: thin, left: thin, right: thin };
const noBorder = {
  style: BorderStyle.NONE,
  size: 0,
  color: "FFFFFF",
};
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thickBottom = {
  style: BorderStyle.SINGLE,
  size: 18,
  color: "1F4E79",
};

const BLUE = "1F4E79";
const LIGHT_BLUE = "D6E3F0";
const LIGHT_GREY = "F2F2F2";
const ACCENT = "2E75B6";

function t(text, opts = {}) {
  return new TextRun({
    text,
    font: "Microsoft JhengHei",
    size: opts.size ?? 21, // 10.5pt
    bold: opts.bold,
    italics: opts.italics,
    color: opts.color,
  });
}

function p(children, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 80, before: opts.before ?? 0, line: opts.line ?? 276 },
    alignment: opts.align,
    ...opts.extra,
    children: Array.isArray(children) ? children : [children],
  });
}

function blankLine(height = 200) {
  return new Paragraph({
    spacing: { after: height },
    children: [t("")],
  });
}

function answerLines(n = 3) {
  const lines = [];
  for (let i = 0; i < n; i++) {
    lines.push(
      new Paragraph({
        spacing: { before: 120, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
        children: [t(" ")],
      })
    );
  }
  return lines;
}

function sectionTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 160 },
    border: { bottom: { ...thickBottom, space: 4 } },
    children: [t(text, { size: 26, bold: true, color: BLUE })],
  });
}

function subTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [t(text, { size: 22, bold: true, color: ACCENT })],
  });
}

function qNum(n, text, marks) {
  return p([
    t(`${n}. `, { bold: true, size: 21 }),
    t(text, { size: 21 }),
    marks ? t(`  [${marks}分]`, { size: 18, color: "666666" }) : t(""),
  ], { before: 140, after: 60 });
}

function choice(letter, text) {
  return p([t(`    ${letter}.  ${text}`)], { after: 40 });
}

function cell(text, width, opts = {}) {
  const runs = Array.isArray(text)
    ? text
    : [t(String(text), { bold: opts.bold, size: opts.size ?? 18, color: opts.color })];
  return new TableCell({
    borders: opts.noBorder ? noBorders : borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill
      ? { fill: opts.fill, type: ShadingType.CLEAR }
      : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.colSpan,
    children: [
      new Paragraph({
        alignment: opts.align ?? AlignmentType.LEFT,
        children: runs,
      }),
    ],
  });
}

function infoBox(title, bodyLines) {
  const w = CONTENT_WIDTH;
  return new Table({
    width: { size: w, type: WidthType.DXA },
    columnWidths: [w],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: w, type: WidthType.DXA },
            shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              p([t(title, { bold: true, size: 20, color: BLUE })], { after: 60 }),
              ...bodyLines.map((line) =>
                p([t(line, { size: 18 })], { after: 40 })
              ),
            ],
          }),
        ],
      }),
    ],
  });
}

function headerFooter(docTitle) {
  return {
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 4 } },
            spacing: { after: 120 },
            children: [
              t("HKDSE 資訊及通訊科技｜中五", { size: 16, color: BLUE, bold: true }),
              t("　　", { size: 16 }),
              t(docTitle, { size: 16, color: "666666" }),
            ],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 4 } },
            alignment: AlignmentType.CENTER,
            spacing: { before: 80 },
            children: [
              t("第 ", { size: 16, color: "666666" }),
              new TextRun({
                children: [PageNumber.CURRENT],
                font: "Microsoft JhengHei",
                size: 16,
                color: "666666",
              }),
              t(" 頁　｜　單元：互聯網及其應用 — 網上威脅及保安", { size: 16, color: "666666" }),
            ],
          }),
        ],
      }),
    },
  };
}

function baseStyles() {
  return {
    default: {
      document: {
        run: { font: "Microsoft JhengHei", size: 21 },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, font: "Microsoft JhengHei", color: BLUE },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 22, bold: true, font: "Microsoft JhengHei", color: ACCENT },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 },
      },
    ],
  };
}

function numberingConfig() {
  return {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
      {
        reference: "notes-bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
      {
        reference: "revise-bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  };
}

/* ===================== STUDENT WORKSHEET ===================== */

function buildStudentDoc() {
  const hf = headerFooter("工作紙（學生版）");
  const children = [
    // Title block
    p([t("資訊及通訊科技科　HKDSE 溫習工作紙", { size: 18, color: ACCENT })], {
      align: AlignmentType.CENTER,
      after: 40,
    }),
    p([t("網上威脅及保安", { size: 36, bold: true, color: BLUE })], {
      align: AlignmentType.CENTER,
      after: 40,
    }),
    p([t("中五　｜　必修單元：互聯網及其應用", { size: 20, color: "444444" })], {
      align: AlignmentType.CENTER,
      after: 160,
    }),

    // Meta table
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2616, 2617, 2616, 2617],
      rows: [
        new TableRow({
          children: [
            cell("姓名：______________", 2616, { size: 18 }),
            cell("班別：________", 2617, { size: 18 }),
            cell("學號：________", 2616, { size: 18 }),
            cell("日期：__________", 2617, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("建議時限：40 分鐘", 2616, { size: 18, fill: LIGHT_GREY }),
            cell("總分：50 分", 2617, { size: 18, fill: LIGHT_GREY }),
            cell("程度：中五／DSE", 2616, { size: 18, fill: LIGHT_GREY }),
            cell("題型：選擇＋短答＋結構", 2617, { size: 18, fill: LIGHT_GREY }),
          ],
        }),
      ],
    }),

    blankLine(120),

    infoBox("學習目標（對應課程指引）", [
      "描述常見網絡保安威脅所造成的潛在風險（病毒、蠕蟲、木馬、間諜軟件、勒索軟件、未經授權存取、攔截、動態網頁入侵、DoS 等）。",
      "提議有效措施以改善網絡保安（防毒軟件、瀏覽器設定、認證與權限控制、防火牆、WPA、VPN 等）。",
      "討論網上潛在私隱威脅，並建議保護私隱的方法。",
    ]),

    blankLine(80),

    // Part A
    sectionTitle("甲部　多項選擇題（每題 1 分，共 10 分）"),
    p([t("請將正確答案填寫在括號內。", { size: 18, italics: true, color: "555555" })]),

    qNum("1", "下列哪一項惡意軟件最可能透過電子郵件附件散播，並會自我複製至其他電腦？　　（　　）"),
    choice("A", "木馬程式（Trojan）"),
    choice("B", "蠕蟲（Worm）"),
    choice("C", "間諜軟件（Spyware）"),
    choice("D", "廣告軟件（Adware）"),

    qNum("2", "勒索軟件（Ransomware）的主要特徵是：　　（　　）"),
    choice("A", "暗中記錄用戶的鍵盤輸入"),
    choice("B", "加密用戶檔案並要求支付贖金才解鎖"),
    choice("C", "偽裝成合法軟件，但不自我複製"),
    choice("D", "只顯示大量彈出式廣告"),

    qNum("3", "防火牆（Firewall）的主要功能是：　　（　　）"),
    choice("A", "掃描並清除硬碟上的病毒"),
    choice("B", "根據規則監控並過濾進出網絡的流量"),
    choice("C", "為無線網絡提供加密密碼"),
    choice("D", "自動備份用戶的重要檔案"),

    qNum("4", "下列哪一項無線安全協議較 WEP 更安全，常被家用路由器採用？　　（　　）"),
    choice("A", "HTTP"),
    choice("B", "FTP"),
    choice("C", "WPA／WPA2／WPA3"),
    choice("D", "SMTP"),

    qNum("5", "虛擬私人網絡（VPN）最能幫助用戶：　　（　　）"),
    choice("A", "加快下載速度"),
    choice("B", "在公共 Wi-Fi 上建立加密通道以保護傳輸數據"),
    choice("C", "取代防毒軟件"),
    choice("D", "自動修復被感染的系統檔案"),

    qNum("6", "仿冒詐騙（Phishing）通常透過以下哪種方式進行？　　（　　）"),
    choice("A", "發送看似來自銀行的電郵，誘導用戶輸入帳號密碼"),
    choice("B", "向伺服器發送大量請求使其癱瘓"),
    choice("C", "在無線網絡上竊聽未加密的數據"),
    choice("D", "利用漏洞令程式緩衝區溢位"),

    qNum("7", "拒絕服務攻擊（DoS）的目的是：　　（　　）"),
    choice("A", "竊取伺服器資料庫中的客戶資料"),
    choice("B", "使目標系統或網絡無法正常提供服務"),
    choice("C", "安裝後門以便日後遙控電腦"),
    choice("D", "修改網頁內容進行宣傳"),

    qNum("8", "下列哪一項最能描述「認證」（Authentication）？　　（　　）"),
    choice("A", "決定用戶可以存取哪些資源"),
    choice("B", "核實用戶身份是否屬實"),
    choice("C", "加密傳輸中的數據"),
    choice("D", "記錄系統的所有操作日誌"),

    qNum("9", "用戶在瀏覽器網址列見到「https://」及掛鎖圖示，通常表示：　　（　　）"),
    choice("A", "該網站一定沒有惡意軟件"),
    choice("B", "與伺服器之間的傳輸已加密（如使用 SSL／TLS）"),
    choice("C", "該網站已通過政府認證"),
    choice("D", "用戶的防毒軟件正在運作"),

    qNum("10", "「最小權限原則」（Principle of Least Privilege）是指：　　（　　）"),
    choice("A", "所有員工應使用同一個系統管理員帳號"),
    choice("B", "用戶只獲授予完成工作所需的最低權限"),
    choice("C", "關閉所有防火牆規則以方便存取"),
    choice("D", "禁止使用任何密碼"),

    // Part B
    sectionTitle("乙部　填充及配對（共 10 分）"),
    subTitle("B1　填充（每空 1 分，共 5 分）"),
    p([t("請填寫適當的詞語。")]),

    qNum(
      "11",
      "____________________ 會偽裝成有用程式誘騙用戶安裝，但本身一般不會自我複製。"
    ),
    blankLine(60),
    qNum(
      "12",
      "____________________ 會在用戶不知情下收集個人資料或上網習慣，並傳送給第三者。"
    ),
    blankLine(60),
    qNum(
      "13",
      "____________________ 是指未獲准許下進入電腦系統或網絡，例如黑客入侵。"
    ),
    blankLine(60),
    qNum(
      "14",
      "在公共場所使用未加密 Wi-Fi 時，攻擊者可能進行 ____________________，竊取傳輸中的敏感資料。"
    ),
    blankLine(60),
    qNum(
      "15",
      "定期更新 ____________________ 定義檔，有助偵測最新的惡意軟件。"
    ),

    subTitle("B2　配對（每項 1 分，共 5 分）"),
    p([t("將左欄威脅／措施與右欄描述配對，把英文字母填在括號內。")]),
    blankLine(40),

    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [4800, 5666],
      rows: [
        new TableRow({
          children: [
            cell("左欄", 4800, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
            cell("右欄", 5666, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            cell("16. 病毒　　（　　）", 4800, { size: 18 }),
            cell("A. 限制不同用戶可讀寫的檔案或功能", 5666, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("17. 存取控制　　（　　）", 4800, { size: 18 }),
            cell("B. 需依附宿主檔案，感染後可破壞或複製", 5666, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("18. 雙重認證（2FA）　　（　　）", 4800, { size: 18 }),
            cell("C. 除密碼外再要求一次性驗證碼等", 5666, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("19. SQL 注入（動態網頁入侵）　　（　　）", 4800, { size: 18 }),
            cell("D. 透過惡意輸入操控資料庫查詢", 5666, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("20. 垃圾電郵（Spam）　　（　　）", 4800, { size: 18 }),
            cell("E. 大量未經索取的推廣或欺詐郵件", 5666, { size: 18 }),
          ],
        }),
      ],
    }),

    // Part C
    sectionTitle("丙部　短答題（共 15 分）"),

    qNum(
      "21",
      "比較「病毒」與「蠕蟲」在散播方式上的兩項主要分別。",
      "4"
    ),
    ...answerLines(4),

    qNum(
      "22",
      "列出三項可改善個人電腦網絡保安的有效措施，並簡述每項如何減低風險。",
      "6"
    ),
    ...answerLines(5),

    qNum(
      "23",
      "解釋認證（Authentication）與授權（Authorization）的分別，並各舉一例。",
      "5"
    ),
    ...answerLines(4),

    // Part D
    sectionTitle("丁部　結構／情境題（共 15 分）"),
    p([
      t("細閱以下情境，然後回答問題。", { italics: true, color: "555555" }),
    ]),
    blankLine(40),

    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders,
              width: { size: CONTENT_WIDTH, type: WidthType.DXA },
              shading: { fill: "FFF8E7", type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [
                p([t("情境", { bold: true, size: 20, color: BLUE })], { after: 60 }),
                p([
                  t(
                    "某中學的「電子學習平台」最近發生事故：多名師生收到聲稱來自「學校資訊科技組」的電郵，要求點擊連結並登入以「更新密碼」。部分同學在公共咖啡店的 Wi-Fi 下登入後，帳戶被盜用；同時平台曾短暫無法開啟。其後發現伺服器日誌出現大量異常請求。資訊科技組亦發現有老師的電腦出現檔案被加密、桌面留下勒索訊息的情況。",
                    { size: 19 }
                  ),
                ], { after: 40 }),
              ],
            }),
          ],
        }),
      ],
    }),

    blankLine(80),

    qNum(
      "24",
      "指出情境中出現的三種不同網絡保安威脅，並分別說明其證據。",
      "6"
    ),
    ...answerLines(5),

    qNum(
      "25",
      "就「公共 Wi-Fi 登入」一事，建議兩項保護措施，並解釋原因。",
      "4"
    ),
    ...answerLines(4),

    qNum(
      "26",
      "學校應如何從技術及管理層面加強平台保安？提出三項建議。",
      "5"
    ),
    ...answerLines(5),

    // Revision checklist
    new Paragraph({ children: [new PageBreak()] }),
    sectionTitle("溫習清單（自我檢測）"),
    p([t("完成工作紙後，請勾選你已掌握的概念：", { size: 19 })]),
    blankLine(40),

    ...[
      "能區分病毒、蠕蟲、木馬、間諜軟件、勒索軟件",
      "明白未經授權存取、攔截、動態網頁入侵、DoS 的含義與影響",
      "能提議防毒、防火牆、認證／權限、瀏覽器設定等措施",
      "理解 WPA 與 VPN 如何保護無線／公共網絡傳輸",
      "能討論釣魚、竊聽、垃圾電郵等私隱威脅及應對",
      "分清認證與授權；認識 HTTPS／加密的基本作用",
      "能就真實情境分析威脅並提出多層防禦建議",
    ].map(
      (item) =>
        new Paragraph({
          numbering: { reference: "revise-bullets", level: 0 },
          spacing: { after: 80 },
          children: [t(`☐  ${item}`, { size: 19 })],
        })
    ),

    blankLine(160),
    infoBox("應試小貼士", [
      "結構題宜用「指出威脅 → 引用情境證據 → 建議措施並解釋」三段式作答。",
      "比較題（如病毒 vs 蠕蟲）要寫「兩方面對比」，避免只描述其中一方。",
      "措施題要寫「做什麼」＋「如何減低風險」，單講名稱通常只能得部分分。",
      "注意課程用語：Authentication／Authorization、Firewall、VPN、WPA、DoS、Phishing。",
    ]),
  ];

  return new Document({
    styles: baseStyles(),
    numbering: numberingConfig(),
    sections: [
      {
        properties: {
          page: {
            size: { width: A4_WIDTH, height: A4_HEIGHT },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        ...hf,
        children,
      },
    ],
  });
}

/* ===================== ANSWER KEY ===================== */

function ans(text) {
  return p([t(text, { size: 20 })], { after: 60 });
}

function markScheme(lines) {
  return lines.map((line) =>
    new Paragraph({
      numbering: { reference: "notes-bullets", level: 0 },
      spacing: { after: 40 },
      children: [t(line, { size: 19 })],
    })
  );
}

function buildAnswerDoc() {
  const hf = headerFooter("教師參考／答案");
  const children = [
    p([t("資訊及通訊科技科　HKDSE 溫習工作紙", { size: 18, color: ACCENT })], {
      align: AlignmentType.CENTER,
      after: 40,
    }),
    p([t("網上威脅及保安　—　答案及評分參考", { size: 32, bold: true, color: BLUE })], {
      align: AlignmentType.CENTER,
      after: 40,
    }),
    p([t("中五　｜　供教師使用　｜　總分 50", { size: 20, color: "444444" })], {
      align: AlignmentType.CENTER,
      after: 160,
    }),

    infoBox("使用說明", [
      "括號內為建議得分；意思正確而用語略異可酌情給分。",
      "結構題鼓勵引用情境證據；空泛答案宜扣分。",
      "建議先讓學生限時完成學生版，再對照本卷討論。",
    ]),

    sectionTitle("甲部　多項選擇題（每題 1 分，共 10 分）"),

    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [1200, 1200, 8066],
      rows: [
        new TableRow({
          children: [
            cell("題號", 1200, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
            cell("答案", 1200, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
            cell("簡析", 8066, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
          ],
        }),
        ...[
          ["1", "B", "蠕蟲可自我複製並經網絡／電郵散播；木馬一般不自我複製。"],
          ["2", "B", "勒索軟件加密檔案並索取贖金。"],
          ["3", "B", "防火牆按規則過濾進出流量。"],
          ["4", "C", "WPA／WPA2／WPA3 為無線加密協議，較 WEP 安全。"],
          ["5", "B", "VPN 建立加密隧道，適合公共 Wi-Fi。"],
          ["6", "A", "釣魚誘導用戶交出憑證。"],
          ["7", "B", "DoS 令服務不可用。"],
          ["8", "B", "認證＝核實身份；授權才是決定權限。"],
          ["9", "B", "HTTPS 表示傳輸加密（SSL／TLS），不保證網站本身無惡意。"],
          ["10", "B", "最小權限：只給完成工作所需的最低權限。"],
        ].map(
          ([n, a, exp]) =>
            new TableRow({
              children: [
                cell(n, 1200, { align: AlignmentType.CENTER, size: 18 }),
                cell(a, 1200, { align: AlignmentType.CENTER, bold: true, size: 18, color: BLUE }),
                cell(exp, 8066, { size: 17 }),
              ],
            })
        ),
      ],
    }),

    sectionTitle("乙部　填充及配對（共 10 分）"),
    subTitle("B1　填充"),
    ans("11. 木馬程式／Trojan（1）"),
    ans("12. 間諜軟件／Spyware（1）"),
    ans("13. 未經授權存取／Unauthorized access（1）"),
    ans("14. 攔截／竊聽／Interception／Eavesdropping（1）"),
    ans("15. 防毒軟件／病毒（定義）／Antivirus（virus definitions）（1）"),

    subTitle("B2　配對"),
    ans("16 → B　　17 → A　　18 → C　　19 → D　　20 → E　　（各 1 分）"),

    sectionTitle("丙部　短答題（共 15 分）"),

    p([t("21. 病毒與蠕蟲的分別（4）", { bold: true, size: 21, color: BLUE })]),
    ...markScheme([
      "病毒需依附宿主檔案／程式才能傳播；蠕蟲可獨立存在並自我複製。（2）",
      "病毒常需用戶開啟受感染檔案才發動；蠕蟲可自動經網絡／電郵漏洞散播。（2）",
      "（其他合理對比亦可，例如對網絡流量影響、傳播速度等。）",
    ]),

    blankLine(60),
    p([t("22. 三項保安措施（6）", { bold: true, size: 21, color: BLUE })]),
    p([t("任答三項，每項「措施＋如何減低風險」各 2 分。示例：", { size: 18, italics: true })]),
    ...markScheme([
      "安裝並定期更新防毒軟件：偵測／隔離惡意軟件。（2）",
      "啟用防火牆：過濾可疑進出連線。（2）",
      "使用強密碼及雙重認證：降低帳戶被盜用風險。（2）",
      "謹慎瀏覽器設定／不隨意下載外掛：減少惡意程式植入。（2）",
      "使用 VPN（尤其公共 Wi-Fi）：加密傳輸，防攔截。（2）",
      "為 Wi-Fi 使用 WPA2／WPA3：防止未授權接入與竊聽。（2）",
      "最小權限／存取控制：限制受損範圍。（2）",
    ]),

    blankLine(60),
    p([t("23. 認證與授權（5）", { bold: true, size: 21, color: BLUE })]),
    ...markScheme([
      "認證：核實「你是誰」（例如輸入密碼、指紋、一次性驗證碼）。（2）",
      "授權：在身份確認後決定「你可以做什麼／存取什麼」（例如學生只能讀筆記、管理員可改設定）。（2）",
      "例子清楚對應兩者各得 0.5～1 分；總分不超過 5。",
    ]),

    sectionTitle("丁部　結構／情境題（共 15 分）"),

    p([t("24. 三種威脅＋證據（6）", { bold: true, size: 21, color: BLUE })]),
    p([t("每項「正確指出威脅＋引用情境」各 2 分。可接受答案：", { size: 18, italics: true })]),
    ...markScheme([
      "仿冒詐騙／釣魚：偽冒「學校資訊科技組」電郵要求點擊連結登入。（2）",
      "攔截／未加密公共 Wi-Fi 風險：同學在咖啡店 Wi-Fi 登入後帳戶被盜。（2）",
      "拒絕服務（DoS／DDoS）：平台短暫無法開啟，日誌有大量異常請求。（2）",
      "勒索軟件：老師電腦檔案被加密並留下勒索訊息。（2）",
      "（答對任何三項即可得滿分 6。）",
    ]),

    blankLine(60),
    p([t("25. 公共 Wi-Fi 保護措施（4）", { bold: true, size: 21, color: BLUE })]),
    p([t("兩項，每項措施＋原因各 2 分。示例：", { size: 18, italics: true })]),
    ...markScheme([
      "使用 VPN：即使流量被截取亦難以閱讀。（2）",
      "避免在公共 Wi-Fi 進行登入／網上銀行等敏感操作；改用流動數據。（2）",
      "確認網站為 HTTPS；不連接可疑開放熱點。（2）",
      "啟用雙重認證：即使密碼外洩亦較難盜用帳戶。（2）",
    ]),

    blankLine(60),
    p([t("26. 學校加強保安建議（5）", { bold: true, size: 21, color: BLUE })]),
    p([t("三項合理建議，技術／管理均可；建議分佈約 2+2+1 或 2+1+2。示例：", { size: 18, italics: true })]),
    ...markScheme([
      "技術：強制 HTTPS、WAF／輸入驗證防 SQL 注入、防火牆／入侵偵測、定期更新系統與防毒、備份。（任選）",
      "管理：保安意識培訓、電郵警示釣魚特徵、事故應變程序、最小權限帳號政策、強制 2FA。",
      "網絡：分隔訪客 Wi-Fi 與內網；監控異常流量以應對 DoS。",
    ]),

    blankLine(120),
    sectionTitle("建議評分總覽"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [3500, 2000, 4966],
      rows: [
        new TableRow({
          children: [
            cell("部分", 3500, { bold: true, fill: LIGHT_BLUE }),
            cell("分數", 2000, { bold: true, fill: LIGHT_BLUE, align: AlignmentType.CENTER }),
            cell("備註", 4966, { bold: true, fill: LIGHT_BLUE }),
          ],
        }),
        new TableRow({
          children: [
            cell("甲部　選擇題", 3500, { size: 18 }),
            cell("10", 2000, { size: 18, align: AlignmentType.CENTER }),
            cell("每題 1 分", 4966, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("乙部　填充＋配對", 3500, { size: 18 }),
            cell("10", 2000, { size: 18, align: AlignmentType.CENTER }),
            cell("每空／項 1 分", 4966, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("丙部　短答", 3500, { size: 18 }),
            cell("15", 2000, { size: 18, align: AlignmentType.CENTER }),
            cell("Q21–23", 4966, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("丁部　情境結構", 3500, { size: 18 }),
            cell("15", 2000, { size: 18, align: AlignmentType.CENTER }),
            cell("Q24–26", 4966, { size: 18 }),
          ],
        }),
        new TableRow({
          children: [
            cell("總分", 3500, { size: 18, bold: true, fill: LIGHT_GREY }),
            cell("50", 2000, { size: 18, bold: true, align: AlignmentType.CENTER, fill: LIGHT_GREY }),
            cell("建議及格參考：25／50", 4966, { size: 18, fill: LIGHT_GREY }),
          ],
        }),
      ],
    }),

    blankLine(160),
    p([
      t("課程對應：教育局《資訊及通訊科技課程及評估指引》— 互聯網及其應用 — 網上威脅及保安。", {
        size: 16,
        color: "666666",
        italics: true,
      }),
    ]),
  ];

  return new Document({
    styles: baseStyles(),
    numbering: numberingConfig(),
    sections: [
      {
        properties: {
          page: {
            size: { width: A4_WIDTH, height: A4_HEIGHT },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        ...hf,
        children,
      },
    ],
  });
}

async function main() {
  const studentPath = path.join(OUT_DIR, "DSE_中五_網絡保安_工作紙_學生版.docx");
  const answerPath = path.join(OUT_DIR, "DSE_中五_網絡保安_工作紙_答案.docx");

  const studentBuf = await Packer.toBuffer(buildStudentDoc());
  fs.writeFileSync(studentPath, studentBuf);
  console.log("Wrote", studentPath);

  const answerBuf = await Packer.toBuffer(buildAnswerDoc());
  fs.writeFileSync(answerPath, answerBuf);
  console.log("Wrote", answerPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
