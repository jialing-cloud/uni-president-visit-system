// ==========================================
// 統一企業參訪活動平台 - 後端主程式
// Uni-President Visit Activity Platform
// ==========================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // ← 請替換成你的 Google Sheet ID

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('統一企業參訪活動專用')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────
// 取得選單項目 / Get menu items
// ─────────────────────────────────────────
function getMenuItems() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('選單設定');
    if (!sheet) return getDefaultMenu();
    const data = sheet.getDataRange().getValues();
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][3] === '啟用') {
        items.push({
          id:          data[i][0],
          name:        data[i][1],
          nameEn:      data[i][2],
          category:    data[i][4] || '飲料',
          emoji:       data[i][5] || '☕',
          tempSelect:  data[i][6] === true || data[i][6] === 'TRUE',
          noIceSelect: data[i][7] === true || data[i][7] === 'TRUE'
        });
      }
    }
    return items.length > 0 ? items : getDefaultMenu();
  } catch (e) {
    return getDefaultMenu();
  }
}

function getMenuItems() {
  // 修改這裡：因為你的圖片目前在 GitHub 根目錄，所以把 images/ 刪掉
  const base = "https://raw.githubusercontent.com/jialing-cloud/uni-president-visit-system/main/";
  return [
    { id: 'm1', name: '焦糖瑪奇朵', nameEn: 'Caramel Macchiato', img: base + 'caramel.png' },
    { id: 'm2', name: '冰蜜柚紅茶', nameEn: 'Iced Shaken Black Tea', img: base + 'iced_tea.png' },
    { id: 'm3', name: '星沁爽', nameEn: 'Strawberry Acai with Lemonade', img: base + 'strawberry.png' },
    { id: 'm4', name: '那堤', nameEn: 'Caffè Latte', img: base + 'latte.png' },
    { id: 'm5', name: '星冰樂', nameEn: 'Java Chip Frappuccino', img: base + 'java_chip.png' },
    { id: 'm6', name: '冷萃咖啡', nameEn: 'Cold Brew Coffee', img: base + 'cold_brew.png' }
  ];
}
// ─────────────────────────────────────────
// 提交訂單 / Submit order
// ─────────────────────────────────────────
function submitOrder(formData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('訂單記錄');

    if (!sheet) {
      sheet = ss.insertSheet('訂單記錄');
      const headers = [
        '提交時間 / Submitted At',
        '參訪日期 / Visit Date',
        '活動名稱 / Event',
        '姓名 / Name',
        '單位 / Organization',
        '類別 / Category',
        '品項 / Item',
        '品項英文 / Item (EN)',
        '溫度 / Temperature',
        '冰塊 / Ice Level',
        '備註 / Notes',
        '來源 / Source'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#1E4620')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const row = [
      new Date(),
      formData.visitDate,
      formData.eventTitle   || '',
      formData.name,
      formData.department   || '',
      formData.itemCategory || '飲料',
      formData.itemName,
      formData.itemNameEn   || '',
      formData.temperature  || '—',
      formData.iceLevel     || '—',
      formData.notes        || '',
      'Web表單'
    ];

    sheet.appendRow(row);
    sheet.autoResizeColumns(1, 12);
    updateSummary(ss);

    return { success: true, message: '已成功送出，謝謝您！\nYour order has been received!' };
  } catch (e) {
    return { success: false, message: '發生錯誤：' + e.message };
  }
}

// ─────────────────────────────────────────
// 更新統計摘要 / Update summary
// ─────────────────────────────────────────
function updateSummary(ss) {
  try {
    let sheet = ss.getSheetByName('統計摘要');
    if (!sheet) {
      sheet = ss.insertSheet('統計摘要');
    }
    const orderSheet = ss.getSheetByName('訂單記錄');
    if (!orderSheet || orderSheet.getLastRow() < 2) return;

    const orders = orderSheet.getDataRange().getValues();
    const countMap = {};

    for (let i = 1; i < orders.length; i++) {
      const date = orders[i][1];
      const item = orders[i][6];
      const temp = orders[i][8];
      const ice  = orders[i][9];
      const key  = `${date}___${item}___${temp}___${ice}`;
      if (!countMap[key]) countMap[key] = { date, item, temp, ice, count: 0 };
      countMap[key].count++;
    }

    sheet.clearContents();
    const headers = ['參訪日期 / Date', '品項 / Item', '溫度 / Temp', '冰塊 / Ice', '數量 / Qty', '最後更新'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E4620').setFontColor('#FFFFFF');

    const rows = Object.values(countMap).map(v => [v.date, v.item, v.temp, v.ice, v.count, new Date()]);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    sheet.autoResizeColumns(1, 6);
  } catch (e) {
    console.log('統計更新失敗：' + e.message);
  }
}

// ─────────────────────────────────────────
// 初始化試算表 / Initialize sheets
// ─────────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let menuSheet = ss.getSheetByName('選單設定');
  if (!menuSheet) {
    menuSheet = ss.insertSheet('選單設定');
    const headers = ['品項ID', '品項名稱', '英文名稱', '狀態', '類別', 'Emoji', '可選溫度(TRUE/FALSE)', '不可選冰塊(TRUE/FALSE)'];
    menuSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    menuSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E4620').setFontColor('#FFFFFF');
    const rows = getDefaultMenu().map(i => [i.id, i.name, i.nameEn, '啟用', i.category, i.emoji, i.tempSelect, i.noIceSelect]);
    menuSheet.getRange(2, 1, rows.length, 8).setValues(rows);
    menuSheet.autoResizeColumns(1, 8);
  }

  let sessionSheet = ss.getSheetByName('活動場次');
  if (!sessionSheet) {
    sessionSheet = ss.insertSheet('活動場次');
    const headers = ['參訪日期', '活動名稱', '來訪單位', '負責人', '備註'];
    sessionSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sessionSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E4620').setFontColor('#FFFFFF');
    sessionSheet.autoResizeColumns(1, 5);
  }

  return '初始化完成！';
}
