const ORDER_SHEET_NAME = "주문서";
const ITEM_SHEET_NAME = "주문상품";
const SETTINGS_PROPERTY_KEY = "GRAPE_ORDER_SETTINGS";

const ORDER_HEADERS = [
  "주문일시", "주문자명", "전화번호", "요청사항", "주문상품",
  "총 박스", "총 금액", "받으실 주소", "입금자명", "주문번호"
];

const ITEM_HEADERS = [
  "주문번호", "주문일시", "주문자명", "전화번호", "받으실 주소",
  "입금자명", "요청사항", "상품명", "단가", "수량(박스)", "소계",
  "주문 총 박스", "주문 총 금액"
];

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "");

  if (action === "settings") {
    return jsonp_(getPublicSettings_(), e && e.parameter && e.parameter.callback);
  }

  return jsonp_({ ok: false, message: "알 수 없는 요청입니다." }, e && e.parameter && e.parameter.callback);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("주문 데이터가 없습니다.");
    }

    const data = JSON.parse(e.postData.contents);
    const action = String((e.parameter && e.parameter.action) || data.action || "order");

    if (action === "settings") {
      saveSettings_(data);
      return json_({ ok: true });
    }

    if (action !== "order") {
      throw new Error("알 수 없는 요청입니다.");
    }

    const customer = data.customer || {};
    const phone = normalizePhone_(customer.phone);
    const items = Array.isArray(data.items) ? data.items : [];
    const storedSettings = getStoredSettings_();
    const soldOutProductIds = new Set(storedSettings.soldOutProductIds);
    const soldOutItems = items.filter(item => soldOutProductIds.has(String(item.id || "")));

    if (soldOutItems.length > 0) {
      throw new Error(
        "품절된 상품이 포함되어 있습니다: " +
        soldOutItems.map(item => String(item.name || item.id || "")).join(", ")
      );
    }

    const orderedAt = new Date();
    const orderId = createOrderId_(orderedAt);
    const totalBoxes = Number(data.totalBoxes) || 0;
    const total = Number(data.total) || 0;
    const itemSummary = items
      .map(item => String(item.name || "") + " x " + (Number(item.quantity) || 0) + "박스")
      .join(", ");

    lock.waitLock(30000);
    hasLock = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = getOrderSheet_(spreadsheet);
    const itemSheet = getItemSheet_(spreadsheet);

    const orderRow = [
      orderedAt,
      customer.name || "",
      phone,
      customer.note || "",
      itemSummary,
      totalBoxes,
      total,
      customer.address || "",
      customer.payerName || "",
      orderId
    ];
    const orderRowIndex = orderSheet.getLastRow() + 1;
    const orderRange = orderSheet.getRange(orderRowIndex, 1, 1, ORDER_HEADERS.length);
    orderRange.setValues([orderRow]);
    orderRange.getCell(1, 3).setNumberFormat("@").setValue(phone);

    const itemRows = items.map(item => [
      orderId,
      orderedAt,
      customer.name || "",
      phone,
      customer.address || "",
      customer.payerName || "",
      customer.note || "",
      item.name || "",
      Number(item.price) || 0,
      Number(item.quantity) || 0,
      Number(item.subtotal) || 0,
      totalBoxes,
      total
    ]);

    if (itemRows.length > 0) {
      const itemStartRow = itemSheet.getLastRow() + 1;
      const itemRange = itemSheet.getRange(itemStartRow, 1, itemRows.length, ITEM_HEADERS.length);
      const phoneRange = itemSheet.getRange(itemStartRow, 4, itemRows.length, 1);
      itemRange.setValues(itemRows);
      phoneRange.setNumberFormat("@").setValues(itemRows.map(() => [phone]));
    }

    SpreadsheetApp.flush();
    return json_({ ok: true, orderId: orderId });
  } catch (error) {
    return json_({ ok: false, message: String(error) });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function getStoredSettings_() {
  const raw = PropertiesService.getScriptProperties().getProperty(SETTINGS_PROPERTY_KEY);

  if (!raw) {
    return { shopName: null, introText: null, soldOutProductIds: [], updatedAt: "" };
  }

  try {
    const parsed = JSON.parse(raw);
    const hasShopName = Object.prototype.hasOwnProperty.call(parsed, "shopName");
    const hasIntroText = Object.prototype.hasOwnProperty.call(parsed, "introText");
    const soldOutProductIds = Array.isArray(parsed.soldOutProductIds)
      ? parsed.soldOutProductIds.map(String)
      : [];

    return {
      shopName: hasShopName ? String(parsed.shopName || "") : null,
      introText: hasIntroText ? String(parsed.introText || "") : null,
      soldOutProductIds: soldOutProductIds,
      updatedAt: String(parsed.updatedAt || "")
    };
  } catch (error) {
    return { shopName: null, introText: null, soldOutProductIds: [], updatedAt: "" };
  }
}

function getPublicSettings_() {
  const settings = getStoredSettings_();
  return {
    ok: true,
    shopName: settings.shopName,
    introText: settings.introText,
    soldOutProductIds: settings.soldOutProductIds,
    updatedAt: settings.updatedAt
  };
}

function saveSettings_(data) {
  const soldOutProductIds = Array.isArray(data.soldOutProductIds)
    ? data.soldOutProductIds.map(String)
    : [];

  PropertiesService.getScriptProperties().setProperty(
    SETTINGS_PROPERTY_KEY,
    JSON.stringify({
      shopName: String(data.shopName || ""),
      introText: String(data.introText || ""),
      soldOutProductIds: soldOutProductIds,
      updatedAt: new Date().toISOString()
    })
  );
}

function getOrderSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ORDER_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.getSheets().find(candidate =>
      candidate.getName() !== ITEM_SHEET_NAME
    ) || spreadsheet.insertSheet(ORDER_SHEET_NAME);

    if (sheet.getName() !== ORDER_SHEET_NAME) {
      sheet.setName(ORDER_SHEET_NAME);
    }
  }

  ensureHeader_(sheet, ORDER_HEADERS);
  return sheet;
}

function getItemSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(ITEM_SHEET_NAME)
    || spreadsheet.insertSheet(ITEM_SHEET_NAME);

  ensureHeader_(sheet, ITEM_HEADERS);
  return sheet;
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const firstCell = String(sheet.getRange(1, 1).getDisplayValue()).trim();
  if (firstCell !== headers[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function normalizePhone_(value) {
  const phone = String(value || "").trim();
  return phone.replace(/^10(?=[0-9-])/, "010");
}

function createOrderId_(orderedAt) {
  const datePart = Utilities.formatDate(orderedAt, Session.getScriptTimeZone(), "yyMMdd");
  const letter = String.fromCharCode(65 + Math.floor(Math.random() * 26));
  const number = String(Math.floor(Math.random() * 1000)).padStart(3, "0");
  return datePart + letter + number;
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(data, callback) {
  const callbackName = String(callback || "").trim();
  const payload = JSON.stringify(data);

  if (/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callbackName)) {
    return ContentService
      .createTextOutput(callbackName + "(" + payload + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return json_(data);
}
