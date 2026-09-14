const ORDER_SHEET_NAME = "주문서";
const ITEM_SHEET_NAME = "주문상품";
const STAFF_ORDER_SHEET_NAME = "[직원용] 주문서";
const STAFF_ITEM_SHEET_NAME = "[직원용] 주문상품";
const SETTINGS_PROPERTY_KEY = "GRAPE_ORDER_SETTINGS";
const STAFF_SETTINGS_PROPERTY_KEY = "GRAPE_STAFF_ORDER_SETTINGS";

const ORDER_HEADERS = [
  "주문일시", "주문자 이름", "주문자 연락처", "택배 받는 사람 이름",
  "택배 받는 사람 연락처", "입금자", "주문상품", "요청사항",
  "택배 받으실 주소", "총박스", "총금액", "주문번호"
];

const ITEM_HEADERS = [
  "주문번호", "주문일시", "주문자 이름", "주문자 연락처",
  "택배 받는 사람 이름", "택배 받는 사람 연락처", "택배 받으실 주소",
  "입금자명", "요청사항", "상품명", "단가", "수량(박스)", "소계",
  "주문 총 박스", "주문 총 금액"
];

const STAFF_ORDER_HEADERS = [
  "주문일시", "주문자 이름", "주문자 연락처", "수령 날짜", "입금자",
  "주문상품", "요청사항", "총박스", "총금액", "주문번호"
];

const STAFF_ITEM_HEADERS = [
  "주문번호", "주문일시", "주문자 이름", "주문자 연락처",
  "수령 날짜", "입금자명", "요청사항", "상품명", "단가", "수량(박스)", "소계", "주문 총 박스", "주문 총 금액"
];

const STAFF_PICKUP_DATES = new Set(["9/14(월)", "9/15(화)"]);

const STAFF_PRODUCTS = {
  "staff-gold-muscat": { name: "골드머스켓 2KG", price: 10000 },
  "staff-italia": { name: "이탈리아 2KG", price: 25000 }
};

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "");
  const orderType = String((e && e.parameter && e.parameter.orderType) || "");

  if (action === "settings") {
    return jsonp_(getPublicSettings_(orderType), e && e.parameter && e.parameter.callback);
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

    const isStaffOrder = String(data.orderType || "") === "staff";
    const customer = data.customer || {};
    const phone = normalizePhone_(customer.phone);
    const pickupDate = String(customer.pickupDate || "").trim();
    const recipientPhone = normalizePhone_(customer.recipientPhone);
    const submittedItems = Array.isArray(data.items) ? data.items : [];
    const items = isStaffOrder ? normalizeStaffItems_(submittedItems) : submittedItems;
    const storedSettings = getStoredSettings_(isStaffOrder ? "staff" : "standard");
    const soldOutProductIds = new Set(storedSettings.soldOutProductIds);
    const hiddenProductIds = new Set(storedSettings.hiddenProductIds);
    const soldOutItems = items.filter(item => soldOutProductIds.has(String(item.id || "")));
    const hiddenItems = items.filter(item => hiddenProductIds.has(String(item.id || "")));

    if (isStaffOrder && !STAFF_PICKUP_DATES.has(pickupDate)) {
      throw new Error("수령 날짜를 선택해 주세요.");
    }

    if (soldOutItems.length > 0) {
      throw new Error(
        "품절된 상품이 포함되어 있습니다: " +
        soldOutItems.map(item => String(item.name || item.id || "")).join(", ")
      );
    }

    if (hiddenItems.length > 0) {
      throw new Error(
        "현재 주문할 수 없는 상품이 포함되어 있습니다: " +
        hiddenItems.map(item => String(item.name || item.id || "")).join(", ")
      );
    }

    const orderedAt = new Date();
    const orderId = createOrderId_(orderedAt);
    const totalBoxes = items.reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);
    const total = items.reduce((sum, item) => sum + (Number(item.subtotal) || 0), 0);
    const itemSummary = items
      .map(item => String(item.name || "") + " x " + (Number(item.quantity) || 0) + "박스")
      .join(", ");

    lock.waitLock(30000);
    hasLock = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = isStaffOrder
      ? getNamedSheet_(spreadsheet, STAFF_ORDER_SHEET_NAME, STAFF_ORDER_HEADERS)
      : getOrderSheet_(spreadsheet);
    const itemSheet = isStaffOrder
      ? getNamedSheet_(spreadsheet, STAFF_ITEM_SHEET_NAME, STAFF_ITEM_HEADERS)
      : getItemSheet_(spreadsheet);

    const orderRow = isStaffOrder
      ? [
          orderedAt,
          customer.name || "",
          phone,
          pickupDate,
          customer.payerName || "",
          itemSummary,
          customer.note || "",
          totalBoxes,
          total,
          orderId
        ]
      : [
          orderedAt,
          customer.name || "",
          phone,
          customer.recipientName || "",
          recipientPhone,
          customer.payerName || "",
          itemSummary,
          customer.note || "",
          customer.address || "",
          totalBoxes,
          total,
          orderId
        ];
    const orderHeaders = isStaffOrder ? STAFF_ORDER_HEADERS : ORDER_HEADERS;
    const orderRowIndex = orderSheet.getLastRow() + 1;
    const orderRange = orderSheet.getRange(orderRowIndex, 1, 1, orderHeaders.length);
    orderRange.setValues([orderRow]);
    orderRange.getCell(1, 3).setNumberFormat("@").setValue(phone);
    if (!isStaffOrder) {
      orderRange.getCell(1, 5).setNumberFormat("@").setValue(recipientPhone);
    }

    const itemRows = items.map(item => isStaffOrder
      ? [
          orderId,
          orderedAt,
          customer.name || "",
          phone,
          pickupDate,
          customer.payerName || "",
          customer.note || "",
          item.name || "",
          Number(item.price) || 0,
          Number(item.quantity) || 0,
          Number(item.subtotal) || 0,
          totalBoxes,
          total
        ]
      : [
          orderId,
          orderedAt,
          customer.name || "",
          phone,
          customer.recipientName || "",
          recipientPhone,
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
      const itemHeaders = isStaffOrder ? STAFF_ITEM_HEADERS : ITEM_HEADERS;
      const itemRange = itemSheet.getRange(itemStartRow, 1, itemRows.length, itemHeaders.length);
      itemRange.setValues(itemRows);
      itemSheet.getRange(itemStartRow, 4, itemRows.length, 1)
        .setNumberFormat("@").setValues(itemRows.map(() => [phone]));
      if (!isStaffOrder) {
        itemSheet.getRange(itemStartRow, 6, itemRows.length, 1)
          .setNumberFormat("@").setValues(itemRows.map(() => [recipientPhone]));
      }
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

function getSettingsPropertyKey_(orderType) {
  return String(orderType || "") === "staff"
    ? STAFF_SETTINGS_PROPERTY_KEY
    : SETTINGS_PROPERTY_KEY;
}

function getStoredSettings_(orderType) {
  const raw = PropertiesService.getScriptProperties().getProperty(
    getSettingsPropertyKey_(orderType)
  );

  if (!raw) {
    return { shopName: null, introText: null, soldOutProductIds: [], hiddenProductIds: [], updatedAt: "" };
  }

  try {
    const parsed = JSON.parse(raw);
    const hasShopName = Object.prototype.hasOwnProperty.call(parsed, "shopName");
    const hasIntroText = Object.prototype.hasOwnProperty.call(parsed, "introText");
    const soldOutProductIds = Array.isArray(parsed.soldOutProductIds)
      ? parsed.soldOutProductIds.map(String)
      : [];
    const hiddenProductIds = Array.isArray(parsed.hiddenProductIds)
      ? parsed.hiddenProductIds.map(String)
      : [];

    return {
      shopName: hasShopName ? String(parsed.shopName || "") : null,
      introText: hasIntroText ? String(parsed.introText || "") : null,
      soldOutProductIds: soldOutProductIds,
      hiddenProductIds: hiddenProductIds,
      updatedAt: String(parsed.updatedAt || "")
    };
  } catch (error) {
    return { shopName: null, introText: null, soldOutProductIds: [], hiddenProductIds: [], updatedAt: "" };
  }
}

function getPublicSettings_(orderType) {
  const settings = getStoredSettings_(orderType);
  return {
    ok: true,
    shopName: settings.shopName,
    introText: settings.introText,
    soldOutProductIds: settings.soldOutProductIds,
    hiddenProductIds: settings.hiddenProductIds,
    updatedAt: settings.updatedAt
  };
}

function saveSettings_(data) {
  const soldOutProductIds = Array.isArray(data.soldOutProductIds)
    ? data.soldOutProductIds.map(String)
    : [];
  const hiddenProductIds = Array.isArray(data.hiddenProductIds)
    ? data.hiddenProductIds.map(String)
    : [];

  PropertiesService.getScriptProperties().setProperty(
    getSettingsPropertyKey_(data.orderType),
    JSON.stringify({
      shopName: String(data.shopName || ""),
      introText: String(data.introText || ""),
      soldOutProductIds: soldOutProductIds,
      hiddenProductIds: hiddenProductIds,
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

function getNamedSheet_(spreadsheet, sheetName, headers) {
  const sheet = spreadsheet.getSheetByName(sheetName)
    || spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() > 0) {
    ["수령 날짜", "요청사항"].forEach(headerName => {
      const currentHeaders = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getDisplayValues()[0]
        .map(value => String(value).trim());

      if (headers.includes(headerName) && !currentHeaders.includes(headerName)) {
        const headerIndex = headers.indexOf(headerName);
        const nextHeader = headers[headerIndex + 1];
        const nextHeaderColumn = currentHeaders.indexOf(nextHeader) + 1;
        sheet.insertColumnBefore(nextHeaderColumn > 0 ? nextHeaderColumn : headerIndex + 1);
      }
    });
  }

  ensureHeader_(sheet, headers, false);
  return sheet;
}

function ensureHeader_(sheet, headers, shouldMigrateRecipientColumns) {
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
    return;
  }

  const currentHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(value => String(value).trim());
  const hasRecipientColumns = currentHeaders.includes("택배 받는 사람 이름");

  if (shouldMigrateRecipientColumns !== false && !hasRecipientColumns && headers[0] === "주문일시") {
    const payerColumn = currentHeaders.indexOf("입금자") + 1;
    sheet.insertColumnsBefore(payerColumn > 0 ? payerColumn : 4, 2);
  }

  if (shouldMigrateRecipientColumns !== false && !hasRecipientColumns && headers[0] === "주문번호") {
    const addressColumn = currentHeaders.findIndex(value =>
      value === "받으실 주소" || value === "택배 받으실 주소"
    ) + 1;
    sheet.insertColumnsBefore(addressColumn > 0 ? addressColumn : 5, 2);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function normalizeStaffItems_(items) {
  if (items.length === 0) {
    throw new Error("주문할 상품을 선택해 주세요.");
  }

  return items.map(item => {
    const product = STAFF_PRODUCTS[String(item.id || "")];
    const quantity = Math.max(0, Math.floor(Number(item.quantity) || 0));

    if (!product || quantity < 1) {
      throw new Error("직원용 주문 상품을 확인해 주세요.");
    }

    return {
      id: String(item.id || ""),
      name: product.name,
      price: product.price,
      quantity: quantity,
      subtotal: product.price * quantity
    };
  });
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
