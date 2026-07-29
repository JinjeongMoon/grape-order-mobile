"use client";

import { useEffect, useMemo, useState } from "react";

type Product = {
  id: string;
  name: string;
  price: number;
};

type Settings = {
  shopName: string;
  introText: string;
  products: Product[];
  bankNotice: string;
  sheetEndpoint: string;
};

type Customer = {
  name: string;
  phone: string;
  address: string;
  payerName: string;
  note: string;
};

const giftSetGuide = [
  { id: "product-1", name: "명품 4KG", composition: "그랑포도 4~6종", price: 165000 },
  { id: "product-2", name: "프리미엄 4KG", composition: "그랑포도 랜덤 3~4종", price: 105000 },
  { id: "product-3", name: "프리미엄 2KG", composition: "그랑포도 랜덤 3~4종", price: 55000 },
  { id: "product-4", name: "베이직 4KG", composition: "그랑포도 랜덤 1~2종", price: 65000 },
  { id: "product-5", name: "베이직 2KG", composition: "그랑포도 랜덤 1~2종", price: 35000 },
];

const defaultSettings: Settings = {
  shopName: "그랑포도",
  introText: `내 가족이 먹는다는 마음으로 깨끗한 자연환경 속에서 한 송이 한 송이 정성껏 키웠습니다.
무농약 인증, 무비료, 무호르몬 처리로 안심하고 드실 수 있습니다.
100% 지하 암반수로만 재배한 최고급 프리미엄 유럽포도를 소중한 사람들과 함께 나누세요.

< 구매 안내 >
입금하신 순서대로 무료배송 해드립니다.
희귀 품종, 한정 물량이라 조기 소진될 경우, 구매가 불가하니 양해 부탁드립니다.

< 주문 문의 >
010-5490-7444`,
  products: giftSetGuide.map(({ id, name, price }) => ({ id, name, price })),
  bankNotice: "계좌이체: 은행명 000-0000-0000 예금주 홍길동",
  sheetEndpoint:
    "https://script.google.com/macros/s/AKfycbw97tf6AbG1mbGDxU3fAdPF1QCaaD8dIS3zlX-v8Ykaf3SDw5AK8Z-WetJAMOt4y4lM4A/exec",
};

const appsScriptCode = `const ORDER_SHEET_NAME = "주문서";
const ITEM_SHEET_NAME = "주문상품";

const ORDER_HEADERS = [
  "주문일시", "주문자명", "전화번호", "요청사항", "주문상품",
  "총 박스", "총 금액", "받으실 주소", "입금자명", "주문번호"
];

const ITEM_HEADERS = [
  "주문번호", "주문일시", "주문자명", "전화번호", "받으실 주소",
  "입금자명", "요청사항", "상품명", "단가", "수량(박스)", "소계",
  "주문 총 박스", "주문 총 금액"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("주문 데이터가 없습니다.");
    }

    const data = JSON.parse(e.postData.contents);
    const customer = data.customer || {};
    const phone = String(customer.phone || "");
    const items = Array.isArray(data.items) ? data.items : [];
    const orderedAt = new Date();
    const orderId = Utilities.getUuid();
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

    // 주문당 한 줄: 기존 주문 요약 시트에 저장합니다.
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
    orderRange.getCell(1, 3).setNumberFormat("@");
    orderRange.setValues([orderRow]);

    // 품목당 한 줄: 주문상품 시트에 저장합니다.
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
      itemSheet.getRange(itemStartRow, 4, itemRows.length, 1).setNumberFormat("@");
      itemSheet
        .getRange(itemStartRow, 1, itemRows.length, ITEM_HEADERS.length)
        .setValues(itemRows);
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

function getOrderSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ORDER_SHEET_NAME);

  if (!sheet) {
    // 기존 주문이 있는 첫 번째 시트를 주문서로 계속 사용합니다.
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
    // 이전 스크립트가 제목 없이 저장한 주문 데이터는 그대로 보존합니다.
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}`;

const won = new Intl.NumberFormat("ko-KR");

function toCurrency(value: number) {
  return `${won.format(value)}원`;
}

function getCopyableBankInfo(bankNotice: string) {
  return bankNotice.replace(/^\s*계좌이체\s*:\s*/, "").trim();
}

function encodeSettings(settings: Settings) {
  const compact = JSON.stringify(settings);
  return btoa(unescape(encodeURIComponent(compact)));
}

function decodeSettings(value: string): Settings | null {
  try {
    const decoded = decodeURIComponent(escape(atob(value)));
    const parsed = JSON.parse(decoded) as Settings;
    return {
      shopName: parsed.shopName || defaultSettings.shopName,
      introText: parsed.introText || defaultSettings.introText,
      products: defaultSettings.products,
      bankNotice: parsed.bankNotice || defaultSettings.bankNotice,
      sheetEndpoint: parsed.sheetEndpoint || defaultSettings.sheetEndpoint,
    };
  } catch {
    return null;
  }
}

function loadSettingsFromUrl() {
  if (typeof window === "undefined") {
    return defaultSettings;
  }

  const url = new URL(window.location.href);
  const shared = url.searchParams.get("config");
  if (shared) {
    return decodeSettings(shared) ?? defaultSettings;
  }

  const saved = window.localStorage.getItem("grape-order-settings");
  if (!saved) {
    return defaultSettings;
  }

  return decodeSettings(saved) ?? defaultSettings;
}

export default function Home() {
  const [settings, setSettings] = useState<Settings>(defaultSettings);
  const [quantities, setQuantities] = useState<Record<string, number>>({});
  const [customer, setCustomer] = useState<Customer>({
    name: "",
    phone: "",
    address: "",
    payerName: "",
    note: "",
  });
  const [mode, setMode] = useState<"order" | "admin">("order");
  const [status, setStatus] = useState("");
  const [toast, setToast] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [shareLink, setShareLink] = useState("");
  const [isAdminUnlocked, setIsAdminUnlocked] = useState(false);
  const [isPasswordOpen, setIsPasswordOpen] = useState(false);
  const [adminPassword, setAdminPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");

  useEffect(() => {
    const next = loadSettingsFromUrl();
    setSettings(next);
    setQuantities(
      Object.fromEntries(next.products.map((product) => [product.id, 0])),
    );
  }, []);

  useEffect(() => {
    if (!toast) {
      return;
    }

    const timer = window.setTimeout(() => setToast(""), 2000);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const activeProducts = useMemo(
    () => settings.products.filter((product) => product.name && product.price > 0),
    [settings.products],
  );

  const cart = useMemo(
    () =>
      activeProducts
        .map((product) => ({
          ...product,
          quantity: quantities[product.id] || 0,
          subtotal: product.price * (quantities[product.id] || 0),
        }))
        .filter((item) => item.quantity > 0),
    [activeProducts, quantities],
  );

  const total = cart.reduce((sum, item) => sum + item.subtotal, 0);
  const totalBoxes = cart.reduce((sum, item) => sum + item.quantity, 0);

  function updateQuantity(productId: string, amount: number) {
    setQuantities((current) => ({
      ...current,
      [productId]: Math.max(0, (current[productId] || 0) + amount),
    }));
  }

  function saveAdminSettings() {
    const normalized = {
      ...settings,
      products: defaultSettings.products,
    };
    const encoded = encodeSettings(normalized);
    window.localStorage.setItem("grape-order-settings", encoded);

    const url = new URL(window.location.href);
    url.searchParams.set("config", encoded);
    setShareLink(url.toString());
    setStatus("설정이 저장됐어요. 아래 공개 링크를 손님에게 보내면 됩니다.");
  }

  function handleModeButton() {
    if (mode === "admin") {
      setMode("order");
      return;
    }

    if (isAdminUnlocked) {
      setMode("admin");
      return;
    }

    setAdminPassword("");
    setPasswordError("");
    setIsPasswordOpen(true);
  }

  function submitAdminPassword(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (adminPassword === "2026") {
      setIsAdminUnlocked(true);
      setIsPasswordOpen(false);
      setMode("admin");
      setPasswordError("");
      return;
    }

    setPasswordError("비밀번호가 맞지 않습니다.");
  }

  async function copyBankNotice() {
    try {
      const bankInfo = getCopyableBankInfo(settings.bankNotice);
      const copyText = `${bankInfo || settings.bankNotice}\n입금 금액: ${toCurrency(total)}`;
      await navigator.clipboard.writeText(copyText);
      setToast("복사되었습니다");
    } catch {
      setToast("복사에 실패했습니다");
    }
  }

  async function submitOrder(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (isSubmitting) {
      return;
    }

    setStatus("");
    setIsSubmitting(true);

    if (!customer.name.trim() || !customer.phone.trim() || !customer.payerName.trim()) {
      setStatus("이름, 전화번호, 입금자명을 입력해 주세요.");
      setIsSubmitting(false);
      return;
    }

    if (cart.length === 0) {
      setStatus("상품을 1박스 이상 담아 주세요.");
      setIsSubmitting(false);
      return;
    }

    if (!settings.sheetEndpoint.trim()) {
      setStatus("주문 내용이 준비됐어요. 관리자 화면에서 구글시트 주소를 연결해 주세요.");
      setIsSubmitting(false);
      return;
    }

    const payload = {
      orderedAt: new Date().toISOString(),
      customer,
      items: cart.map(({ id, name, price, quantity, subtotal }) => ({
        id,
        name,
        price,
        quantity,
        subtotal,
      })),
      totalBoxes,
      total,
    };

    try {
      await fetch(settings.sheetEndpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload),
      });
      setCustomer({ name: "", phone: "", address: "", payerName: "", note: "" });
      setQuantities(
        Object.fromEntries(settings.products.map((product) => [product.id, 0])),
      );
      setStatus("주문이 접수됐어요. 계좌이체 후 확인 문자를 기다려 주세요.");
    } catch {
      setStatus("주문 전송에 실패했어요. 잠시 후 다시 시도해 주세요.");
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <main className="min-h-screen bg-[#f6f3eb] text-[#202016]">
      <section className="mx-auto flex min-h-screen w-full max-w-md flex-col px-4 py-5">
        <header className="flex items-center justify-between pb-4">
          <div>
            <p className="text-sm font-semibold text-[#6d6a55]">최고급 프리미엄 유럽 포도 선물세트</p>
            <h1 className="text-3xl font-black tracking-normal">{settings.shopName}</h1>
          </div>
          <button
            className={
              mode === "order"
                ? "px-1.5 py-1 text-xs font-medium text-[#aaa491] underline-offset-2 hover:text-[#6d6a55] hover:underline"
                : "rounded-full border border-[#d6ccb6] bg-white px-3 py-2 text-sm font-bold"
            }
            onClick={handleModeButton}
            type="button"
          >
            {mode === "order" ? "관리" : "주문"}
          </button>
        </header>

        {mode === "order" ? (
          <form className="flex flex-1 flex-col gap-4" onSubmit={submitOrder}>
            {settings.introText.trim() ? (
              <section className="rounded-lg border border-[#d6ccb6] bg-[#fffaf0] p-4">
                <h2 className="text-base font-black text-[#426b2f]">그랑포도 안내</h2>
                <p className="mt-3 whitespace-pre-line text-sm font-medium leading-6 text-[#4d4939]">
                  {settings.introText}
                </p>
              </section>
            ) : null}

            <section className="overflow-hidden rounded-lg border border-[#e1d7bd] bg-white">
              <h2 className="bg-[#8e294c] px-4 py-3 text-lg font-black text-white">
                그랑포도 선물 세트
              </h2>
              <table className="w-full table-fixed border-collapse text-left">
                <colgroup>
                  <col className="w-[30%]" />
                  <col className="w-[43%]" />
                  <col className="w-[27%]" />
                </colgroup>
                <thead className="bg-[#fffaf0] text-sm font-black text-[#9a6c25]">
                  <tr>
                    <th className="px-3 py-3">상품명</th>
                    <th className="border-l border-[#eee5d0] px-3 py-3">상품 구성</th>
                    <th className="border-l border-[#eee5d0] px-3 py-3 text-right">가격</th>
                  </tr>
                </thead>
                <tbody className="text-sm font-semibold text-[#28251f]">
                  {giftSetGuide.map((giftSet) => (
                    <tr className="border-t border-[#eee9df]" key={giftSet.name}>
                      <td className="px-3 py-3">{giftSet.name}</td>
                      <td className="border-l border-[#f0ebe1] px-3 py-3 leading-5">
                        {giftSet.composition}
                      </td>
                      <td className="border-l border-[#f0ebe1] px-3 py-3 text-right whitespace-nowrap">
                        {toCurrency(giftSet.price)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </section>

            <section className="grid gap-3">
              {activeProducts.map((product) => {
                const quantity = quantities[product.id] || 0;
                return (
                  <article
                    className="grid grid-cols-[1fr_auto] gap-3 rounded-lg border border-[#e1d7bd] bg-white p-4 shadow-sm"
                    key={product.id}
                  >
                    <div>
                      <h2 className="text-lg font-extrabold">{product.name}</h2>
                      <p className="mt-1 text-sm font-semibold text-[#6b5f45]">
                        {toCurrency(product.price)} / 1박스
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        aria-label={`${product.name} 수량 빼기`}
                        className="h-9 w-9 rounded-full border border-[#cabf9e] text-xl font-bold disabled:opacity-40"
                        disabled={isSubmitting}
                        onClick={() => updateQuantity(product.id, -1)}
                        type="button"
                      >
                        -
                      </button>
                      <span className="w-8 text-center text-lg font-black">{quantity}</span>
                      <button
                        aria-label={`${product.name} 수량 더하기`}
                        className="h-9 w-9 rounded-full bg-[#426b2f] text-xl font-bold text-white disabled:opacity-40"
                        disabled={isSubmitting}
                        onClick={() => updateQuantity(product.id, 1)}
                        type="button"
                      >
                        +
                      </button>
                    </div>
                  </article>
                );
              })}
            </section>

            <section className="rounded-lg bg-[#202016] p-4 text-white">
              <div className="flex items-center justify-between">
                <span className="text-sm font-semibold text-[#d8d0b9]">장바구니</span>
                <strong>{totalBoxes}박스</strong>
              </div>
              <div className="mt-3 grid gap-2">
                {cart.length === 0 ? (
                  <p className="text-sm text-[#d8d0b9]">담은 상품이 없습니다.</p>
                ) : (
                  cart.map((item) => (
                    <div className="flex justify-between text-sm" key={item.id}>
                      <span>
                        {item.name} x {item.quantity}
                      </span>
                      <span>{toCurrency(item.subtotal)}</span>
                    </div>
                  ))
                )}
              </div>
              <div className="mt-4 flex items-end justify-between border-t border-white/20 pt-3">
                <span className="font-semibold">총 결제금액</span>
                <strong className="text-2xl">{toCurrency(total)}</strong>
              </div>
            </section>

            <section className="grid gap-3 rounded-lg bg-white p-4">
              <label className="grid gap-1 text-sm font-bold">
                이름
                <input
                  autoComplete="street-address"
                  className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  disabled={isSubmitting}
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, name: event.target.value }))
                  }
                  placeholder="홍길동"
                  value={customer.name}
                />
              </label>
              <label className="grid gap-1 text-sm font-bold">
                전화번호
                <input
                  className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  disabled={isSubmitting}
                  inputMode="tel"
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, phone: event.target.value }))
                  }
                  placeholder="010-0000-0000"
                  value={customer.phone}
                />
              </label>
              <label className="grid gap-1 text-sm font-bold">
                받으실 주소 <span className="font-medium text-[#6d6a55]">(선택)</span>
                <input
                  className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  disabled={isSubmitting}
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, address: event.target.value }))
                  }
                  placeholder="주소를 입력해 주세요"
                  value={customer.address}
                />
              </label>
              <label className="grid gap-1 text-sm font-bold">
                입금자명 <span className="font-medium text-[#6d6a55]">(필수)</span>
                <input
                  autoComplete="name"
                  className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  disabled={isSubmitting}
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, payerName: event.target.value }))
                  }
                  placeholder="입금하실 분의 이름"
                  value={customer.payerName}
                />
              </label>
              <label className="grid gap-1 text-sm font-bold">
                요청사항
                <textarea
                  className="min-h-24 rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  disabled={isSubmitting}
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, note: event.target.value }))
                  }
                  placeholder="기타 요청사항이 있으면 적어주세요. 예시) 포도농장에 와서 직접 받겠습니다."
                  value={customer.note}
                />
              </label>
            </section>

            <section className="rounded-lg border border-[#d6ccb6] bg-[#fffaf0] p-4">
              <div className="flex items-start justify-between gap-3">
                <div className="grid gap-2">
                  <p className="text-sm font-semibold leading-6">{settings.bankNotice}</p>
                  <p className="text-sm text-[#6b5f45]">
                    입금 금액 <strong className="ml-2 text-base text-[#202016]">{toCurrency(total)}</strong>
                  </p>
                </div>
                <button
                  className="shrink-0 rounded-md bg-white px-3 py-2 text-sm font-black text-[#426b2f] shadow-sm"
                  onClick={copyBankNotice}
                  type="button"
                >
                  계좌·금액 복사
                </button>
              </div>
            </section>

            {status ? (
              <p className="rounded-md bg-white px-3 py-2 text-sm font-bold text-[#426b2f]">
                {status}
              </p>
            ) : null}

            <button
              className="sticky bottom-4 mt-auto rounded-lg bg-[#426b2f] px-4 py-4 text-lg font-black text-white shadow-lg disabled:cursor-not-allowed disabled:bg-[#8ba07f]"
              disabled={isSubmitting}
              type="submit"
            >
              {isSubmitting ? "전송 중..." : "주문서 보내기"}
            </button>
          </form>
        ) : (
          <section className="grid gap-4">
            <div className="rounded-lg bg-white p-4">
              <h2 className="text-xl font-black">관리자 설정</h2>
              <p className="mt-1 text-sm font-semibold text-[#6d6a55]">
                상품과 가격은 고정되어 있으며, 안내문과 계좌정보를 저장할 수 있습니다.
              </p>
            </div>

            <label className="grid gap-1 text-sm font-bold">
              상점 이름
              <input
                className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                onChange={(event) =>
                  setSettings((current) => ({ ...current, shopName: event.target.value }))
                }
                value={settings.shopName}
              />
            </label>

            <label className="grid gap-1 text-sm font-bold">
              페이지 상단 설명문
              <textarea
                className="min-h-64 rounded-md border border-[#d8cfba] px-3 py-3 text-base leading-6"
                onChange={(event) =>
                  setSettings((current) => ({ ...current, introText: event.target.value }))
                }
                placeholder="상품과 구매 안내를 입력해 주세요."
                value={settings.introText}
              />
            </label>

            <label className="grid gap-1 text-sm font-bold">
              계좌이체 안내 문구
              <textarea
                className="min-h-20 rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                onChange={(event) =>
                  setSettings((current) => ({
                    ...current,
                    bankNotice: event.target.value,
                  }))
                }
                value={settings.bankNotice}
              />
            </label>

            <label className="grid gap-1 text-sm font-bold">
              Google Apps Script 웹앱 URL
              <input
                className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                onChange={(event) =>
                  setSettings((current) => ({
                    ...current,
                    sheetEndpoint: event.target.value,
                  }))
                }
                placeholder="https://script.google.com/macros/s/..."
                value={settings.sheetEndpoint}
              />
            </label>

            <button
              className="rounded-lg bg-[#202016] px-4 py-4 text-lg font-black text-white"
              onClick={saveAdminSettings}
              type="button"
            >
              설정 저장하고 공개 링크 만들기
            </button>

            {shareLink ? (
              <div className="grid gap-2 rounded-lg bg-white p-4">
                <strong>공개 링크</strong>
                <textarea
                  className="min-h-28 rounded-md border border-[#d8cfba] p-3 text-xs"
                  readOnly
                  value={shareLink}
                />
              </div>
            ) : null}

            <details className="rounded-lg border border-[#d6ccb6] bg-[#fffaf0] p-4 text-sm leading-6">
              <summary className="cursor-pointer font-black">구글시트 연결 코드</summary>
              <p className="mt-3">
                기존 코드를 아래 전체 코드로 교체하면 주문서와 주문상품 시트에 동시에
                저장됩니다. 저장한 뒤 웹앱 배포를 새 버전으로 업데이트하세요.
              </p>
              <pre className="mt-3 overflow-auto rounded-md bg-[#202016] p-3 text-xs text-white">
                {appsScriptCode}
              </pre>
            </details>

            {status ? (
              <p className="rounded-md bg-white px-3 py-2 text-sm font-bold text-[#426b2f]">
                {status}
              </p>
            ) : null}
          </section>
        )}

        {isPasswordOpen ? (
          <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/40 px-4">
            <form
              className="w-full max-w-sm rounded-lg bg-white p-5 shadow-xl"
              onSubmit={submitAdminPassword}
            >
              <h2 className="text-xl font-black">관리자 비밀번호</h2>
              <p className="mt-1 text-sm font-semibold text-[#6d6a55]">
                4자리 비밀번호를 입력해 주세요.
              </p>
              <input
                autoFocus
                className="mt-4 w-full rounded-md border border-[#d8cfba] px-3 py-3 text-center text-2xl font-black tracking-[0.3em]"
                inputMode="numeric"
                maxLength={4}
                onChange={(event) => {
                  setAdminPassword(event.target.value.replace(/\D/g, "").slice(0, 4));
                  setPasswordError("");
                }}
                type="password"
                value={adminPassword}
              />
              {passwordError ? (
                <p className="mt-2 text-sm font-bold text-[#b33a2b]">{passwordError}</p>
              ) : null}
              <div className="mt-4 grid grid-cols-2 gap-2">
                <button
                  className="rounded-lg border border-[#d6ccb6] px-4 py-3 font-black"
                  onClick={() => setIsPasswordOpen(false)}
                  type="button"
                >
                  취소
                </button>
                <button
                  className="rounded-lg bg-[#202016] px-4 py-3 font-black text-white"
                  type="submit"
                >
                  확인
                </button>
              </div>
            </form>
          </div>
        ) : null}

        {toast ? (
          <div className="fixed left-1/2 top-5 z-50 -translate-x-1/2 rounded-full bg-[#202016] px-5 py-3 text-sm font-black text-white shadow-xl">
            {toast}
          </div>
        ) : null}
      </section>
    </main>
  );
}
