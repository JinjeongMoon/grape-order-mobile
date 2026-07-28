"use client";

import { useEffect, useMemo, useState } from "react";

type Product = {
  id: string;
  name: string;
  price: number;
};

type Settings = {
  shopName: string;
  products: Product[];
  bankNotice: string;
  sheetEndpoint: string;
};

type Customer = {
  name: string;
  phone: string;
  note: string;
};

const emptyProducts: Product[] = Array.from({ length: 5 }, (_, index) => ({
  id: `product-${index + 1}`,
  name: "",
  price: 0,
}));

const defaultSettings: Settings = {
  shopName: "그랑포도",
  products: [
    { id: "product-1", name: "샤인머스캣 1박스", price: 35000 },
    { id: "product-2", name: "캠벨포도 1박스", price: 25000 },
    { id: "product-3", name: "거봉 1박스", price: 30000 },
    { id: "product-4", name: "혼합 선물세트", price: 45000 },
    { id: "product-5", name: "프리미엄 포도박스", price: 55000 },
  ],
  bankNotice: "계좌이체: 은행명 000-0000-0000 예금주 홍길동",
  sheetEndpoint: "",
};

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
      products: normalizeProducts(parsed.products),
      bankNotice: parsed.bankNotice || defaultSettings.bankNotice,
      sheetEndpoint: parsed.sheetEndpoint || "",
    };
  } catch {
    return null;
  }
}

function normalizeProducts(products: Product[] = emptyProducts) {
  return Array.from({ length: 5 }, (_, index) => {
    const item = products[index] ?? emptyProducts[index];
    return {
      id: item.id || `product-${index + 1}`,
      name: item.name || "",
      price: Number(item.price) || 0,
    };
  });
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
    note: "",
  });
  const [mode, setMode] = useState<"order" | "admin">("order");
  const [status, setStatus] = useState("");
  const [toast, setToast] = useState("");
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

  function updateProduct(index: number, field: "name" | "price", value: string) {
    setSettings((current) => ({
      ...current,
      products: current.products.map((product, productIndex) =>
        productIndex === index
          ? {
              ...product,
              [field]: field === "price" ? Number(value) || 0 : value,
            }
          : product,
      ),
    }));
  }

  function saveAdminSettings() {
    const normalized = {
      ...settings,
      products: normalizeProducts(settings.products),
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
      await navigator.clipboard.writeText(bankInfo || settings.bankNotice);
      setToast("복사되었습니다");
    } catch {
      setToast("복사에 실패했습니다");
    }
  }

  async function submitOrder(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setStatus("");

    if (!customer.name.trim() || !customer.phone.trim()) {
      setStatus("이름과 전화번호를 입력해 주세요.");
      return;
    }

    if (cart.length === 0) {
      setStatus("상품을 1박스 이상 담아 주세요.");
      return;
    }

    if (!settings.sheetEndpoint.trim()) {
      setStatus("주문 내용이 준비됐어요. 관리자 화면에서 구글시트 주소를 연결해 주세요.");
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
      setCustomer({ name: "", phone: "", note: "" });
      setQuantities(
        Object.fromEntries(settings.products.map((product) => [product.id, 0])),
      );
      setStatus("주문이 접수됐어요. 계좌이체 후 확인 문자를 기다려 주세요.");
    } catch {
      setStatus("주문 전송에 실패했어요. 잠시 후 다시 시도해 주세요.");
    }
  }

  return (
    <main className="min-h-screen bg-[#f6f3eb] text-[#202016]">
      <section className="mx-auto flex min-h-screen w-full max-w-md flex-col px-4 py-5">
        <header className="flex items-center justify-between pb-4">
          <div>
            <p className="text-sm font-semibold text-[#6d6a55]">1박스 단위 주문</p>
            <h1 className="text-3xl font-black tracking-normal">{settings.shopName}</h1>
          </div>
          <button
            className="rounded-full border border-[#d6ccb6] bg-white px-3 py-2 text-sm font-bold"
            onClick={handleModeButton}
            type="button"
          >
            {mode === "order" ? "관리" : "주문"}
          </button>
        </header>

        {mode === "order" ? (
          <form className="flex flex-1 flex-col gap-4" onSubmit={submitOrder}>
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
                        className="h-9 w-9 rounded-full border border-[#cabf9e] text-xl font-bold"
                        onClick={() => updateQuantity(product.id, -1)}
                        type="button"
                      >
                        -
                      </button>
                      <span className="w-8 text-center text-lg font-black">{quantity}</span>
                      <button
                        aria-label={`${product.name} 수량 더하기`}
                        className="h-9 w-9 rounded-full bg-[#426b2f] text-xl font-bold text-white"
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
                  className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
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
                  inputMode="tel"
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, phone: event.target.value }))
                  }
                  placeholder="010-0000-0000"
                  value={customer.phone}
                />
              </label>
              <label className="grid gap-1 text-sm font-bold">
                요청사항
                <textarea
                  className="min-h-24 rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                  onChange={(event) =>
                    setCustomer((current) => ({ ...current, note: event.target.value }))
                  }
                  placeholder="배송 요청, 선물 포장 등"
                  value={customer.note}
                />
              </label>
            </section>

            <section className="rounded-lg border border-[#d6ccb6] bg-[#fffaf0] p-4">
              <div className="flex items-start justify-between gap-3">
                <p className="text-sm font-semibold leading-6">{settings.bankNotice}</p>
                <button
                  className="shrink-0 rounded-md bg-white px-3 py-2 text-sm font-black text-[#426b2f] shadow-sm"
                  onClick={copyBankNotice}
                  type="button"
                >
                  복사
                </button>
              </div>
            </section>

            {status ? (
              <p className="rounded-md bg-white px-3 py-2 text-sm font-bold text-[#426b2f]">
                {status}
              </p>
            ) : null}

            <button
              className="sticky bottom-4 mt-auto rounded-lg bg-[#426b2f] px-4 py-4 text-lg font-black text-white shadow-lg"
              type="submit"
            >
              주문서 보내기
            </button>
          </form>
        ) : (
          <section className="grid gap-4">
            <div className="rounded-lg bg-white p-4">
              <h2 className="text-xl font-black">관리자 설정</h2>
              <p className="mt-1 text-sm font-semibold text-[#6d6a55]">
                상품과 가격을 저장한 뒤 공개 링크를 손님에게 보내세요.
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

            <section className="grid gap-3">
              {settings.products.map((product, index) => (
                <div className="grid gap-2 rounded-lg bg-white p-4" key={product.id}>
                  <strong>상품 {index + 1}</strong>
                  <input
                    className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                    onChange={(event) => updateProduct(index, "name", event.target.value)}
                    placeholder="상품명"
                    value={product.name}
                  />
                  <input
                    className="rounded-md border border-[#d8cfba] px-3 py-3 text-base"
                    inputMode="numeric"
                    onChange={(event) => updateProduct(index, "price", event.target.value)}
                    placeholder="가격"
                    type="number"
                    value={product.price || ""}
                  />
                </div>
              ))}
            </section>

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
                구글시트에서 확장 프로그램, Apps Script를 열고 아래 코드를 붙여 넣은 뒤
                웹앱으로 배포하세요.
              </p>
              <pre className="mt-3 overflow-auto rounded-md bg-[#202016] p-3 text-xs text-white">
                {`function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);
  const items = data.items.map(item =>
    item.name + " x " + item.quantity + "박스"
  ).join(", ");

  sheet.appendRow([
    new Date(),
    data.customer.name,
    data.customer.phone,
    data.customer.note,
    items,
    data.totalBoxes,
    data.total
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}`}
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
