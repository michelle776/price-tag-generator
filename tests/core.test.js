// 價格牌產生器自動測試：執行 `node --test "tests/*.test.js"`（不需安裝任何套件）
// 從 index.html 取出主程式，在 Node 的 vm 沙盒裡執行，透過 window.__PTG_TEST__ 取得內部函式。
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const HTML = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const APP_SRC = (() => {
  const blocks = [...HTML.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  assert.equal(blocks.length, 1, "index.html 應只有一段 inline <script>");
  return blocks[0];
})();
const XLSX_SRC = fs.readFileSync(path.join(ROOT, "vendor", "xlsx.full.min.js"), "utf8");

// SheetJS 只載入一次（約 900KB），測試內用它產生 .xlsx 檔
const xlsxCtx = vm.createContext({});
vm.runInContext(XLSX_SRC, xlsxCtx);
const XLSX = xlsxCtx.XLSX;

// 每次呼叫都給一個全新的 app 與假 localStorage，避免測試互相影響
function loadApp(saved) {
  const store = new Map();
  const localStorage = {
    getItem: k => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => store.set(k, String(v)),
  };
  const window = { __PTG_TEST__: true };
  const ctx = vm.createContext({
    window, localStorage, XLSX,
    document: { addEventListener() {} },
  });
  if (saved !== undefined) {
    store.set("priceTagGeneratorState_v1", typeof saved === "string" ? saved : JSON.stringify(saved));
  }
  vm.runInContext(APP_SRC, ctx);
  const api = window.__PTG_TEST__;
  assert.equal(typeof api, "object", "index.html 的 __PTG_TEST__ 測試 hook 不見了");
  return api;
}

// vm 沙盒裡的物件原型不同，比較前先轉成一般物件
const plain = o => JSON.parse(JSON.stringify(o));

function workbook(rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "S1");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}

/* ---------------- 開啟網頁時讀取、轉換舊存檔（loadDeck） ---------------- */

test("loadDeck：沒有存檔時給預設大巨蛋卡", () => {
  const app = loadApp();
  const d = app.loadDeck();
  assert.equal(d.cards.length, 1);
  assert.equal(d.cards[0].template, "dome");
  assert.equal(d.activeIndex, 0);
  assert.equal(d.gap, 0.2);
});

test("loadDeck：存檔損壞時不當掉，回到預設", () => {
  const d = loadApp("{這不是 JSON").loadDeck();
  assert.equal(d.cards[0].template, "dome");
});

test("loadDeck：空清單視為沒有存檔", () => {
  const d = loadApp({ cards: [], activeIndex: 0 }).loadDeck();
  assert.equal(d.cards.length, 1);
  assert.equal(d.cards[0].template, "dome");
});

test("loadDeck：保留卡片、間距，選取索引超出範圍時夾到最後一張", () => {
  const app = loadApp();
  const cards = [plain(app.DEFAULTS.combo), plain(app.DEFAULTS.specs)];
  const d = loadApp({ cards, activeIndex: 9, gap: 0.5 }).loadDeck();
  assert.equal(d.cards.length, 2);
  assert.equal(d.activeIndex, 1);
  assert.equal(d.gap, 0.5);
  assert.equal(loadApp({ cards, activeIndex: -3 }).loadDeck().activeIndex, 0);
  assert.equal(loadApp({ cards }).loadDeck().gap, 0.2);
});

test("loadDeck：舊版單張卡片格式包成清單並套用轉換", () => {
  const app = loadApp();
  const old = plain(app.DEFAULTS.taimall);
  old.size = { w: 21, h: 29.7 };
  const d = loadApp(old).loadDeck();
  assert.equal(d.cards.length, 1);
  assert.deepEqual(plain(d.cards[0].size), { w: 10.5, h: 14.85 });
});

test("loadDeck：大巨蛋 11×16 舊版改成 A4；其他尺寸不動", () => {
  const app = loadApp();
  const a = Object.assign(plain(app.DEFAULTS.dome), { domeBg: "11x16", size: { w: 11, h: 16 } });
  const b = Object.assign(plain(app.DEFAULTS.dome), { domeBg: "11x16", size: { w: 12, h: 8 } });
  const [ca, cb] = loadApp({ cards: [a, b] }).loadDeck().cards;
  assert.equal(ca.domeBg, "A4");
  assert.deepEqual(plain(ca.size), { w: 21, h: 29.7 });
  assert.equal(cb.domeBg, "A4");
  assert.deepEqual(plain(cb.size), { w: 12, h: 8 });
});

test("loadDeck：價目表舊價格欄位轉成價格行，字級微調跟著搬到第 0 行", () => {
  const app = loadApp();
  const card = Object.assign(plain(app.DEFAULTS.dome), {
    domeTableRows: [
      { name: "A", price: "10,000", priceRed: true, specialPrice: "8,000", specialRed: false },
      { name: "B", price: "5,000", specialPrice: "  " },
      { name: "C", prices: [{ label: "自訂", value: "1", unit: "", highlight: false }] },
    ],
    lineStyles: { "dome.table.0.price": { fontScale: 1.3 } },
  });
  const out = loadApp({ cards: [card] }).loadDeck().cards[0];
  const [r0, r1, r2] = plain(out.domeTableRows);
  assert.deepEqual(r0.prices, [
    { label: "售價", value: "10,000", unit: "", highlight: true },
    { label: "特價", value: "8,000", unit: "", highlight: false },
  ]);
  assert.equal("price" in r0 || "specialPrice" in r0 || "priceRed" in r0 || "specialRed" in r0, false);
  assert.deepEqual(r1.prices, [{ label: "售價", value: "5,000", unit: "", highlight: false }]);
  assert.deepEqual(r2.prices, [{ label: "自訂", value: "1", unit: "", highlight: false }]);
  assert.deepEqual(plain(out.lineStyles["dome.table.0.price.0"]), { fontScale: 1.3 });
  assert.equal(out.lineStyles["dome.table.0.price"], undefined);
});

test("loadDeck：台茂卡 A4 改回 10.5×14.85，補齊預設字級但不覆蓋使用者調過的", () => {
  const app = loadApp();
  const defaults = plain(app.DEFAULTS.taimall.lineStyles);
  const keys = Object.keys(defaults);
  assert.ok(keys.length >= 2, "台茂預設字級應至少有兩項");
  const custom = { fontScale: 9.9, custom: true };
  const card = Object.assign(plain(app.DEFAULTS.taimall), {
    size: { w: 21, h: 29.7 },
    lineStyles: { [keys[0]]: custom },
  });
  const out = plain(loadApp({ cards: [card] }).loadDeck().cards[0]);
  assert.deepEqual(out.size, { w: 10.5, h: 14.85 });
  assert.deepEqual(out.lineStyles[keys[0]], custom);
  for (const k of keys.slice(1)) assert.deepEqual(out.lineStyles[k], defaults[k]);

  // 其他尺寸的台茂卡不改尺寸；沒有 lineStyles 也能補上
  const small = Object.assign(plain(app.DEFAULTS.taimall), { size: { w: 7, h: 4 } });
  delete small.lineStyles;
  const out2 = plain(loadApp({ cards: [small] }).loadDeck().cards[0]);
  assert.deepEqual(out2.size, { w: 7, h: 4 });
  assert.deepEqual(out2.lineStyles, defaults);
});

/* ---------------- 台茂 Excel 批次匯入 ---------------- */

test("matchBrandId：關鍵字、品牌名、包含關鍵字、對不到", () => {
  const app = loadApp();
  assert.equal(app.matchBrandId("LG"), "lg-color");
  assert.equal(app.matchBrandId(" 樂金 "), "lg-color");
  assert.equal(app.matchBrandId("SONY"), "sony");
  assert.equal(app.matchBrandId("Audio-Technica"), "audiotechnica");
  assert.equal(app.matchBrandId("B&O"), "bo");
  assert.equal(app.matchBrandId("三星電子"), "samsung");
  assert.equal(app.matchBrandId("完全沒有這個牌子"), "");
  assert.equal(app.matchBrandId(""), "");
  assert.equal(app.matchBrandId(null), "");
});

test("splitTaimallIntro：換行、、；; 分隔，去掉項目符號與空白", () => {
  const app = loadApp();
  assert.deepEqual(plain(app.splitTaimallIntro("A\r\nB\n・C、D；E;F\n\n  - G ")),
    ["A", "B", "C", "D", "E", "F", "G"]);
  assert.deepEqual(plain(app.splitTaimallIntro(null)), []);
});

test("cleanTaimallPrice：去掉 NT$ / $ / 元 / 台 / 限量", () => {
  const app = loadApp();
  assert.equal(app.cleanTaimallPrice("NT$36,800元"), "36,800");
  assert.equal(app.cleanTaimallPrice("nt 1,000"), "1,000");
  assert.equal(app.cleanTaimallPrice("$990"), "990");
  assert.equal(app.cleanTaimallPrice("限量5台"), "5");
  assert.equal(app.cleanTaimallPrice(29900), "29900");
  assert.equal(app.cleanTaimallPrice(null), "");
});

test("parseTaimallWorkbook：略過標題列與空白列，欄位正確對應", () => {
  const app = loadApp();
  const buf = workbook([
    ["品牌", "商品名稱", "型號", "商品介紹", "原價", "新品特價", "限量"],
    ["LG", " 空氣清淨機 ", "AS101", "360°淨化\n寵物模式、多重過濾", "NT$36,800元", "29,900", "限量10台"],
    ["", "", "", "", "", "", ""],
    ["Sony", "耳機", "WH-1000XM6", "", 12990, "", ""],
    ["只有品牌", "", "", "", "", "", ""],
  ]);
  const cards = plain(app.parseTaimallWorkbook(buf));
  assert.equal(cards.length, 3);
  const [a, b, c] = cards;
  assert.equal(a.template, "taimall");
  assert.equal(a.taimallBrand, "lg-color");
  assert.equal(a.taimallName, "空氣清淨機");
  assert.equal(a.taimallModel, "AS101");
  assert.deepEqual(a.taimallIntro, ["360°淨化", "寵物模式", "多重過濾"]);
  assert.equal(a.taimallOrigPrice, "36,800");
  assert.equal(a.taimallSpecialPrice, "29,900");
  assert.equal(a.taimallLimit, "10");
  assert.deepEqual(a.lineStyles, plain(app.DEFAULTS.taimall.lineStyles));
  assert.deepEqual(a.size, { w: 10.5, h: 14.85 });
  assert.equal(b.taimallBrand, "sony");
  assert.equal(b.taimallOrigPrice, "12990");
  assert.deepEqual(b.taimallIntro, []);
  assert.equal(c.taimallBrand, "");
  assert.equal(c.taimallName, "");
});

test("parseTaimallWorkbook：第一列是資料（沒有標題列）時不會被略過", () => {
  const app = loadApp();
  const cards = app.parseTaimallWorkbook(workbook([["Bose", "喇叭", "S1", "", "1,000", "", ""]]));
  assert.equal(cards.length, 1);
  assert.equal(cards[0].taimallName, "喇叭");
});

test("parseTaimallWorkbook：網頁下載的匯入範本，匯回來剛好一張卡", () => {
  const app = loadApp();
  // 與 downloadTaimallTemplate() 相同的標題列與範例列
  const headers = ["品牌", "商品名稱", "型號", "商品介紹（同一格 Alt+Enter 換行）", "原價", "新品特價", "限量（選填）"];
  for (const h of headers) {
    assert.ok(APP_SRC.includes(JSON.stringify(h)), "downloadTaimallTemplate 的標題「" + h + "」已變更，請同步更新此測試");
  }
  const example = ["LG", "PuriCare™ 360°空氣清淨機", "AS101DBY0", "CASR值:12.34\n360°強力淨化", "36,800", "29,900", ""];
  const cards = plain(app.parseTaimallWorkbook(workbook([headers, example])));
  assert.equal(cards.length, 1);
  assert.equal(cards[0].taimallModel, "AS101DBY0");
  assert.deepEqual(cards[0].taimallIntro, ["CASR值:12.34", "360°強力淨化"]);
});

/* ---------------- 列印排版（buildPages）：畫面預覽＝實際列印 ---------------- */

function cardOf(w, h, copies) {
  return { template: "compare", size: { w, h }, copies: copies || 1 };
}

test("a4TileGrid：能整除 A4 的尺寸才走無邊界拼版", () => {
  const app = loadApp();
  assert.deepEqual(plain(app.a4TileGrid(cardOf(10.5, 14.85))), { cols: 2, rows: 2, cap: 4 });
  assert.equal(app.a4TileGrid(cardOf(21, 29.7)), null); // 單張滿版
  assert.equal(app.a4TileGrid(cardOf(7, 4)), null);
  assert.equal(app.a4TileGrid(cardOf(0, 4)), null);
});

test("buildPages：台茂 4 張排一頁（2×2），第 5 張換新頁；份數會展開", () => {
  const app = loadApp();
  const t = plain(app.DEFAULTS.taimall);
  app.setDeck([Object.assign({}, t, { copies: 3 }), Object.assign({}, t, { copies: 2 })]);
  const pages = plain(app.buildPages());
  assert.deepEqual(pages.map(p => [p.tile, p.cols, p.items.length]), [[true, 2, 4], [true, 2, 1]]);
  assert.deepEqual(pages[0].items.map(i => i.cardIndex), [0, 0, 0, 1]);
});

test("buildPages：一般尺寸依可用寬高換行換頁（7×4 每頁 2 欄 × 6 列）", () => {
  const app = loadApp();
  app.setDeck([cardOf(7, 4, 13)], 0.2);
  const pages = plain(app.buildPages());
  assert.deepEqual(pages.map(p => [p.tile, p.items.length]), [[false, 12], [false, 1]]);
});

test("buildPages：一般卡與拼版卡混排時分開頁面，順序保留", () => {
  const app = loadApp();
  const t = plain(app.DEFAULTS.taimall);
  app.setDeck([cardOf(7, 4), t, cardOf(7, 4)]);
  const pages = plain(app.buildPages());
  assert.deepEqual(pages.map(p => [p.tile, p.items.map(i => i.cardIndex)]),
    [[false, [0]], [true, [1]], [false, [2]]]);
});

test("buildPages：份數限制在 1–200，空清單仍有一頁", () => {
  const app = loadApp();
  app.setDeck([cardOf(7, 4, 0), cardOf(7, 4, "abc")]);
  assert.equal(app.buildPages()[0].items.length, 2);
  app.setDeck([cardOf(1, 1, 999)]);
  assert.equal(app.buildPages().reduce((n, p) => n + p.items.length, 0), 200);
  app.setDeck([]);
  assert.equal(app.buildPages().length, 1);
});

/* ---------------- 天母SOGO 三種版面 ---------------- */

test("天母SOGO：組合卡 15×10、規格價目表 10.8×15、小張電視價格表 10×7", () => {
  const app = loadApp();
  assert.deepEqual(plain(app.DEFAULTS.combo.size), { w: 15, h: 10 });
  assert.deepEqual(plain(app.DEFAULTS.specs.size), { w: 10.8, h: 15 });
  const tv = plain(app.DEFAULTS.sogotv);
  assert.equal(tv.template, "sogotv");
  assert.deepEqual(tv.size, { w: 10, h: 7 });
});

test("loadDeck：舊的 11×15 規格表存檔尺寸不被改動", () => {
  const app = loadApp();
  const old = Object.assign(plain(app.DEFAULTS.specs), { size: { w: 11, h: 15 } });
  assert.deepEqual(plain(loadApp({ cards: [old] }).loadDeck().cards[0].size), { w: 11, h: 15 });
});

test("buildPages：10×7 電視卡縮小頁邊置中，一頁 2 欄 × 4 列＝8 張", () => {
  const app = loadApp();
  const tv = plain(app.DEFAULTS.sogotv);
  assert.equal(app.a4TileGrid(tv), null);
  assert.deepEqual(plain(app.compactGrid(tv, 0.2)), { cols: 2, rows: 4, cap: 8, x: 0.4, y: 0.55 });
  app.setDeck([Object.assign({}, tv, { copies: 9 }), cardOf(7, 4), Object.assign({}, tv, { copies: 2 })], 0.2);
  const pages = plain(app.buildPages());
  assert.deepEqual(pages.map(p => [!!p.compact, p.items.map(i => i.cardIndex).length]),
    [[true, 8], [true, 1], [false, 1], [true, 2]]);
  assert.deepEqual(pages[0].inset, { x: 0.4, y: 0.55 });
});

test("compactGrid：間距調大放不下兩欄時自動變一欄；其他版型不套用", () => {
  const app = loadApp();
  const tv = plain(app.DEFAULTS.sogotv);
  assert.deepEqual(plain(app.compactGrid(tv, 0.5)).cols, 1);
  assert.equal(app.compactGrid(plain(app.DEFAULTS.combo), 0.2), null);
});

test("addCardsToFillPage：10×7 補滿 A4 是新增 7 張可各自編輯的卡，而非重複份數", () => {
  const app = loadApp();
  const tv = Object.assign(plain(app.DEFAULTS.sogotv), { copies: 5 });
  app.setDeck([cardOf(7, 4), tv, cardOf(7, 4)], 0.2);
  assert.equal(app.addCardsToFillPage(1), 7);
  const cards = app.getDeck().cards;
  assert.equal(cards.length, 10);
  assert.deepEqual(cards.slice(1, 9).map(c => [c.template, c.copies]), Array(8).fill(["sogotv", 1]));
  assert.notEqual(cards[1], cards[2], "新增的應是獨立物件，改一張不影響另一張");
  assert.equal(cards[9].template, "compare");
  const pages = plain(app.buildPages());
  assert.deepEqual(pages.map(p => p.items.length), [1, 8, 1]);
  // 已經滿了再按不會多加
  assert.equal(app.addCardsToFillPage(1), 0);
});

/* ---------------- 儲存清單成檔案 / 開啟清單檔案 ---------------- */

test("儲存清單：下載的檔案內容再開啟，清單、選取、間距完整還原", () => {
  const app = loadApp();
  const cards = [plain(app.DEFAULTS.sogotv), Object.assign(plain(app.DEFAULTS.taimall), { copies: 3 })];
  cards[0].sogotvSpecialPrice = "159,900";
  cards[0].sogotvBrandUpload = "data:image/png;base64,AAAA"; // 上傳的 Logo 也要一起存
  app.setDeck(cards, 0.6);
  const file = JSON.parse(JSON.stringify(app.deckFileData()));
  assert.equal(file.app, "price-tag-generator");
  const d = plain(app.normalizeDeck(file));
  assert.deepEqual(d.cards, cards);
  assert.equal(d.gap, 0.6);
});

test("開啟清單：不是清單檔案、或卡片版型看不懂的，回傳 null 不取代目前清單", () => {
  const app = loadApp();
  assert.equal(app.normalizeDeck(null), null);
  assert.equal(app.normalizeDeck({ foo: 1 }), null);
  assert.equal(app.normalizeDeck({ cards: [] }), null);
  assert.equal(app.normalizeDeck({ cards: [{ template: "不存在的版型", size: { w: 1, h: 1 } }] }), null);
  // 混有壞卡片時只保留看得懂的
  const ok = plain(app.DEFAULTS.combo);
  assert.equal(app.normalizeDeck({ cards: [ok, { template: "x" }, null] }).cards.length, 1);
});

test("開啟舊檔案時同樣套用舊資料轉換（例如台茂 A4 → 10.5×14.85）", () => {
  const app = loadApp();
  const old = Object.assign(plain(app.DEFAULTS.taimall), { size: { w: 21, h: 29.7 } });
  assert.deepEqual(plain(app.normalizeDeck({ cards: [old] }).cards[0].size), { w: 10.5, h: 14.85 });
});

test("儲存檔名含日期時間", () => {
  const app = loadApp();
  assert.equal(app.deckFileName(new Date(2026, 8, 24, 9, 5)), "價格牌清單_20260924-0905.json");
});
