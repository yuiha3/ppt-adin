const COLORS = {
  BLUE:  "#0070c0",
  GREEN: "#00b050",
  BLACK: "#000000"
};

const UI = {
  result:     "result",
  copyButton: "copyResultButton"
};

const DEFECT_GROUPS = [
  {
    label: "タイル面",
    items: [
      { content: "割れ",       unit: "枚数" },
      { content: "浮き",       unit: "枚数" },
      { content: "陶片浮き",   unit: "枚数" },
      { content: "下地浮き",   unit: "枚数" },
      { content: "欠損",       unit: "枚数" },
      { content: "剝落",       unit: "枚数" },
    ]
  },
  {
    label: "塗装面",
    items: [
      { content: "ひび割れ 0.3mm未満", unit: "長さ（m）"  },
      { content: "ひび割れ 0.3mm以上", unit: "長さ（m）"  },
      { content: "亀甲割れ",           unit: "面積（㎡）" },
      { content: "塗膜浮き",           unit: "面積（㎡）" },
      { content: "塗膜剥離",           unit: "面積（㎡）" },
      { content: "モルタル浮き",       unit: "面積（㎡）" },
    ]
  },
  {
    label: "その他",
    items: [
      { content: "エフロ",     unit: "箇所" },
      { content: "爆裂",       unit: "箇所" },
      { content: "サビ",       unit: "箇所" },
      { content: "目地欠損",   unit: "箇所" },
      { content: "シール劣化", unit: "箇所" },
      { content: "シール状況", unit: "箇所" },
      { content: "その他",     unit: "",    noUnit: true },
    ]
  },
  {
    label: "写真番号",
    items: [
      { content: "写真番号", unit: "", noUnit: true, pinBottom: true },
    ]
  },
];

Office.onReady(() => {
  initializeTabs();
  initializeTableBuilder();
  initializeSummaryButtons();
});

// ─── 初期化 ──────────────────────────────────────────────

function initializeSummaryButtons() {
  bindClick("sumBlueButton",  () => sumNumbersByTextColor(COLORS.BLUE,  "青文字 RGB(0,112,192)"));
  bindClick("sumGreenButton", () => sumNumbersByTextColor(COLORS.GREEN, "緑文字 RGB(0,176,80)"));
  bindClick("sumSelectedColorButton",           sumNumbersBySelectedTextColor);
  bindClick("sumSelectedTextAndFillColorButton", sumNumbersBySelectedTextColorAndFillColor);
  bindClick("formatBlackCodeButton",            () => formatCodesByTextColor(COLORS.BLACK, false));
  bindClick("formatAllSlidesBlackCodeButton",   () => formatCodesByTextColor(COLORS.BLACK, true));

  const copyBtn = document.getElementById(UI.copyButton);
  if (!copyBtn) return;

  copyBtn.disabled = true;
  copyBtn.addEventListener("click", async () => {
    const value = copyBtn.dataset.copyValue;
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      const original = copyBtn.textContent;
      copyBtn.textContent = "✓ コピー済";
      setTimeout(() => { copyBtn.textContent = original; }, 1500);
    } catch {
      copyBtn.textContent = "失敗";
      setTimeout(() => { copyBtn.textContent = "コピー"; }, 1500);
    }
  });
}

function initializeTabs() {
  document.querySelectorAll(".tabButton").forEach((btn) => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tabTarget));
  });
}

function activateTab(tabId) {
  if (!tabId) return;
  document.querySelectorAll(".tabButton").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tabTarget === tabId);
  });
  document.querySelectorAll(".tabContent").forEach((el) => {
    el.classList.toggle("active", el.id === tabId);
  });
}

function initializeTableBuilder() {
  renderDefectButtons();
  bindClick("outputTableButton",    outputTableToSlide);
  bindClick("clearTableRowsButton", clearTableRows);
}

function renderDefectButtons() {
  const container = document.getElementById("defectButtonGrid");
  if (!container) return;

  DEFECT_GROUPS.forEach(({ label, items }) => {
    const groupLabel = document.createElement("p");
    groupLabel.className = "defect-group-label";
    groupLabel.textContent = label;
    container.appendChild(groupLabel);

    items.forEach((item) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "defectButton";
      button.textContent = item.content;
      button.addEventListener("click", () => addTableRow(item));
      container.appendChild(button);
    });
  });

  shrinkOverflowButtons(container);
  new ResizeObserver(() => shrinkOverflowButtons(container)).observe(container);
}

function shrinkOverflowButtons(container) {
  container.querySelectorAll(".defectButton").forEach((btn) => {
    btn.style.fontSize = "";
    if (btn.scrollHeight > btn.clientHeight + 2) btn.style.fontSize = "10px";
  });
}

function addTableRow({ content, unit, noUnit = false, pinBottom = false }) {
  const tableRows = document.getElementById("tableRows");
  if (!tableRows) return;

  const row = document.createElement("div");
  row.className = noUnit ? "tableInputRow noUnitRow" : "tableInputRow";
  row.draggable = true;
  if (pinBottom) row.dataset.pinBottom = "true";

  const handle = document.createElement("span");
  handle.className = "drag-handle";
  handle.textContent = "⠿";
  handle.setAttribute("aria-hidden", "true");

  const quantityEl = noUnit
    ? Object.assign(document.createElement("textarea"), {
        className:   "tableInput quantityInput quantityTextarea",
        placeholder: "入力",
        rows:        2
      })
    : createTableInput({ className: "quantityInput", placeholder: "数量", inputmode: "numeric" });

  const deleteButton = Object.assign(document.createElement("button"), {
    type:        "button",
    className:   "row-delete-btn",
    textContent: "×"
  });
  deleteButton.addEventListener("click", () => row.remove());

  row.append(
    handle,
    createTableInput({ value: content, className: "contentInput" }),
    quantityEl,
    ...(noUnit ? [] : [createTableInput({ value: unit, className: "unitInput" })]),
    deleteButton
  );

  const firstPinned = tableRows.querySelector("[data-pin-bottom]");
  tableRows.insertBefore(row, firstPinned ?? null);

  setupDragAndDrop(row, tableRows);
  quantityEl.focus();
}

// ─── ドラッグ&ドロップ ────────────────────────────────────
let dragSrc    = null;
let dragOverEl = null;  // 前回ハイライトした要素を記憶し querySelectorAll を省く

function setupDragAndDrop(row, tableRows) {
  row.addEventListener("dragstart", (e) => {
    dragSrc = row;
    row.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  });

  row.addEventListener("dragend", () => {
    row.classList.remove("dragging");
    dragOverEl?.classList.remove("drag-over");
    dragOverEl = null;
    dragSrc    = null;
    enforcePinnedRows(tableRows);
  });

  row.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (!dragSrc || dragSrc === row) return;
    if (dragOverEl !== row) {
      dragOverEl?.classList.remove("drag-over");
      row.classList.add("drag-over");
      dragOverEl = row;
    }
  });

  row.addEventListener("drop", (e) => {
    e.preventDefault();
    if (!dragSrc || dragSrc === row) return;
    const children = [...tableRows.children];
    const after = children.indexOf(dragSrc) < children.indexOf(row);
    tableRows.insertBefore(dragSrc, after ? row.nextSibling : row);
  });
}

function enforcePinnedRows(tableRows) {
  tableRows.querySelectorAll("[data-pin-bottom]").forEach((el) => tableRows.appendChild(el));
}

function createTableInput({ value = "", className = "", placeholder = "", inputmode = "" }) {
  const input = document.createElement("input");
  input.type = "text";
  input.className = ["tableInput", className].filter(Boolean).join(" ");
  input.value = value;
  input.placeholder = placeholder;
  if (inputmode) input.inputMode = inputmode;
  return input;
}

function clearTableRows() {
  document.getElementById("tableRows")?.replaceChildren();
}

// ─── 表出力 ──────────────────────────────────────────────

/**
 * tableRows の内容を読み取り、現在のスライドに表を出力する。
 *
 * 表仕様（VBAコード準拠）
 *   列：内容 80pt / 数量 60pt / 単位 105pt
 *   1行目：ヘッダー（下罫線のみ）
 *   2行目以降：全罫線 黒 1pt
 *   全セル：背景白、余白 上下0 左5pt、垂直中央、フォント9pt 黒 太字 メイリオ
 *   数量列(2列目)：左余白0、横中央揃え
 *   行高さ：13.5pt
 */
async function outputTableToSlide() {
  const rows = collectTableRows();

  if (rows.length === 0) {
    setResult({ text: "行がありません。" });
    return;
  }

  setResult({ text: "出力中..." });

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const rowCount  = rows.length + 1;
    const headers   = ["内容", "数量", "単位"];
    const colWidths = [80, 60, 105]; // pt
    const rowHeight = 13.5;          // pt

    const solidBorder = { color: "000000", dashStyle: "solid", weight: 1 };
    const noBorder    = { color: "000000", dashStyle: "solid", weight: 0 };

    // セルプロパティを行列で組み立てる
    const specificCellProperties = Array.from({ length: rowCount }, (_, r) => {
      const isHeader = r === 0;
      const rowData  = isHeader ? null : rows[r - 1];
      const isNoUnit = !isHeader && rowData.noUnit;

      return Array.from({ length: 3 }, (_, c) => {
        const isQty      = c === 1;
        const isUnitCell = c === 2;

        // noUnit行の単位セル（c=2）：結合のダミーセルなので空＋ボーダーなし
        if (isNoUnit && isUnitCell) {
          return {
            text: "",
            fill: { color: "FFFFFF" },
            font: { name: "Meiryo", size: 9, bold: true, color: "000000" },
            horizontalAlignment: "Left",
            verticalAlignment:   "MiddleCentered",
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            borders: { top: solidBorder, left: noBorder, right: solidBorder, bottom: solidBorder }
          };
        }

        const text = isHeader
          ? headers[c]
          : [rowData.content, rowData.quantity, rowData.unit][c];

        // noUnit行の数量セル（c=1）：columnSpan:2 で単位列まで結合、左揃え
        const colSpan = (isNoUnit && isQty) ? 2 : undefined;

        return {
          text: text ?? "",
          fill: { color: "FFFFFF" },
          font: { name: "Meiryo", size: 9, bold: true, color: "000000" },
          horizontalAlignment: "Left",
          verticalAlignment:   "MiddleCentered",
          margins: { top: 0, bottom: 0, left: 5, right: 0 },
          borders: isHeader
            ? { top: noBorder, left: noBorder, right: noBorder, bottom: solidBorder }
            : { top: solidBorder, left: solidBorder, right: solidBorder, bottom: solidBorder },
          ...(colSpan !== undefined && { columnSpan: colSpan })
        };
      });
    });

    const tableShape = slide.shapes.addTable(rowCount, 3, { specificCellProperties });
    await context.sync();

    tableShape.left = 30;
    tableShape.top  = 120;

    const table = tableShape.getTable();
    table.load();
    await context.sync();

    colWidths.forEach((w, c) => { table.columns.getItemAt(c).width = w; });
    Array.from({ length: rowCount }, (_, r) => { table.rows.getItemAt(r).height = rowHeight; });

    await context.sync();
    setResult({ text: "スライドに出力しました" });
  });
}

/**
 * tableRows DOM から行データを収集する。
 */
function collectTableRows() {
  const container = document.getElementById("tableRows");
  if (!container) return [];

  return [...container.querySelectorAll(".tableInputRow")].map((row) => {
    const inputs   = row.querySelectorAll("input.tableInput, textarea.tableInput");
    const noUnit   = row.classList.contains("noUnitRow");
    const content  = inputs[0]?.value ?? "";
    const quantity = inputs[1]?.value ?? "";
    const unit     = noUnit ? "" : (inputs[2]?.value ?? "");
    return { content, quantity, unit, noUnit };
  }).filter((r) => r.content || r.quantity);
}

function bindClick(elementId, handler) {
  const element = document.getElementById(elementId);
  if (!element) return;
  element.addEventListener("click", async () => {
    try {
      await handler();
    } catch (error) {
      console.error(error);
      const parts = [
        error?.message,
        error?.code,
        error?.name,
        error?.debugInfo  ? JSON.stringify(error.debugInfo)  : null,
        error?.innerError ? JSON.stringify(error.innerError) : null,
        error?.traceMessages?.join(" / "),
      ].filter(Boolean);
      const msg = parts.join(" | ") || String(error);
      setResult({ text: "エラー：" + msg });
      await placeErrorTextOnSlide(msg);
    }
  });
}

/**
 * エラーメッセージをスライド上に図形として配置する（デバッグ用）。
 */
async function placeErrorTextOnSlide(message) {
  await PowerPoint.run(async (context) => {
    let slide;
    try {
      const selected = context.presentation.getSelectedSlides();
      selected.load("items");
      await context.sync();
      slide = selected.items[0];
    } catch {
      const allSlides = context.presentation.slides;
      allSlides.load("items");
      await context.sync();
      slide = allSlides.items[0];
    }

    const shape = slide.shapes.addGeometricShape("rectangle");
    Object.assign(shape, { left: 50, top: 50, width: 600, height: 120 });
    shape.fill.setSolidColor("FFEEEE");
    await context.sync();

    const font = shape.textFrame.textRange.font;
    shape.textFrame.textRange.text = "DEBUG ERROR:\n" + message;
    Object.assign(font, { size: 10, color: "CC0000", bold: true, name: "Meiryo" });
    await context.sync();
  });
}

// ─── 集計処理 ─────────────────────────────────────────────

/**
 * 指定した文字色の数字を合計する。
 */
async function sumNumbersByTextColor(targetTextColor, label) {
  setResult({ text: "集計中..." });

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const numbers = [];
    for (const item of await getTextShapes(context, slide)) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      numbers.push(...extractNumbers(coloredText));
    }

    const total = String(sum(numbers));
    setResult({ text: total, color: targetTextColor, copyValue: total });
  });
}

/**
 * 選択中テキストの文字色で数字を合計する。
 */
async function sumNumbersBySelectedTextColor() {
  setResult({ text: "選択中の文字色を取得中..." });
  const selected = await getSelectedTextInfo();
  if (!selected.ok) { setResult({ text: selected.message }); return; }
  await sumNumbersByTextColor(selected.textColor, `選択中の文字色 ${selected.textColor}`);
}

/**
 * 選択中テキストの文字色＋図形の背景色で数字を合計する。
 */
async function sumNumbersBySelectedTextColorAndFillColor() {
  setResult({ text: "選択中の文字色と背景色を取得中..." });
  const selected = await getSelectedTextAndFillInfo();
  if (!selected.ok) { setResult({ text: selected.message }); return; }

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const numbers = [];
    const textShapes = await getTextShapes(context, slide, { includeFillColor: true });
    for (const item of textShapes.filter((s) => s.fillColor === selected.fillColor)) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, selected.textColor);
      numbers.push(...extractNumbers(coloredText));
    }

    const total = String(sum(numbers));
    setResult({ text: total, color: selected.textColor, copyValue: total });
  });
}

/**
 * 指定色の A-1 形式コードをアルファベットごとに連番圧縮する。
 */
async function formatCodesByTextColor(targetTextColor, allSlides = false) {
  setResult({ text: allSlides ? "全スライドを整理中..." : "整理中..." });

  await PowerPoint.run(async (context) => {
    let slides;
    if (allSlides) {
      const col = context.presentation.slides;
      col.load("items");
      await context.sync();
      slides = col.items;
    } else {
      const current = await getCurrentSlide(context);
      if (!current) return;
      slides = [current];
    }

    const grouped = {};
    for (const slide of slides) {
      for (const item of await getTextShapes(context, slide)) {
        const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
        for (const { letter, number } of extractLetterCodes(coloredText)) {
          (grouped[letter] ??= []).push(number);
        }
      }
    }

    const outputLines = Object.keys(grouped).sort()
      .map((letter) => `${letter}-${compressNumberRanges(grouped[letter])}`);
    const outputText = outputLines.length > 0 ? outputLines.join("\n") : "該当なし";

    setResult({
      html:      `<pre class="outputText" style="color:${targetTextColor}">${escapeHtml(outputText)}</pre>`,
      color:     targetTextColor,
      copyValue: outputText
    });
  });
}

// ─── PowerPoint API ───────────────────────────────────────

/**
 * テキスト選択情報の共通取得ロジック。
 * テキスト範囲が有効かどうかを検証して返す。
 */
async function loadSelectedTextRange(context) {
  const selectedTextRange = context.presentation.getSelectedTextRangeOrNullObject();
  selectedTextRange.load("text");
  selectedTextRange.font.load("color");
  await context.sync();

  if (selectedTextRange.isNullObject || !selectedTextRange.text?.trim()) {
    return { ok: false, message: "テキストが選択されていません。色を取得したい文字を選択してください。" };
  }

  const textColor = normalizeColor(selectedTextRange.font.color);
  if (!textColor) {
    return { ok: false, message: "選択中テキストの文字色を取得できませんでした。1色の文字だけを選択してください。" };
  }

  return { ok: true, textColor };
}

/**
 * 現在選択中のテキスト情報を取得する。
 */
async function getSelectedTextInfo() {
  return PowerPoint.run(async (context) => {
    const count = context.presentation.getSelectedShapes().getCount();
    await context.sync();

    if (count.value > 1) {
      return { ok: false, message: "複数選択されています。色を取得したいテキストを1つだけ選択してください。" };
    }

    return loadSelectedTextRange(context);
  });
}

/**
 * 現在選択中のテキスト色と図形の塗りつぶし色を取得する。
 */
async function getSelectedTextAndFillInfo() {
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    const count = selectedShapes.getCount();
    await context.sync();

    if (count.value !== 1) {
      return { ok: false, message: "文字色と背景色を取得したいテキストボックスを1つだけ選択してください。" };
    }

    const textResult = await loadSelectedTextRange(context);
    if (!textResult.ok) return textResult;

    const fillColor = await getShapeFillColor(context, selectedShapes.items[0]);
    if (!fillColor) {
      return { ok: false, message: "選択中テキストボックスの背景色を取得できませんでした。単色の塗りつぶし色が設定されたテキストボックスを選択してください。" };
    }

    return { ok: true, textColor: textResult.textColor, fillColor };
  });
}

/**
 * 現在選択中のスライドを1枚取得する。
 */
async function getCurrentSlide(context) {
  const selectedSlides = context.presentation.getSelectedSlides();
  selectedSlides.load("items");
  const slideCount = selectedSlides.getCount();
  await context.sync();

  if (slideCount.value === 0) {
    setResult({ text: "現在のスライドを取得できませんでした。" });
    return null;
  }

  return selectedSlides.items[0];
}

/**
 * スライド内のテキストを持つ図形を取得する。
 */
async function getTextShapes(context, slide, options = {}) {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();

  const candidates = [];
  for (const shape of shapes.items) {
    try {
      const textFrame = shape.getTextFrameOrNullObject();
      textFrame.load("hasText");
      candidates.push({ shape, textFrame });
    } catch (error) {
      console.warn("テキスト枠を取得できない図形をスキップしました:", error);
    }
  }
  await context.sync();

  const textShapes = [];
  for (const item of candidates) {
    if (item.textFrame.isNullObject || !item.textFrame.hasText) continue;
    const textRange = item.textFrame.textRange;
    textRange.load("text");
    textShapes.push({ ...item, textRange });
  }
  await context.sync();

  for (const item of textShapes) {
    item.text = item.textRange.text || "";
    if (options.includeFillColor) {
      item.fillColor = await getShapeFillColor(context, item.shape);
    }
  }

  return textShapes.filter((item) => item.text.length > 0);
}

/**
 * 指定色の文字だけを抽出し、それ以外をスペースに置換する。
 */
async function extractTextByColor(context, textRange, text, targetColor) {
  const charRanges = Array.from({ length: text.length }, (_, i) => {
    const charRange = textRange.getSubstring(i, 1);
    charRange.load("text");
    charRange.font.load("color");
    return charRange;
  });
  await context.sync();

  return charRanges.map((cr) =>
    normalizeColor(cr.font.color) === targetColor ? (cr.text || "") : " "
  ).join("");
}

async function getShapeFillColor(context, shape) {
  try {
    shape.fill.load("foregroundColor");
    await context.sync();
    return normalizeColor(shape.fill.foregroundColor);
  } catch (error) {
    console.warn("背景色を取得できない図形をスキップしました:", error);
    return "";
  }
}

// ─── テキスト処理ユーティリティ ───────────────────────────

function extractNumbers(text) {
  return (text.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number).filter((v) => !Number.isNaN(v));
}

function extractLetterCodes(text) {
  return [...text.matchAll(/\b([A-Z])-(\d+)\b/g)].map((m) => ({
    letter: m[1],
    number: Number(m[2])
  }));
}

function compressNumberRanges(numbers) {
  if (!numbers?.length) return "";

  const sorted = [...new Set(numbers)].sort((a, b) => a - b);
  const ranges = [];
  let start = sorted[0];
  let prev  = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === prev + 1) { prev = sorted[i]; continue; }
    ranges.push(formatRange(start, prev));
    start = prev = sorted[i];
  }
  ranges.push(formatRange(start, prev));

  return ranges.join(", ");
}

function formatRange(start, end) {
  return start === end ? String(start) : `${start}~${end}`;
}

function sum(numbers) {
  return numbers.reduce((total, v) => total + v, 0);
}

function normalizeColor(color) {
  if (!color) return "";
  const s = String(color).trim().toLowerCase();
  return /^[0-9a-f]{6}$/.test(s) ? "#" + s : s;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// ─── UI 更新 ──────────────────────────────────────────────

/**
 * 結果欄を一括更新する。
 * @param {object}      opts
 * @param {string}      [opts.text]      テキスト表示
 * @param {string}      [opts.html]      HTML表示（text より優先）
 * @param {string|null} [opts.color]     文字色（省略時リセット）
 * @param {string|null} [opts.copyValue] コピーボタンに渡す値（省略時は無効化）
 */
function setResult({ text, html, color = null, copyValue = null }) {
  const el = document.getElementById(UI.result);
  if (el) {
    if (html !== undefined) { el.innerHTML = html; }
    else                    { el.textContent = text ?? ""; }
    el.style.color = color ?? "";
  }

  const btn = document.getElementById(UI.copyButton);
  if (btn) {
    btn.disabled = copyValue === null;
    btn.dataset.copyValue = copyValue ?? "";
  }

  if (copyValue !== null) navigator.clipboard.writeText(copyValue).catch(() => {});
}
