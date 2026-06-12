const COLORS = {
  BLUE:  "#0070c0",
  GREEN: "#00b050",
  BLACK: "#000000"
};

const UI = {
  result:      "result",
  copyButton:  "copyResultButton"
};

const DEFECT_ITEMS = [
  { content: "タイル割れ",  unit: "枚数" },
  { content: "タイル浮き",  unit: "枚数" },
  { content: "モルタル浮き", unit: "㎡"  },
  { content: "塗膜浮き",   unit: "㎡"  }
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
  bindClick("sumSelectedColorButton",          sumNumbersBySelectedTextColor);
  bindClick("sumSelectedTextAndFillColorButton", sumNumbersBySelectedTextColorAndFillColor);
  bindClick("formatBlackCodeButton",         () => formatCodesByTextColor(COLORS.BLACK, false));
  bindClick("formatAllSlidesBlackCodeButton", () => formatCodesByTextColor(COLORS.BLACK, true));

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
  document.querySelectorAll(".tabButton").forEach((button) => {
    button.addEventListener("click", () => activateTab(button.dataset.tabTarget));
  });
}

function activateTab(tabId) {
  if (!tabId) return;
  document.querySelectorAll(".tabButton").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tabTarget === tabId);
  });
  document.querySelectorAll(".tabContent").forEach((content) => {
    content.classList.toggle("active", content.id === tabId);
  });
}

function initializeTableBuilder() {
  renderDefectButtons();
  bindClick("clearTableRowsButton", clearTableRows);
}

function renderDefectButtons() {
  const container = document.getElementById("defectButtonGrid");
  if (!container) return;

  DEFECT_ITEMS.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "defectButton";
    button.textContent = item.content;
    button.addEventListener("click", () => addTableRow(item));
    container.appendChild(button);
  });
}

function addTableRow({ content, unit }) {
  const tableRows = document.getElementById("tableRows");
  if (!tableRows) return;

  const row = document.createElement("div");
  row.className = "tableInputRow";

  const quantityInput = createTableInput({ type: "number", className: "quantityInput", placeholder: "数量" });

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "row-delete-btn";
  deleteButton.textContent = "×";
  deleteButton.addEventListener("click", () => row.remove());

  row.append(
    createTableInput({ value: content, className: "contentInput" }),
    quantityInput,
    createTableInput({ value: unit, className: "unitInput" }),
    deleteButton
  );

  tableRows.appendChild(row);
  quantityInput.focus();
}

function createTableInput({ type = "text", value = "", className = "", placeholder = "" }) {
  const input = document.createElement("input");
  input.type = type;
  input.className = ["tableInput", className].filter(Boolean).join(" ");
  input.value = value;
  input.placeholder = placeholder;
  return input;
}

function clearTableRows() {
  document.getElementById("tableRows")?.replaceChildren();
}

function bindClick(elementId, handler) {
  const element = document.getElementById(elementId);
  if (!element) return;
  element.addEventListener("click", async () => {
    try {
      await handler();
    } catch (error) {
      console.error(error);
      setResult({ text: "エラー：" + (error?.message ?? String(error)) });
    }
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

    const textShapes = await getTextShapes(context, slide);
    const numbers = [];

    for (const item of textShapes) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      numbers.push(...extractNumbers(coloredText));
    }

    setResult({ text: String(sum(numbers)), color: targetTextColor, copyValue: String(sum(numbers)) });
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

    const textShapes = await getTextShapes(context, slide, { includeFillColor: true });
    const targetShapes = textShapes.filter((item) => item.fillColor === selected.fillColor);
    const numbers = [];

    for (const item of targetShapes) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, selected.textColor);
      numbers.push(...extractNumbers(coloredText));
    }

    const total = sum(numbers);
    setResult({ text: String(total), color: selected.textColor, copyValue: String(total) });
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
      const allSlideCollection = context.presentation.slides;
      allSlideCollection.load("items");
      await context.sync();
      slides = allSlideCollection.items;
    } else {
      const current = await getCurrentSlide(context);
      if (!current) return;
      slides = [current];
    }

    const grouped = {};

    for (const slide of slides) {
      const textShapes = await getTextShapes(context, slide);
      for (const item of textShapes) {
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
      html: `<pre class="outputText" style="color:${targetTextColor}">${escapeHtml(outputText)}</pre>`,
      color: targetTextColor,
      copyValue: outputText
    });
  });
}

// ─── PowerPoint API ───────────────────────────────────────

/**
 * 現在選択中のテキスト情報を取得する。
 */
async function getSelectedTextInfo() {
  return PowerPoint.run(async (context) => {
    const selectedShapes   = context.presentation.getSelectedShapes();
    const selectedShapeCount = selectedShapes.getCount();
    const selectedTextRange  = context.presentation.getSelectedTextRangeOrNullObject();

    await context.sync();

    if (selectedShapeCount.value > 1) {
      return { ok: false, message: "複数選択されています。色を取得したいテキストを1つだけ選択してください。" };
    }

    selectedTextRange.load("text");
    selectedTextRange.font.load("color");
    await context.sync();

    if (selectedTextRange.isNullObject || !selectedTextRange.text?.trim()) {
      return { ok: false, message: "テキストが選択されていません。色を取得したい文字を1つ選択してください。" };
    }

    const textColor = normalizeColor(selectedTextRange.font.color);
    if (!textColor) {
      return { ok: false, message: "選択中テキストの文字色を取得できませんでした。1色の文字だけを選択してください。" };
    }

    return { ok: true, textColor };
  });
}

/**
 * 現在選択中のテキスト色と図形の塗りつぶし色を取得する。
 */
async function getSelectedTextAndFillInfo() {
  return PowerPoint.run(async (context) => {
    const selectedShapes   = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    const selectedShapeCount = selectedShapes.getCount();
    const selectedTextRange  = context.presentation.getSelectedTextRangeOrNullObject();

    await context.sync();

    if (selectedShapeCount.value !== 1) {
      return { ok: false, message: "文字色と背景色を取得したいテキストボックスを1つだけ選択してください。" };
    }

    selectedTextRange.load("text");
    selectedTextRange.font.load("color");
    await context.sync();

    if (selectedTextRange.isNullObject || !selectedTextRange.text?.trim()) {
      return { ok: false, message: "テキストが選択されていません。色を取得したい文字を選択してください。" };
    }

    const textColor = normalizeColor(selectedTextRange.font.color);
    if (!textColor) {
      return { ok: false, message: "選択中テキストの文字色を取得できませんでした。" };
    }

    const fillColor = await getShapeFillColor(context, selectedShapes.items[0]);
    if (!fillColor) {
      return { ok: false, message: "選択中テキストボックスの背景色を取得できませんでした。単色の塗りつぶし色が設定されたテキストボックスを選択してください。" };
    }

    return { ok: true, textColor, fillColor };
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
  const charRanges = [];
  for (let i = 0; i < text.length; i++) {
    const charRange = textRange.getSubstring(i, 1);
    charRange.load("text");
    charRange.font.load("color");
    charRanges.push(charRange);
  }
  await context.sync();

  return charRanges.map((charRange) => {
    const color = normalizeColor(charRange.font.color);
    return color === targetColor ? (charRange.text || "") : " ";
  }).join("");
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
  // キャプチャグループで直接分解し、再マッチを省く
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
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// ─── UI 更新 ──────────────────────────────────────────────

/**
 * 結果欄を一括更新する。
 * @param {object} opts
 * @param {string}      [opts.text]      テキスト表示（htmlより優先度低）
 * @param {string}      [opts.html]      HTML表示
 * @param {string|null} [opts.color]     文字色（省略時リセット）
 * @param {string|null} [opts.copyValue] コピーボタンに渡す値（省略時は無効化）
 */
function setResult({ text, html, color = null, copyValue = null }) {
  const el = document.getElementById(UI.result);
  if (el) {
    if (html !== undefined) {
      el.innerHTML = html;
    } else {
      el.textContent = text ?? "";
    }
    el.style.color = color ?? "";
  }

  const btn = document.getElementById(UI.copyButton);
  if (btn) {
    btn.disabled = copyValue === null;
    btn.dataset.copyValue = copyValue ?? "";
  }

  if (copyValue !== null) {
    navigator.clipboard.writeText(copyValue).catch(() => {});
  }
}
