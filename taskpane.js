const COLORS = {
  BLUE: "#0070c0",
  GREEN: "#00b050",
  BLACK: "#000000"
};

const UI = {
  result: "result",
  targetColor: "targetColor"
};

const TAB_IDS = {
  summary: "summaryTab",
  table: "tableTab"
};

const DEFECT_ITEMS = [
  { content: "タイル割れ", unit: "枚数" },
  { content: "タイル浮き", unit: "枚数" },
  { content: "モルタル浮き", unit: "㎡" },
  { content: "塗膜浮き", unit: "㎡" }
];

Office.onReady(() => {
  initializeTabs();
  initializeTableBuilder();
  initializeSummaryButtons();
});

function initializeSummaryButtons() {
  bindClick("sumBlueButton", () =>
    sumNumbersByTextColor(COLORS.BLUE, "青文字 RGB(0,112,192)")
  );

  bindClick("sumGreenButton", () =>
    sumNumbersByTextColor(COLORS.GREEN, "緑文字 RGB(0,176,80)")
  );

  bindClick("sumSelectedColorButton", sumNumbersBySelectedTextColor);

  bindClick(
    "sumSelectedTextAndFillColorButton",
    sumNumbersBySelectedTextColorAndFillColor
  );

  bindClick("formatBlackCodeButton", () =>
    formatCodesByTextColor(COLORS.BLACK, "黒文字 RGB(0,0,0)")
  );

  const copyBtn = document.getElementById("copyResultButton");
  if (copyBtn) {
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
}

function initializeTabs() {
  const tabButtons = Array.from(document.querySelectorAll(".tabButton"));

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activateTab(button.dataset.tabTarget);
    });
  });
}

function activateTab(tabId) {
  if (!tabId) return;

  document.querySelectorAll(".tabButton").forEach((button) => {
    button.classList.toggle("active", button.dataset.tabTarget === tabId);
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

  container.innerHTML = "";

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

  const row = createElement("div", {
    className: "tableInputRow"
  });

  const quantityInput = createTableInput({
    type: "number",
    className: "quantityInput",
    placeholder: "数量"
  });

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "row-delete-btn";
  deleteButton.textContent = "×";
  deleteButton.addEventListener("click", () => row.remove());

  row.append(
    createTableInput({
      value: content,
      className: "contentInput"
    }),
    quantityInput,
    createTableInput({
      value: unit,
      className: "unitInput"
    }),
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
  const tableRows = document.getElementById("tableRows");
  if (tableRows) tableRows.replaceChildren();
}

function bindClick(elementId, handler) {
  const element = document.getElementById(elementId);
  if (!element) return;

  element.addEventListener("click", async () => {
    try {
      await handler();
    } catch (error) {
      console.error(error);
      setResultText("エラー：" + getErrorMessage(error));
    }
  });
}

/**
 * 指定した文字色の数字を、現在のスライドから取得して合計する。
 */
async function sumNumbersByTextColor(targetTextColor, label) {
  setStatus(`対象色：${label}`, "集計中...");

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const textShapes = await getTextShapes(context, slide);
    const numbers = [];

    for (const item of textShapes) {
      const coloredText = await extractTextByColor(
        context,
        item.textRange,
        item.text,
        targetTextColor
      );

      numbers.push(...extractNumbers(coloredText));
    }

    renderNumberSummary({
      total: sum(numbers),
      numbers,
      checkedShapeCount: textShapes.length,
      color: targetTextColor
    });
  });
}

/**
 * 選択中テキストの文字色を取得し、その文字色の数字を現在のスライドから合計する。
 */
async function sumNumbersBySelectedTextColor() {
  setStatus("対象色：選択中の文字色", "選択中の文字色を取得中...");

  const selected = await getSelectedTextInfo();

  if (!selected.ok) {
    setResultText(selected.message);
    return;
  }

  await sumNumbersByTextColor(
    selected.textColor,
    `選択中の文字色 ${selected.textColor}`
  );
}

/**
 * 選択中テキストの文字色と、選択中図形の塗りつぶし色を取得し、
 * 両方が一致する図形内の数字を合計する。
 */
async function sumNumbersBySelectedTextColorAndFillColor() {
  setStatus("対象：選択中の文字色＋背景色", "選択中の文字色と背景色を取得中...");

  const selected = await getSelectedTextAndFillInfo();

  if (!selected.ok) {
    setResultText(selected.message);
    return;
  }

  await sumNumbersByTextAndFillColor({
    targetTextColor: selected.textColor,
    targetFillColor: selected.fillColor
  });
}

/**
 * 指定した文字色かつ指定した背景色の数字を合計する。
 */
async function sumNumbersByTextAndFillColor({ targetTextColor, targetFillColor }) {
  setStatus(
    `対象：文字色 ${targetTextColor} / 背景色 ${targetFillColor}`,
    "集計中..."
  );

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const textShapes = await getTextShapes(context, slide, {
      includeFillColor: true
    });

    const targetShapes = textShapes.filter(
      (item) => item.fillColor === targetFillColor
    );

    const numbers = [];

    for (const item of targetShapes) {
      const coloredText = await extractTextByColor(
        context,
        item.textRange,
        item.text,
        targetTextColor
      );

      numbers.push(...extractNumbers(coloredText));
    }

    renderNumberSummary({
      total: sum(numbers),
      numbers,
      checkedShapeCount: textShapes.length,
      matchedShapeCount: targetShapes.length,
      color: targetTextColor
    });
  });
}

/**
 * 指定色の A-1 形式コードを、アルファベットごとに連番圧縮する。
 */
async function formatCodesByTextColor(targetTextColor, label) {
  setStatus(`対象色：${label}`, "整理中...");

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const textShapes = await getTextShapes(context, slide);
    const grouped = {};
    let matchedCodeCount = 0;

    for (const item of textShapes) {
      const coloredText = await extractTextByColor(
        context,
        item.textRange,
        item.text,
        targetTextColor
      );

      for (const code of extractLetterCodes(coloredText)) {
        if (!grouped[code.letter]) grouped[code.letter] = [];
        grouped[code.letter].push(code.number);
        matchedCodeCount++;
      }
    }

    renderCodeSummary({
      grouped,
      matchedCodeCount,
      checkedShapeCount: textShapes.length,
      color: targetTextColor
    });
  });
}

/**
 * 現在選択中のテキスト情報を取得する。
 * 複数図形選択時はエラーとして返す。
 */
async function getSelectedTextInfo() {
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    const selectedShapeCount = selectedShapes.getCount();

    const selectedTextRange =
      context.presentation.getSelectedTextRangeOrNullObject();

    await context.sync();

    if (selectedShapeCount.value > 1) {
      return {
        ok: false,
        message: "複数選択されています。色を取得したいテキストを1つだけ選択してください。"
      };
    }

    selectedTextRange.load("text");
    selectedTextRange.font.load("color");

    await context.sync();

    if (selectedTextRange.isNullObject) {
      return {
        ok: false,
        message: "テキストが選択されていません。色を取得したい文字を1つ選択してください。"
      };
    }

    const selectedText = selectedTextRange.text || "";
    const textColor = normalizeColor(selectedTextRange.font.color);

    if (!selectedText.trim()) {
      return {
        ok: false,
        message: "選択中のテキストが空です。色を取得したい文字を選択してください。"
      };
    }

    if (!textColor) {
      return {
        ok: false,
        message: "選択中テキストの文字色を取得できませんでした。1色の文字だけを選択してください。"
      };
    }

    return { ok: true, textColor };
  });
}

/**
 * 現在選択中のテキスト色と図形の塗りつぶし色を取得する。
 */
async function getSelectedTextAndFillInfo() {
  return PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");

    const selectedShapeCount = selectedShapes.getCount();
    const selectedTextRange =
      context.presentation.getSelectedTextRangeOrNullObject();

    await context.sync();

    if (selectedShapeCount.value !== 1) {
      return {
        ok: false,
        message: "文字色と背景色を取得したいテキストボックスを1つだけ選択してください。"
      };
    }

    selectedTextRange.load("text");
    selectedTextRange.font.load("color");

    await context.sync();

    if (selectedTextRange.isNullObject) {
      return {
        ok: false,
        message: "テキストが選択されていません。色を取得したい文字を選択してください。"
      };
    }

    const selectedText = selectedTextRange.text || "";
    const textColor = normalizeColor(selectedTextRange.font.color);

    if (!selectedText.trim()) {
      return {
        ok: false,
        message: "選択中のテキストが空です。色を取得したい文字を選択してください。"
      };
    }

    if (!textColor) {
      return {
        ok: false,
        message: "選択中テキストの文字色を取得できませんでした。"
      };
    }

    const fillColor = await getShapeFillColor(context, selectedShapes.items[0]);

    if (!fillColor) {
      return {
        ok: false,
        message: "選択中テキストボックスの背景色を取得できませんでした。単色の塗りつぶし色が設定されたテキストボックスを選択してください。"
      };
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
    setResultText("現在のスライドを取得できませんでした。");
    return null;
  }

  return selectedSlides.items[0];
}

/**
 * スライド内のテキストを持つ図形を取得する。
 * includeFillColor=true の場合は、図形の塗りつぶし色も取得する。
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

    textShapes.push({
      ...item,
      textRange
    });
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

  return charRanges
    .map((charRange) => {
      const charText = charRange.text || "";
      const color = normalizeColor(charRange.font.color);

      return color === targetColor ? charText : " ";
    })
    .join("");
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

function extractNumbers(text) {
  const matches = text.match(/-?\d+(?:\.\d+)?/g);

  return matches
    ? matches.map(Number).filter((value) => !Number.isNaN(value))
    : [];
}

function extractLetterCodes(text) {
  const matches = text.match(/\b[A-Z]-\d+\b/g);
  if (!matches) return [];

  return matches
    .map((code) => code.match(/^([A-Z])-(\d+)$/))
    .filter(Boolean)
    .map((match) => ({
      letter: match[1],
      number: Number(match[2])
    }));
}

function compressNumberRanges(numbers) {
  if (!numbers || numbers.length === 0) return "";

  const sorted = [...new Set(numbers)].sort((a, b) => a - b);
  const ranges = [];

  let start = sorted[0];
  let previous = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const current = sorted[i];

    if (current === previous + 1) {
      previous = current;
      continue;
    }

    ranges.push(formatRange(start, previous));
    start = current;
    previous = current;
  }

  ranges.push(formatRange(start, previous));

  return ranges.join(", ");
}

function formatRange(start, end) {
  return start === end ? String(start) : `${start}~${end}`;
}

function sum(numbers) {
  return numbers.reduce((total, value) => total + value, 0);
}

function normalizeColor(color) {
  if (!color) return "";

  let normalized = String(color).trim().toLowerCase();

  if (/^[0-9a-f]{6}$/.test(normalized)) {
    normalized = "#" + normalized;
  }

  return normalized;
}

function renderNumberSummary({
  total,
  numbers,
  checkedShapeCount,
  matchedShapeCount = null,
  color = null
}) {
  setResultValue(total, color);
}

function renderCodeSummary({ grouped, matchedCodeCount, checkedShapeCount, color = null }) {
  setResultValue(matchedCodeCount, color);
}

function createElement(tagName, options = {}) {
  const element = document.createElement(tagName);

  if (options.className) element.className = options.className;
  if (options.textContent) element.textContent = options.textContent;

  return element;
}

function setStatus(targetColorText, resultText) {
  setElementText(UI.targetColor, targetColorText);
  setResultText(resultText);
}

function setResultText(text) {
  setElementText(UI.result, text);
  updateCopyButton(null);
}

function setResultValue(value, color = null) {
  const element = document.getElementById(UI.result);
  if (element) {
    element.textContent = typeof value === "number" ? String(value) : value;
    element.style.color = color || "";
  }
  updateCopyButton(typeof value === "number" ? value : null);
}

function setResultHtml(html) {
  const element = document.getElementById(UI.result);
  if (element) element.innerHTML = html;
  updateCopyButton(null);
}

function updateCopyButton(value) {
  const btn = document.getElementById("copyResultButton");
  if (!btn) return;
  if (value === null) {
    btn.disabled = true;
    btn.dataset.copyValue = "";
  } else {
    btn.disabled = false;
    btn.dataset.copyValue = String(value);
  }
}

function setElementText(elementId, text) {
  const element = document.getElementById(elementId);
  if (element) element.textContent = text;
}

function getErrorMessage(error) {
  if (!error) return "不明なエラー";
  return error.message || String(error);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
