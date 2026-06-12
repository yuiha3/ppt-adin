const COLORS = {
  BLUE:  "#0070c0",
  GREEN: "#00b050",
  BLACK: "#000000"
};

const UI = {
  result:      "result",
  copyButton:  "copyResultButton"
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
  bindClick("outputTableButton",   outputTableToSlide);
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

  // テキストが折り返すボタンのフォントサイズを縮小
  shrinkOverflowButtons(container);
  new ResizeObserver(() => shrinkOverflowButtons(container)).observe(container);
}

function shrinkOverflowButtons(container) {
  container.querySelectorAll(".defectButton").forEach((btn) => {
    btn.style.fontSize = "";            // いったんリセット
    if (btn.scrollHeight > btn.clientHeight + 2) {
      btn.style.fontSize = "10px";
    }
  });
}

function addTableRow({ content, unit, noUnit = false, pinBottom = false }) {
  const tableRows = document.getElementById("tableRows");
  if (!tableRows) return;

  const row = document.createElement("div");
  row.className = noUnit ? "tableInputRow noUnitRow" : "tableInputRow";
  row.draggable = true;
  if (pinBottom) row.dataset.pinBottom = "true";

  // ドラッグハンドル
  const handle = document.createElement("span");
  handle.className = "drag-handle";
  handle.textContent = "⠿";
  handle.setAttribute("aria-hidden", "true");

  // 数量欄：noUnitはtextarea、通常はtext input
  let quantityEl;
  if (noUnit) {
    quantityEl = document.createElement("textarea");
    quantityEl.className = "tableInput quantityInput quantityTextarea";
    quantityEl.placeholder = "入力";
    quantityEl.rows = 2;
  } else {
    quantityEl = createTableInput({ className: "quantityInput", placeholder: "数量", inputmode: "numeric" });
  }

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "row-delete-btn";
  deleteButton.textContent = "×";
  deleteButton.addEventListener("click", () => row.remove());

  if (noUnit) {
    row.append(
      handle,
      createTableInput({ value: content, className: "contentInput" }),
      quantityEl,
      deleteButton
    );
  } else {
    row.append(
      handle,
      createTableInput({ value: content, className: "contentInput" }),
      quantityEl,
      createTableInput({ value: unit, className: "unitInput" }),
      deleteButton
    );
  }

  // pinBottom行（写真番号）は常に最下部、それ以外は最初のpinBottom行の直前に挿入
  const firstPinned = tableRows.querySelector("[data-pin-bottom]");
  if (firstPinned) {
    tableRows.insertBefore(row, firstPinned);
  } else {
    tableRows.appendChild(row);
  }

  setupDragAndDrop(row, tableRows);
  quantityEl.focus();
}

// ── ドラッグ&ドロップ ────────────────────────────────────
let dragSrc = null;

function setupDragAndDrop(row, tableRows) {
  row.addEventListener("dragstart", (e) => {
    dragSrc = row;
    row.classList.add("dragging");
    e.dataTransfer.effectAllowed = "move";
  });

  row.addEventListener("dragend", () => {
    row.classList.remove("dragging");
    tableRows.querySelectorAll(".drag-over").forEach((el) => el.classList.remove("drag-over"));
    dragSrc = null;
    enforcePinnedRows(tableRows);
  });

  row.addEventListener("dragover", (e) => {
    e.preventDefault();
    if (!dragSrc || dragSrc === row) return;
    e.dataTransfer.dropEffect = "move";
    tableRows.querySelectorAll(".drag-over").forEach((el) => el.classList.remove("drag-over"));
    row.classList.add("drag-over");
  });

  row.addEventListener("drop", (e) => {
    e.preventDefault();
    if (!dragSrc || dragSrc === row) return;
    const rows = [...tableRows.children];
    const srcIdx = rows.indexOf(dragSrc);
    const tgtIdx = rows.indexOf(row);
    if (srcIdx < tgtIdx) {
      tableRows.insertBefore(dragSrc, row.nextSibling);
    } else {
      tableRows.insertBefore(dragSrc, row);
    }
  });
}

// noUnit行が末尾に来るよう整列
function enforcePinnedRows(tableRows) {
  const pinned = [...tableRows.querySelectorAll("[data-pin-bottom]")];
  pinned.forEach((el) => tableRows.appendChild(el));
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

    // Step1: addTable で表を作成
    slide.shapes.addTable(rowCount, 3);
    slide.shapes.load("items");
    await context.sync();

    // addTable で追加した表は shapes の末尾に入る
    const tableShape = slide.shapes.items[slide.shapes.items.length - 1];

    // Step2: 位置を設定
    tableShape.left = 30;
    tableShape.top  = 120;

    // Step3: 表のOOXMLを取得してスタイル適用後に上書き
    const ooxmlProxy = tableShape.getOoxml();
    await context.sync();

    const styledXml = buildStyledTableOoxml(
      ooxmlProxy.value, rows, headers, colWidths, rowHeight
    );
    tableShape.setOoxml(styledXml);
    await context.sync();

    setResult({ text: "スライドに出力しました" });
  });
}


/**
 * 取得したOOXMLのテーブル部分をVBA仕様に沿って書き換える。
 * - テキスト・罫線・フォント・余白・背景色・行高・列幅をすべてXMLで設定
 */
function buildStyledTableOoxml(baseXml, rows, headers, colWidths, rowHeight) {
  const PT  = 12700; // 1pt = 12700 EMU
  const W   = 12700; // 罫線 1pt
  const B   = `<a:ln w="${W}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`;
  const N   = `<a:ln w="0"><a:noFill/></a:ln>`;

  const rowCount = rows.length + 1;

  // 列幅グリッド
  const gridCols = colWidths.map(w => `<a:gridCol w="${w * PT}"/>`).join("");

  // セルXMLを生成
  const makeCell = (text, rowIdx, colIdx) => {
    const isHeader = rowIdx === 0;
    const isQty    = colIdx === 1;
    const leftMarg = isQty ? 0 : 5 * PT;
    const algn     = isQty ? "ctr" : "l";
    const lnL = isHeader ? N : B;
    const lnR = isHeader ? N : B;
    const lnT = isHeader ? N : B;
    const lnB = B; // 全行下罫線あり
    const safeText = String(text ?? "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");

    return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${algn}"/><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Meiryo"/><a:ea typeface="Meiryo"/></a:rPr><a:t>${safeText}</a:t></a:r></a:p></a:txBody><a:tcPr marL="${leftMarg}" marR="0" marT="0" marB="0" anchor="ctr"><a:lnL>${lnL}</a:lnL><a:lnR>${lnR}</a:lnR><a:lnT>${lnT}</a:lnT><a:lnB>${lnB}</a:lnB><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr></a:tc>`;
  };

  // 全行のXMLを生成
  const rowsXml = [];
  for (let r = 0; r < rowCount; r++) {
    const cells = [0, 1, 2].map(c => {
      const text = r === 0
        ? headers[c]
        : (c === 0 ? rows[r-1].content : c === 1 ? rows[r-1].quantity : rows[r-1].unit);
      return makeCell(text, r, c);
    });
    rowsXml.push(
      `<a:tr h="${Math.round(rowHeight * PT)}">${cells.join("")}</a:tr>`
    );
  }

  // baseXml の <p:graphicFrame> から xfrm（位置情報）だけ流用し tbl を丸ごと置換
  const tblXml = `<a:tbl><a:tblPr firstRow="0" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${gridCols}</a:tblGrid>${rowsXml.join("")}</a:tbl>`;

  // baseXml 内の既存 <a:tbl>...</a:tbl> を置換
  const result = baseXml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, tblXml);

  // 置換できなかった場合（まれに tbl がない）はそのまま返す
  return result.includes("<a:tbl>") ? result : baseXml;
}


/**
 * tableRows DOM から行データを収集する。
 */
function collectTableRows() {
  const container = document.getElementById("tableRows");
  if (!container) return [];

  return [...container.querySelectorAll(".tableInputRow")].map((row) => {
    const inputs   = row.querySelectorAll("input.tableInput, textarea.tableInput");
    const content  = inputs[0]?.value ?? "";
    const quantity = inputs[1]?.value ?? "";
    const unit     = row.classList.contains("noUnitRow") ? "" : (inputs[2]?.value ?? "");
    return { content, quantity, unit };
  }).filter((r) => r.content || r.quantity);
}

function escapeXml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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
        error?.message     ?? "",
        error?.code        ?? "",
        error?.name        ?? "",
        error?.debugInfo   ? JSON.stringify(error.debugInfo)   : "",
        error?.innerError  ? JSON.stringify(error.innerError)  : "",
        error?.traceMessages ? error.traceMessages.join(" / ") : "",
      ].filter(Boolean);
      const msg = parts.join(" | ") || String(error);
      setResult({ text: "エラー：" + msg });
      await placeErrorTextOnSlide(msg);
    }
  });
}

/**
 * エラーメッセージをスライド上に図形テキストとして配置する。
 * デバッグ用。エラー内容の把握に使用する。
 */
async function placeErrorTextOnSlide(message) {
  await PowerPoint.run(async (context) => {
    // 現在のスライドを取得
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

    // 四角形図形にテキストを入れる（addTextBoxより互換性が高い）
    const shape = slide.shapes.addGeometricShape("rectangle");
    shape.left   = 50;
    shape.top    = 50;
    shape.width  = 600;
    shape.height = 120;
    shape.fill.setSolidColor("FFEEEE");

    await context.sync();

    shape.textFrame.textRange.text       = "DEBUG ERROR:\n" + message;
    shape.textFrame.textRange.font.size  = 10;
    shape.textFrame.textRange.font.color = "CC0000";
    shape.textFrame.textRange.font.bold  = true;
    shape.textFrame.textRange.font.name  = "Meiryo";

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
