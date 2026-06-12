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
      { label: "割れ",     content: "タイル割れ",     unit: "枚数" },
      { label: "浮き",     content: "タイル浮き",     unit: "枚数" },
      { label: "陶片浮き", content: "タイル陶片浮き", unit: "枚数" },
      { label: "下地浮き", content: "タイル下地浮き", unit: "枚数" },
      { label: "欠損",     content: "タイル欠損",     unit: "枚数" },
      { label: "剝落",     content: "タイル剝落",     unit: "枚数" },
    ]
  },
  {
    label: "塗装面",
    items: [
      { label: "ひび割れ 0.3mm未満", content: "塗装面ひび割れ 0.3mm未満", unit: "長さ（m）"  },
      { label: "ひび割れ 0.3mm以上", content: "塗装面ひび割れ 0.3mm以上", unit: "長さ（m）"  },
      { label: "亀甲割れ",           content: "塗装面亀甲割れ",           unit: "面積（㎡）" },
      { label: "塗膜浮き",           content: "塗膜浮き",                 unit: "面積（㎡）" },
      { label: "塗膜剥離",           content: "塗膜剥離",                 unit: "面積（㎡）" },
      { label: "モルタル浮き",       content: "塗装面モルタル浮き",       unit: "面積（㎡）" },
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

Office.onReady(async () => {
  initializeTabs();
  initializeTableBuilder();
  initializeSummaryButtons();

  // ── API バージョン確認用（確認後に削除） ──
  const v18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  const v19 = Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
  await placeApiVersionOnSlide(v18, v19);
});

async function placeApiVersionOnSlide(v18, v19) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      const slide = slides.items[0];

      const shape = slide.shapes.addGeometricShape("rectangle");
      Object.assign(shape, { left: 50, top: 50, width: 300, height: 60 });
      shape.fill.setSolidColor("E8F4FD");
      await context.sync();

      shape.textFrame.textRange.text =
        "API v1.8: " + (v18 ? "✓" : "✗") + "  /  v1.9: " + (v19 ? "✓" : "✗");
      Object.assign(shape.textFrame.textRange.font, {
        size: 11, color: "003366", bold: true, name: "Meiryo"
      });
      await context.sync();
    });
  } catch { /* スライドが開いていない場合は無視 */ }
}

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
  bindClick("outputTableButton",    () => outputTableToSlide(false));
  bindClick("outputSumTableButton", () => outputTableToSlide(true));
  bindClick("autoSumButton",        openAutoSumPopup);
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
      button.textContent = item.label ?? item.content;
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

  // ハンドル列：上下2つの色選択丸ボタン
  const handleCol = document.createElement("div");
  handleCol.className = "handle-col";

  const dotTop = document.createElement("button");
  dotTop.type = "button";
  dotTop.className = "color-dot";
  dotTop.title = "フォント色を選択";
  dotTop.dataset.colorType = "font";
  dotTop.addEventListener("click", () => openColorPicker(dotTop, "font"));

  const dotBottom = document.createElement("button");
  dotBottom.type = "button";
  dotBottom.className = "color-dot";
  dotBottom.title = "背景色を選択";
  dotBottom.dataset.colorType = "fill";
  dotBottom.addEventListener("click", () => openColorPicker(dotBottom, "fill"));

  handleCol.append(dotTop, dotBottom);
  if (noUnit) handleCol.classList.add("handle-col--hidden");

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
    handleCol,
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

// ─── 自動集計ポップアップ ─────────────────────────────────

function openAutoSumPopup() {
  const rows = collectTableRows();
  if (rows.length === 0) { setResult({ text: "行がありません。" }); return; }

  closeAutoSumPopup();

  const overlay = document.createElement("div");
  overlay.id = "autoSumPopup";
  overlay.className = "autosum-popup";
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeAutoSumPopup();
  });

  const inner = document.createElement("div");
  inner.className = "autosum-popup__inner";

  // ヘッダー
  const header = document.createElement("div");
  header.className = "autosum-popup__header";
  header.textContent = "自動集計";
  inner.appendChild(header);

  // 本体：行ごとに丸2つ + 内容テキスト + 削除ボタン
  const body = document.createElement("div");
  body.className = "autosum-popup__body";

  // DOM行との対応を保持（丸の色を参照するため）
  const tableRowEls = [...document.querySelectorAll(".tableInputRow")];

  rows.forEach((rowData, i) => {
    const domRow    = tableRowEls[i];
    const dotTop    = domRow?.querySelector(".color-dot[data-color-type='font']");
    const dotBot    = domRow?.querySelector(".color-dot[data-color-type='fill']");
    const fontColor = dotTop?.dataset.color || null;
    const fillColor = dotBot?.dataset.color || null;

    const rowEl = document.createElement("div");
    rowEl.className = "autosum-row";

    // noUnit行は丸を非表示
    const showDots = !rowData.noUnit;

    const dots = document.createElement("div");
    dots.className = showDots ? "autosum-dots" : "autosum-dots autosum-dots--hidden";

    const makeDot = (type, color) => {
      const d = document.createElement("button");
      d.type = "button";
      d.className = "color-dot";
      d.dataset.colorType = type;
      d.title = type === "font" ? "文字色を変更" : "背景色を変更";
      // 色未設定なら背景は設定しない（白+ボーダーのデフォルト外観を保つ）
      if (color) { d.style.background = color; d.dataset.color = color; }
      if (showDots) {
        d.addEventListener("click", () => {
          openColorPicker(d, type, (selected) => {
            const target = domRow?.querySelector(`.color-dot[data-color-type='${type}']`);
            if (target) { target.style.background = selected; target.dataset.color = selected; }
            const contentEl = rowEl.querySelector(".autosum-content");
            if (contentEl) {
              if (type === "font") contentEl.style.color = selected;
              else contentEl.style.background = selected;
            }
          });
        });
      }
      return d;
    };

    dots.append(makeDot("font", fontColor), makeDot("fill", fillColor));
    rowEl.appendChild(dots);

    // 内容テキスト
    const content = document.createElement("span");
    content.className = "autosum-content";
    content.textContent = rowData.content;
    if (fontColor) content.style.color = fontColor;
    if (fillColor) content.style.background = fillColor;
    rowEl.appendChild(content);

    // ポップアップ内削除ボタン
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "autosum-delete-btn";
    delBtn.textContent = "×";
    delBtn.addEventListener("click", () => rowEl.remove());
    rowEl.appendChild(delBtn);

    body.appendChild(rowEl);
  });
  inner.appendChild(body);

  // フッター
  const footer = document.createElement("div");
  footer.className = "autosum-popup__footer";

  const runBtn = document.createElement("button");
  runBtn.type = "button";
  runBtn.className = "autosum-run-btn";
  runBtn.textContent = "集計実行";
  runBtn.addEventListener("click", async () => {
    closeAutoSumPopup();
    await runAutoSum(rows, tableRowEls);
  });

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "autosum-close-btn";
  closeBtn.textContent = "閉じる";
  closeBtn.addEventListener("click", closeAutoSumPopup);

  footer.append(runBtn, closeBtn);
  inner.appendChild(footer);

  overlay.appendChild(inner);
  document.body.appendChild(overlay);
}

function closeAutoSumPopup() {
  document.getElementById("autoSumPopup")?.remove();
}

/**
 * 自動集計実行：各行の丸の色を使って集計し、数量欄に結果を入力する。
 */
async function runAutoSum(rows, tableRowEls) {
  setResult({ text: "自動集計中..." });

  for (let i = 0; i < rows.length; i++) {
    const domRow   = tableRowEls[i];
    const dotTop   = domRow?.querySelector(".color-dot[data-color-type='font']");
    const dotBot   = domRow?.querySelector(".color-dot[data-color-type='fill']");
    const fontColor = dotTop?.dataset.color;
    const fillColor = dotBot?.dataset.color;
    const quantityInput = domRow?.querySelector(".quantityInput");

    if (!fontColor || !quantityInput) continue;

    try {
      let result;
      if (fillColor) {
        result = await sumNumbersByColorDirect(fontColor, fillColor);
      } else {
        result = await sumNumbersByColorDirect(fontColor, null);
      }
      quantityInput.value = result !== null ? String(result) : "";
    } catch (e) {
      console.error("集計エラー:", e);
    }
  }

  setResult({ text: "自動集計が完了しました" });
}

/**
 * 指定色（フォント色 + オプションで背景色）で数字を合計する。
 */
async function sumNumbersByColorDirect(targetTextColor, targetFillColor) {
  return PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return null;

    const textShapes = await getTextShapes(
      context, slide,
      { includeFillColor: !!targetFillColor }
    );

    const targets = targetFillColor
      ? textShapes.filter((s) => s.fillColor === targetFillColor)
      : textShapes;

    const numbers = [];
    for (const item of targets) {
      const coloredText = await extractTextByColor(
        context, item.textRange, item.text, targetTextColor
      );
      numbers.push(...extractNumbers(coloredText));
    }
    return sum(numbers);
  });
}

// ─── カラーピッカー ──────────────────────────────────────

/**
 * スライド上の色を収集してポップアップを表示し、選択色を丸ボタンに反映する。
 * @param {HTMLElement} dotEl  - クリックされた丸ボタン
 * @param {"font"|"fill"} colorType - 取得する色の種類
 */
async function openColorPicker(dotEl, colorType, onSelect = null) {
  const title = colorType === "font"
    ? "集計したい文字色を選択"
    : "集計したい背景色を選択";

  // 既存のポップアップを閉じる
  closeColorPicker();

  let colors = [];
  try {
    colors = await collectSlideColors(colorType);
  } catch (e) {
    console.error("色の取得に失敗しました:", e);
  }

  // ポップアップを生成
  const popup = document.createElement("div");
  popup.id = "colorPickerPopup";
  popup.className = "color-picker-popup";

  const titleEl = document.createElement("p");
  titleEl.className = "color-picker-title";
  titleEl.textContent = title;
  popup.appendChild(titleEl);

  const grid = document.createElement("div");
  grid.className = "color-picker-grid";

  if (colors.length === 0) {
    const empty = document.createElement("p");
    empty.className = "color-picker-empty";
    empty.textContent = "色が見つかりませんでした";
    popup.appendChild(empty);
  } else {
    colors.forEach((color) => {
      const swatch = document.createElement("button");
      swatch.type = "button";
      swatch.className = "color-swatch";
      swatch.title = color;
      swatch.style.background = color;
      swatch.addEventListener("click", () => {
        dotEl.style.background = color;
        dotEl.dataset.color = color;
        closeColorPicker();
        if (onSelect) onSelect(color);
      });
      grid.appendChild(swatch);
    });
    popup.appendChild(grid);
  }

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "color-picker-close";
  closeBtn.textContent = "閉じる";
  closeBtn.addEventListener("click", closeColorPicker);
  popup.appendChild(closeBtn);

  // ポップアップをdotElの近くに配置
  document.body.appendChild(popup);
  const rect = dotEl.getBoundingClientRect();
  popup.style.top  = (rect.bottom + window.scrollY + 4) + "px";
  popup.style.left = Math.max(4, rect.left + window.scrollX - popup.offsetWidth / 2) + "px";

  // 外側クリックで閉じる
  setTimeout(() => {
    document.addEventListener("click", onOutsideClick);
  }, 0);
}

function onOutsideClick(e) {
  const popup = document.getElementById("colorPickerPopup");
  if (popup && !popup.contains(e.target)) {
    closeColorPicker();
  }
}

function closeColorPicker() {
  document.getElementById("colorPickerPopup")?.remove();
  document.removeEventListener("click", onOutsideClick);
}

/**
 * 現在のスライド上のテキストボックスからフォント色または背景色を収集する。
 */
async function collectSlideColors(colorType) {
  return PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return [];

    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();

    const candidates = [];
    for (const shape of shapes.items) {
      if (colorType === "fill") {
        shape.fill.load("foregroundColor");
        candidates.push({ shape, type: "fill" });
      } else {
        try {
          const tf = shape.getTextFrameOrNullObject();
          tf.load("hasText");
          candidates.push({ shape, tf, type: "font" });
        } catch { /* スキップ */ }
      }
    }
    await context.sync();

    const colorSet = new Set();

    for (const item of candidates) {
      if (item.type === "fill") {
        const c = normalizeColor(item.shape.fill.foregroundColor);
        if (c && c !== "#ffffff" && c !== "#000000") colorSet.add(c);
      } else {
        if (item.tf.isNullObject || !item.tf.hasText) continue;
        // テキストを文字単位で色取得
        const textRange = item.tf.textRange;
        textRange.load("text");
        await context.sync();
        const text = textRange.text || "";
        const charRanges = Array.from({ length: Math.min(text.length, 200) }, (_, i) => {
          const cr = textRange.getSubstring(i, 1);
          cr.font.load("color");
          return cr;
        });
        await context.sync();
        charRanges.forEach((cr) => {
          const c = normalizeColor(cr.font.color);
          if (c) colorSet.add(c);
        });
      }
    }

    return [...colorSet];
  });
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
async function outputTableToSlide(includeQuantity = false) {
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

    // mergedAreas：noUnit行の数量列(c=1)〜単位列(c=2) を結合
    const mergedAreas = [];
    rows.forEach((row, i) => {
      if (row.noUnit) {
        mergedAreas.push({ rowIndex: i + 1, columnIndex: 1, rowCount: 1, columnCount: 2 });
      }
    });

    // values：テキストを2次元配列で渡す
    const values = Array.from({ length: rowCount }, (_, r) => {
      if (r === 0) return [...headers];
      const d = rows[r - 1];
      // 凡例表出力（includeQuantity=false）：数量は空欄。写真番号行のみ固定テキスト
      let qty = "";
      if (includeQuantity) {
        qty = d.quantity;
      } else if (d.pinBottom) {
        qty = "A-1, B-1, C-1...";
      }
      return [d.content, qty, d.unit ?? ""];
    });

    // specificCellProperties：フォント・罫線・背景・余白を設定する
    const specificCellProperties = Array.from({ length: rowCount }, (_, r) => {
      const isHeader = r === 0;
      const isNoUnit = !isHeader && rows[r - 1].noUnit;

      return Array.from({ length: 3 }, (_, c) => {
        if (isNoUnit && c === 2) return {};

        const borders = isHeader
          ? { top: noBorder, left: noBorder, right: noBorder, bottom: solidBorder }
          : { top: solidBorder, left: solidBorder, right: solidBorder, bottom: solidBorder };

        return {
          fill:    { color: "FFFFFF" },
          font:    { name: "Meiryo", size: 9, bold: true, color: "000000" },
          margins: { top: 0, bottom: 0, left: 5, right: 0 },
          borders
        };
      });
    });

    const tableShape = slide.shapes.addTable(rowCount, 3, {
      left:    30,
      top:     120,
      columns: colWidths.map((w) => ({ columnWidth: w })),
      rows:    Array.from({ length: rowCount }, () => ({ rowHeight })),
      values,
      mergedAreas,
      specificCellProperties
    });

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
    const inputs    = row.querySelectorAll("input.tableInput, textarea.tableInput");
    const noUnit    = row.classList.contains("noUnitRow");
    const pinBottom = row.dataset.pinBottom === "true";
    const content   = inputs[0]?.value ?? "";
    const quantity  = inputs[1]?.value ?? "";
    const unit      = noUnit ? "" : (inputs[2]?.value ?? "");
    return { content, quantity, unit, noUnit, pinBottom };
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
