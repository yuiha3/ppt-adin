const COLORS = {
  BLUE:  "#0070c0",
  GREEN: "#00b050",
  BLACK: "#000000",
  RED:   "#ff0000"
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

// pdf.js CDN ワーカー設定（index.htmlでスクリプト読み込み後に設定される）
const PDFJS_VERSION = "3.11.174";
// PDF挿入修正版: contain配置で左上タイトル・端部の欠けを防止

Office.onReady(() => {
  initializeTabs();
  initializeTableBuilder();
  initializeSummaryButtons();
  initializePdfImport();
});

// ─── 初期化 ──────────────────────────────────────────────

// ─── PDF取込 ─────────────────────────────────────────────

let pdfPageImages = []; // {base64: string, width: number, height: number}[]

function initializePdfImport() {
  const dropZone    = document.getElementById("pdfDropZone");
  const fileInput   = document.getElementById("pdfFileInput");
  const insertBtn   = document.getElementById("pdfInsertButton");
  if (!dropZone || !fileInput || !insertBtn) return;

  // ドラッグ&ドロップ
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("pdf-drop-zone--active");
  });
  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("pdf-drop-zone--active");
  });
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("pdf-drop-zone--active");
    const file = e.dataTransfer.files[0];
    if (file?.type === "application/pdf") loadPdf(file);
    else showPdfStatus("PDFファイルをドロップしてください。", "error");
  });

  // ファイル選択
  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    if (file) loadPdf(file);
  });

  // スライド挿入ボタン
  insertBtn.addEventListener("click", insertPdfToSlides);
}

function showPdfStatus(message, type = "info") {
  const el = document.getElementById("pdfStatus");
  if (!el) return;
  el.textContent = message;
  el.className = `pdf-status pdf-status--${type}`;
  el.style.display = "block";
}

async function loadPdf(file) {
  showPdfStatus("PDFを読み込んでいます...", "info");
  document.getElementById("pdfPreviewArea").style.display = "none";
  document.getElementById("pdfPreviewList").innerHTML = "";
  pdfPageImages = [];

  try {
    // pdf.js が読み込まれているか確認
    if (typeof pdfjsLib === "undefined") {
      showPdfStatus("pdf.jsが読み込まれていません。", "error");
      return;
    }
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.worker.min.js`;

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({
      data: arrayBuffer,
      // CJKフォント（日本語・中国語・韓国語）を含むPDFのレンダリングに必要
      cMapUrl:    `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/cmaps/`,
      cMapPacked: true,
    }).promise;
    const numPages = pdf.numPages;

    showPdfStatus(`${numPages}ページを変換中...`, "info");

    const previewList = document.getElementById("pdfPreviewList");

    for (let i = 1; i <= numPages; i++) {
      showPdfStatus(`変換中: ${i} / ${numPages} ページ`, "info");

      const page = await pdf.getPage(i);
      // scale: 8.0 = 576dpi相当（A3サイズ対応）
      const TARGET_SCALE = 8.0;
      const viewport = page.getViewport({ scale: TARGET_SCALE });

      // Canvasサイズをviewportと完全一致させる（transform不要）
      // Math.floorによるわずかなずれもなくすためceil→整数化
      const canvasWidth  = Math.ceil(viewport.width);
      const canvasHeight = Math.ceil(viewport.height);

      const canvas = document.createElement("canvas");
      canvas.width  = canvasWidth;
      canvas.height = canvasHeight;

      await page.render({
        canvasContext: canvas.getContext("2d"),
        viewport
      }).promise;

      const base64 = canvas.toDataURL("image/png").split(",")[1];
      pdfPageImages.push({ base64, width: viewport.width, height: viewport.height });

      // サムネイル表示（scale: 0.15で縮小）
      const thumbCanvas = document.createElement("canvas");
      const thumbScale  = 0.15;
      thumbCanvas.width  = viewport.width  * thumbScale;
      thumbCanvas.height = viewport.height * thumbScale;
      const thumbCtx = thumbCanvas.getContext("2d");
      thumbCtx.drawImage(canvas, 0, 0, thumbCanvas.width, thumbCanvas.height);

      const thumb = document.createElement("div");
      thumb.className = "pdf-thumb";
      const img = document.createElement("img");
      img.src = thumbCanvas.toDataURL("image/png");
      img.alt = `${i}ページ`;
      const label = document.createElement("p");
      label.textContent = `${i} / ${numPages}`;
      thumb.append(img, label);
      previewList.appendChild(thumb);
    }

    showPdfStatus(`${numPages}ページの変換が完了しました。`, "success");
    document.getElementById("pdfPreviewArea").style.display = "block";

  } catch (err) {
    console.error(err);
    showPdfStatus("PDFの読み込みに失敗しました: " + err.message, "error");
  }
}

async function insertPdfToSlides() {
  if (pdfPageImages.length === 0) {
    showPdfStatus("先にPDFを読み込んでください。", "error");
    return;
  }

  showPdfStatus("スライドに挿入しています...", "info");
  document.getElementById("pdfInsertButton").disabled = true;

  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();

      // スライドサイズを pageSetup（API 1.10）で取得。
      // 取得できない環境ではPowerPoint標準の16:9ワイドを仮値として使用する。
      // ※旧コードのA3縦固定は、横長PDFで左上タイトルが切れる原因になり得るため変更。
      let slideW = 960;   // 16:9 ワイド 13.333in × 7.5in 相当（pt）
      let slideH = 540;
      try {
        const pageSetup = presentation.pageSetup;
        pageSetup.load("slideWidth,slideHeight");
        await context.sync();
        if (pageSetup.slideWidth && pageSetup.slideHeight) {
          slideW = pageSetup.slideWidth;
          slideH = pageSetup.slideHeight;
        }
      } catch {
        // API 1.10未満の環境では16:9ワイド仮値を使用
      }

      for (let i = 0; i < pdfPageImages.length; i++) {
        showPdfStatus(`挿入中: ${i + 1} / ${pdfPageImages.length} ページ`, "info");

        // 追加前のスライド数を取得（add()はvoidを返すため、インデックスで取得する）
        const slideCount = slides.getCount();
        await context.sync();

        slides.add();
        await context.sync();

        // 追加後のスライドは追加前の件数のインデックスに存在する（0-based）
        const newSlide = slides.getItemAt(slideCount.value);
        newSlide.load("shapes");
        await context.sync();

        // PDFページのアスペクト比を保ちつつ、スライド内に全体を収める（contain）。
        // 旧コードのcover（Math.max）は上下左右の一部がスライド外へ出るため、左上タイトルが消える原因になる。
        const imgW = pdfPageImages[i].width;
        const imgH = pdfPageImages[i].height;
        const margin = 10; // 端の文字・線が欠けないように安全余白を確保（pt）

        const availableW = Math.max(1, slideW - margin * 2);
        const availableH = Math.max(1, slideH - margin * 2);
        const scaleW = availableW / imgW;
        const scaleH = availableH / imgH;

        // 全体が切れないように小さい方の倍率を使用
        const containScale = Math.min(scaleW, scaleH);
        const fitW = imgW * containScale;
        const fitH = imgH * containScale;

        // 中央配置。containのためleft/topは原則マイナスにならない。
        const left = Math.max(margin, (slideW - fitW) / 2);
        const top  = Math.max(margin, (slideH - fitH) / 2);

        // 可能なら画像オブジェクトとして挿入する。
        // addImageが使えない環境では、従来通り四角形の画像塗りつぶしにフォールバックする。
        if (typeof newSlide.shapes.addImage === "function") {
          const shape = newSlide.shapes.addImage(pdfPageImages[i].base64);
          shape.left = left;
          shape.top = top;
          shape.width = fitW;
          shape.height = fitH;
          await context.sync();
        } else {
          const shape = newSlide.shapes.addGeometricShape("rectangle", {
            left,
            top,
            width: fitW,
            height: fitH
          });
          await context.sync();

          shape.lineFormat.visible = false;
          shape.fill.setImage(pdfPageImages[i].base64);
          await context.sync();
        }
      }
    });

    showPdfStatus(`${pdfPageImages.length}枚のスライドを追加しました。`, "success");
  } catch (err) {
    console.error(err);
    showPdfStatus("挿入に失敗しました: " + err.message, "error");
  } finally {
    document.getElementById("pdfInsertButton").disabled = false;
  }
}

function initializeSummaryButtons() {
  bindClick("sumBlueButton",  () => sumNumbersByTextColor(COLORS.BLUE));
  bindClick("sumGreenButton", () => sumNumbersByTextColor(COLORS.GREEN));
  bindClick("sumSelectedColorButton",           sumNumbersBySelectedTextColor);
  bindClick("sumSelectedTextAndFillColorButton", sumNumbersBySelectedTextColorAndFillColor);
  bindClick("formatBlackCodeButton",            () => formatCodesByTextColor(COLORS.BLACK, false));
  bindClick("formatAllSlidesBlackCodeButton",   () => formatCodesByTextColor(COLORS.BLACK, true));
  bindClick("sumAreaByColorButton",             sumAreaBySelectedTextColor);
  bindClick("collectRedTextButton",             collectRedTextFromSlide);

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
    const details = document.createElement("details");
    details.className = "defect-group";
    if (label !== "塗装面") details.open = true;

    const summary = document.createElement("summary");
    summary.textContent = label;
    details.appendChild(summary);

    const btnGrid = document.createElement("div");
    btnGrid.className = "defect-group__buttons";

    items.forEach((item) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "defectButton";
      button.textContent = item.label ?? item.content;
      button.addEventListener("click", () => addTableRow(item));
      btnGrid.appendChild(button);
    });

    details.appendChild(btnGrid);
    container.appendChild(details);
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

// 色選択丸ボタンを生成するユーティリティ
function makeColorDot(type, color = null, onClick = null) {
  const d = document.createElement("button");
  d.type = "button";
  d.className = "color-dot";
  d.dataset.colorType = type;
  d.title = type === "font" ? "文字色を選択" : "背景色を選択";
  if (color) { d.style.background = color; d.dataset.color = color; }
  if (onClick) d.addEventListener("click", onClick);
  return d;
}

function addTableRow({ content, unit, noUnit = false, pinBottom = false }) {
  const tableRows = document.getElementById("tableRows");
  if (!tableRows) return;

  const row = document.createElement("div");
  row.className = noUnit ? "tableInputRow noUnitRow" : "tableInputRow";
  row.draggable = true;
  if (pinBottom) row.dataset.pinBottom = "true";

  const handleCol = document.createElement("div");
  handleCol.className = "handle-col";
  if (noUnit) handleCol.classList.add("handle-col--hidden");

  handleCol.append(
    makeColorDot("font", null, (e) => openColorPicker(e.currentTarget, "font")),
    makeColorDot("fill", null, (e) => openColorPicker(e.currentTarget, "fill"))
  );

  const quantityEl = noUnit
    ? Object.assign(document.createElement("textarea"), {
        className:   "tableInput quantityInput quantityTextarea",
        placeholder: "入力",
        rows:        2
      })
    : createTableInput({ className: "quantityInput", placeholder: "数量", inputmode: "numeric" });

  const deleteButton = Object.assign(document.createElement("button"), {
    type: "button", className: "row-delete-btn", textContent: "×"
  });
  deleteButton.addEventListener("click", () => row.remove());

  row.append(
    handleCol,
    createTableInput({ value: content, className: "contentInput" }),
    quantityEl,
    ...(noUnit ? [] : [createTableInput({ value: unit, className: "unitInput" })]),
    deleteButton
  );

  // 挿入位置：通常行 → その他(noUnit) → 写真番号(pinBottom) の順
  const anchor = pinBottom ? null
    : noUnit ? (tableRows.querySelector("[data-pin-bottom]") ?? null)
    : (tableRows.querySelector(".noUnitRow:not([data-pin-bottom])")
        ?? tableRows.querySelector("[data-pin-bottom]")
        ?? null);
  tableRows.insertBefore(row, anchor);

  setupDragAndDrop(row, tableRows);
  quantityEl.focus();
}

// ─── ドラッグ&ドロップ ────────────────────────────────────
let dragSrc    = null;
let dragOverEl = null;

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
  const container = document.getElementById("tableRows");
  if (!container || container.children.length === 0) return;

  showConfirmDialog(
    "行をすべて削除しますか？",
    "追加された行がすべて削除されます。この操作は元に戻せません。",
    "削除する",
    () => container.replaceChildren()
  );
}

/**
 * 確認ダイアログを表示する。
 * @param {string}   title        - タイトル
 * @param {string}   message      - 説明文
 * @param {string}   confirmLabel - 確認ボタンのラベル
 * @param {Function} onConfirm    - 確認ボタン押下時のコールバック
 */
function showConfirmDialog(title, message, confirmLabel, onConfirm) {
  document.getElementById("confirmDialog")?.remove();

  const overlay = document.createElement("div");
  overlay.id = "confirmDialog";
  overlay.className = "confirm-overlay";
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });

  const cancelBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "confirm-cancel-btn", textContent: "キャンセル"
  });
  cancelBtn.addEventListener("click", () => overlay.remove());

  const confirmBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "confirm-ok-btn", textContent: confirmLabel
  });
  confirmBtn.addEventListener("click", () => { overlay.remove(); onConfirm(); });

  const btnRow = document.createElement("div");
  btnRow.className = "confirm-dialog__buttons";
  btnRow.append(cancelBtn, confirmBtn);

  const dialog = document.createElement("div");
  dialog.className = "confirm-dialog";
  dialog.append(
    Object.assign(document.createElement("p"), { className: "confirm-dialog__title",   textContent: title }),
    Object.assign(document.createElement("p"), { className: "confirm-dialog__message", textContent: message }),
    btnRow
  );

  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
}

// ─── 自動集計ポップアップ ─────────────────────────────────

function openAutoSumPopup() {
  const rows = collectTableRows();
  if (rows.length === 0) { setResult({ text: "行がありません。" }); return; }

  closeAutoSumPopup();

  const tableRowEls = [...document.querySelectorAll(".tableInputRow")];

  const overlay = document.createElement("div");
  overlay.id = "autoSumPopup";
  overlay.className = "autosum-popup";
  overlay.addEventListener("click", (e) => { if (e.target === overlay) closeAutoSumPopup(); });

  const inner = document.createElement("div");
  inner.className = "autosum-popup__inner";

  inner.append(
    Object.assign(document.createElement("div"), { className: "autosum-popup__header", textContent: "自動集計" })
  );

  const body = document.createElement("div");
  body.className = "autosum-popup__body";
  rows.forEach((rowData, i) => body.appendChild(makeAutoSumRow(rowData, i, tableRowEls)));
  inner.appendChild(body);

  const runBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "autosum-run-btn", textContent: "集計実行"
  });
  runBtn.addEventListener("click", async () => {
    const activePopupRows = [...body.querySelectorAll(".autosum-row")];
    closeAutoSumPopup();
    await runAutoSum(activePopupRows, tableRowEls);
  });

  const closeBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "autosum-close-btn", textContent: "閉じる"
  });
  closeBtn.addEventListener("click", closeAutoSumPopup);

  const footer = document.createElement("div");
  footer.className = "autosum-popup__footer";
  footer.append(runBtn, closeBtn);

  inner.appendChild(footer);
  overlay.appendChild(inner);
  document.body.appendChild(overlay);
}

/**
 * 自動集計ポップアップの1行要素を生成する。
 */
function makeAutoSumRow(rowData, index, tableRowEls) {
  const domRow    = tableRowEls[index];
  const fontColor = domRow?.querySelector(".color-dot[data-color-type='font']")?.dataset.color ?? null;
  const fillColor = domRow?.querySelector(".color-dot[data-color-type='fill']")?.dataset.color ?? null;
  const showDots  = !rowData.noUnit;

  const rowEl = document.createElement("div");
  rowEl.className = "autosum-row";
  rowEl.dataset.rowIndex = String(index);

  const dots = document.createElement("div");
  dots.className = showDots ? "autosum-dots" : "autosum-dots autosum-dots--hidden";

  const makeDotWithSync = (type, color) =>
    makeColorDot(type, color, showDots ? () => {
      openColorPicker(dots.querySelector(`[data-color-type='${type}']`), type, (selected) => {
        // 元DOM行の丸にも反映
        const target = domRow?.querySelector(`.color-dot[data-color-type='${type}']`);
        if (target) { target.style.background = selected; target.dataset.color = selected; }
        // テキスト表示を更新
        const contentEl = rowEl.querySelector(".autosum-content");
        if (contentEl) {
          if (type === "font") contentEl.style.color = selected;
          else contentEl.style.background = selected;
        }
      });
    } : null);

  dots.append(makeDotWithSync("font", fontColor), makeDotWithSync("fill", fillColor));

  const content = Object.assign(document.createElement("span"), {
    className: "autosum-content", textContent: rowData.content
  });
  if (fontColor) content.style.color = fontColor;
  if (fillColor) content.style.background = fillColor;

  // 面積チェックボックス（noUnit行は非表示）
  const areaLabel = document.createElement("label");
  areaLabel.className = showDots ? "autosum-area-check" : "autosum-area-check autosum-area-check--hidden";
  areaLabel.title = "面積として計算する";
  const areaCheckbox = document.createElement("input");
  areaCheckbox.type = "checkbox";
  areaCheckbox.className = "autosum-area-checkbox";
  areaLabel.appendChild(areaCheckbox);

  const delBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "autosum-delete-btn", textContent: "×"
  });
  delBtn.addEventListener("click", () => rowEl.remove());

  rowEl.append(dots, content, areaLabel, delBtn);
  return rowEl;
}

function closeAutoSumPopup() {
  document.getElementById("autoSumPopup")?.remove();
}

/**
 * 自動集計実行。ポップアップ内で現在表示中の行から色を読み、数量欄に書き込む。
 */
async function runAutoSum(activePopupRows, tableRowEls) {
  setResult({ text: "自動集計中..." });

  for (const popupRow of activePopupRows) {
    const rowIndex = Number(popupRow.dataset.rowIndex);
    if (isNaN(rowIndex)) continue;

    const domRow        = tableRowEls[rowIndex];
    const quantityInput = domRow?.querySelector(".quantityInput");
    if (!quantityInput) continue;

    // 写真番号行：写真番号集計を実行
    if (domRow?.dataset.pinBottom === "true") {
      try { quantityInput.value = await collectPhotoCodesFromSlide(); }
      catch (e) { console.error("写真番号集計エラー:", e); }
      continue;
    }

    const fontColor  = popupRow.querySelector(".color-dot[data-color-type='font']")?.dataset.color ?? null;
    const fillColor  = popupRow.querySelector(".color-dot[data-color-type='fill']")?.dataset.color ?? null;
    const isAreaMode = popupRow.querySelector(".autosum-area-checkbox")?.checked ?? false;
    if (!fontColor) continue;

    try {
      if (isAreaMode) {
        const result = await sumAreaByColorDirect(fontColor, fillColor);
        if (result !== null) {
          const rounded = Math.ceil(result * 100) / 100;
          quantityInput.value = rounded.toFixed(2);
        }
      } else {
        const result = await sumNumbersByColorDirect(fontColor, fillColor);
        quantityInput.value = result !== null ? String(result) : "";
      }
    } catch (e) { console.error("集計エラー:", e); }
  }

  setResult({ text: "自動集計が完了しました" });
}

// ─── カラーピッカー ──────────────────────────────────────

async function openColorPicker(dotEl, colorType, onSelect = null) {
  const title = colorType === "font" ? "集計したい文字色を選択" : "集計したい背景色を選択";

  closeColorPicker();

  let colors = [];
  try { colors = await collectSlideColors(colorType); }
  catch (e) { console.error("色の取得に失敗しました:", e); }

  const popup = document.createElement("div");
  popup.id = "colorPickerPopup";
  popup.className = "color-picker-popup";
  popup.append(
    Object.assign(document.createElement("p"), { className: "color-picker-title", textContent: title })
  );

  if (colors.length === 0) {
    popup.append(
      Object.assign(document.createElement("p"), { className: "color-picker-empty", textContent: "色が見つかりませんでした" })
    );
  } else {
    const grid = document.createElement("div");
    grid.className = "color-picker-grid";
    colors.forEach((color) => {
      const swatch = Object.assign(document.createElement("button"), {
        type: "button", className: "color-swatch", title: color
      });
      swatch.style.background = color;
      swatch.addEventListener("click", () => {
        dotEl.style.background = color;
        dotEl.dataset.color = color;
        closeColorPicker();
        onSelect?.(color);
      });
      grid.appendChild(swatch);
    });
    popup.appendChild(grid);
  }

  const closeBtn = Object.assign(document.createElement("button"), {
    type: "button", className: "color-picker-close", textContent: "閉じる"
  });
  closeBtn.addEventListener("click", closeColorPicker);
  popup.appendChild(closeBtn);

  document.body.appendChild(popup);
  const rect = dotEl.getBoundingClientRect();
  popup.style.top  = (rect.bottom + window.scrollY + 4) + "px";
  popup.style.left = Math.max(4, rect.left + window.scrollX - popup.offsetWidth / 2) + "px";

  setTimeout(() => document.addEventListener("click", onOutsideClick), 0);
}

function onOutsideClick(e) {
  const popup = document.getElementById("colorPickerPopup");
  if (popup && !popup.contains(e.target)) closeColorPicker();
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

    const colorSet = new Set();

    if (colorType === "fill") {
      shapes.items.forEach((s) => s.fill.load("foregroundColor"));
      await context.sync();
      shapes.items.forEach((s) => {
        const c = normalizeColor(s.fill.foregroundColor);
        if (c && c !== "#ffffff" && c !== "#000000") colorSet.add(c);
      });
    } else {
      for (const item of await getTextShapes(context, slide)) {
        const charRanges = Array.from({ length: Math.min(item.text.length, 200) }, (_, i) => {
          const cr = item.textRange.getSubstring(i, 1);
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
 *   全セル：背景白、余白 上下0 左5pt、下寄せ、フォント9pt 黒 太字 メイリオ
 *   行高さ：13.5pt
 */
async function outputTableToSlide(includeQuantity = false) {
  const rows = collectTableRows();

  if (rows.length === 0) { setResult({ text: "行がありません。" }); return; }

  setResult({ text: "出力中..." });

  // 集計表出力時：各DOM行の上の丸（フォント色）を取得しておく
  const fontColors = includeQuantity
    ? [...document.querySelectorAll(".tableInputRow")].map((domRow) =>
        domRow.querySelector(".color-dot[data-color-type='font']")?.dataset.color ?? null
      )
    : [];

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const rowCount  = rows.length + 1;
    const colWidths = [80, 60, 105];
    const rowHeight = 13.5;

    const solidBorder = { color: "000000", dashStyle: "solid", weight: 1 };
    const noBorder    = { color: "000000", dashStyle: "solid", weight: 0 };

    const mergedAreas = rows
      .map((r, i) => r.noUnit ? { rowIndex: i + 1, columnIndex: 1, rowCount: 1, columnCount: 2 } : null)
      .filter(Boolean);

    const values = Array.from({ length: rowCount }, (_, r) => {
      if (r === 0) return [includeQuantity ? "＜集計＞" : "＜凡例＞", "", ""];
      const d = rows[r - 1];
      const qty = includeQuantity ? d.quantity : (d.pinBottom ? "A-1, B-1, C-1..." : "");
      return [d.content, qty, d.unit ?? ""];
    });

    const baseStyle = {
      fill:              { color: "FFFFFF" },
      font:              { name: "Meiryo", size: 9, bold: true, color: "000000" },
      margins:           { top: 0, bottom: 0, left: 5, right: 0 },
      verticalAlignment: "Bottom"
    };

    const specificCellProperties = Array.from({ length: rowCount }, (_, r) => {
      const isHeader    = r === 0;
      const isNoUnit    = !isHeader && rows[r - 1].noUnit;
      const isPinBottom = !isHeader && rows[r - 1].pinBottom;

      return Array.from({ length: 3 }, (_, c) => {
        // 結合エリア内の非左上セル
        if (isNoUnit && c === 2) return {};

        const borders = isHeader
          ? { top: noBorder, left: noBorder, right: noBorder, bottom: solidBorder }
          : { top: solidBorder, left: solidBorder, right: solidBorder, bottom: solidBorder };

        // 数量列（c=1）：写真番号行以外は中央寄せ
        const hAlign = (c === 1 && !isHeader && !isPinBottom) ? "Center" : undefined;

        // 集計表出力のデータ行数量列：丸の色をフォントカラーに
        const fontColor = (includeQuantity && !isHeader && c === 1)
          ? (fontColors[r - 1]?.replace("#", "") ?? "000000")
          : undefined;

        return {
          ...baseStyle,
          ...(hAlign    ? { horizontalAlignment: hAlign } : {}),
          ...(fontColor ? { font: { ...baseStyle.font, color: fontColor } } : {}),
          borders
        };
      });
    });

    const tableShape = slide.shapes.addTable(rowCount, 3, {
      left: 30, top: 120,
      columns: colWidths.map((w) => ({ columnWidth: w })),
      rows:    Array.from({ length: rowCount }, () => ({ rowHeight })),
      values, mergedAreas, specificCellProperties
    });

    await context.sync();

    // specificCellProperties の horizontalAlignment が効かない環境向けの保険（API 1.9）
    try {
      const table = tableShape.getTable();
      table.load();
      await context.sync();
      for (let r = 1; r < rowCount; r++) {
        if (rows[r - 1].pinBottom) continue;
        const cell = table.getCellOrNullObject(r, 1);
        if (!cell.isNullObject) cell.horizontalAlignment = "Center";
      }
      await context.sync();
    } catch { /* 1.9未満の環境では無視 */ }

    setResult({ text: "スライドに出力しました" });
  });
}

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

    shape.textFrame.textRange.text = "DEBUG ERROR:\n" + message;
    Object.assign(shape.textFrame.textRange.font, { size: 10, color: "CC0000", bold: true, name: "Meiryo" });
    await context.sync();
  });
}

// ─── 集計処理 ─────────────────────────────────────────────

async function sumNumbersByTextColor(targetTextColor, decimalPlaces = null) {
  setResult({ text: "集計中..." });

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const numbers = [];
    const hitIds  = [];
    for (const item of await getTextShapes(context, slide)) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      const found = extractNumbers(coloredText);
      if (found.length > 0) { numbers.push(...found); hitIds.push(item.shapeId); }
    }

    if (hitIds.length > 0) { slide.setSelectedShapes(hitIds); await context.sync(); }

    const total     = sum(numbers);
    const totalStr  = decimalPlaces !== null ? total.toFixed(decimalPlaces) : String(total);
    setResult({ text: totalStr, color: targetTextColor, copyValue: totalStr });
  });
}

async function sumNumbersBySelectedTextColor() {
  setResult({ text: "選択中の文字色を取得中..." });
  const selected = await getSelectedTextInfo();
  if (!selected.ok) { setResult({ text: selected.message }); return; }
  await sumNumbersByTextColor(selected.textColor, selected.decimalPlaces ?? null);
}

/**
 * 現在のスライドの赤文字（R255,G0,B0）を全て収集して「, 」つなぎで表示する。
 */
async function collectRedTextFromSlide() {
  setResult({ text: "赤文字を収集中..." });

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const texts = [];
    for (const item of await getTextShapes(context, slide)) {
      const coloredText = await extractTextByColor(
        context, item.textRange, item.text, COLORS.RED
      );
      // 色一致部分（スペース以外）をトークン単位で収集
      coloredText.split(/\s+/)
        .map((t) => t.trim())
        .filter((t) => t.length > 0)
        .forEach((t) => texts.push(t));
    }

    const outputText = texts.length > 0 ? texts.join(", ") : "該当なし";
    setResult({ text: outputText, color: COLORS.RED, copyValue: outputText });
  });
}

/**
 * 選択中テキストの文字色で「0.5x1.5」形式の面積を合計する。
 * 必ず4文字目（index=3）が「x」であることを前提に前半・後半を掛け算する。
 */
async function sumAreaBySelectedTextColor() {
  setResult({ text: "選択中の文字色を取得中..." });
  const selected = await getSelectedTextInfo();
  if (!selected.ok) { setResult({ text: selected.message }); return; }

  setResult({ text: "面積を集計中..." });

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    let total = 0;
    const hitIds = [];
    for (const item of await getTextShapes(context, slide)) {
      const coloredText = await extractTextByColor(
        context, item.textRange, item.text, selected.textColor
      );
      const found = extractAreaValues(coloredText);
      if (found > 0) { total += found; hitIds.push(item.shapeId); }
    }

    if (hitIds.length > 0) { slide.setSelectedShapes(hitIds); await context.sync(); }

    const result    = Math.ceil(total * 100) / 100;
    const resultStr = result.toFixed(2);
    setResult({ text: resultStr, color: selected.textColor, copyValue: resultStr });
  });
}

/**
 * テキストから「0.5x1.5」形式（4文字目がx）を全て抽出して掛け算し合計する。
 */
function extractAreaValues(text) {
  let total = 0;
  // 4文字目（index=3）が x であるパターン：前半3文字 x 後半
  const matches = [...text.matchAll(/(\S{3})x(\S+)/g)];
  for (const m of matches) {
    const a = parseFloat(m[1]);
    const b = parseFloat(m[2]);
    if (!isNaN(a) && !isNaN(b)) total += a * b;
  }
  return total;
}

async function sumNumbersBySelectedTextColorAndFillColor() {
  setResult({ text: "選択中の文字色と背景色を取得中..." });
  const selected = await getSelectedTextAndFillInfo();
  if (!selected.ok) { setResult({ text: selected.message }); return; }

  await PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return;

    const numbers = [];
    const hitIds  = [];
    for (const item of (await getTextShapes(context, slide, { includeFillColor: true }))
        .filter((s) => s.fillColor === selected.fillColor)) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, selected.textColor);
      const found = extractNumbers(coloredText);
      if (found.length > 0) { numbers.push(...found); hitIds.push(item.shapeId); }
    }

    if (hitIds.length > 0) { slide.setSelectedShapes(hitIds); await context.sync(); }

    const total = String(sum(numbers));
    setResult({ text: total, color: selected.textColor, copyValue: total });
  });
}

async function sumNumbersByColorDirect(targetTextColor, targetFillColor) {
  return PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return null;

    const textShapes = await getTextShapes(context, slide, { includeFillColor: !!targetFillColor });
    const targets = targetFillColor
      ? textShapes.filter((s) => s.fillColor === targetFillColor)
      : textShapes;

    const numbers = [];
    const hitIds  = [];
    for (const item of targets) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      const found = extractNumbers(coloredText);
      if (found.length > 0) { numbers.push(...found); hitIds.push(item.shapeId); }
    }

    if (hitIds.length > 0) { slide.setSelectedShapes(hitIds); await context.sync(); }

    return sum(numbers);
  });
}

/**
 * 指定色（フォント色 + オプションで背景色）で面積（数値x数値形式）を合計する（自動集計用）。
 */
async function sumAreaByColorDirect(targetTextColor, targetFillColor) {
  return PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return null;

    const textShapes = await getTextShapes(context, slide, { includeFillColor: !!targetFillColor });
    const targets = targetFillColor
      ? textShapes.filter((s) => s.fillColor === targetFillColor)
      : textShapes;

    let total = 0;
    for (const item of targets) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      total += extractAreaValues(coloredText);
    }
    return total;
  });
}

/**
 * A-1形式コードをスライドから収集してアルファベット別に連番圧縮する。
 * formatCodesByTextColor・collectPhotoCodesFromSlide の共通ロジック。
 */
async function collectGroupedCodes(context, slides, targetTextColor) {
  const grouped = {};
  for (const slide of slides) {
    for (const item of await getTextShapes(context, slide)) {
      const coloredText = await extractTextByColor(context, item.textRange, item.text, targetTextColor);
      for (const { letter, number } of extractLetterCodes(coloredText)) {
        (grouped[letter] ??= []).push(number);
      }
    }
  }
  return Object.keys(grouped).sort()
    .map((letter) => `${letter}-${compressNumberRanges(grouped[letter])}`);
}

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

    const outputLines = await collectGroupedCodes(context, slides, targetTextColor);
    const outputText  = outputLines.length > 0 ? outputLines.join("\n") : "該当なし";

    setResult({
      html:      `<pre class="outputText" style="color:${targetTextColor}">${escapeHtml(outputText)}</pre>`,
      color:     targetTextColor,
      copyValue: outputText
    });
  });
}

async function collectPhotoCodesFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = await getCurrentSlide(context);
    if (!slide) return "";
    const outputLines = await collectGroupedCodes(context, [slide], COLORS.BLACK);
    return outputLines.join("\n");
  });
}

// ─── PowerPoint API ───────────────────────────────────────

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

  // 選択テキストの小数桁数を取得（例: "1.50" → 2, "3" → 0）
  const selectedText = selectedTextRange.text.trim();
  const decimalPlaces = getDecimalPlaces(selectedText);

  return { ok: true, textColor, decimalPlaces };
}

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

  // shapeId を事前にロード
  textShapes.forEach((item) => item.shape.load("id"));
  await context.sync();

  for (const item of textShapes) {
    item.text    = item.textRange.text || "";
    item.shapeId = item.shape.id;
    if (options.includeFillColor) {
      item.fillColor = await getShapeFillColor(context, item.shape);
    }
  }

  return textShapes.filter((item) => item.text.length > 0);
}

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

/**
 * テキストから小数桁数を取得する。
 * 数値に変換できるテキストの小数点以下の桁数を返す。
 * 変換できない場合は null を返す。
 */
function getDecimalPlaces(text) {
  const trimmed = text.trim();
  const num = Number(trimmed);
  if (isNaN(num)) return null;
  const dotIndex = trimmed.indexOf(".");
  if (dotIndex === -1) return 0;
  return trimmed.length - dotIndex - 1;
}

function extractNumbers(text) {
  // extractTextByColor は色違い文字をスペースに置換する。
  // 「1.0x1.4」はトークン全体が数値でないので除外されるが、
  // x が別色の場合「1.0 1.4」になるケースも防ぐため、
  // まず「数値 x 数値」「数値x数値」の面積形式パターンを除去してから評価する。
  const cleaned = text.replace(/\d+(?:\.\d+)?\s*[xX]\s*\d+(?:\.\d+)?/g, " ");
  return cleaned.split(/\s+/)
    .filter((t) => t.length > 0)
    .map((t) => Number(t))
    .filter((v) => !isNaN(v));
}

function extractLetterCodes(text) {
  return [...text.matchAll(/\b([A-Z])-(\d+)\b/g)].map((m) => ({
    letter: m[1],
    number: Number(m[2])
  }));
}

function compressNumberRanges(numbers) {
  if (!numbers?.length) return "";

  // 重複はそのまま保持してソート
  const sorted = [...numbers].sort((a, b) => a - b);
  const ranges = [];
  let i = 0;

  while (i < sorted.length) {
    // 重複している値は1つずつ個別出力
    if (i + 1 < sorted.length && sorted[i] === sorted[i + 1]) {
      ranges.push(String(sorted[i]));
      i++;
      continue;
    }
    // 重複なし：連番を探して圧縮
    // ただし次が重複値（sorted[i+1] === sorted[i+2]）の場合はここでは連番に含めない
    let start = sorted[i];
    let prev  = sorted[i];
    while (
      i + 1 < sorted.length &&
      sorted[i + 1] === prev + 1 &&
      // 次の値が更にその次と同じ（重複）なら連番に含めず止める
      sorted[i + 1] !== sorted[i + 2]
    ) {
      i++;
      prev = sorted[i];
    }
    ranges.push(formatRange(start, prev));
    i++;
  }

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
