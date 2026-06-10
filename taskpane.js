Office.onReady(() => {
  const blueButton = document.getElementById("sumButton");
  const greenButton = document.getElementById("sumGreenButton");
  const selectedColorButton = document.getElementById("sumSelectedColorButton");
  const formatBlackCodeButton = document.getElementById("formatBlackCodeButton");

  if (blueButton) {
    blueButton.addEventListener("click", async () => {
      await sumNumbersByColorOnCurrentSlide("#0070c0", "青文字 RGB(0,112,192)");
    });
  }

  if (greenButton) {
    greenButton.addEventListener("click", async () => {
      await sumNumbersByColorOnCurrentSlide("#00b050", "緑文字 RGB(0,176,80)");
    });
  }

  if (selectedColorButton) {
    selectedColorButton.addEventListener("click", async () => {
      await sumNumbersBySelectedTextColor();
    });
  }

  if (formatBlackCodeButton) {
    formatBlackCodeButton.addEventListener("click", async () => {
      await formatCodesByColorOnCurrentSlide("#000000", "黒文字 RGB(0,0,0)");
    });
  }
});

async function sumNumbersBySelectedTextColor() {
  const result = document.getElementById("result");
  const targetColorElement = document.getElementById("targetColor");

  result.textContent = "選択中の文字色を取得中...";
  targetColorElement.textContent = "対象色：選択中の文字色";

  try {
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      const selectedShapeCount = selectedShapes.getCount();
      const selectedTextRange = context.presentation.getSelectedTextRangeOrNullObject();

      await context.sync();

      if (selectedShapeCount.value > 1) {
        result.textContent = "複数選択されています。色を取得したいテキストを1つだけ選択してください。";
        return;
      }

      selectedTextRange.load("text");
      selectedTextRange.font.load("color");

      await context.sync();

      if (selectedTextRange.isNullObject) {
        result.textContent = "テキストが選択されていません。色を取得したい文字を1つ選択してください。";
        return;
      }

      const selectedText = selectedTextRange.text || "";
      const selectedColor = normalizeColor(selectedTextRange.font.color);

      if (!selectedText.trim()) {
        result.textContent = "選択中のテキストが空です。色を取得したい文字を選択してください。";
        return;
      }

      if (!selectedColor) {
        result.textContent = "選択中テキストの文字色を取得できませんでした。1色の文字だけを選択してください。";
        return;
      }

      await sumNumbersByColorOnCurrentSlide(selectedColor, `選択中の文字色 ${selectedColor}`);
    });
  } catch (error) {
    console.error(error);
    result.textContent = "エラー：" + error.message;
  }
}

async function sumNumbersByColorOnCurrentSlide(targetColor, colorLabel) {
  const result = document.getElementById("result");
  const targetColorElement = document.getElementById("targetColor");

  result.textContent = "集計中...";
  targetColorElement.textContent = `対象色：${colorLabel}`;

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");

      const slideCount = selectedSlides.getCount();

      await context.sync();

      if (slideCount.value === 0) {
        result.textContent = "現在のスライドを取得できませんでした。";
        return;
      }

      const slide = selectedSlides.items[0];
      const shapes = slide.shapes;
      shapes.load("items");

      await context.sync();

      let total = 0;
      const matchedNumbers = [];
      let checkedShapeCount = 0;

      for (const shape of shapes.items) {
        try {
          const textFrame = shape.getTextFrameOrNullObject();
          textFrame.load("hasText");

          await context.sync();

          if (textFrame.isNullObject || !textFrame.hasText) continue;

          checkedShapeCount++;

          const textRange = textFrame.textRange;
          textRange.load("text");

          await context.sync();

          const text = textRange.text || "";
          if (!text) continue;

          const targetColorText = await extractTextByColor(context, textRange, text, targetColor);
          const numbers = targetColorText.match(/-?\d+(?:\.\d+)?/g);

          if (!numbers) continue;

          for (const numText of numbers) {
            const value = Number(numText);
            if (!Number.isNaN(value)) {
              total += value;
              matchedNumbers.push(value);
            }
          }
        } catch (shapeError) {
          console.warn("この図形はスキップしました:", shapeError);
        }
      }

      result.innerHTML = `
        <div><strong>合計：</strong>${total}</div>
        <div><strong>取得した数字：</strong>${matchedNumbers.length ? matchedNumbers.join(", ") : "なし"}</div>
        <div><strong>確認したテキスト図形数：</strong>${checkedShapeCount}</div>
      `;
    });
  } catch (error) {
    console.error(error);
    result.textContent = "エラー：" + error.message;
  }
}

async function formatCodesByColorOnCurrentSlide(targetColor, colorLabel) {
  const result = document.getElementById("result");
  const targetColorElement = document.getElementById("targetColor");

  result.textContent = "整理中...";
  targetColorElement.textContent = `対象色：${colorLabel}`;

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");

      const slideCount = selectedSlides.getCount();

      await context.sync();

      if (slideCount.value === 0) {
        result.textContent = "現在のスライドを取得できませんでした。";
        return;
      }

      const slide = selectedSlides.items[0];
      const shapes = slide.shapes;
      shapes.load("items");

      await context.sync();

      const grouped = {};
      let checkedShapeCount = 0;
      let matchedCodeCount = 0;

      for (const shape of shapes.items) {
        try {
          const textFrame = shape.getTextFrameOrNullObject();
          textFrame.load("hasText");

          await context.sync();

          if (textFrame.isNullObject || !textFrame.hasText) continue;

          checkedShapeCount++;

          const textRange = textFrame.textRange;
          textRange.load("text");

          await context.sync();

          const text = textRange.text || "";
          if (!text) continue;

          const targetColorText = await extractTextByColor(context, textRange, text, targetColor);
          const codeMatches = targetColorText.match(/\b[A-Z]-\d+\b/g);

          if (!codeMatches) continue;

          for (const code of codeMatches) {
            const match = code.match(/^([A-Z])-(\d+)$/);
            if (!match) continue;

            const letter = match[1];
            const number = Number(match[2]);

            if (!grouped[letter]) grouped[letter] = [];

            grouped[letter].push(number);
            matchedCodeCount++;
          }
        } catch (shapeError) {
          console.warn("この図形はスキップしました:", shapeError);
        }
      }

      const outputLines = [];
      const letters = Object.keys(grouped).sort();

      for (const letter of letters) {
        const uniqueSortedNumbers = [...new Set(grouped[letter])].sort((a, b) => a - b);
        const compressed = compressNumberRanges(uniqueSortedNumbers);
        outputLines.push(`${letter}-${compressed}`);
      }

      if (outputLines.length === 0) {
        result.innerHTML = `
          <div><strong>結果：</strong>該当する番号はありません。</div>
          <div><strong>確認したテキスト図形数：</strong>${checkedShapeCount}</div>
        `;
        return;
      }

      result.innerHTML = `
        <div><strong>出力結果：</strong></div>
        <pre class="outputText">${escapeHtml(outputLines.join("\n"))}</pre>
        <div><strong>取得した番号数：</strong>${matchedCodeCount}</div>
        <div><strong>確認したテキスト図形数：</strong>${checkedShapeCount}</div>
      `;
    });
  } catch (error) {
    console.error(error);
    result.textContent = "エラー：" + error.message;
  }
}

async function extractTextByColor(context, textRange, text, targetColor) {
  const charRanges = [];

  for (let i = 0; i < text.length; i++) {
    const charRange = textRange.getSubstring(i, 1);
    charRange.load("text");
    charRange.font.load("color");
    charRanges.push(charRange);
  }

  await context.sync();

  let targetColorText = "";

  for (const charRange of charRanges) {
    const charText = charRange.text || "";
    const color = normalizeColor(charRange.font.color);

    if (color === targetColor) {
      targetColorText += charText;
    } else {
      targetColorText += " ";
    }
  }

  return targetColorText;
}

function compressNumberRanges(numbers) {
  if (!numbers || numbers.length === 0) return "";

  const ranges = [];
  let start = numbers[0];
  let prev = numbers[0];

  for (let i = 1; i < numbers.length; i++) {
    const current = numbers[i];

    if (current === prev + 1) {
      prev = current;
      continue;
    }

    ranges.push(formatRange(start, prev));
    start = current;
    prev = current;
  }

  ranges.push(formatRange(start, prev));

  return ranges.join(", ");
}

function formatRange(start, end) {
  if (start === end) return String(start);
  return `${start}~${end}`;
}

function normalizeColor(color) {
  if (!color) return "";

  let c = String(color).trim().toLowerCase();

  if (/^[0-9a-f]{6}$/.test(c)) c = "#" + c;

  return c;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
