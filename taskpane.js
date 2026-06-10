Office.onReady(() => {
  const blueButton = document.getElementById("sumButton");
  const greenButton = document.getElementById("sumGreenButton");
  const selectedColorButton = document.getElementById("sumSelectedColorButton");

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
});

/**
 * 選択中テキストの文字色を取得し、
 * 現在のスライド内で同じ文字色の数字を合計する
 */
async function sumNumbersBySelectedTextColor() {
  const result = document.getElementById("result");
  const targetColorElement = document.getElementById("targetColor");

  result.textContent = "選択中の文字色を取得中...";
  targetColorElement.textContent = "対象色：選択中の文字色";

  try {
    await PowerPoint.run(async (context) => {
      /*
        複数選択チェック。
        図形が複数選択されている場合は中止。
      */
      const selectedShapes = context.presentation.getSelectedShapes();
      const selectedShapeCount = selectedShapes.getCount();

      /*
        選択中のテキスト範囲を取得。
        テキストが選択されていない場合は NullObject になる。
      */
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

      await sumNumbersByColorOnCurrentSlide(
        selectedColor,
        `選択中の文字色 ${selectedColor}`
      );
    });
  } catch (error) {
    console.error(error);
    result.textContent = "エラー：" + error.message;
  }
}

/**
 * 現在のスライド内で、指定色の数字を合計する
 */
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

          if (textFrame.isNullObject || !textFrame.hasText) {
            continue;
          }

          checkedShapeCount++;

          const textRange = textFrame.textRange;
          textRange.load("text");

          await context.sync();

          const text = textRange.text || "";

          if (!text) {
            continue;
          }

          /*
            文字単位で色を見る。
            これにより、1つのテキストボックス内に複数色が混在しても、
            対象色の文字だけを抜き出せる。
          */
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

          const numbers = targetColorText.match(/-?\d+(?:\.\d+)?/g);

          if (!numbers) {
            continue;
          }

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

/**
 * Office.js の色表記を #xxxxxx に正規化
 */
function normalizeColor(color) {
  if (!color) {
    return "";
  }

  let c = String(color).trim().toLowerCase();

  if (/^[0-9a-f]{6}$/.test(c)) {
    c = "#" + c;
  }

  return c;
}