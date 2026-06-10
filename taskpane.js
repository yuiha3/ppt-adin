Office.onReady(() => {
  const button = document.getElementById("sumButton");

  button.addEventListener("click", async () => {
    await sumBlueNumbersOnCurrentSlide();
  });
});

async function sumBlueNumbersOnCurrentSlide() {
  const result = document.getElementById("result");
  result.textContent = "集計中...";

  const targetColor = "#0070c0";

  try {
    await PowerPoint.run(async (context) => {
      // 現在表示中のスライドを取得
      const selectedSlides = context.presentation.getSelectedSlides();
      const slideCount = selectedSlides.getCount();
      selectedSlides.load("items");

      await context.sync();

      if (slideCount.value === 0) {
        result.textContent = "現在のスライドを取得できませんでした。";
        return;
      }

      const slide = selectedSlides.items[0];

      // 現在のスライド内の図形を取得
      const shapes = slide.shapes;
      shapes.load("items");

      await context.sync();

      let total = 0;
      let matchedNumbers = [];
      let checkedShapeCount = 0;

      for (const shape of shapes.items) {
        try {
          const textFrame = shape.textFrame;
          textFrame.load("hasText");

          await context.sync();

          if (!textFrame.hasText) {
            continue;
          }

          checkedShapeCount++;

          const textRange = textFrame.textRange;
          textRange.load("text");

          await context.sync();

          const text = textRange.text || "";

          if (text.length === 0) {
            continue;
          }

          /*
            文字単位で色を確認します。
            これにより、1つのテキストボックス内に
            青文字と黒文字が混在していても対応できます。
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
              // 数字同士がつながらないように区切る
              targetColorText += " ";
            }
          }

          // 整数・小数・マイナスに対応
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
          // テキストを持たない特殊な図形などは無視
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

function normalizeColor(color) {
  if (!color) {
    return "";
  }

  let c = String(color).trim().toLowerCase();

  // "0070c0" のように # がない場合に対応
  if (/^[0-9a-f]{6}$/.test(c)) {
    c = "#" + c;
  }

  // "#0070C0" → "#0070c0"
  return c;
}