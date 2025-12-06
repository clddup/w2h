import mammoth from "mammoth";

async function convertDocxToHtml(inputPath: string) {
  if (!inputPath) {
    console.error("错误：请提供 Word 文件路径");
    console.log("用法：bun word.js <docx文件路径>");
    process.exit(1);
  }

  const file = Bun.file(inputPath);

  if (!(await file.exists())) {
    console.error(`错误：文件不存在 - ${inputPath}`);
    process.exit(1);
  }

  const arrayBuffer = await file.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  const outputPath = inputPath.replace(/\.docx$/i, ".html");

  const { value: html, messages } = await mammoth.convertToHtml(
    { buffer },
    {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Heading 5'] => h5:fresh",
        "p[style-name='Heading 6'] => h6:fresh",
      ],
      ignoreEmptyParagraphs: false,
    }
  );

  const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>文档</title>
  <style>
    body {
      max-width: 800px;
      margin: 40px auto;
      padding: 0 20px;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      line-height: 1.6;
      color: #333;
    }
    h1, h2, h3, h4, h5, h6 {
      margin-top: 24px;
      margin-bottom: 16px;
      font-weight: 600;
      line-height: 1.25;
    }
    p {
      margin-bottom: 16px;
    }
    /* 让只包含加粗文字的段落显示为标题样式 */
    p:has(> strong:only-child) {
      font-size: 1.5em;
      font-weight: 600;
      margin-top: 32px;
      margin-bottom: 16px;
      line-height: 1.25;
    }
    /* 主标题（第一个加粗段落）更大 */
    body > p:first-of-type:has(> strong:only-child) {
      font-size: 2em;
      text-align: center;
      margin-bottom: 24px;
    }
  </style>
</head>
<body>
${html}
</body>
</html>`;

  await Bun.write(outputPath, fullHtml);

  console.log(`转换完成：${outputPath}`);

  const errors = messages.filter(msg => msg.type === "error");

  if (errors.length > 0) {
    console.error("错误：");
    errors.forEach(msg => console.error(`  - ${msg.message}`));
  }
}

const inputFile = Bun.argv[2];
convertDocxToHtml(inputFile || '');
