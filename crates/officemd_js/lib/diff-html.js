/**
 * HTML diff generation using the `diff` package.
 *
 * Generates a self-contained HTML file and opens it in the browser.
 */

import { createTwoFilesPatch } from "diff";
import { spawnSync } from "node:child_process";
import { writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

export async function generateHtmlDiff(oldMd, newMd, outputPath) {
  const patch = createTwoFilesPatch("a", "b", oldMd, newMd, undefined, undefined, {
    context: 3,
  });

  const lines = patch.split("\n");
  const htmlLines = lines
    .map((line) => {
      const escapedLine = line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

      if (line.startsWith("+") && !line.startsWith("+++")) {
        return `<span class="added">${escapedLine}</span>`;
      }
      if (line.startsWith("-") && !line.startsWith("---")) {
        return `<span class="removed">${escapedLine}</span>`;
      }
      if (line.startsWith("@@")) {
        return `<span class="hunk">${escapedLine}</span>`;
      }
      return `<span class="context">${escapedLine}</span>`;
    })
    .join("\n");

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>office-md diff</title>
  <style>
    body {
      font-family: 'SF Mono', 'Menlo', 'Monaco', 'Courier New', monospace;
      background: #1e1e2e;
      color: #cdd6f4;
      margin: 0;
      padding: 2rem;
    }
    h1 {
      font-size: 1.2rem;
      color: #89b4fa;
      margin-bottom: 1.5rem;
    }
    pre {
      background: #181825;
      border: 1px solid #313244;
      border-radius: 8px;
      padding: 1.5rem;
      overflow-x: auto;
      line-height: 1.6;
      font-size: 0.85rem;
    }
    .added {
      color: #a6e3a1;
      background: rgba(166, 227, 161, 0.1);
      display: block;
    }
    .removed {
      color: #f38ba8;
      background: rgba(243, 139, 168, 0.1);
      display: block;
    }
    .hunk {
      color: #89b4fa;
      display: block;
      margin-top: 0.5rem;
    }
    .context {
      display: block;
    }
  </style>
</head>
<body>
  <h1>office-md markdown diff</h1>
  <pre>${htmlLines}</pre>
</body>
</html>`;

  const outPath = outputPath || join(tmpdir(), `officemd-diff-${Date.now()}.html`);
  writeFileSync(outPath, html, "utf-8");
  return outPath;
}

export async function openInBrowser(filePath) {
  const candidates =
    process.platform === "darwin"
      ? [["open", filePath]]
      : process.platform === "win32"
        ? [["cmd", "/c", "start", "", filePath]]
        : [
            ["xdg-open", filePath],
            ["open", filePath],
          ];

  for (const command of candidates) {
    const result = spawnSync(command[0], command.slice(1), {
      stdio: "ignore",
    });
    if (!result.error && result.status === 0) {
      return;
    }
  }

  console.warn(`Could not auto-open browser. Open manually: ${filePath}`);
}
