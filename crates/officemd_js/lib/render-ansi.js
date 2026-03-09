/**
 * ANSI terminal markdown rendering.
 *
 * Uses Bun's markdown renderer when available; falls back to lightweight
 * ANSI styling in plain Node environments.
 */

const RESET = "\x1b[0m";
const BOLD = "\x1b[1m";
const DIM = "\x1b[2m";
const BLUE = "\x1b[34m";
const CYAN = "\x1b[36m";
const YELLOW = "\x1b[33m";

function fallbackRender(markdown) {
  const lines = markdown.split("\n");
  let inCode = false;

  const rendered = lines.map((line) => {
    if (line.startsWith("```")) {
      inCode = !inCode;
      return `${DIM}${line}${RESET}`;
    }
    if (inCode) {
      return `${DIM}${line}${RESET}`;
    }

    const headingMatch = /^(#{1,6})\s+(.*)$/.exec(line);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const color = level <= 2 ? BLUE : CYAN;
      return `${color}${BOLD}${line}${RESET}`;
    }

    if (/^\s*[-*+]\s+/.test(line)) {
      return line.replace(/^(\s*[-*+])\s+/, `$1 ${YELLOW}`);
    }

    return line;
  });

  return rendered.join("\n");
}

export function renderMarkdownToTerminal(markdown) {
  // Bun.markdown.render can collapse block spacing for our generated markdown.
  // Keep a deterministic line-preserving renderer across Node and Bun.
  return fallbackRender(markdown);
}
