/**
 * Terminal diff rendering using the `diff` package.
 */

import { diffLines } from "diff";

const GREEN = "\x1b[32m";
const RED = "\x1b[31m";
const GRAY = "\x1b[90m";
const RESET = "\x1b[0m";

function pushDiffLines(parts, prefix, color, value) {
  const lines = value.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isTerminalSplit = i === lines.length - 1 && line === "";
    if (isTerminalSplit) {
      continue;
    }
    parts.push(`${color}${prefix}${line}${RESET}`);
  }
}

export function renderTerminalDiff(a, b) {
  const changes = diffLines(a, b);
  const parts = [];

  for (const part of changes) {
    if (part.added) {
      pushDiffLines(parts, "+ ", GREEN, part.value);
    } else if (part.removed) {
      pushDiffLines(parts, "- ", RED, part.value);
    } else {
      pushDiffLines(parts, "  ", GRAY, part.value);
    }
  }

  return parts.join("\n") + "\n";
}
