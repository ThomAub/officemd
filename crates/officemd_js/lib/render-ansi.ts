/**
 * ANSI terminal markdown rendering using Bun.markdown.render().
 *
 * Zero external dependencies - uses raw ANSI escape codes.
 */

const RESET = "\x1b[0m";
const BOLD = "\x1b[1m";
const BOLD_OFF = "\x1b[22m";
const DIM = "\x1b[2m";
const DIM_OFF = "\x1b[22m";
const ITALIC = "\x1b[3m";
const ITALIC_OFF = "\x1b[23m";
const UNDERLINE = "\x1b[4m";
const UNDERLINE_OFF = "\x1b[24m";
const REVERSE = "\x1b[7m";
const REVERSE_OFF = "\x1b[27m";
const STRIKETHROUGH = "\x1b[9m";
const STRIKETHROUGH_OFF = "\x1b[29m";

// Colors
const RED = "\x1b[31m";
const GREEN = "\x1b[32m";
const YELLOW = "\x1b[33m";
const BLUE = "\x1b[34m";
const MAGENTA = "\x1b[35m";
const CYAN = "\x1b[36m";
const GRAY = "\x1b[90m";

const HEADING_COLORS = [RED, GREEN, YELLOW, BLUE, MAGENTA, CYAN];

function headingColor(level: number): string {
  return HEADING_COLORS[Math.min(level - 1, HEADING_COLORS.length - 1)];
}

export function renderMarkdownToTerminal(markdown: string): string {
  // @ts-expect-error Bun.markdown is a Bun-specific API
  return Bun.markdown.render(markdown, {
    heading(content: string, level: number): string {
      const color = headingColor(level);
      const prefix = "#".repeat(level) + " ";
      return `\n${color}${BOLD}${UNDERLINE}${prefix}${content}${UNDERLINE_OFF}${BOLD_OFF}${RESET}\n\n`;
    },

    paragraph(content: string): string {
      return `${content}\n\n`;
    },

    strong(content: string): string {
      return `${BOLD}${content}${BOLD_OFF}`;
    },

    emphasis(content: string): string {
      return `${ITALIC}${content}${ITALIC_OFF}`;
    },

    strikethrough(content: string): string {
      return `${STRIKETHROUGH}${content}${STRIKETHROUGH_OFF}`;
    },

    codespan(content: string): string {
      return `${DIM}${REVERSE} ${content} ${REVERSE_OFF}${DIM_OFF}`;
    },

    code(code: string, language: string | undefined): string {
      const lang = language ? `${DIM}[${language}]${DIM_OFF}\n` : "";
      const border = `${DIM}${"─".repeat(40)}${DIM_OFF}`;
      return `\n${border}\n${lang}${code}\n${border}\n\n`;
    },

    link(href: string, _title: string | null, content: string): string {
      // OSC 8 hyperlink: \x1b]8;;URL\x1b\\TEXT\x1b]8;;\x1b\\
      return `\x1b]8;;${href}\x1b\\${BLUE}${UNDERLINE}${content}${UNDERLINE_OFF}${RESET}\x1b]8;;\x1b\\`;
    },

    image(_src: string, _title: string | null, alt: string): string {
      return `${DIM}[image: ${alt}]${DIM_OFF}`;
    },

    list(content: string, ordered: boolean): string {
      if (ordered) {
        let index = 0;
        return content.replace(/^• /gm, () => {
          index++;
          return `${YELLOW}${index}.${RESET} `;
        });
      }
      return content;
    },

    listItem(content: string): string {
      return `${YELLOW}•${RESET} ${content.trim()}\n`;
    },

    blockquote(content: string): string {
      const lines = content
        .split("\n")
        .map((line) => (line ? `${DIM}${GREEN}│${RESET} ${DIM}${line}${DIM_OFF}` : ""));
      return lines.join("\n") + "\n";
    },

    hr(): string {
      return `\n${DIM}${"─".repeat(40)}${DIM_OFF}\n\n`;
    },

    table(header: string, body: string): string {
      return `\n${header}${body}\n`;
    },

    tablerow(content: string): string {
      return `${GRAY}│${RESET}${content}\n`;
    },

    tablecell(content: string, flags: { header: boolean; align: string | null }): string {
      const padded = ` ${content} `;
      if (flags.header) {
        return `${BOLD}${padded}${BOLD_OFF}${GRAY}│${RESET}`;
      }
      return `${padded}${GRAY}│${RESET}`;
    },

    html(raw: string): string {
      return `${DIM}${raw}${DIM_OFF}`;
    },
  });
}
