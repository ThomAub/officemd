//! Markdown cleanup and post-processing.

use regex::Regex;

use super::MarkdownOptions;

/// Clean up markdown output with post-processing
pub(crate) fn clean_markdown(mut text: String, options: &MarkdownOptions) -> String {
    // Collapse dot leaders (e.g. TOC entries: "Introduction...............................1")
    text = collapse_dot_leaders(&text);

    // Fix hyphenation first (before other processing)
    if options.fix_hyphenation {
        text = fix_hyphenation(&text);
    }

    text = remove_chart_noise_lines(&text);
    text = split_dense_bold_metric_lines(&text);

    // Remove standalone page numbers
    if options.remove_page_numbers {
        text = remove_page_numbers(&text);
    }

    // Format URLs as markdown links
    if options.format_urls {
        text = format_urls(&text);
    }

    // Remove excessive newlines (more than 2 in a row)
    while text.contains("\n\n\n") {
        text = text.replace("\n\n\n", "\n\n");
    }

    // Trim leading and trailing whitespace, ensure ends with single newline
    text = text.trim().to_string();
    text.push('\n');

    text
}

/// Collapse dot leaders (runs of 4+ dots) into " ... "
/// Common in tables of contents: "Introduction...............................1" -> "Introduction ... 1"
fn collapse_dot_leaders(text: &str) -> String {
    use once_cell::sync::Lazy;
    static DOT_LEADER_RE: Lazy<Regex> = Lazy::new(|| Regex::new(r"\.{4,}").unwrap());

    DOT_LEADER_RE.replace_all(text, " ... ").to_string()
}

/// Fix words broken across lines with spaces before the continuation
/// e.g., "Limoeiro do Nort e" -> "Limoeiro do Norte"
fn fix_hyphenation(text: &str) -> String {
    use once_cell::sync::Lazy;

    // Fix "word - word" patterns that should be "word-word" (compound words)
    // But be careful not to break list items (which start with "- ")
    static SPACED_HYPHEN_RE: Lazy<Regex> = Lazy::new(|| {
        Regex::new(r"([a-zA-ZáàâãéèêíïóôõöúçñÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]) - ([a-zA-ZáàâãéèêíïóôõöúçñÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ])").unwrap()
    });

    let result = SPACED_HYPHEN_RE
        .replace_all(text, |caps: &regex::Captures| {
            format!("{}-{}", &caps[1], &caps[2])
        })
        .to_string();

    result
}

fn is_numeric_tick_token(token: &str) -> bool {
    let trimmed = token.trim_matches(|c: char| matches!(c, ',' | ';' | ':' | '.'));
    let trimmed = trimmed.strip_prefix('-').unwrap_or(trimmed);
    !trimmed.is_empty() && trimmed.chars().all(|c| c.is_ascii_digit())
}

fn is_year_marker_token(token: &str) -> bool {
    let trimmed = token.trim_matches(|c: char| matches!(c, ',' | ';' | ':' | '.'));
    trimmed.eq_ignore_ascii_case("exercice") || matches!(trimmed, "N" | "N-1" | "N-2" | "N-3" | "E")
}

fn is_chart_noise_line(trimmed: &str) -> bool {
    let tokens: Vec<&str> = trimmed.split_whitespace().collect();
    if tokens.len() < 5 {
        return false;
    }

    let numeric_tokens = tokens.iter().filter(|t| is_numeric_tick_token(t)).count();
    let year_marker_tokens = tokens.iter().filter(|t| is_year_marker_token(t)).count();
    let alpha_tokens = tokens
        .iter()
        .filter(|token| token.chars().any(|c| c.is_alphabetic()))
        .count();
    let short_or_abbrev_tokens = tokens
        .iter()
        .filter(|token| {
            let cleaned: String = token.chars().filter(|c| c.is_alphabetic()).collect();
            !cleaned.is_empty() && (cleaned.chars().count() <= 5 || token.contains('.'))
        })
        .count();

    (year_marker_tokens >= 2 && numeric_tokens >= 8)
        || (numeric_tokens >= 14 && alpha_tokens <= 6)
        || (numeric_tokens == 0
            && alpha_tokens >= 5
            && short_or_abbrev_tokens >= alpha_tokens.saturating_sub(1))
}

fn remove_chart_noise_lines(text: &str) -> String {
    text.lines()
        .filter(|line| !is_chart_noise_line(line.trim()))
        .collect::<Vec<_>>()
        .join("\n")
}

fn split_dense_bold_metric_lines(text: &str) -> String {
    use once_cell::sync::Lazy;

    static BOLD_SEGMENT_RE: Lazy<Regex> = Lazy::new(|| Regex::new(r"\*\*\*[^*].*?\*\*\*").unwrap());

    fn strip_bold(segment: &str) -> &str {
        segment
            .strip_prefix("***")
            .and_then(|s| s.strip_suffix("***"))
            .unwrap_or(segment)
            .trim()
    }

    fn is_numeric_heavy(segment: &str) -> bool {
        let text = strip_bold(segment);
        let digits = text.chars().filter(|c| c.is_ascii_digit()).count();
        let alpha = text.chars().filter(|c| c.is_alphabetic()).count();
        digits >= 3 && alpha <= 4
    }

    fn is_metric_label(segment: &str) -> bool {
        let text = strip_bold(segment).to_lowercase();
        text.starts_with("% sur ") || text.contains("capacité") || text.contains("résultat")
    }

    text.lines()
        .flat_map(|line| {
            let trimmed = line.trim();
            let segments: Vec<&str> = BOLD_SEGMENT_RE
                .find_iter(trimmed)
                .map(|m| m.as_str())
                .collect();
            let should_split = segments.len() >= 3
                && trimmed.len() >= 80
                && trimmed.chars().any(|c| c.is_ascii_digit());

            if !should_split {
                return vec![line.to_string()];
            }

            let mut rebuilt = Vec::new();
            let mut i = 0usize;
            while i < segments.len() {
                if i + 1 < segments.len()
                    && is_metric_label(segments[i])
                    && is_numeric_heavy(segments[i + 1])
                {
                    rebuilt.push(format!("{} {}", segments[i], segments[i + 1]));
                    i += 2;
                } else {
                    rebuilt.push(segments[i].to_string());
                    i += 1;
                }
            }

            if rebuilt.is_empty() {
                vec![line.to_string()]
            } else {
                rebuilt
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Remove standalone page numbers (lines that are just 1-4 digit numbers)
fn remove_page_numbers(text: &str) -> String {
    let mut result = Vec::new();
    let lines: Vec<&str> = text.lines().collect();

    for (i, line) in lines.iter().enumerate() {
        let trimmed = line.trim();

        // Check for page number patterns
        if is_page_number_line(trimmed) {
            // Check context to determine if this is isolated
            let prev_is_break = i > 0 && lines[i - 1].trim() == "---";
            let next_is_break = i + 1 < lines.len() && lines[i + 1].trim() == "---";
            let prev_is_empty = i > 0 && lines[i - 1].trim().is_empty();
            let next_is_empty = i + 1 < lines.len() && lines[i + 1].trim().is_empty();

            // Check if it's on its own line (surrounded by empty lines or page breaks)
            let is_isolated = (prev_is_break || prev_is_empty || i == 0)
                && (next_is_break || next_is_empty || i + 1 == lines.len());

            // Also remove numbers that appear right before a page break
            let before_break = i + 1 < lines.len()
                && (lines[i + 1].trim() == "---"
                    || (i + 2 < lines.len()
                        && lines[i + 1].trim().is_empty()
                        && lines[i + 2].trim() == "---"));

            if is_isolated || before_break {
                continue;
            }
        }

        result.push(*line);
    }

    result.join("\n")
}

/// Check if a line looks like a page number
fn is_page_number_line(trimmed: &str) -> bool {
    // Empty lines are not page numbers
    if trimmed.is_empty() {
        return false;
    }

    // Pattern 1: Just a number (1-4 digits)
    if trimmed.len() <= 4 && trimmed.chars().all(|c| c.is_ascii_digit()) {
        return true;
    }

    // Pattern 2: "Page X of Y" or "Page X" or "Page   of" (placeholder)
    let lower = trimmed.to_lowercase();
    if let Some(rest) = lower.strip_prefix("page") {
        let rest = rest.trim();
        // "Page   of" (empty page numbers)
        if rest == "of" || rest.starts_with("of ") {
            return true;
        }
        // "Page X" or "Page X of Y"
        if rest
            .chars()
            .next()
            .map(|c| c.is_ascii_digit())
            .unwrap_or(false)
        {
            return true;
        }
        // Just "Page" followed by whitespace and maybe "of"
        if rest.is_empty()
            || rest
                .split_whitespace()
                .all(|w| w == "of" || w.chars().all(|c| c.is_ascii_digit()))
        {
            return true;
        }
    }

    // Pattern 3: "X of Y" where X and Y are numbers
    if let Some(of_idx) = trimmed.find(" of ") {
        let before = trimmed[..of_idx].trim();
        let after = trimmed[of_idx + 4..].trim();
        if before.chars().all(|c| c.is_ascii_digit())
            && after.chars().all(|c| c.is_ascii_digit())
            && !before.is_empty()
            && !after.is_empty()
        {
            return true;
        }
    }

    // Pattern 4: "- X -" centered page number
    if trimmed.len() >= 3 && trimmed.starts_with('-') && trimmed.ends_with('-') {
        let inner = trimmed[1..trimmed.len() - 1].trim();
        if inner.chars().all(|c| c.is_ascii_digit()) && !inner.is_empty() {
            return true;
        }
    }

    false
}

/// Convert URLs to markdown links
fn format_urls(text: &str) -> String {
    use once_cell::sync::Lazy;

    // Match URLs - we'll check context manually to avoid formatting already-linked URLs
    static URL_RE: Lazy<Regex> =
        Lazy::new(|| Regex::new(r"https?://[^\s<>\)\]]+[^\s<>\)\]\.\,;]").unwrap());

    let mut result = String::with_capacity(text.len());
    let mut last_end = 0;

    for mat in URL_RE.find_iter(text) {
        let start = mat.start();
        let url = mat.as_str();

        // Check if this URL is already in a markdown link by looking at preceding chars
        // Use safe character boundary checking for multi-byte UTF-8
        let before = {
            let mut check_start = start.saturating_sub(2);
            // Find a valid character boundary
            while check_start > 0 && !text.is_char_boundary(check_start) {
                check_start -= 1;
            }
            if check_start < start && text.is_char_boundary(start) {
                &text[check_start..start]
            } else {
                ""
            }
        };
        let already_linked = before.ends_with("](") || before.ends_with("](");

        // Also check if it's inside square brackets (link text)
        // Ensure we're slicing at a valid char boundary
        let prefix = if text.is_char_boundary(start) {
            &text[..start]
        } else {
            // Find the nearest valid boundary before start
            let mut safe_start = start;
            while safe_start > 0 && !text.is_char_boundary(safe_start) {
                safe_start -= 1;
            }
            &text[..safe_start]
        };
        let open_brackets = prefix.matches('[').count();
        let close_brackets = prefix.matches(']').count();
        let inside_link_text = open_brackets > close_brackets;

        // Ensure mat boundaries are valid char boundaries
        let safe_last_end = if text.is_char_boundary(last_end) {
            last_end
        } else {
            let mut pos = last_end;
            while pos < text.len() && !text.is_char_boundary(pos) {
                pos += 1;
            }
            pos
        };
        let safe_start = if text.is_char_boundary(start) {
            start
        } else {
            let mut pos = start;
            while pos < text.len() && !text.is_char_boundary(pos) {
                pos += 1;
            }
            pos
        };
        let safe_end = if text.is_char_boundary(mat.end()) {
            mat.end()
        } else {
            let mut pos = mat.end();
            while pos < text.len() && !text.is_char_boundary(pos) {
                pos += 1;
            }
            pos
        };

        if already_linked || inside_link_text {
            // Already formatted, keep as-is
            if safe_last_end <= safe_end {
                result.push_str(&text[safe_last_end..safe_end]);
            }
        } else {
            // Add text before this URL
            if safe_last_end <= safe_start {
                result.push_str(&text[safe_last_end..safe_start]);
            }
            // Format as markdown link
            result.push_str(&format!("[{}]({})", url, url));
        }
        last_end = safe_end;
    }

    // Add remaining text (ensure valid char boundary)
    let safe_last_end = if text.is_char_boundary(last_end) {
        last_end
    } else {
        let mut pos = last_end;
        while pos < text.len() && !text.is_char_boundary(pos) {
            pos += 1;
        }
        pos
    };
    if safe_last_end < text.len() {
        result.push_str(&text[safe_last_end..]);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn remove_chart_noise_lines_filters_axis_rows() {
        let input = "Rotations (en jours)\nExercice N-2 34 000 N-1 32 000 Exercice N 30 000 28 000 26 000 24 000\nCrédit clients 161,77 193,50\nC.A. Marge M.B.P. V.A. E.B.E. Résultat C.A.F.";
        let cleaned = remove_chart_noise_lines(input);
        assert!(cleaned.contains("Rotations (en jours)"));
        assert!(cleaned.contains("Crédit clients 161,77 193,50"));
        assert!(!cleaned.contains("Exercice N-2 34 000"));
        assert!(!cleaned.contains("C.A. Marge"));
    }

    #[test]
    fn split_dense_bold_metric_lines_breaks_blob_into_rows() {
        let input = "***Marge brute de production (282) (9 682) 9 400-97,08*** ***% sur production-818,17 N/S*** ***Valeur ajoutée (14 071) (24 490) 10 419-42,54***";
        let cleaned = split_dense_bold_metric_lines(input);
        let lines: Vec<&str> = cleaned.lines().collect();
        assert!(lines.len() >= 3);
        assert_eq!(
            lines[0],
            "***Marge brute de production (282) (9 682) 9 400-97,08***"
        );
        assert_eq!(lines[1], "***% sur production-818,17 N/S***");
    }
}
