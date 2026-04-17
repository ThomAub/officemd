use std::collections::HashMap;

use officemd_core::opc::OpcPackage;
use quick_xml::Reader as XmlReader;
use quick_xml::events::{BytesStart, BytesText, Event};

use crate::error::XlsxError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ValueRenderMode {
    LegacyDefault,
    StyleAware,
}

#[derive(Debug, Default)]
pub struct StyleContext {
    shared_strings: Vec<String>,
    styles: StyleTable,
}

impl StyleContext {
    pub fn load(package: &mut OpcPackage<'_>) -> Result<Self, XlsxError> {
        let shared_strings = load_shared_strings(package)?;
        let mut styles = load_styles(package)?;
        styles.uses_1904_dates = load_workbook_uses_1904_dates(package)?;
        Ok(Self {
            shared_strings,
            styles,
        })
    }

    pub(crate) fn render_cell_text(
        &self,
        cell_type: Option<&str>,
        raw_value: &str,
        inline_text: &str,
        style_index: Option<usize>,
        mode: ValueRenderMode,
    ) -> String {
        let cell = CellValueRef {
            style_index,
            cell_type,
            raw_value,
            inline_text,
        };
        render_cell_value_with_mode(&cell, self, mode)
    }
}

#[derive(Debug, Default)]
struct StyleTable {
    custom_num_formats: HashMap<u32, String>,
    cell_xf_num_fmt_ids: Vec<u32>,
    cell_xf_hints: Vec<Option<FormatHint>>,
    uses_1904_dates: bool,
}

impl StyleTable {
    fn hint_for_style(&self, style_index: usize) -> Option<FormatHint> {
        if let Some(hint) = self.cell_xf_hints.get(style_index) {
            return *hint;
        }
        let num_fmt_id = *self.cell_xf_num_fmt_ids.get(style_index)?;
        self.classify_num_fmt_id(num_fmt_id)
    }

    fn classify_num_fmt_id(&self, num_fmt_id: u32) -> Option<FormatHint> {
        let format_code = self
            .custom_num_formats
            .get(&num_fmt_id)
            .map(String::as_str)
            .or_else(|| builtin_num_format_code(num_fmt_id));
        classify_format_hint(num_fmt_id, format_code)
    }

    fn rebuild_hints(&mut self) {
        self.cell_xf_hints = self
            .cell_xf_num_fmt_ids
            .iter()
            .map(|&num_fmt_id| self.classify_num_fmt_id(num_fmt_id))
            .collect();
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormatHint {
    DateTime { has_time: bool, show_seconds: bool },
    Percent { decimals: usize },
    Currency { decimals: usize, grouping: bool },
    Number { decimals: usize, grouping: bool },
}

#[derive(Debug, Default)]
struct CellValueRef<'a> {
    style_index: Option<usize>,
    cell_type: Option<&'a str>,
    raw_value: &'a str,
    inline_text: &'a str,
}

fn render_cell_value_with_mode(
    cell: &CellValueRef<'_>,
    context: &StyleContext,
    mode: ValueRenderMode,
) -> String {
    match cell.cell_type {
        Some("inlineStr") => cell.inline_text.to_string(),
        Some("s") => {
            let idx = cell.raw_value.trim().parse::<usize>().ok();
            idx.and_then(|i| context.shared_strings.get(i).cloned())
                .unwrap_or_else(|| cell.raw_value.to_string())
        }
        Some("b") => {
            if cell.raw_value.trim() == "1" {
                "true".to_string()
            } else {
                "false".to_string()
            }
        }
        Some("str" | "e") => cell.raw_value.to_string(),
        _ => render_numeric_or_raw(cell, context, mode),
    }
}

fn render_numeric_or_raw(
    cell: &CellValueRef<'_>,
    context: &StyleContext,
    mode: ValueRenderMode,
) -> String {
    let raw = cell.raw_value.trim();
    if raw.is_empty() {
        return String::new();
    }
    let Ok(value) = raw.parse::<f64>() else {
        return cell.raw_value.to_string();
    };

    let style_index = cell.style_index.unwrap_or(0);
    match mode {
        ValueRenderMode::StyleAware => {
            let Some(hint) = context.styles.hint_for_style(style_index) else {
                return cell.raw_value.to_string();
            };
            format_numeric_value(value, hint, context.styles.uses_1904_dates)
        }
        ValueRenderMode::LegacyDefault => {
            if matches!(
                context.styles.hint_for_style(style_index),
                Some(FormatHint::DateTime { .. })
            ) {
                String::new()
            } else {
                cell.raw_value.to_string()
            }
        }
    }
}

fn format_numeric_value(value: f64, hint: FormatHint, uses_1904_dates: bool) -> String {
    match hint {
        FormatHint::DateTime {
            has_time,
            show_seconds,
        } => format_excel_datetime(value, has_time, show_seconds, uses_1904_dates),
        FormatHint::Percent { decimals } => {
            let formatted = format_number(value * 100.0, decimals, false);
            format!("{formatted}%")
        }
        FormatHint::Currency { decimals, grouping } => {
            let formatted = format_number(value, decimals, grouping);
            format!("${formatted}")
        }
        FormatHint::Number { decimals, grouping } => format_number(value, decimals, grouping),
    }
}

fn load_shared_strings(package: &mut OpcPackage<'_>) -> Result<Vec<String>, XlsxError> {
    let Some(xml) = package
        .read_part_string("xl/sharedStrings.xml")
        .map_err(XlsxError::from)?
    else {
        return Ok(Vec::new());
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(false);

    let mut values = Vec::new();
    let mut in_si = false;
    let mut in_t = false;
    let mut current = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                b"si" => {
                    in_si = true;
                    current.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => {
                if local_name(e.name().as_ref()) == b"si" {
                    values.push(String::new());
                }
            }
            Ok(Event::Text(t)) => {
                if in_si && in_t {
                    current.push_str(&unescape_text(&t)?);
                }
            }
            Ok(Event::CData(t)) => {
                if in_si
                    && in_t
                    && let Ok(text) = std::str::from_utf8(t.as_ref())
                {
                    current.push_str(text);
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()) {
                b"t" => in_t = false,
                b"si" => {
                    in_si = false;
                    values.push(std::mem::take(&mut current));
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
    }

    Ok(values)
}

fn load_styles(package: &mut OpcPackage<'_>) -> Result<StyleTable, XlsxError> {
    let Some(xml) = package
        .read_part_string("xl/styles.xml")
        .map_err(XlsxError::from)?
    else {
        return Ok(StyleTable::default());
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut table = StyleTable::default();
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match local_name(e.name().as_ref()) {
                b"cellXfs" => in_cell_xfs = true,
                b"numFmt" => {
                    if let (Some(id), Some(code)) =
                        (attr_u32(e, b"numFmtId"), attr_string(e, b"formatCode"))
                    {
                        table.custom_num_formats.insert(id, code);
                    }
                }
                b"xf" if in_cell_xfs => {
                    table
                        .cell_xf_num_fmt_ids
                        .push(attr_u32(e, b"numFmtId").unwrap_or(0));
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match local_name(e.name().as_ref()) {
                b"numFmt" => {
                    if let (Some(id), Some(code)) =
                        (attr_u32(e, b"numFmtId"), attr_string(e, b"formatCode"))
                    {
                        table.custom_num_formats.insert(id, code);
                    }
                }
                b"xf" if in_cell_xfs => {
                    table
                        .cell_xf_num_fmt_ids
                        .push(attr_u32(e, b"numFmtId").unwrap_or(0));
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == b"cellXfs" {
                    in_cell_xfs = false;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
    }

    table.rebuild_hints();
    Ok(table)
}

fn load_workbook_uses_1904_dates(package: &mut OpcPackage<'_>) -> Result<bool, XlsxError> {
    let Some(xml) = package
        .read_part_string("xl/workbook.xml")
        .map_err(XlsxError::from)?
    else {
        return Ok(false);
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                if local_name(e.name().as_ref()) == b"workbookPr" {
                    let uses_1904_dates =
                        attr_string(e, b"date1904").is_some_and(|value| parse_xml_bool(&value));
                    return Ok(uses_1904_dates);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(XlsxError::Xml(e.to_string())),
        }
    }

    Ok(false)
}

fn classify_format_hint(num_fmt_id: u32, format_code: Option<&str>) -> Option<FormatHint> {
    if format_code.is_none() {
        if is_builtin_date_num_fmt_id(num_fmt_id) {
            return Some(FormatHint::DateTime {
                has_time: false,
                show_seconds: false,
            });
        }
        return None;
    }
    let code = format_code.unwrap_or_default();
    if is_date_like_format(code) || is_builtin_date_num_fmt_id(num_fmt_id) {
        return Some(FormatHint::DateTime {
            has_time: format_has_time(code),
            show_seconds: format_has_seconds(code),
        });
    }
    if code.contains('%') {
        return Some(FormatHint::Percent {
            decimals: decimal_places(code),
        });
    }
    if code.contains('$') || code.contains("[$") {
        return Some(FormatHint::Currency {
            decimals: decimal_places(code),
            grouping: code.contains(','),
        });
    }
    if code.contains('0') || code.contains('#') {
        return Some(FormatHint::Number {
            decimals: decimal_places(code),
            grouping: code.contains(','),
        });
    }
    None
}

fn format_number(value: f64, decimals: usize, grouping: bool) -> String {
    let is_negative = value.is_sign_negative();
    let abs_value = value.abs();
    let mut rendered = format!("{abs_value:.decimals$}");

    if grouping {
        if let Some(dot_idx) = rendered.find('.') {
            let int_part = add_grouping(&rendered[..dot_idx]);
            rendered = format!("{}{}", int_part, &rendered[dot_idx..]);
        } else {
            rendered = add_grouping(&rendered);
        }
    }

    if is_negative && rendered != "0" && rendered != "0.0" && rendered != "0.00" {
        format!("-{rendered}")
    } else {
        rendered
    }
}

fn add_grouping(int_part: &str) -> String {
    let len = int_part.len();
    if len <= 3 {
        return int_part.to_string();
    }
    let bytes = int_part.as_bytes();
    let mut out = String::with_capacity(len + (len - 1) / 3);
    for (i, &b) in bytes.iter().enumerate() {
        let remaining = len - i;
        if i > 0 && remaining.is_multiple_of(3) {
            out.push(',');
        }
        out.push(b as char);
    }
    out
}

fn format_excel_datetime(
    value: f64,
    has_time: bool,
    show_seconds: bool,
    uses_1904_dates: bool,
) -> String {
    // Excel serial dates are well within i64 range; precision loss on
    // round-trip to f64 is acceptable for sub-second calendar math.
    #[allow(clippy::cast_possible_truncation)]
    let mut whole_days = value.floor() as i64;
    #[allow(clippy::cast_possible_truncation, clippy::cast_precision_loss)]
    let mut seconds = ((value - whole_days as f64) * 86_400.0).round() as i64;
    if seconds >= 86_400 {
        whole_days += 1;
        seconds = 0;
    }
    if seconds < 0 {
        seconds = 0;
    }

    // Excel serial day 25569 (1900 system) and 24107 (1904 system) map to 1970-01-01.
    let epoch_offset = if uses_1904_dates { 24_107 } else { 25_569 };
    let unix_days = whole_days - epoch_offset;
    let (year, month, day) = civil_from_days(unix_days);
    let hour = seconds / 3_600;
    let minute = (seconds % 3_600) / 60;
    let second = seconds % 60;

    if has_time {
        if show_seconds {
            format!("{year:04}-{month:02}-{day:02} {hour:02}:{minute:02}:{second:02}")
        } else {
            format!("{year:04}-{month:02}-{day:02} {hour:02}:{minute:02}")
        }
    } else {
        format!("{year:04}-{month:02}-{day:02}")
    }
}

fn civil_from_days(days_from_unix_epoch: i64) -> (i64, i64, i64) {
    // Howard Hinnant's civil-from-days algorithm.
    let z = days_from_unix_epoch + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    let year = y + i64::from(m <= 2);
    (year, m, d)
}

fn decimal_places(format_code: &str) -> usize {
    let primary = primary_format_section(format_code);
    let cleaned = strip_quoted_and_bracketed(primary);
    let Some(dot_idx) = cleaned.find('.') else {
        return 0;
    };
    cleaned[dot_idx + 1..]
        .chars()
        .take_while(|ch| *ch == '0' || *ch == '#')
        .count()
}

fn is_date_like_format(format_code: &str) -> bool {
    let cleaned = strip_quoted_and_bracketed(primary_format_section(format_code)).to_lowercase();
    if cleaned.is_empty() {
        return false;
    }
    cleaned.contains("yy")
        || cleaned.contains("dd")
        || cleaned.contains("mm/")
        || cleaned.contains("/mm")
        || cleaned.contains("m/d")
        || cleaned.contains("d/m")
        || cleaned.contains("am/pm")
        || cleaned.contains('h')
}

fn format_has_time(format_code: &str) -> bool {
    let cleaned = strip_quoted_and_bracketed(primary_format_section(format_code)).to_lowercase();
    cleaned.contains('h') || cleaned.contains("am/pm")
}

fn format_has_seconds(format_code: &str) -> bool {
    let cleaned = strip_quoted_and_bracketed(primary_format_section(format_code)).to_lowercase();
    cleaned.contains('s')
}

fn primary_format_section(format_code: &str) -> &str {
    format_code.split(';').next().unwrap_or(format_code)
}

fn strip_quoted_and_bracketed(s: &str) -> String {
    let mut out = String::new();
    let mut in_quote = false;
    let mut bracket_depth = 0usize;
    let mut escape_next = false;
    for ch in s.chars() {
        if escape_next {
            escape_next = false;
            continue;
        }
        match ch {
            '\\' => {
                escape_next = true;
            }
            '"' => in_quote = !in_quote,
            '[' if !in_quote => bracket_depth += 1,
            ']' if !in_quote && bracket_depth > 0 => bracket_depth -= 1,
            _ => {
                if !in_quote && bracket_depth == 0 {
                    out.push(ch);
                }
            }
        }
    }
    out
}

fn builtin_num_format_code(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        49 => Some("@"),
        _ => None,
    }
}

fn is_builtin_date_num_fmt_id(id: u32) -> bool {
    matches!(id, 14..=22 | 45..=47)
}

pub(crate) fn parse_cell_ref(cell_ref: &str) -> Option<(usize, usize)> {
    if cell_ref.is_empty() {
        return None;
    }
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let row = digits.parse::<usize>().ok()?.checked_sub(1)?;
    let mut col = 0usize;
    for ch in letters.chars() {
        col = col * 26 + (ch as usize - 'A' as usize + 1);
    }
    Some((row, col.checked_sub(1)?))
}

fn unescape_text(t: &BytesText<'_>) -> Result<String, XlsxError> {
    t.unescape()
        .map(std::borrow::Cow::into_owned)
        .map_err(|e| XlsxError::Xml(e.to_string()))
}

fn local_name(name: &[u8]) -> &[u8] {
    if let Some(idx) = name.iter().rposition(|b| *b == b':') {
        &name[idx + 1..]
    } else if let Some(idx) = name.iter().rposition(|b| *b == b'}') {
        &name[idx + 1..]
    } else {
        name
    }
}

fn attr_string(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let attr_key = local_name(attr.key.as_ref());
        if attr_key == key {
            if let Ok(value) = attr.unescape_value() {
                return Some(value.into_owned());
            }
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

fn attr_u32(e: &BytesStart<'_>, key: &[u8]) -> Option<u32> {
    attr_string(e, key)?.parse::<u32>().ok()
}

fn parse_xml_bool(value: &str) -> bool {
    value == "1" || value.eq_ignore_ascii_case("true")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("C12"), Some((11, 2)));
        assert_eq!(parse_cell_ref("AA3"), Some((2, 26)));
    }

    #[test]
    fn formats_excel_datetime_date_only() {
        assert_eq!(
            format_numeric_value(
                45292.0,
                FormatHint::DateTime {
                    has_time: false,
                    show_seconds: false
                },
                false,
            ),
            "2024-01-01"
        );
    }

    #[test]
    fn formats_excel_datetime_with_1904_date_system() {
        assert_eq!(
            format_numeric_value(
                43830.0,
                FormatHint::DateTime {
                    has_time: false,
                    show_seconds: false
                },
                true,
            ),
            "2024-01-01"
        );
    }

    #[test]
    fn formats_percent() {
        assert_eq!(
            format_numeric_value(0.125, FormatHint::Percent { decimals: 2 }, false),
            "12.50%"
        );
    }

    #[test]
    fn legacy_default_date_is_blank() {
        let mut styles = StyleTable::default();
        styles.cell_xf_num_fmt_ids.push(14);
        let ctx = StyleContext {
            shared_strings: Vec::new(),
            styles,
        };

        let out = ctx.render_cell_text(None, "45292", "", Some(0), ValueRenderMode::LegacyDefault);
        assert_eq!(out, "");
    }
}
