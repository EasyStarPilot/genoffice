//! Native .xlsx -> .ods (OpenDocument Spreadsheet) writer, the save-side
//! counterpart to ods_import: reads a workbook through the crate's own
//! mature native OOXML pipeline (`WorkbookSessions` — the same reader every
//! other open in this app goes through) and serializes what it returns into
//! real ODF spreadsheet XML (content.xml/styles.xml/meta.xml/manifest.xml/
//! mimetype).
//!
//! No byte-fidelity and a deliberately bounded feature set, matching
//! ods_import's own stated scope: values/formulas/basic cell styles
//! (bold/italic/underline/strike/font/size/colors/fill/number-format/
//! horizontal-alignment)/column widths/merged cells round-trip; rich
//! per-run text formatting within one cell (only the plain concatenated
//! text survives), legacy CSE array formulas (written as a plain formula,
//! without the array marking), row heights, and every Excel-specific
//! feature ods_import already doesn't model (pivot tables, sparklines,
//! structured tables, charts, conditional formatting, data validation,
//! comments, defined names, freeze panes, page setup) do not.

use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::convert::ConvertResult;
use crate::ods_formula::a1_to_openformula;
use crate::{CellRange, CellStyle, CellValue, MergedRange, RangeResult, SheetMetadata, WorkbookSessions};
use crate::SidecarError;

const ODS_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const MAX_RANGE_CELLS: usize = 100_000;
/// A sheet needing more than this many row-bands to cover is truncated —
/// see `export_sheet`. 20 bands * 100,000 cells/band = 2,000,000 cells,
/// comfortably past any real-world spreadsheet this engine is likely to see.
const MAX_BANDS: usize = 20;
/// Per `read_range` call, how many times to re-poll (each poll internally
/// waits up to ~750ms for the background indexer) before giving up and
/// using whatever is indexed so far rather than hanging indefinitely.
const MAX_POLLS: usize = 20;

pub fn convert_xlsx_to_ods(source: &Path, target: &Path) -> Result<ConvertResult, SidecarError> {
    let mut sessions = WorkbookSessions::new();
    let metadata = sessions.open(source)?;
    let session_id = metadata.session_id.clone();

    let mut acc = StyleAcc::new();
    let mut sheets_xml: Vec<(String, String)> = Vec::new();
    let mut total_cells = 0usize;
    let export_result = (|| -> Result<(), SidecarError> {
        for sheet in &metadata.sheets {
            let (xml, cells) =
                export_sheet(&mut sessions, &session_id, sheet, &metadata.styles, &mut acc)?;
            total_cells += cells;
            sheets_xml.push((sheet.name.clone(), xml));
        }
        Ok(())
    })();
    // best-effort cleanup: a failed close must not shadow the real export error
    let _ = sessions.close(&session_id);
    export_result?;

    let out = File::create(target)?;
    let mut writer = ZipWriter::new(out);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    // mimetype must be the first entry, stored uncompressed — the ODF package convention
    writer.start_file("mimetype", stored)?;
    writer.write_all(ODS_MIME.as_bytes())?;
    let mut add = |name: &str, content: &str| -> Result<(), SidecarError> {
        writer.start_file(name, deflated)?;
        writer.write_all(content.as_bytes())?;
        Ok(())
    };
    add("META-INF/manifest.xml", &manifest_xml())?;
    add("meta.xml", &meta_xml())?;
    add("styles.xml", &styles_xml())?;
    add("content.xml", &content_xml(&sheets_xml, &acc))?;
    writer.finish()?.sync_all()?;

    Ok(ConvertResult {
        sheets: sheets_xml.len(),
        cells: total_cells,
    })
}

/// `read_range`'s own internal wait only blocks on row-level progress
/// (`indexed_through_row` reaching the requested range) — it returns as soon
/// as that's satisfied even when `indexing_complete` is still false, because
/// the streaming indexer parses a worksheet's XML in document order and
/// sheet-wide extras (merges, conditional formatting, ...) live in elements
/// that come *after* `<sheetData>`, only available once the whole sheet is
/// done. For a small sheet, row data is ready almost instantly, so a bare
/// retry loop here would spin through every poll in a fraction of a
/// millisecond — real wall-clock time the background indexer thread never
/// gets scheduled into, especially under load. A short sleep between polls
/// is what actually gives it a chance to reach the rest of the document.
fn read_range_complete(
    sessions: &mut WorkbookSessions,
    session_id: &str,
    sheet_id: &str,
    range: &CellRange,
) -> Result<RangeResult, SidecarError> {
    let mut result = sessions.read_range(session_id, sheet_id, range)?;
    let mut polls = 0;
    while !result.indexing_complete && polls < MAX_POLLS {
        std::thread::sleep(std::time::Duration::from_millis(25));
        result = sessions.read_range(session_id, sheet_id, range)?;
        polls += 1;
    }
    Ok(result)
}

fn export_sheet(
    sessions: &mut WorkbookSessions,
    session_id: &str,
    sheet: &SheetMetadata,
    styles: &[CellStyle],
    acc: &mut StyleAcc,
) -> Result<(String, usize), SidecarError> {
    let name_attr = escape_xml_attr(&sheet.name);
    if sheet.row_count == 0 || sheet.column_count == 0 {
        return Ok((format!(r#"<table:table table:name="{name_attr}"/>"#), 0));
    }

    let band_rows = (MAX_RANGE_CELLS / sheet.column_count).max(1);
    let mut cells = Vec::new();
    let mut merges: Vec<MergedRange> = Vec::new();
    let mut seen_merges = std::collections::HashSet::new();
    let mut start_row = 0usize;
    let mut band = 0usize;
    while start_row < sheet.row_count && band < MAX_BANDS {
        let end_row = (start_row + band_rows - 1).min(sheet.row_count - 1);
        let range = CellRange {
            start_row,
            end_row,
            start_column: 0,
            end_column: sheet.column_count - 1,
        };
        let result = read_range_complete(sessions, session_id, &sheet.id, &range)?;
        cells.extend(result.cells);
        for merge in result.merges {
            if seen_merges.insert((merge.start_row, merge.start_column, merge.end_row, merge.end_column)) {
                merges.push(merge);
            }
        }
        start_row = end_row + 1;
        band += 1;
    }

    // merge-continuation cells (covered by another cell's span) need a
    // table:covered-table-cell instead of a normal cell, and the merge's
    // starting cell needs its own span attributes
    let mut span_start: HashMap<(usize, usize), (usize, usize)> = HashMap::new();
    let mut covered: std::collections::HashSet<(usize, usize)> = std::collections::HashSet::new();
    for merge in &merges {
        let rows = merge.end_row - merge.start_row + 1;
        let cols = merge.end_column - merge.start_column + 1;
        span_start.insert((merge.start_row, merge.start_column), (rows, cols));
        for r in merge.start_row..=merge.end_row {
            for c in merge.start_column..=merge.end_column {
                if (r, c) != (merge.start_row, merge.start_column) {
                    covered.insert((r, c));
                }
            }
        }
    }

    let mut by_row: HashMap<usize, Vec<crate::CellRecord>> = HashMap::new();
    for cell in cells {
        by_row.entry(cell.row).or_default().push(cell);
    }
    let mut cell_count = 0usize;
    let mut rows_xml = String::new();
    for row in 0..sheet.row_count {
        let mut line = by_row.remove(&row).unwrap_or_default();
        line.sort_unstable_by_key(|c| c.column);
        let mut by_column: HashMap<usize, crate::CellRecord> =
            line.into_iter().map(|c| (c.column, c)).collect();
        rows_xml.push_str("<table:table-row>");
        for column in 0..sheet.column_count {
            if covered.contains(&(row, column)) {
                rows_xml.push_str("<table:covered-table-cell/>");
                continue;
            }
            let span = span_start.get(&(row, column)).copied();
            match by_column.remove(&column) {
                Some(cell) if cell.value.is_some() || cell.formula.is_some() || cell.style_index.is_some() => {
                    cell_count += 1;
                    rows_xml.push_str(&table_cell_xml(&cell, styles, acc, span));
                }
                _ => {
                    if let Some((rows, cols)) = span {
                        rows_xml.push_str(&format!(
                            r#"<table:table-cell table:number-rows-spanned="{rows}" table:number-columns-spanned="{cols}"/>"#,
                        ));
                    } else {
                        rows_xml.push_str("<table:table-cell/>");
                    }
                }
            }
        }
        rows_xml.push_str("</table:table-row>");
    }

    let columns_xml: String = (0..sheet.column_count)
        .map(|index| {
            let width = sheet
                .column_widths
                .iter()
                .find(|w| index >= w.start_column && index <= w.end_column)
                .and_then(|w| w.width);
            match width {
                Some(width) => {
                    let style = acc.column_style(width);
                    format!(r#"<table:table-column table:style-name="{style}"/>"#)
                }
                None => "<table:table-column/>".to_string(),
            }
        })
        .collect();

    let xml = format!(
        r#"<table:table table:name="{name_attr}">{columns_xml}{rows_xml}</table:table>"#,
    );
    Ok((xml, cell_count))
}

fn table_cell_xml(
    cell: &crate::CellRecord,
    styles: &[CellStyle],
    acc: &mut StyleAcc,
    span: Option<(usize, usize)>,
) -> String {
    let style_name = cell
        .style_index
        .and_then(|index| styles.get(index))
        .map(|style| acc.cell_style(cell.style_index.unwrap(), style));
    let style_attr = style_name
        .as_deref()
        .map(|name| format!(r#" table:style-name="{name}""#))
        .unwrap_or_default();
    let span_attr = span
        .map(|(rows, cols)| {
            format!(r#" table:number-rows-spanned="{rows}" table:number-columns-spanned="{cols}""#)
        })
        .unwrap_or_default();
    let formula_attr = cell
        .formula
        .as_deref()
        // CellRecord.formula always carries a leading `=` (worksheet.rs's own
        // convention, worksheet.rs:1139) — a1_to_openformula adds its own `of:=`.
        .map(|f| {
            format!(
                r#" table:formula="{}""#,
                escape_xml_attr(&a1_to_openformula(f.trim_start_matches('='))),
            )
        })
        .unwrap_or_default();

    let text = match &cell.value {
        Some(CellValue::String(text)) => text.clone(),
        Some(CellValue::Number(number)) => format_number_text(*number),
        Some(CellValue::Boolean(flag)) => (if *flag { "TRUE" } else { "FALSE" }).to_string(),
        None => cell
            .rich
            .as_ref()
            .map(|runs| runs.iter().map(|r| r.text.as_str()).collect::<String>())
            .unwrap_or_default(),
    };
    let (value_attrs, text) = match &cell.value {
        Some(CellValue::Number(number)) => (
            format!(r#" office:value-type="float" office:value="{number}""#),
            text,
        ),
        Some(CellValue::Boolean(flag)) => (
            format!(
                r#" office:value-type="boolean" office:boolean-value="{}""#,
                if *flag { "true" } else { "false" },
            ),
            text,
        ),
        Some(CellValue::String(_)) => (r#" office:value-type="string""#.to_string(), text),
        None if !text.is_empty() => (r#" office:value-type="string""#.to_string(), text),
        None => (String::new(), text),
    };
    let paragraph = if text.is_empty() {
        String::new()
    } else {
        format!("<text:p>{}</text:p>", escape_xml_text(&text))
    };
    format!(
        r#"<table:table-cell{style_attr}{span_attr}{formula_attr}{value_attrs}>{paragraph}</table:table-cell>"#,
    )
}

/// Excel's `<v>` values are already decimal text; ODF's `office:value` wants
/// the same, formatted without unnecessary trailing zeros/exponents for the
/// common case.
fn format_number_text(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{value:.0}")
    } else {
        format!("{value}")
    }
}

// ── style accumulation: xlsx CellStyle -> ODF style:style / number-format ──

struct StyleAcc {
    fragments: Vec<String>,
    n: usize,
    cell_style_names: HashMap<usize, String>,
    num_fmt_names: HashMap<String, String>,
    column_style_names: HashMap<u64, String>,
}

impl StyleAcc {
    fn new() -> Self {
        Self {
            fragments: Vec::new(),
            n: 0,
            cell_style_names: HashMap::new(),
            num_fmt_names: HashMap::new(),
            column_style_names: HashMap::new(),
        }
    }

    fn id(&mut self, prefix: &str) -> String {
        self.n += 1;
        format!("{prefix}{}", self.n)
    }

    fn number_format(&mut self, code: &str) -> Option<String> {
        if let Some(name) = self.num_fmt_names.get(code) {
            return Some(name.clone());
        }
        let name = self.id("N");
        let xml = odf_number_style_xml(code, &name)?;
        self.fragments.push(xml);
        self.num_fmt_names.insert(code.to_string(), name.clone());
        Some(name)
    }

    fn cell_style(&mut self, style_index: usize, style: &CellStyle) -> String {
        if let Some(name) = self.cell_style_names.get(&style_index) {
            return name.clone();
        }
        let name = self.id("ce");
        let data_style = style
            .number_format
            .as_deref()
            .filter(|code| !code.is_empty())
            .and_then(|code| self.number_format(code));
        let data_style_attr = data_style
            .as_deref()
            .map(|n| format!(r#" style:data-style-name="{n}""#))
            .unwrap_or_default();

        let mut text_attrs = Vec::new();
        if style.bold {
            text_attrs.push(r#"fo:font-weight="bold""#.to_string());
        }
        if style.italic {
            text_attrs.push(r#"fo:font-style="italic""#.to_string());
        }
        if style.underline {
            text_attrs.push(
                r#"style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color""#
                    .to_string(),
            );
        }
        if style.strikethrough {
            text_attrs.push(
                r#"style:text-line-through-style="solid" style:text-line-through-type="single""#
                    .to_string(),
            );
        }
        if let Some(size) = style.font_size {
            text_attrs.push(format!(r#"fo:font-size="{size}pt""#));
        }
        if let Some(color) = &style.font_color {
            text_attrs.push(format!("fo:color=\"#{}\"", color.trim_start_matches('#')));
        }
        if let Some(font) = &style.font_family {
            text_attrs.push(format!(r#"style:font-name="{}""#, escape_xml_attr(font)));
        }
        let text_props = if text_attrs.is_empty() {
            String::new()
        } else {
            format!("<style:text-properties {}/>", text_attrs.join(" "))
        };

        let cell_props = style
            .fill_color
            .as_ref()
            .map(|color| {
                format!(
                    "<style:table-cell-properties fo:background-color=\"#{}\"/>",
                    color.trim_start_matches('#'),
                )
            })
            .unwrap_or_default();

        let align = style.horizontal_alignment.as_deref().and_then(|a| match a {
            "center" | "centerContinuous" => Some("center"),
            "right" => Some("end"),
            "left" => Some("start"),
            "justify" => Some("justify"),
            _ => None,
        });
        let para_props = align
            .map(|a| format!(r#"<style:paragraph-properties fo:text-align="{a}"/>"#))
            .unwrap_or_default();

        self.fragments.push(format!(
            r#"<style:style style:name="{name}" style:family="table-cell"{data_style_attr}>{cell_props}{para_props}{text_props}</style:style>"#,
        ));
        self.cell_style_names.insert(style_index, name.clone());
        name
    }

    fn column_style(&mut self, width_chars: f64) -> String {
        let key = width_chars.to_bits();
        if let Some(name) = self.column_style_names.get(&key) {
            return name.clone();
        }
        let name = self.id("co");
        // Excel character-width unit -> pixels (~7px/char + 5px padding, the
        // inverse of ods_import's own column-width formula) -> cm @96dpi
        let pixels = width_chars * 7.0 + 5.0;
        let cm = pixels * 2.54 / 96.0;
        self.fragments.push(format!(
            r#"<style:style style:name="{name}" style:family="table-column"><style:table-column-properties style:column-width="{cm:.4}cm"/></style:style>"#,
        ));
        self.column_style_names.insert(key, name.clone());
        name
    }

    fn to_xml(&self) -> String {
        self.fragments.join("")
    }
}

/// Best-effort Excel format-code -> ODF number-format style. Excel's format
/// codes are themselves a small pattern language (`0.00%`, `$#,##0.00`,
/// `yyyy-mm-dd`, ...); this recognizes the common shapes by their token
/// characters rather than fully parsing the pattern; dates/times always
/// render as an ISO-ordered (year-month-day / hour-minute-second) style
/// regardless of the source pattern's exact token order or separators.
/// Returns `None` for "General" or anything unrecognized — the cell then
/// carries no `style:data-style-name` and uses ODF's own default display.
fn odf_number_style_xml(code: &str, name: &str) -> Option<String> {
    if code.is_empty() || code.eq_ignore_ascii_case("general") || code == "@" {
        return None;
    }
    let lower = code.to_ascii_lowercase();
    let has_year_or_day = lower.contains('y') || lower.contains('d');
    let has_hour_or_second = lower.contains('h') || lower.contains('s');
    if has_year_or_day && !has_hour_or_second {
        return Some(date_style_xml(name));
    }
    if has_hour_or_second && !has_year_or_day {
        return Some(time_style_xml(name));
    }
    let decimals = decimal_places_in(code);
    if lower.contains('%') {
        return Some(percentage_style_xml(name, decimals));
    }
    if let Some(symbol) = leading_currency_symbol(code) {
        return Some(currency_style_xml(name, &symbol, decimals));
    }
    let grouping = code.contains(',');
    Some(number_style_xml(name, decimals, grouping))
}

fn decimal_places_in(code: &str) -> usize {
    match code.split(['%', ';']).next().unwrap_or(code).split_once('.') {
        Some((_, after)) => after.chars().take_while(|c| *c == '0' || *c == '#').count(),
        None => 0,
    }
}

/// A currency format conventionally opens with the symbol, possibly after a
/// bracketed locale marker (`[$€-407]`) or immediately (`$#,##0.00`).
fn leading_currency_symbol(code: &str) -> Option<String> {
    if let Some(rest) = code.strip_prefix("[$") {
        let symbol: String = rest.chars().take_while(|&c| c != '-' && c != ']').collect();
        if !symbol.is_empty() {
            return Some(symbol);
        }
    }
    let first = code.chars().next()?;
    if matches!(first, '$' | '\u{20AC}' | '\u{00A3}' | '\u{00A5}') {
        Some(first.to_string())
    } else {
        None
    }
}

fn number_style_xml(name: &str, decimals: usize, grouping: bool) -> String {
    let grouping_attr = if grouping { r#" number:grouping="true""# } else { "" };
    format!(
        r#"<number:number-style style:name="{name}"><number:number number:decimal-places="{decimals}" number:min-integer-digits="1"{grouping_attr}/></number:number-style>"#,
    )
}

fn percentage_style_xml(name: &str, decimals: usize) -> String {
    format!(
        r#"<number:percentage-style style:name="{name}"><number:number number:decimal-places="{decimals}" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>"#,
    )
}

fn currency_style_xml(name: &str, symbol: &str, decimals: usize) -> String {
    format!(
        r#"<number:currency-style style:name="{name}"><number:currency-symbol>{}</number:currency-symbol><number:number number:decimal-places="{decimals}" number:min-integer-digits="1"/></number:currency-style>"#,
        escape_xml_text(symbol),
    )
}

fn date_style_xml(name: &str) -> String {
    format!(
        r#"<number:date-style style:name="{name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>"#,
    )
}

fn time_style_xml(name: &str) -> String {
    format!(
        r#"<number:time-style style:name="{name}"><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>"#,
    )
}

// ── package assembly ──

fn escape_xml_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_xml_attr(text: &str) -> String {
    escape_xml_text(text).replace('"', "&quot;")
}

fn content_xml(sheets: &[(String, String)], acc: &StyleAcc) -> String {
    let tables: String = sheets.iter().map(|(_, xml)| xml.as_str()).collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content {NAMESPACES} office:version="1.2">
<office:automatic-styles>{}</office:automatic-styles>
<office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body>
</office:document-content>"#,
        acc.to_xml(),
    )
}

const NAMESPACES: &str = concat!(
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" "#,
    r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""#,
);

fn styles_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
<office:styles/>
<office:automatic-styles>
<style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="1.905cm" fo:margin-bottom="1.905cm" fo:margin-left="1.778cm" fo:margin-right="1.778cm"/></style:page-layout>
</office:automatic-styles>
<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>
</office:document-styles>"#
        .to_string()
}

fn meta_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">
<office:meta><meta:generator>GenOffice</meta:generator></office:meta>
</office:document-meta>"#
        .to_string()
}

fn manifest_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="{ODS_MIME}"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ods_import::convert_ods_to_xlsx;
    use std::io::Read as _;

    /// Mirrors tests.rs's own `open_fixture` helper (a minimal xlsx
    /// `WorkbookSessions::open` accepts — no `[Content_Types].xml` or
    /// `_rels/.rels` required, confirmed by that module's own fixtures).
    fn write_xlsx_fixture(path: &Path, entries: &[(&str, &str)]) {
        let mut writer = ZipWriter::new(File::create(path).unwrap());
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, content) in entries {
            writer.start_file(*name, options).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
    }

    const WORKBOOK_XML: &str = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    const WORKBOOK_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    const STYLES_XML: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"#;

    fn read_entry(path: &Path, name: &str) -> String {
        let mut archive = zip::ZipArchive::new(File::open(path).unwrap()).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut content = String::new();
        entry.read_to_string(&mut content).unwrap();
        content
    }

    #[test]
    fn exports_values_a_formula_and_a_bold_style() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        write_xlsx_fixture(
            &source,
            &[
                ("xl/workbook.xml", WORKBOOK_XML),
                ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
                ("xl/styles.xml", STYLES_XML),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Name &amp; Co</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2"><f>B1*2</f><v>84</v></c></row></sheetData></worksheet>"#,
                ),
            ],
        );
        let target = dir.path().join("out.ods");

        let result = convert_xlsx_to_ods(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert_eq!(result.cells, 4);

        let content = read_entry(&target, "content.xml");
        assert!(content.contains(r#"table:name="Data""#), "{content}");
        assert!(content.contains("Name &amp; Co"), "{content}");
        assert!(content.contains(r#"office:value-type="float" office:value="42""#), "{content}");
        assert!(content.contains(r#"office:value-type="boolean" office:boolean-value="true""#), "{content}");
        assert!(content.contains(r#"table:formula="of:=[.B1]*2""#), "{content}");
        assert!(content.contains(r#"fo:font-weight="bold""#), "{content}");

        let manifest = read_entry(&target, "META-INF/manifest.xml");
        assert!(manifest.contains("opendocument.spreadsheet"), "{manifest}");
    }

    #[test]
    fn exports_a_merged_cell() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        write_xlsx_fixture(
            &source,
            &[
                ("xl/workbook.xml", WORKBOOK_XML),
                ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Merged</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#,
                ),
            ],
        );
        let target = dir.path().join("out.ods");
        convert_xlsx_to_ods(&source, &target).unwrap();
        let content = read_entry(&target, "content.xml");
        assert!(
            content.contains(r#"table:number-rows-spanned="1" table:number-columns-spanned="2""#),
            "{content}",
        );
        assert!(content.contains("<table:covered-table-cell/>"), "{content}");
    }

    #[test]
    fn round_trips_through_ods_xlsx_ods_without_losing_values() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.xlsx");
        write_xlsx_fixture(
            &source,
            &[
                ("xl/workbook.xml", WORKBOOK_XML),
                ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
                (
                    "xl/worksheets/sheet1.xml",
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#,
                ),
            ],
        );
        let ods_path = dir.path().join("round.ods");
        convert_xlsx_to_ods(&source, &ods_path).unwrap();
        let back_to_xlsx = dir.path().join("back.xlsx");
        convert_ods_to_xlsx(&ods_path, &back_to_xlsx).unwrap();
        let sheet = read_entry(&back_to_xlsx, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains("<v>7</v>"), "{sheet}");
    }
}
