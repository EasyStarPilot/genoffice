//! Native .ods (OpenDocument Spreadsheet) reader: parses real ODF
//! content.xml/styles.xml (table:table/table:table-row/table:table-cell,
//! OpenFormula formulas via ods_formula, style:style with
//! style:family="table-cell", ODF number-format styles) and writes a
//! higher-fidelity .xlsx — values, translated A1 formulas, basic cell
//! styles (bold/italic/underline/strike/font/size/colors/fill/number-format/
//! horizontal-alignment), column widths, and merged cells — that the
//! existing native OOXML pipeline (`WorkbookSessions::open_with_locale`)
//! then opens completely unchanged: the same "convert to real xlsx, then
//! open normally" shape `convert.rs` already uses for legacy .xls (see that
//! module's own doc comment).
//!
//! No byte-fidelity bookkeeping and a deliberately bounded feature set —
//! matching odt-engine/odp-engine's own stated non-goals for the sibling ODF
//! formats: pivot tables, sparklines, structured tables, charts, conditional
//! formatting, data validation, comments, defined names, freeze panes, and
//! page setup are not modeled. A values/formulas/basic-styles/merges/
//! column-widths workbook round-trips with real fidelity; a workbook
//! leaning on any of the above opens with those features simply absent.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use roxmltree::{Document, Node};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::convert::ConvertResult;
use crate::ods_formula::openformula_to_a1;
use crate::SidecarError;

/// A cell/column repeat count above this is treated as "the rest of the
/// sheet is empty padding" and capped rather than materialized — real ODF
/// producers write a single trailing row/cell with a huge repeat count
/// (often the full 1,048,576-row grid) to mean exactly that.
const MAX_REPEAT: u32 = 500;

pub fn convert_ods_to_xlsx(source: &Path, target: &Path) -> Result<ConvertResult, SidecarError> {
    let file = File::open(source)?;
    let mut archive = ZipArchive::new(file)?;
    let content_xml = read_zip_string(&mut archive, "content.xml")?;
    let styles_xml = read_zip_string(&mut archive, "styles.xml").ok();

    let content_doc = Document::parse(&content_xml)
        .map_err(|error| SidecarError::Workbook(format!("Malformed content.xml: {error}")))?;
    let styles_doc = styles_xml
        .as_deref()
        .and_then(|xml| Document::parse(xml).ok());

    let mut styles = StyleTable::new();
    let office_styles = styles_doc
        .as_ref()
        .and_then(|doc| find_child(doc.root_element(), "styles"));
    if let Some(office_styles) = office_styles {
        styles.collect_number_formats(office_styles);
    }
    let auto_styles = find_child(content_doc.root_element(), "automatic-styles");
    if let Some(auto_styles) = auto_styles {
        styles.collect_number_formats(auto_styles);
    }
    if let Some(office_styles) = office_styles {
        styles.collect_cell_styles(office_styles);
    }
    if let Some(auto_styles) = auto_styles {
        styles.collect_cell_styles(auto_styles);
        styles.collect_column_widths(auto_styles);
    }

    let body = find_child(content_doc.root_element(), "body")
        .ok_or_else(|| SidecarError::Workbook("odt: missing office:body".into()))?;
    let spreadsheet = find_child(body, "spreadsheet")
        .ok_or_else(|| SidecarError::Workbook("Not a spreadsheet document.".into()))?;

    let mut sheet_names = Vec::new();
    let mut sheet_xmls = Vec::new();
    let mut total_cells = 0usize;
    let mut xfs = XfTable::new();
    for table in spreadsheet.children().filter(|n| local_name(n) == "table") {
        let name = table.attribute((ODF_TABLE_NS, "name")).unwrap_or("Sheet").to_string();
        let (xml, cell_count) = sheet_xml(table, &styles, &mut xfs);
        total_cells += cell_count;
        sheet_names.push(name);
        sheet_xmls.push(xml);
    }
    if sheet_names.is_empty() {
        return Err(SidecarError::Workbook("The workbook has no sheets.".into()));
    }

    let out = File::create(target)?;
    let mut writer = ZipWriter::new(out);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut add = |name: &str, content: &str| -> Result<(), SidecarError> {
        writer.start_file(name, options)?;
        writer.write_all(content.as_bytes())?;
        Ok(())
    };
    add("[Content_Types].xml", &content_types_xml(sheet_names.len()))?;
    add("_rels/.rels", ROOT_RELS)?;
    add("xl/workbook.xml", &workbook_xml(&sheet_names))?;
    add("xl/_rels/workbook.xml.rels", &workbook_rels_xml(sheet_names.len()))?;
    add("xl/styles.xml", &xfs.to_xml())?;
    for (index, xml) in sheet_xmls.iter().enumerate() {
        add(&format!("xl/worksheets/sheet{}.xml", index + 1), xml)?;
    }
    writer.finish()?.sync_all()?;
    Ok(ConvertResult {
        sheets: sheet_names.len(),
        cells: total_cells,
    })
}

fn read_zip_string(archive: &mut ZipArchive<File>, name: &str) -> Result<String, SidecarError> {
    let mut entry = archive
        .by_name(name)
        .map_err(|_| SidecarError::Workbook(format!("odt: missing {name}")))?;
    let mut text = String::new();
    entry.read_to_string(&mut text)?;
    Ok(text)
}

const ODF_TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const ODF_OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const ODF_STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const ODF_FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const ODF_NUMBER_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/// Local (unprefixed) tag name — ODF elements are always namespace-prefixed
/// (`table:table-cell`, `office:body`, ...) and this crate has no need to
/// distinguish same-named elements across different ODF namespaces.
fn local_name<'a>(node: &Node<'a, 'a>) -> &'a str {
    node.tag_name().name()
}

fn find_child<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children().find(|child| local_name(child) == name)
}

fn attr<'a>(node: &Node<'a, '_>, ns: &str, name: &str) -> Option<&'a str> {
    node.attribute((ns, name))
}

// ── number-format styles: style:name -> an Excel number-format code ──

struct StyleTable {
    /// data-style name -> resolved Excel format code
    number_formats: HashMap<String, String>,
    /// cell style name -> resolved style
    cell_styles: HashMap<String, ResolvedCellStyle>,
    /// column style name -> width in Excel's character-width column unit
    column_widths: HashMap<String, f64>,
}

#[derive(Clone, Default, PartialEq, Eq, Hash)]
struct ResolvedCellStyle {
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    font: Option<String>,
    /// half-points, matching docx-engine's own convention elsewhere in this repo
    size_half_points: Option<u32>,
    color: Option<String>,
    bg_color: Option<String>,
    halign: Option<&'static str>,
    number_format: Option<String>,
}

impl StyleTable {
    fn new() -> Self {
        Self {
            number_formats: HashMap::new(),
            cell_styles: HashMap::new(),
            column_widths: HashMap::new(),
        }
    }

    fn collect_number_formats(&mut self, styles_root: Node) {
        for node in styles_root.children() {
            let format = match local_name(&node) {
                "number-style" | "currency-style" | "percentage-style" => {
                    number_format_code(node, local_name(&node))
                }
                "date-style" => date_or_time_format_code(node),
                "time-style" => date_or_time_format_code(node),
                _ => None,
            };
            if let (Some(format), Some(name)) = (format, attr(&node, ODF_STYLE_NS, "name")) {
                self.number_formats.insert(name.to_string(), format);
            }
        }
    }

    fn collect_cell_styles(&mut self, styles_root: Node) {
        for node in styles_root.children() {
            if local_name(&node) != "style" || attr(&node, ODF_STYLE_NS, "family") != Some("table-cell")
            {
                continue;
            }
            let Some(name) = attr(&node, ODF_STYLE_NS, "name") else { continue };
            let mut style = ResolvedCellStyle::default();
            if let Some(data_style) = attr(&node, ODF_STYLE_NS, "data-style-name") {
                style.number_format = self.number_formats.get(data_style).cloned();
            }
            if let Some(props) = find_child(node, "table-cell-properties") {
                style.bg_color = attr(&props, ODF_FO_NS, "background-color")
                    .filter(|c| *c != "transparent")
                    .map(|c| c.trim_start_matches('#').to_uppercase());
            }
            if let Some(props) = find_child(node, "paragraph-properties") {
                style.halign = match attr(&props, ODF_FO_NS, "text-align") {
                    Some("center") => Some("center"),
                    Some("end") => Some("right"),
                    Some("start") | Some("left") => Some("left"),
                    _ => None,
                };
            }
            if let Some(props) = find_child(node, "text-properties") {
                style.bold = attr(&props, ODF_FO_NS, "font-weight") == Some("bold");
                style.italic = attr(&props, ODF_FO_NS, "font-style") == Some("italic");
                style.underline = attr(&props, ODF_STYLE_NS, "text-underline-style")
                    .is_some_and(|v| v != "none");
                style.strike = attr(&props, ODF_STYLE_NS, "text-line-through-style")
                    .is_some_and(|v| v != "none");
                style.font = attr(&props, ODF_STYLE_NS, "font-name").map(str::to_string);
                style.size_half_points = attr(&props, ODF_FO_NS, "font-size")
                    .and_then(parse_pt)
                    .map(|pt| (pt * 2.0).round() as u32);
                style.color = attr(&props, ODF_FO_NS, "color")
                    .map(|c| c.trim_start_matches('#').to_uppercase());
            }
            self.cell_styles.insert(name.to_string(), style);
        }
    }

    fn collect_column_widths(&mut self, auto_styles: Node) {
        for node in auto_styles.children() {
            if local_name(&node) != "style" || attr(&node, ODF_STYLE_NS, "family") != Some("table-column")
            {
                continue;
            }
            let (Some(name), Some(props)) =
                (attr(&node, ODF_STYLE_NS, "name"), find_child(node, "table-column-properties"))
            else {
                continue;
            };
            if let Some(width) = attr(&props, ODF_STYLE_NS, "column-width").and_then(parse_cm) {
                // cm -> pixels (96dpi) -> Excel's character-width unit (~7px/char, Calibri 11 default)
                let pixels = width * 96.0 / 2.54;
                let chars = ((pixels - 5.0) / 7.0 * 256.0).round() / 256.0;
                self.column_widths.insert(name.to_string(), chars.max(0.0));
            }
        }
    }
}

/// `"12pt"` -> `12.0`. ODF also allows other CSS length units here in
/// principle; only points (the overwhelmingly common producer output) are
/// handled, matching this module's bounded scope.
fn parse_pt(value: &str) -> Option<f64> {
    value.strip_suffix("pt")?.parse().ok()
}

fn parse_cm(value: &str) -> Option<f64> {
    if let Some(cm) = value.strip_suffix("cm") {
        return cm.parse().ok();
    }
    if let Some(inch) = value.strip_suffix('"') {
        return inch.parse::<f64>().ok().map(|v| v * 2.54);
    }
    if let Some(inch) = value.strip_suffix("in") {
        return inch.parse::<f64>().ok().map(|v| v * 2.54);
    }
    None
}

/// `number:number-style` / `number:percentage-style` / `number:currency-style`:
/// reads the (at most one) `number:number` child for decimal places and
/// grouping, plus a currency symbol / trailing `%` literal from sibling
/// `number:text`/`number:currency-symbol` elements.
fn number_format_code(node: Node, kind: &str) -> Option<String> {
    let number = node.children().find(|c| local_name(c) == "number")?;
    let decimals: usize = attr(&number, ODF_NUMBER_NS, "decimal-places")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let grouping = attr(&number, ODF_NUMBER_NS, "grouping") == Some("true");
    let integer_part = if grouping { "#,##0" } else { "0" };
    let mut code = if decimals > 0 {
        format!("{integer_part}.{}", "0".repeat(decimals))
    } else {
        integer_part.to_string()
    };
    match kind {
        "percentage-style" => code.push('%'),
        "currency-style" => {
            let symbol = node
                .children()
                .find(|c| local_name(c) == "currency-symbol")
                .and_then(|c| c.text())
                .unwrap_or("$");
            code = format!("{symbol}{code}");
        }
        _ => {}
    }
    Some(code)
}

/// `number:date-style` / `number:time-style`: walks children in document
/// order — `number:year`/`month`/`day`/`hours`/`minutes`/`seconds` become
/// the matching Excel token, `number:text` literal separators pass through.
fn date_or_time_format_code(node: Node) -> Option<String> {
    let mut code = String::new();
    for child in node.children() {
        let long = attr(&child, ODF_NUMBER_NS, "style") == Some("long");
        match local_name(&child) {
            "year" => code.push_str(if long { "yyyy" } else { "yy" }),
            "month" => code.push_str(if long { "mm" } else { "m" }),
            "day" => code.push_str(if long { "dd" } else { "d" }),
            "hours" => code.push_str(if long { "hh" } else { "h" }),
            "minutes" => code.push_str(if long { "mm" } else { "m" }),
            "seconds" => code.push_str(if long { "ss" } else { "s" }),
            "am-pm" => code.push_str("AM/PM"),
            "text" => code.push_str(child.text().unwrap_or("")),
            _ => {}
        }
    }
    if code.is_empty() { None } else { Some(code) }
}

// ── xlsx styles.xml assembly: dedupe resolved cell styles into cellXfs ──

struct XfTable {
    fonts: Vec<FontXml>,
    fills: Vec<String>,
    num_fmts: Vec<(u32, String)>,
    xfs: Vec<(usize, usize, Option<u32>, Option<&'static str>)>, // (fontId, fillId, numFmtId, halign)
    index_of: HashMap<ResolvedCellStyle, usize>,
    next_num_fmt_id: u32,
}

#[derive(Clone, PartialEq, Eq, Hash)]
struct FontXml {
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    name: String,
    size_half_points: u32,
    color: Option<String>,
}

impl XfTable {
    fn new() -> Self {
        Self {
            fonts: vec![FontXml {
                bold: false,
                italic: false,
                underline: false,
                strike: false,
                name: "Calibri".into(),
                size_half_points: 22,
                color: None,
            }],
            fills: vec!["none".into(), "gray125".into()],
            num_fmts: Vec::new(),
            xfs: vec![(0, 0, None, None)],
            index_of: HashMap::new(),
            next_num_fmt_id: 176, // first free custom numFmtId per the OOXML spec (< 164 are builtin/reserved)
        }
    }

    /// Returns the cellXf index (`s="N"`) for a resolved style, reusing an
    /// existing xf when the same combination was already registered.
    fn intern(&mut self, style: &ResolvedCellStyle) -> usize {
        if let Some(&index) = self.index_of.get(style) {
            return index;
        }
        let font_id = self.intern_font(style);
        let fill_id = style.bg_color.as_ref().map(|hex| self.intern_fill(hex)).unwrap_or(0);
        let num_fmt_id = style.number_format.as_ref().map(|code| self.intern_num_fmt(code));
        let index = self.xfs.len();
        self.xfs.push((font_id, fill_id, num_fmt_id, style.halign));
        self.index_of.insert(style.clone(), index);
        index
    }

    fn intern_font(&mut self, style: &ResolvedCellStyle) -> usize {
        let font = FontXml {
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strike: style.strike,
            name: style.font.clone().unwrap_or_else(|| "Calibri".into()),
            size_half_points: style.size_half_points.unwrap_or(22),
            color: style.color.clone(),
        };
        if let Some(index) = self.fonts.iter().position(|f| *f == font) {
            return index;
        }
        self.fonts.push(font);
        self.fonts.len() - 1
    }

    fn intern_fill(&mut self, hex: &str) -> usize {
        if let Some(index) = self.fills.iter().position(|f| f == hex) {
            return index;
        }
        self.fills.push(hex.to_string());
        self.fills.len() - 1
    }

    fn intern_num_fmt(&mut self, code: &str) -> u32 {
        if let Some((id, _)) = self.num_fmts.iter().find(|(_, c)| c == code) {
            return *id;
        }
        let id = self.next_num_fmt_id;
        self.next_num_fmt_id += 1;
        self.num_fmts.push((id, code.to_string()));
        id
    }

    fn to_xml(&self) -> String {
        let num_fmts = if self.num_fmts.is_empty() {
            String::new()
        } else {
            let entries: String = self
                .num_fmts
                .iter()
                .map(|(id, code)| format!(r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#, escape_xml(code)))
                .collect();
            format!(r#"<numFmts count="{}">{entries}</numFmts>"#, self.num_fmts.len())
        };
        let fonts: String = self
            .fonts
            .iter()
            .map(|font| {
                let bold = if font.bold { "<b/>" } else { "" };
                let italic = if font.italic { "<i/>" } else { "" };
                let underline = if font.underline { r#"<u/>"# } else { "" };
                let strike = if font.strike { "<strike/>" } else { "" };
                let color = font
                    .color
                    .as_ref()
                    .map(|c| format!(r#"<color rgb="FF{c}"/>"#))
                    .unwrap_or_default();
                format!(
                    r#"<font>{bold}{italic}{underline}{strike}<sz val="{}"/>{color}<name val="{}"/></font>"#,
                    font.size_half_points as f64 / 2.0,
                    escape_xml(&font.name),
                )
            })
            .collect();
        let fills: String = self
            .fills
            .iter()
            .enumerate()
            .map(|(index, fill)| match index {
                0 => r#"<fill><patternFill patternType="none"/></fill>"#.to_string(),
                1 => r#"<fill><patternFill patternType="gray125"/></fill>"#.to_string(),
                _ => format!(
                    r#"<fill><patternFill patternType="solid"><fgColor rgb="FF{fill}"/><bgColor indexed="64"/></patternFill></fill>"#,
                ),
            })
            .collect();
        let cell_xfs: String = self
            .xfs
            .iter()
            .map(|(font_id, fill_id, num_fmt_id, halign)| {
                let num_fmt_id = num_fmt_id.unwrap_or(0);
                let apply_num_fmt = if num_fmt_id != 0 { r#" applyNumberFormat="1""# } else { "" };
                let align = halign
                    .map(|h| format!(r#"<alignment horizontal="{h}"/>"#))
                    .unwrap_or_default();
                let apply_align = if halign.is_some() { r#" applyAlignment="1""# } else { "" };
                format!(
                    r#"<xf numFmtId="{num_fmt_id}" fontId="{font_id}" fillId="{fill_id}" borderId="0" xfId="0"{apply_num_fmt}{apply_align}>{align}</xf>"#,
                )
            })
            .collect();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{num_fmts}<fonts count="{}">{fonts}</fonts><fills count="{}">{fills}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="{}">{cell_xfs}</cellXfs></styleSheet>"#,
            self.fonts.len(),
            self.fills.len(),
            self.xfs.len(),
        )
    }
}

// ── one sheet: table:table -> worksheet XML ──

/// A cell's XML content, everything except the `r="..."` reference — which
/// depends on which materialized row it ends up at (a `table:table-row`
/// with `table:number-rows-repeated` renders the same parsed cell content at
/// several different absolute rows), so it is resolved once, at
/// materialization time in `sheet_xml`, not while parsing the row.
#[derive(Clone)]
struct CellBody {
    type_attr: &'static str,
    style_attr: String,
    inner: String,
}

impl CellBody {
    fn render(&self, reference: &str) -> String {
        format!(
            r#"<c r="{reference}"{}{}>{}</c>"#,
            self.type_attr, self.style_attr, self.inner,
        )
    }
}

struct PendingCell {
    row: u32,
    column: u32,
    column_letters: String,
    body: CellBody,
}

fn sheet_xml(table: Node, styles: &StyleTable, xfs: &mut XfTable) -> (String, usize) {
    let mut column_widths: Vec<(u32, f64)> = Vec::new(); // (1-based column index, width)
    let mut column_cursor = 0u32;
    for col in table.children().filter(|n| local_name(n) == "table-column") {
        let repeat = attr(&col, ODF_TABLE_NS, "number-columns-repeated")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1)
            .min(MAX_REPEAT);
        if let Some(width) = attr(&col, ODF_TABLE_NS, "style-name").and_then(|name| styles.column_widths.get(name))
        {
            for i in 0..repeat {
                column_widths.push((column_cursor + i + 1, *width));
            }
        }
        column_cursor += repeat;
    }

    let mut cells: Vec<PendingCell> = Vec::new();
    let mut merges: Vec<String> = Vec::new();
    let mut row_cursor = 0u32;
    let mut max_row = 0u32;
    let mut max_col = 0u32;
    let mut cell_count = 0usize;

    for row in table.children().filter(|n| local_name(n) == "table-row") {
        let row_repeat = attr(&row, ODF_TABLE_NS, "number-rows-repeated")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1);
        let row_cells = row_cell_records(row, styles, xfs);
        let row_has_content = !row_cells.is_empty();
        let materialize = if row_has_content { row_repeat.min(MAX_REPEAT) } else { 1 };
        for repeat_index in 0..materialize {
            let this_row = row_cursor + repeat_index;
            for cell in &row_cells {
                let absolute_col = cell.column;
                max_row = max_row.max(this_row);
                max_col = max_col.max(absolute_col);
                if let Some((row_span, col_span)) = cell.span {
                    merges.push(format!(
                        "{}:{}",
                        cell_reference(this_row, absolute_col),
                        cell_reference(this_row + row_span - 1, absolute_col + col_span - 1),
                    ));
                }
                cells.push(PendingCell {
                    row: this_row,
                    column: absolute_col,
                    column_letters: cell.column_letters.clone(),
                    body: cell.body.clone(),
                });
                cell_count += 1;
            }
        }
        row_cursor += row_repeat.min(MAX_REPEAT);
    }

    let mut by_row: HashMap<u32, Vec<PendingCell>> = HashMap::new();
    for cell in cells {
        by_row.entry(cell.row).or_default().push(cell);
    }
    let mut row_numbers: Vec<u32> = by_row.keys().copied().collect();
    row_numbers.sort_unstable();
    let mut body = String::new();
    for row in row_numbers {
        let mut line = by_row.remove(&row).unwrap_or_default();
        line.sort_unstable_by_key(|cell| cell.column);
        body.push_str(&format!(r#"<row r="{}">"#, row + 1));
        for cell in line {
            let reference = format!("{}{}", cell.column_letters, cell.row + 1);
            body.push_str(&cell.body.render(&reference));
        }
        body.push_str("</row>");
    }

    let dimension = if cells_is_empty(&body) {
        "A1:A1".to_string()
    } else {
        format!("A1:{}", cell_reference(max_row, max_col))
    };
    let cols_xml = if column_widths.is_empty() {
        String::new()
    } else {
        let entries: String = column_widths
            .iter()
            .map(|(index, width)| {
                format!(r#"<col min="{index}" max="{index}" width="{width}" customWidth="1"/>"#)
            })
            .collect();
        format!("<cols>{entries}</cols>")
    };
    let merge_xml = if merges.is_empty() {
        String::new()
    } else {
        let entries: String = merges.iter().map(|range| format!(r#"<mergeCell ref="{range}"/>"#)).collect();
        format!(r#"<mergeCells count="{}">{entries}</mergeCells>"#, merges.len())
    };
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="{dimension}"/>{cols_xml}<sheetData>{body}</sheetData>{merge_xml}</worksheet>"#,
    );
    (xml, cell_count)
}

fn cells_is_empty(body: &str) -> bool {
    body.is_empty()
}

struct RowCell {
    column: u32,
    column_letters: String,
    body: CellBody,
    /// (row_span, col_span) when this cell starts a merge wider/taller than 1x1
    span: Option<(u32, u32)>,
}

/// One `table:table-row`'s cells, with `table:number-columns-repeated`
/// expanded (capped) and `table:covered-table-cell` merge continuations
/// consuming a column position without producing their own output cell.
fn row_cell_records(row: Node, styles: &StyleTable, xfs: &mut XfTable) -> Vec<RowCell> {
    let mut out = Vec::new();
    let mut column = 0u32;
    for cell in row.children() {
        match local_name(&cell) {
            "covered-table-cell" => {
                column += 1;
            }
            "table-cell" => {
                let repeat = attr(&cell, ODF_TABLE_NS, "number-columns-repeated")
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(1);
                let row_span: u32 = attr(&cell, ODF_TABLE_NS, "number-rows-spanned")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1);
                let col_span: u32 = attr(&cell, ODF_TABLE_NS, "number-columns-spanned")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1);
                if let Some(body) = cell_body(&cell, styles, xfs) {
                    out.push(RowCell {
                        column,
                        column_letters: column_letters(column),
                        body,
                        span: (row_span > 1 || col_span > 1).then_some((row_span, col_span)),
                    });
                }
                // a spanning cell's own repeated-column-cursor advance only makes sense for
                // repeat=1 in practice (a spanning cell repeating itself would overlap its own
                // span) — real producers never combine the two, so span is only honored once
                let materialize = repeat.clamp(1, MAX_REPEAT);
                column += materialize;
            }
            _ => {}
        }
    }
    out
}

fn cell_body(cell: &Node, styles: &StyleTable, xfs: &mut XfTable) -> Option<CellBody> {
    let value_type = attr(cell, ODF_OFFICE_NS, "value-type");
    let formula = attr(cell, ODF_TABLE_NS, "formula").map(openformula_to_a1);
    let style_name = attr(cell, ODF_TABLE_NS, "style-name");
    let resolved_style = style_name.and_then(|name| styles.cell_styles.get(name));
    let style_index = resolved_style.map(|style| xfs.intern(style));

    let text_content = || -> String {
        cell.children()
            .filter(|n| local_name(n) == "p")
            .map(|p| p.text().unwrap_or("").to_string())
            .collect::<Vec<_>>()
            .join("\n")
    };

    let value_xml = match value_type {
        Some("float") | Some("percentage") | Some("currency") => {
            let value = attr(cell, ODF_OFFICE_NS, "value")?;
            Some(format!("<v>{}</v>", escape_xml(value)))
        }
        Some("date") => {
            let date = attr(cell, ODF_OFFICE_NS, "date-value")?;
            let serial = excel_date_serial(date)?;
            Some(format!("<v>{serial}</v>"))
        }
        Some("time") => {
            let time = attr(cell, ODF_OFFICE_NS, "time-value")?;
            let fraction = excel_time_fraction(time)?;
            Some(format!("<v>{fraction}</v>"))
        }
        Some("boolean") => {
            let flag = attr(cell, ODF_OFFICE_NS, "boolean-value") == Some("true");
            Some(format!("<v>{}</v>", if flag { 1 } else { 0 }))
        }
        Some("string") => {
            let text = attr(cell, ODF_OFFICE_NS, "string-value")
                .map(str::to_string)
                .unwrap_or_else(text_content);
            return Some(inline_string_cell(style_index, formula.as_deref(), &text));
        }
        _ => {
            let text = text_content();
            if text.is_empty() {
                None
            } else {
                return Some(inline_string_cell(style_index, formula.as_deref(), &text));
            }
        }
    };

    if value_xml.is_none() && formula.is_none() && style_index.is_none() {
        return None;
    }
    let type_attr = match value_type {
        Some("boolean") => r#" t="b""#,
        _ => "",
    };
    let style_attr = style_index.map(|s| format!(r#" s="{s}""#)).unwrap_or_default();
    let formula_xml = formula
        .as_deref()
        .map(|text| format!("<f>{}</f>", escape_xml(text.trim_start_matches('='))))
        .unwrap_or_default();
    Some(CellBody {
        type_attr,
        style_attr,
        inner: format!("{formula_xml}{}", value_xml.unwrap_or_default()),
    })
}

fn inline_string_cell(style_index: Option<usize>, formula: Option<&str>, text: &str) -> CellBody {
    let style_attr = style_index.map(|s| format!(r#" s="{s}""#)).unwrap_or_default();
    let formula_xml = formula
        .map(|text| format!("<f>{}</f>", escape_xml(text.trim_start_matches('='))))
        .unwrap_or_default();
    CellBody {
        type_attr: r#" t="inlineStr""#,
        style_attr,
        inner: format!(
            r#"{formula_xml}<is><t xml:space="preserve">{}</t></is>"#,
            escape_xml(text),
        ),
    }
}

/// ODF dates are `YYYY-MM-DD` (optionally with a time part); Excel's serial
/// epoch is 1899-12-30 (day 0), the same epoch this converter's sibling
/// `convert.rs` assumes via calamine's own date handling.
fn excel_date_serial(date: &str) -> Option<f64> {
    let (date_part, time_part) = date.split_once('T').unwrap_or((date, ""));
    let mut parts = date_part.split('-');
    let year: i64 = parts.next()?.parse().ok()?;
    let month: i64 = parts.next()?.parse().ok()?;
    let day: i64 = parts.next()?.parse().ok()?;
    let days = days_from_civil(year, month, day) - days_from_civil(1899, 12, 30);
    let fraction = if time_part.is_empty() {
        0.0
    } else {
        excel_time_fraction(time_part).unwrap_or(0.0)
    };
    Some(days as f64 + fraction)
}

fn excel_time_fraction(time: &str) -> Option<f64> {
    // "PT13H30M00S" (duration form) or a plain "13:30:00"
    if let Some(rest) = time.strip_prefix("PT") {
        let mut hours = 0.0;
        let mut minutes = 0.0;
        let mut seconds = 0.0;
        let mut number = String::new();
        for ch in rest.chars() {
            match ch {
                '0'..='9' | '.' => number.push(ch),
                'H' => {
                    hours = number.parse().unwrap_or(0.0);
                    number.clear();
                }
                'M' => {
                    minutes = number.parse().unwrap_or(0.0);
                    number.clear();
                }
                'S' => {
                    seconds = number.parse().unwrap_or(0.0);
                    number.clear();
                }
                _ => {}
            }
        }
        return Some((hours * 3600.0 + minutes * 60.0 + seconds) / 86400.0);
    }
    let mut parts = time.split(':');
    let hours: f64 = parts.next()?.parse().ok()?;
    let minutes: f64 = parts.next().unwrap_or("0").parse().ok()?;
    let seconds: f64 = parts.next().unwrap_or("0").parse().ok()?;
    Some((hours * 3600.0 + minutes * 60.0 + seconds) / 86400.0)
}

/// Howard Hinnant's civil-calendar day-count algorithm (proleptic Gregorian,
/// days since 0000-03-01) — used only as a common reference point to
/// subtract two dates, never displayed directly.
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let y = if month <= 2 { year - 1 } else { year };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = (month + 9) % 12;
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn column_letters(column: u32) -> String {
    let mut letters = String::new();
    let mut remaining = column + 1;
    while remaining > 0 {
        remaining -= 1;
        letters.insert(0, char::from(b'A' + (remaining % 26) as u8));
        remaining /= 26;
    }
    letters
}

fn cell_reference(row: u32, column: u32) -> String {
    format!("{}{}", column_letters(column), row + 1)
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn content_types_xml(sheet_count: usize) -> String {
    let overrides: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        ))
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{overrides}</Types>"#,
    )
}

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

fn workbook_xml(names: &[String]) -> String {
    let sheets: String = names
        .iter()
        .enumerate()
        .map(|(index, name)| format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            escape_xml(name),
            index + 1,
            index + 1,
        ))
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#,
    )
}

fn workbook_rels_xml(sheet_count: usize) -> String {
    let mut relationships: String = (1..=sheet_count)
        .map(|index| format!(
            r#"<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>"#,
        ))
        .collect();
    relationships.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        sheet_count + 1,
    ));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn write_fixture(path: &Path, content_xml: &str) {
        let file = File::create(path).unwrap();
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default();
        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer.start_file("content.xml", options).unwrap();
        writer.write_all(content_xml.as_bytes()).unwrap();
        writer.finish().unwrap();
    }

    fn read_entry(path: &Path, name: &str) -> String {
        let mut archive = zip::ZipArchive::new(File::open(path).unwrap()).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut content = String::new();
        entry.read_to_string(&mut content).unwrap();
        content
    }

    const NS: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""#;

    #[test]
    fn converts_values_formulas_and_a_bold_style() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.ods");
        let content = format!(
            r#"<?xml version="1.0"?>
<office:document-content {NS}>
<office:automatic-styles>
<style:style style:name="ce1" style:family="table-cell">
  <style:text-properties fo:font-weight="bold"/>
</style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
  <table:table-cell table:style-name="ce1" office:value-type="string"><text:p>Name &amp; Co</text:p></table:table-cell>
  <table:table-cell office:value-type="float" office:value="42"><text:p>42</text:p></table:table-cell>
</table:table-row>
<table:table-row>
  <table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>
  <table:table-cell office:value-type="float" office:value="84" table:formula="of:=[.B1]*2"><text:p>84</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>"#
        );
        write_fixture(&source, &content);
        let target = dir.path().join("converted.xlsx");

        let result = convert_ods_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.sheets, 1);
        assert_eq!(result.cells, 4);

        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<c r="A1" t="inlineStr" s="1">"#), "{sheet}");
        assert!(sheet.contains("Name &amp; Co"), "{sheet}");
        assert!(sheet.contains(r#"<c r="B1"><v>42</v></c>"#), "{sheet}");
        assert!(sheet.contains(r#"<c r="A2" t="b"><v>1</v></c>"#), "{sheet}");
        assert!(sheet.contains(r#"<f>B1*2</f>"#), "{sheet}");

        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains("<b/>"), "{styles}");

        let workbook = read_entry(&target, "xl/workbook.xml");
        assert!(workbook.contains(r#"<sheet name="Data" sheetId="1" r:id="rId1"/>"#));
    }

    #[test]
    fn expands_a_merged_cell_into_a_merge_cell_entry() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.ods");
        let content = format!(
            r#"<?xml version="1.0"?>
<office:document-content {NS}>
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
  <table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>Merged</text:p></table:table-cell>
  <table:covered-table-cell/>
</table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>"#
        );
        write_fixture(&source, &content);
        let target = dir.path().join("converted.xlsx");
        convert_ods_to_xlsx(&source, &target).unwrap();
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<mergeCell ref="A1:B1"/>"#), "{sheet}");
    }

    #[test]
    fn does_not_materialize_a_huge_trailing_empty_row_repeat() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.ods");
        let content = format!(
            r#"<?xml version="1.0"?>
<office:document-content {NS}>
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
  <table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
</table:table-row>
<table:table-row table:number-rows-repeated="1048572"/>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>"#
        );
        write_fixture(&source, &content);
        let target = dir.path().join("converted.xlsx");
        let result = convert_ods_to_xlsx(&source, &target).unwrap();
        assert_eq!(result.cells, 1);
    }

    #[test]
    fn resolves_a_percentage_number_format_and_a_column_width() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.ods");
        let content = format!(
            r#"<?xml version="1.0"?>
<office:document-content {NS}>
<office:automatic-styles>
<number:percentage-style xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" style:name="N1">
  <number:number number:decimal-places="1"/>
  <number:text>%</number:text>
</number:percentage-style>
<style:style style:name="ce1" style:family="table-cell" style:data-style-name="N1"/>
<style:style style:name="co1" style:family="table-column">
  <style:table-column-properties style:column-width="2.5cm"/>
</style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-column table:style-name="co1"/>
<table:table-row>
  <table:table-cell table:style-name="ce1" office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>"#
        );
        write_fixture(&source, &content);
        let target = dir.path().join("converted.xlsx");
        convert_ods_to_xlsx(&source, &target).unwrap();

        let styles = read_entry(&target, "xl/styles.xml");
        assert!(styles.contains(r#"formatCode="0.0%""#), "{styles}");
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        // 2.5cm -> 94.49px @96dpi -> (94.49 - 5) / 7 ≈ 12.785 Excel character-width units
        assert!(sheet.contains(r#"<col min="1" max="1" width="12.785"#), "{sheet}");
    }

    #[test]
    fn translates_a_date_value_to_an_excel_serial() {
        let dir = tempfile::tempdir().unwrap();
        let source = dir.path().join("source.ods");
        let content = format!(
            r#"<?xml version="1.0"?>
<office:document-content {NS}>
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
  <table:table-cell office:value-type="date" office:date-value="2024-01-01"><text:p>2024-01-01</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>"#
        );
        write_fixture(&source, &content);
        let target = dir.path().join("converted.xlsx");
        convert_ods_to_xlsx(&source, &target).unwrap();
        let sheet = read_entry(&target, "xl/worksheets/sheet1.xml");
        // 2024-01-01 is Excel serial 45292
        assert!(sheet.contains("<v>45292</v>"), "{sheet}");
    }
}
