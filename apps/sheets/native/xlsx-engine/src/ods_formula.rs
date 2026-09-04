//! OpenFormula (ODFF) -> Excel A1 formula text translation.
//!
//! ODS stores formulas as `table:formula="of:=SUM([.A1:.A2])"`: an `of:`
//! namespace prefix, then an `=`-led expression whose cell/range references
//! are bracketed and dot-separated (`[.A1]`, `[.A1:.B2]`, `[Sheet2.A1]`,
//! `[$Sheet2.$A$1:.$B$2]`) instead of Excel's bare `A1` / `Sheet2!A1`. Function
//! names are not translated — OpenFormula defines them to match Excel's for
//! every common function (SUM, IF, VLOOKUP, ...), and this crate has no
//! existing table of the rare exceptions to remap. `ironcalc` (the recalc
//! engine) and every existing formula consumer in this crate expect plain A1
//! text, so this translation must run before an ODS-origin formula reaches
//! `xl/worksheets/sheetN.xml`.
//!
//! A small character scanner, not a general parser: it only has to find
//! bracketed references and string literals (so a `[` inside a quoted string
//! is never mistaken for one) and copy everything else through unchanged.

/// Translates one `table:formula` value into Excel A1 formula text (still
/// including the leading `=`). Returns the input unchanged, `=`-prefixed if
/// it wasn't already, when nothing bracket-shaped is found — a formula with
/// no cell references (e.g. `=1+1`) is already valid A1 text as-is.
pub fn openformula_to_a1(formula: &str) -> String {
    let body = formula.strip_prefix("of:").unwrap_or(formula);
    let body = body.strip_prefix('=').unwrap_or(body);
    let mut out = String::with_capacity(body.len());
    let chars: Vec<char> = body.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                // copy the whole string literal (with its escaped "" pairs) verbatim
                out.push('"');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    let is_quote = chars[i] == '"';
                    i += 1;
                    if is_quote && chars.get(i) != Some(&'"') {
                        break;
                    }
                    if is_quote {
                        // escaped "" inside the literal: consume the second quote too
                        out.push(chars[i]);
                        i += 1;
                    }
                }
            }
            '[' => {
                if let Some((reference, next)) = translate_bracket_ref(&chars, i) {
                    out.push_str(&reference);
                    i = next;
                } else {
                    // not a recognizable [...] reference (or unterminated) — copy as-is
                    out.push('[');
                    i += 1;
                }
            }
            other => {
                out.push(other);
                i += 1;
            }
        }
    }
    format!("={out}")
}

/// `chars[start]` is the opening `[` of a candidate reference. Returns the
/// translated A1 text and the index just past the closing `]`, or `None` if
/// this bracket group isn't a well-formed cell/range reference (left as-is —
/// OpenFormula also uses `[...]` for a handful of other constructs this
/// translator does not model, e.g. database-range references).
fn translate_bracket_ref(chars: &[char], start: usize) -> Option<(String, usize)> {
    let close = chars[start + 1..].iter().position(|&c| c == ']')? + start + 1;
    let inner: String = chars[start + 1..close].iter().collect();
    let translated = translate_reference_body(&inner)?;
    Some((translated, close + 1))
}

/// `body` is the text between `[` and `]`, e.g. `.A1`, `.A1:.B2`,
/// `Sheet2.A1`, `$Sheet 1$.$A$1:.$B$2`, `'Quoted Sheet'.A1`.
///
/// Known gap: a genuine 3D range spanning different start/end sheets
/// (`[Sheet2.A1:Sheet3.A1]`, "A1 on every sheet from 2 through 3") has no
/// direct A1 equivalent for a single cell reference and would need Excel's
/// distinct `Sheet2:Sheet3!A1` 3D-reference syntax; this translator instead
/// keeps only the start side's sheet, narrowing the reference to that one
/// sheet. Rare in practice (most producers only vary the sheet when both
/// sides already agree) and produces a valid, if narrower, formula rather
/// than an invalid one.
fn translate_reference_body(body: &str) -> Option<String> {
    let (first, second) = match body.split_once(':') {
        Some((a, b)) => (a, Some(b)),
        None => (body, None),
    };
    let start = translate_one_ref(first)?;
    match second {
        None => Some(start.text),
        Some(second) => {
            let end = translate_one_ref(second)?;
            // a range only carries the sheet prefix once, on the start side
            Some(format!("{}:{}", start.text, end.cell))
        }
    }
}

struct RefPart {
    /// full A1 text including a `Sheet!` prefix when this side names one
    text: String,
    /// just the cell/column/row part, no sheet prefix (used for a range's end side)
    cell: String,
}

/// One side of a reference: `.A1`, `$Sheet 1$.$A$1`, `'Quoted'.$A1`,
/// `.A:.A` (column) or `.1:.1` (row) is handled by the caller splitting on `:`
/// first, so this only ever sees one `sheet.cell` or `.cell` segment.
fn translate_one_ref(part: &str) -> Option<RefPart> {
    // The overwhelmingly common case is a same-sheet reference, which ODFF
    // still writes with a leading dot (`.A1`, not bare `A1`) — rsplit_once
    // on that yields Some(("", "A1")), an empty sheet side, not None, so it
    // must be checked for explicitly rather than trusted as "there is a
    // sheet name here".
    let (sheet, cell) = match part.rsplit_once('.') {
        Some(("", cell)) => (None, cell),
        Some((sheet, cell)) => (Some(sheet), cell),
        None => (None, part),
    };
    // OpenFormula's own absolute markers ($ before the sheet and/or before
    // the column/row) are a superset of Excel's — Excel only ever uses $ on
    // the cell side, so a sheet-side $ is simply dropped.
    let cell = cell.to_string();
    if cell.is_empty() {
        return None;
    }
    let text = match sheet {
        None => cell.clone(),
        Some(sheet) => {
            // OpenFormula quotes a sheet name needing it (spaces/specials) the
            // same way Excel does — 'It''s Sheet', doubled quotes escaping a
            // literal one — so the quoting is passed through unchanged; only
            // OpenFormula's own sheet-absolute `$` marker (which Excel has no
            // equivalent for) is stripped.
            let name = sheet.trim_start_matches('$');
            format!("{name}!{cell}")
        }
    };
    Some(RefPart { text, cell })
}

/// Translates an Excel A1 formula (as stored in `CellRecord.formula` — no
/// leading `=`, matching how OOXML's own `<f>` element never carries one)
/// into ODF's `table:formula` value, `of:=`-prefixed with every bare cell/
/// range reference re-bracketed and dot-separated. The reverse of
/// [`openformula_to_a1`], and the harder direction: A1 has no delimiter
/// marking where a reference starts, so this has to tell a reference token
/// apart from a defined name, a function name, or plain text by shape and
/// context alone (word boundaries on both sides, and never claiming a token
/// immediately followed by `(` — that is a function call, not a reference).
pub fn a1_to_openformula(formula: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                out.push('"');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    let is_quote = chars[i] == '"';
                    i += 1;
                    if is_quote && chars.get(i) != Some(&'"') {
                        break;
                    }
                    if is_quote {
                        out.push(chars[i]);
                        i += 1;
                    }
                }
            }
            c if is_reference_start(c, i, &chars) => {
                if let Some((token, next)) = try_parse_reference(&chars, i) {
                    out.push_str(&token);
                    i = next;
                } else {
                    out.push(c);
                    i += 1;
                }
            }
            other => {
                out.push(other);
                i += 1;
            }
        }
    }
    format!("of:={out}")
}

/// A reference (or its leading sheet-name quote) may only start where the
/// previous character isn't itself part of an identifier — otherwise this
/// would misfire inside a longer name like a defined name or function.
fn is_reference_start(c: char, index: usize, chars: &[char]) -> bool {
    if !(c.is_ascii_alphabetic() || c == '\'' || c == '$') {
        return false;
    }
    index == 0 || !is_ident_char(chars[index - 1])
}

fn is_ident_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '$'
}

/// `chars[start]` begins a candidate `[Sheet!]A1[:B2]` reference. Returns
/// the bracketed OpenFormula text and the index just past it, or `None` if
/// this position isn't actually a well-formed, properly-bounded reference
/// (left as plain text — e.g. a defined name, or `LOG10(` where the digits
/// belong to a function name, never a cell address, because it's followed
/// by `(`).
fn try_parse_reference(chars: &[char], start: usize) -> Option<(String, usize)> {
    let mut pos = start;
    let sheet = parse_sheet_prefix(chars, &mut pos);
    let (col1, row1, after_first) = parse_cell_address(chars, pos)?;
    let (end, second_cell) = if chars.get(after_first) == Some(&':') {
        let (col2, row2, after_second) = parse_cell_address(chars, after_first + 1)?;
        (after_second, Some((col2, row2)))
    } else {
        (after_first, None)
    };
    // a reference is never immediately followed by `(` (that's a call) or by
    // another identifier character (that's a longer name, e.g. a defined
    // name that happens to start with something reference-shaped)
    if chars.get(end).is_some_and(|&c| c == '(' || is_ident_char(c)) {
        return None;
    }
    let sheet_prefix = sheet.map(|s| format!("{s}.")).unwrap_or_else(|| ".".to_string());
    let text = match second_cell {
        None => format!("[{sheet_prefix}{col1}{row1}]"),
        Some((col2, row2)) => format!("[{sheet_prefix}{col1}{row1}:.{col2}{row2}]"),
    };
    Some((text, end))
}

/// An optional `SheetName!` or `'Quoted Name'!` prefix at `chars[*pos]`,
/// advancing `*pos` past it (including the `!`) only when a real prefix was
/// found — otherwise `*pos` is left untouched for the cell-address parse
/// that follows.
fn parse_sheet_prefix(chars: &[char], pos: &mut usize) -> Option<String> {
    let start = *pos;
    if chars.get(start) == Some(&'\'') {
        let mut i = start + 1;
        let mut name = String::new();
        loop {
            let c = *chars.get(i)?;
            i += 1;
            if c == '\'' {
                if chars.get(i) == Some(&'\'') {
                    name.push('\'');
                    i += 1;
                    continue;
                }
                break;
            }
            name.push(c);
        }
        if chars.get(i) == Some(&'!') {
            *pos = i + 1;
            return Some(format!("'{}'", name.replace('\'', "''")));
        }
        return None;
    }
    let mut i = start;
    while chars.get(i).is_some_and(|&c| c.is_ascii_alphanumeric() || c == '_' || c == '.') {
        i += 1;
    }
    if i > start && chars.get(i) == Some(&'!') {
        let name: String = chars[start..i].iter().collect();
        *pos = i + 1;
        return Some(name);
    }
    None
}

/// `[$]?[A-Za-z]{1,3}[$]?[0-9]+` at `chars[pos]` — returns the column
/// letters, the row digits (with any `$` markers kept, since OpenFormula
/// uses the identical convention Excel does), and the index just past it.
fn parse_cell_address(chars: &[char], pos: usize) -> Option<(String, String, usize)> {
    let mut i = pos;
    let col_dollar = if chars.get(i) == Some(&'$') {
        i += 1;
        "$"
    } else {
        ""
    };
    let col_start = i;
    while chars.get(i).is_some_and(|c| c.is_ascii_alphabetic()) && i - col_start < 3 {
        i += 1;
    }
    if i == col_start {
        return None;
    }
    let letters: String = chars[col_start..i].iter().collect();
    if !letters.chars().all(|c| c.is_ascii_uppercase()) && !letters.chars().all(|c| c.is_ascii_lowercase()) {
        // Excel column letters are never mixed-case as a single token in
        // practice; this is almost certainly the start of a longer identifier
        return None;
    }
    let row_dollar = if chars.get(i) == Some(&'$') {
        i += 1;
        "$"
    } else {
        ""
    };
    let row_start = i;
    while chars.get(i).is_some_and(|c| c.is_ascii_digit()) {
        i += 1;
    }
    if i == row_start {
        return None;
    }
    let digits: String = chars[row_start..i].iter().collect();
    Some((format!("{col_dollar}{letters}"), format!("{row_dollar}{digits}"), i))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_the_of_prefix_and_equals_sign() {
        assert_eq!(openformula_to_a1("of:=1+1"), "=1+1");
        assert_eq!(openformula_to_a1("=1+1"), "=1+1");
    }

    #[test]
    fn translates_a_simple_cell_reference() {
        assert_eq!(openformula_to_a1("of:=[.A1]+1"), "=A1+1");
    }

    #[test]
    fn translates_a_range_reference() {
        assert_eq!(openformula_to_a1("of:=SUM([.A1:.A2])"), "=SUM(A1:A2)");
    }

    #[test]
    fn translates_absolute_references() {
        assert_eq!(openformula_to_a1("of:=[.$A$1]"), "=$A$1");
    }

    #[test]
    fn translates_a_cross_sheet_reference() {
        assert_eq!(openformula_to_a1("of:=SUM([Sheet2.A1:.B2])"), "=SUM(Sheet2!A1:B2)");
    }

    #[test]
    fn translates_a_quoted_sheet_name_with_a_space() {
        assert_eq!(
            openformula_to_a1("of:=['My Sheet'.A1]"),
            "='My Sheet'!A1",
        );
    }

    #[test]
    fn leaves_string_literals_untouched_even_with_brackets_inside() {
        assert_eq!(
            openformula_to_a1(r#"of:=IF([.A1]="[x]","yes","no")"#),
            r#"=IF(A1="[x]","yes","no")"#
        );
    }

    #[test]
    fn leaves_a_formula_with_no_references_unchanged_besides_the_prefix() {
        assert_eq!(openformula_to_a1("of:=1+2*3"), "=1+2*3");
    }

    #[test]
    fn a1_translates_a_simple_reference() {
        assert_eq!(a1_to_openformula("A1+1"), "of:=[.A1]+1");
    }

    #[test]
    fn a1_translates_a_range() {
        assert_eq!(a1_to_openformula("SUM(A1:A2)"), "of:=SUM([.A1:.A2])");
    }

    #[test]
    fn a1_translates_absolute_references() {
        assert_eq!(a1_to_openformula("$A$1"), "of:=[.$A$1]");
    }

    #[test]
    fn a1_translates_a_cross_sheet_range() {
        assert_eq!(a1_to_openformula("SUM(Sheet2!A1:B2)"), "of:=SUM([Sheet2.A1:.B2])");
    }

    #[test]
    fn a1_translates_a_quoted_sheet_name_with_a_space() {
        assert_eq!(a1_to_openformula("'My Sheet'!A1"), "of:=['My Sheet'.A1]");
    }

    #[test]
    fn a1_leaves_string_literals_untouched() {
        assert_eq!(
            a1_to_openformula(r#"IF(A1="B2","yes","no")"#),
            r#"of:=IF([.A1]="B2","yes","no")"#
        );
    }

    #[test]
    fn a1_does_not_mistake_a_function_name_for_a_reference() {
        // LOG10( looks reference-shaped (letters then digits) but is a call
        assert_eq!(a1_to_openformula("LOG10(A1)"), "of:=LOG10([.A1])");
        assert_eq!(a1_to_openformula("ATAN2(A1,B1)"), "of:=ATAN2([.A1],[.B1])");
    }

    #[test]
    fn a1_does_not_mistake_a_longer_identifier_for_a_reference() {
        // neither a defined name containing a reference-shaped substring...
        assert_eq!(a1_to_openformula("TOTAL_A1"), "of:=TOTAL_A1");
        // ...nor one immediately followed by more identifier characters
        assert_eq!(a1_to_openformula("A1B"), "of:=A1B");
    }

    #[test]
    fn a1_round_trips_through_openformula_to_a1() {
        let original = "SUM(A1:A2)+Sheet2!$B$3";
        let openformula = a1_to_openformula(original);
        assert_eq!(openformula_to_a1(&openformula), format!("={original}"));
    }
}
