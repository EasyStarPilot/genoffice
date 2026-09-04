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
}
