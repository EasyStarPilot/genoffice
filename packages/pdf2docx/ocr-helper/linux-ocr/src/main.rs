//! Linux system-OCR helper for pdf2docx (see ../../src/ocr.ts) — the Linux
//! counterpart of vision-ocr.swift / win-ocr.cs, speaking the exact same
//! protocol:
//!
//! Reads a PNG from stdin, writes recognized lines as JSON to stdout:
//!   {"paper":0.93,"lines":[{"t":"text","c":0.83,"b":[x0,y0,x1,y1],"chars":[{"t":"a","b":[...]}]}]}
//! Boxes are normalized 0-1, origin bottom-left, y up (PDF page space).
//! "paper" is the near-white pixel share (photo-vs-document signal).
//!
//! Unlike macOS/Windows, Linux has no universal built-in OCR API, so this
//! shells out to the system `tesseract` binary (the ubiquitous open-source
//! engine — `tesseract-ocr` on every major distro's package manager) instead
//! of bundling a recognizer. tesseract reports only word-level boxes (no
//! per-character geometry), so — exactly like win-ocr.cs — each word's box
//! is split evenly across its characters, and inter-word gaps in the line
//! text get zero boxes.
//!
//! argv[1] (optional): comma-separated language hints, e.g. "zh-Hans,en-US"
//! (BCP-47-ish, matching the mac/Windows helpers' own argv convention) —
//! mapped to tesseract's three-letter codes and filtered to what's actually
//! installed (`tesseract --list-langs`). Without a usable hint, every
//! installed language is passed together (tesseract's `-l a+b+c` multi-
//! language mode) — the closest analog this engine has to Vision's
//! auto-detect or Windows' user-profile default.
//!
//! Exit codes: 0 ok; 2 cannot decode input; 3 recognition failed;
//! 4 no OCR available (tesseract not installed, or no language data at
//! all) — callers treat any non-zero exit as "no engine" and keep the
//! bitmap fallback, the same contract win-ocr.cs's exit 4 already uses for
//! a machine with no OCR language.
//!
//! Build: cargo build --release (see build-linux.mjs / electron-builder.cjs).

use std::env;
use std::fs;
use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::process::Command;

const ERR_DECODE: i32 = 2;
const ERR_RECOGNIZE: i32 = 3;
const ERR_NO_ENGINE: i32 = 4;

struct Image {
    width: u32,
    height: u32,
    /// tightly packed RGBA8, row-major, top-left origin (pixel space)
    rgba: Vec<u8>,
}

/// One recognized word, in pixel space (top-left origin) at the source image's resolution.
struct Word {
    text: String,
    conf: f32,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
}

/// One reconstructed line — tesseract's TSV has no line-text row, so this is
/// built from consecutive same-(block,par,line) word rows, mirroring
/// win-ocr.cs's own line-from-words reconstruction.
struct Line {
    text: String,
    /// 0-1, averaged over the line's words (tesseract's own scale is 0-100)
    confidence: f64,
    px0: f64,
    py0: f64,
    px1: f64,
    py1: f64,
    /// (text, px0, py0, px1, py1); an all-zero box marks an inter-word gap
    chars: Vec<(String, f64, f64, f64, f64)>,
}

fn main() {
    std::process::exit(run());
}

fn run() -> i32 {
    let mut png_bytes = Vec::new();
    if io::stdin().read_to_end(&mut png_bytes).is_err() || png_bytes.is_empty() {
        eprintln!("empty input");
        return ERR_DECODE;
    }
    let image = match decode_png(&png_bytes) {
        Some(image) => image,
        None => {
            eprintln!("cannot decode input image");
            return ERR_DECODE;
        }
    };

    let installed = match installed_languages() {
        Some(langs) if !langs.is_empty() => langs,
        _ => {
            eprintln!("tesseract not installed or no language data available");
            return ERR_NO_ENGINE;
        }
    };
    let hints = env::args().nth(1).unwrap_or_default();
    let lang_arg = resolve_languages(&hints, &installed);

    let workdir = match make_tempdir() {
        Some(dir) => dir,
        None => {
            eprintln!("cannot create temp directory");
            return ERR_RECOGNIZE;
        }
    };
    let input_path = workdir.join("page.png");
    let output_base = workdir.join("page");
    if fs::write(&input_path, &png_bytes).is_err() {
        eprintln!("cannot write temp input file");
        let _ = fs::remove_dir_all(&workdir);
        return ERR_RECOGNIZE;
    }

    let run_result = Command::new("tesseract")
        .arg(&input_path)
        .arg(&output_base)
        .arg("-l")
        .arg(&lang_arg)
        .arg("tsv")
        .output();
    match run_result {
        Ok(output) if output.status.success() => {}
        Ok(output) => {
            eprintln!("recognition failed: {}", String::from_utf8_lossy(&output.stderr));
            let _ = fs::remove_dir_all(&workdir);
            return ERR_RECOGNIZE;
        }
        Err(err) => {
            eprintln!("recognition failed: {err}");
            let _ = fs::remove_dir_all(&workdir);
            return ERR_RECOGNIZE;
        }
    }

    let tsv = fs::read_to_string(workdir.join("page.tsv")).unwrap_or_default();
    let _ = fs::remove_dir_all(&workdir);

    let lines = parse_tsv(&tsv);
    let paper = paper_share(&image);
    print_json(&lines, image.width, image.height, paper);
    0
}

// ── tesseract process helpers ──

fn installed_languages() -> Option<Vec<String>> {
    let output = Command::new("tesseract").arg("--list-langs").output().ok()?;
    if !output.status.success() {
        return None;
    }
    // tesseract prints "List of available languages (N):" then one per line;
    // read both streams since the exact stream has varied across versions.
    let combined = format!(
        "{}\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
    let langs: Vec<String> = combined
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with("List of available languages"))
        // osd = orientation/script-detection data, not a recognizable language
        .filter(|&line| line != "osd")
        .map(str::to_string)
        .collect();
    Some(langs)
}

/// `hints` is the raw argv[1] (comma-separated BCP-47-ish tags, possibly
/// empty). Maps each to a tesseract code, keeps only what's installed
/// (preserving hint order, deduped), and falls back to every installed
/// language together when no hint survives that filter.
fn resolve_languages(hints: &str, installed: &[String]) -> String {
    let mut chosen: Vec<String> = Vec::new();
    for tag in hints.split(',') {
        let tag = tag.trim();
        if tag.is_empty() {
            continue;
        }
        if let Some(code) = bcp47_to_tesseract(tag) {
            if installed.iter().any(|l| l == code) && !chosen.iter().any(|c| c == code) {
                chosen.push(code.to_string());
            }
        }
    }
    if chosen.is_empty() {
        chosen = installed.to_vec();
    }
    chosen.join("+")
}

fn bcp47_to_tesseract(tag: &str) -> Option<&'static str> {
    let lower = tag.to_ascii_lowercase();
    if lower.starts_with("zh") {
        return Some(if lower.contains("hant") || lower.contains("tw") || lower.contains("hk") {
            "chi_tra"
        } else {
            "chi_sim"
        });
    }
    let primary = lower.split('-').next().unwrap_or(&lower);
    Some(match primary {
        "en" => "eng",
        "ja" => "jpn",
        "ko" => "kor",
        "de" => "deu",
        "es" => "spa",
        "fr" => "fra",
        "it" => "ita",
        "pt" => "por",
        "ru" => "rus",
        "ar" => "ara",
        "hi" => "hin",
        "th" => "tha",
        "vi" => "vie",
        "id" => "ind",
        "nl" => "nld",
        "pl" => "pol",
        "tr" => "tur",
        "he" => "heb",
        _ => return None,
    })
}

fn make_tempdir() -> Option<PathBuf> {
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .ok()?
        .as_nanos();
    let dir = env::temp_dir().join(format!("linux-ocr-{}-{nanos}", std::process::id()));
    fs::create_dir_all(&dir).ok()?;
    Some(dir)
}

// ── TSV -> lines ──

/// tesseract TSV columns: level page_num block_num par_num line_num word_num
/// left top width height conf text — level 5 rows are words; this groups
/// consecutive same-(page,block,par,line) word rows into a line, the same
/// reconstruction win-ocr.cs does from its own OcrLine/OcrWord objects.
fn parse_tsv(tsv: &str) -> Vec<Line> {
    let mut lines = Vec::new();
    let mut current_key: Option<(i64, i64, i64, i64)> = None;
    let mut current_words: Vec<Word> = Vec::new();

    for row in tsv.lines().skip(1) {
        let cols: Vec<&str> = row.splitn(12, '\t').collect();
        if cols.len() < 12 || cols[0] != "5" {
            continue;
        }
        let key = (
            cols[1].parse().unwrap_or(0),
            cols[2].parse().unwrap_or(0),
            cols[3].parse().unwrap_or(0),
            cols[4].parse().unwrap_or(0),
        );
        let (left, top, width, height, conf) = (
            cols[6].parse().unwrap_or(0.0),
            cols[7].parse().unwrap_or(0.0),
            cols[8].parse().unwrap_or(0.0),
            cols[9].parse().unwrap_or(0.0),
            cols[10].parse().unwrap_or(-1.0),
        );
        let text = cols[11].to_string();
        if text.trim().is_empty() {
            continue;
        }
        if current_key.is_some() && current_key != Some(key) {
            if let Some(line) = build_line(&current_words) {
                lines.push(line);
            }
            current_words.clear();
        }
        current_key = Some(key);
        current_words.push(Word { text, conf, left, top, width, height });
    }
    if let Some(line) = build_line(&current_words) {
        lines.push(line);
    }
    lines
}

fn build_line(words: &[Word]) -> Option<Line> {
    if words.is_empty() {
        return None;
    }
    let text = words.iter().map(|w| w.text.as_str()).collect::<Vec<_>>().join(" ");
    if text.trim().is_empty() {
        return None;
    }
    let mut px0 = f64::MAX;
    let mut py0 = f64::MAX;
    let mut px1 = f64::MIN;
    let mut py1 = f64::MIN;
    let mut chars: Vec<(String, f64, f64, f64, f64)> = Vec::new();
    for (i, word) in words.iter().enumerate() {
        if i > 0 {
            chars.push((" ".to_string(), 0.0, 0.0, 0.0, 0.0));
        }
        px0 = px0.min(word.left);
        py0 = py0.min(word.top);
        px1 = px1.max(word.left + word.width);
        py1 = py1.max(word.top + word.height);
        let n = (word.text.chars().count().max(1)) as f64;
        for (i, ch) in word.text.chars().enumerate() {
            let cx0 = word.left + (word.width * i as f64) / n;
            let cx1 = word.left + (word.width * (i as f64 + 1.0)) / n;
            chars.push((ch.to_string(), cx0, word.top, cx1, word.top + word.height));
        }
    }
    if px1 <= px0 || py1 <= py0 {
        return None;
    }
    let confidence = words.iter().map(|w| w.conf.max(0.0) as f64).sum::<f64>()
        / words.len() as f64
        / 100.0;
    Some(Line { text, confidence, px0, py0, px1, py1, chars })
}

// ── image decode + paper-tone sample ──

fn decode_png(bytes: &[u8]) -> Option<Image> {
    let mut decoder = png::Decoder::new(bytes);
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
    let mut reader = decoder.read_info().ok()?;
    let mut buf = vec![0u8; reader.output_buffer_size()];
    let info = reader.next_frame(&mut buf).ok()?;
    let (width, height) = (info.width, info.height);
    let data = &buf[..info.buffer_size()];
    let rgba = match info.color_type {
        png::ColorType::Rgba => data.to_vec(),
        png::ColorType::Rgb => {
            let mut out = Vec::with_capacity(data.len() / 3 * 4);
            for px in data.chunks_exact(3) {
                out.extend_from_slice(px);
                out.push(255);
            }
            out
        }
        png::ColorType::GrayscaleAlpha => {
            let mut out = Vec::with_capacity(data.len() * 2);
            for px in data.chunks_exact(2) {
                out.extend_from_slice(&[px[0], px[0], px[0], px[1]]);
            }
            out
        }
        png::ColorType::Grayscale => {
            let mut out = Vec::with_capacity(data.len() * 4);
            for &g in data {
                out.extend_from_slice(&[g, g, g, 255]);
            }
            out
        }
        // EXPAND already lifts Indexed to Rgb/Rgba; anything else is unexpected
        _ => return None,
    };
    Some(Image { width, height, rgba })
}

/// Near-white pixel share over strided samples — same signal, threshold and
/// stride formula as vision-ocr.swift / win-ocr.cs's own paperShare/PaperShare,
/// so the TS policy layer's gate behaves identically across platforms.
/// Transparent/partial-alpha pixels are blended toward white first (straight,
/// non-premultiplied alpha — PNG's default), matching Vision's own
/// white-fill-then-composite handling of transparent sources.
fn paper_share(image: &Image) -> f64 {
    let total = (image.width as usize) * (image.height as usize);
    if total == 0 {
        return 1.0;
    }
    let stride = (total / 200_000).max(1);
    let blend_to_white = |c: u8, a: u8| -> u32 {
        (c as u32 * a as u32 + 255 * (255 - a as u32)) / 255
    };
    let mut paper = 0usize;
    let mut sampled = 0usize;
    let mut i = 0usize;
    while i < total {
        let o = i * 4;
        let (r, g, b, a) = (image.rgba[o], image.rgba[o + 1], image.rgba[o + 2], image.rgba[o + 3]);
        let (r, g, b) = if a == 255 {
            (r as u32, g as u32, b as u32)
        } else {
            (blend_to_white(r, a), blend_to_white(g, a), blend_to_white(b, a))
        };
        if r.min(g).min(b) >= 200 {
            paper += 1;
        }
        sampled += 1;
        i += stride;
    }
    if sampled > 0 {
        paper as f64 / sampled as f64
    } else {
        1.0
    }
}

// ── JSON output ──

fn print_json(lines: &[Line], width: u32, height: u32, paper: f64) {
    let mut out = String::new();
    out.push_str("{\"paper\":");
    out.push_str(&num(paper));
    out.push_str(",\"lines\":[");
    for (i, line) in lines.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        out.push_str("{\"t\":");
        out.push_str(&escape_json(&line.text));
        out.push_str(",\"c\":");
        out.push_str(&num(line.confidence));
        out.push_str(",\"b\":");
        out.push_str(&box_json(line.px0, line.py0, line.px1, line.py1, width, height));
        out.push_str(",\"chars\":[");
        for (j, (text, cx0, cy0, cx1, cy1)) in line.chars.iter().enumerate() {
            if j > 0 {
                out.push(',');
            }
            out.push_str("{\"t\":");
            out.push_str(&escape_json(text));
            out.push_str(",\"b\":");
            if *cx0 == 0.0 && *cy0 == 0.0 && *cx1 == 0.0 && *cy1 == 0.0 {
                out.push_str("[0,0,0,0]");
            } else {
                out.push_str(&box_json(*cx0, *cy0, *cx1, *cy1, width, height));
            }
            out.push('}');
        }
        out.push_str("]}");
    }
    out.push_str("]}");
    let _ = io::stdout().write_all(out.as_bytes());
}

/// pixel rect (origin top-left) -> normalized PDF-space box (origin bottom-left)
fn box_json(px0: f64, py0: f64, px1: f64, py1: f64, width: u32, height: u32) -> String {
    let (w, h) = (width as f64, height as f64);
    format!(
        "[{},{},{},{}]",
        num(px0 / w),
        num(1.0 - py1 / h),
        num(px1 / w),
        num(1.0 - py0 / h),
    )
}

fn num(v: f64) -> String {
    let mut s = format!("{v:.6}");
    while s.ends_with('0') {
        s.pop();
    }
    if s.ends_with('.') {
        s.pop();
    }
    if s.is_empty() || s == "-" {
        s = "0".to_string();
    }
    s
}

fn escape_json(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}
