use serde::Serialize;
use serde_json::{json, Map, Number, Value};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::env;
use std::error::Error;
use std::fs;
use std::path::Path;

const FONT: &str = "맑은 고딕";
const C_NAVY: &str = "1F3864";
const HEAD_FILL: &str = "D6E0F2";
const SUB_FILL: &str = "EDF2FA";
const C_WHITE: &str = "FFFFFF";
const C_GRID: &str = "C9D2E0";
const C_HEAVY: &str = "404040";
const C_TEXT: &str = "1A1A1A";
const C_MUTE: &str = "44546A";
const THIN: &str = "thin";
const HEAVY: &str = "medium";

#[derive(Clone, Copy)]
struct Al {
    h: &'static str,
    ind: usize,
    wrap: bool,
}

const A_CENTER: Al = Al {
    h: "center",
    ind: 0,
    wrap: true,
};
const A_LEFT: Al = Al {
    h: "left",
    ind: 0,
    wrap: true,
};
const A_RIGHT: Al = Al {
    h: "right",
    ind: 0,
    wrap: false,
};
const A_LEFT_IND: Al = Al {
    h: "left",
    ind: 1,
    wrap: false,
};

#[derive(Clone)]
struct Theme {
    navy: Option<&'static str>,
    head: Option<&'static str>,
    sub: Option<&'static str>,
    white: &'static str,
    text: &'static str,
    mute: &'static str,
    grid: &'static str,
    heavy: &'static str,
}

fn theme(name: &str) -> Result<Theme, Box<dyn Error>> {
    match name {
        "color" => Ok(Theme {
            navy: Some(C_NAVY),
            head: Some(HEAD_FILL),
            sub: Some(SUB_FILL),
            white: C_WHITE,
            text: C_TEXT,
            mute: C_MUTE,
            grid: C_GRID,
            heavy: C_HEAVY,
        }),
        "grey" => Ok(Theme {
            navy: Some("595959"),
            head: Some("D9D9D9"),
            sub: Some("F2F2F2"),
            white: C_WHITE,
            text: "1A1A1A",
            mute: "404040",
            grid: "BFBFBF",
            heavy: "262626",
        }),
        "mono" => Ok(Theme {
            navy: None,
            head: None,
            sub: None,
            white: "000000",
            text: "000000",
            mute: "333333",
            grid: "808080",
            heavy: "000000",
        }),
        _ => Err(format!("theme must be color, grey or mono: {name}").into()),
    }
}

#[derive(Clone, Debug, Serialize)]
struct GCell {
    value: Option<Value>,
    text: String,
    numfmt: Option<String>,
    fill: Option<String>,
    bold: bool,
    size: usize,
    halign: Option<String>,
    indent: usize,
    wrap: bool,
    color: String,
    top: Option<String>,
    left: Option<String>,
    right: Option<String>,
    bottom: Option<String>,
}

impl Default for GCell {
    fn default() -> Self {
        Self {
            value: None,
            text: String::new(),
            numfmt: None,
            fill: None,
            bold: false,
            size: 10,
            halign: None,
            indent: 0,
            wrap: false,
            color: C_TEXT.to_string(),
            top: None,
            left: None,
            right: None,
            bottom: None,
        }
    }
}

#[derive(Clone, Debug)]
struct Grid {
    g: BTreeMap<(usize, usize), GCell>,
    merges: Vec<(usize, usize, usize, usize)>,
    soft_merges: Vec<(usize, usize, usize, usize)>,
    widths: BTreeMap<usize, f64>,
    heights: BTreeMap<usize, f64>,
    outline_rows: BTreeMap<usize, usize>,
    outline_cols: BTreeMap<usize, usize>,
    summary_below: Option<bool>,
    summary_right: Option<bool>,
    freeze: Option<(usize, usize)>,
    nr: usize,
    nc: usize,
    depth: Option<usize>,
}

impl Grid {
    fn new() -> Self {
        Self {
            g: BTreeMap::new(),
            merges: Vec::new(),
            soft_merges: Vec::new(),
            widths: BTreeMap::new(),
            heights: BTreeMap::new(),
            outline_rows: BTreeMap::new(),
            outline_cols: BTreeMap::new(),
            summary_below: None,
            summary_right: None,
            freeze: None,
            nr: 0,
            nc: 0,
            depth: None,
        }
    }

    fn cell(&mut self, r: usize, c: usize) -> &mut GCell {
        self.nr = self.nr.max(r);
        self.nc = self.nc.max(c);
        self.g.entry((r, c)).or_default()
    }
}

#[derive(Clone, Debug)]
struct Node {
    label: String,
    summary: Option<Vec<Value>>,
    values: Option<Vec<Value>>,
    children: Vec<Node>,
}

fn py_str(v: &Value) -> String {
    match v {
        Value::Null => "None".to_string(),
        Value::Bool(b) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
        Value::Number(n) => n.to_string(),
        Value::String(s) => s.clone(),
        _ => v.to_string(),
    }
}

fn obj<'a>(v: &'a Value, key: &str) -> Option<&'a Value> {
    v.as_object().and_then(|m| m.get(key))
}

fn req<'a>(v: &'a Value, key: &str) -> Result<&'a Value, Box<dyn Error>> {
    obj(v, key).ok_or_else(|| format!("missing key: {key}").into())
}

fn arr<'a>(v: &'a Value, key: &str) -> Result<&'a Vec<Value>, Box<dyn Error>> {
    req(v, key)?
        .as_array()
        .ok_or_else(|| format!("{key} must be an array").into())
}

fn opt_array(v: &Value, key: &str) -> Option<Vec<Value>> {
    obj(v, key).and_then(|x| x.as_array()).cloned()
}

fn label(v: &Value) -> Result<String, Box<dyn Error>> {
    Ok(py_str(req(v, "label")?))
}

fn parse_node(v: &Value) -> Result<Node, Box<dyn Error>> {
    let children = match obj(v, "children").and_then(|x| x.as_array()) {
        Some(items) => {
            let mut out = Vec::new();
            for ch in items {
                out.push(parse_node(ch)?);
            }
            out
        }
        None => Vec::new(),
    };
    Ok(Node {
        label: label(v)?,
        summary: opt_array(v, "summary"),
        values: opt_array(v, "values"),
        children,
    })
}

fn normalize_items(m: &Value) -> Result<Vec<Node>, Box<dyn Error>> {
    let mut out = Vec::new();
    for it in arr(m, "items")? {
        let mut children = Vec::new();
        if let Some(details) = obj(it, "details").and_then(|x| x.as_array()) {
            for d in details {
                children.push(Node {
                    label: label(d)?,
                    summary: None,
                    values: opt_array(d, "values"),
                    children: Vec::new(),
                });
            }
        }
        out.push(Node {
            label: label(it)?,
            summary: opt_array(it, "summary"),
            values: None,
            children,
        });
    }
    Ok(out)
}

fn nodes_of(m: &Value) -> Result<Vec<Node>, Box<dyn Error>> {
    if let Some(nodes) = obj(m, "nodes").and_then(|x| x.as_array()) {
        let mut out = Vec::new();
        for n in nodes {
            out.push(parse_node(n)?);
        }
        Ok(out)
    } else {
        normalize_items(m)
    }
}

fn max_depth(nodes: &[Node]) -> usize {
    nodes
        .iter()
        .map(|n| {
            if n.children.is_empty() {
                1
            } else {
                1 + max_depth(&n.children)
            }
        })
        .max()
        .unwrap_or(0)
}

fn tree_size(nodes: &[Node]) -> usize {
    nodes
        .iter()
        .map(|n| 1 + if n.children.is_empty() { 0 } else { tree_size(&n.children) })
        .sum()
}

fn looks_num(v: &str) -> bool {
    let s = v.trim().replace(',', "").trim_end_matches('%').to_string();
    s.parse::<f64>().is_ok()
}

fn value_number(v: f64) -> Value {
    Number::from_f64(v).map(Value::Number).unwrap_or(Value::Null)
}

fn comma_int(s: &str) -> String {
    let mut sign = "";
    let mut digits = s;
    if let Some(rest) = s.strip_prefix('-') {
        sign = "-";
        digits = rest;
    }
    let mut out = String::new();
    for (i, ch) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    let body: String = out.chars().rev().collect();
    format!("{sign}{body}")
}

fn comma_number(n: &Number) -> String {
    let s = n.to_string();
    if let Some((intp, frac)) = s.split_once('.') {
        format!("{}.{}", comma_int(intp), frac)
    } else {
        comma_int(&s)
    }
}

fn disp(value: &Value) -> (Option<Value>, Option<String>, String, Option<String>) {
    match value {
        Value::String(s) => {
            let trimmed = s.trim();
            if let Some(num) = trimmed.strip_suffix('%') {
                let body = num.trim();
                let valid = {
                    let mut chars = body.chars();
                    let first_ok = matches!(chars.next(), Some('-') | Some('0'..='9'));
                    first_ok && body.parse::<f64>().is_ok()
                };
                if valid && !body.is_empty() {
                    let dec = body
                        .split_once('.')
                        .map(|(_, frac)| frac.chars().count())
                        .unwrap_or(0);
                    let f = body.parse::<f64>().unwrap_or(0.0) / 100.0;
                    let fmt = if dec == 0 {
                        "0%".to_string()
                    } else {
                        format!("0.{}%", "0".repeat(dec))
                    };
                    return (
                        Some(value_number(f)),
                        Some(fmt),
                        trimmed.to_string(),
                        Some("right".to_string()),
                    );
                }
            }
            (
                Some(Value::String(s.clone())),
                None,
                s.clone(),
                Some(if looks_num(s) { "right" } else { "left" }.to_string()),
            )
        }
        Value::Bool(b) => (
            Some(Value::Bool(*b)),
            None,
            if *b { "True" } else { "False" }.to_string(),
            Some("left".to_string()),
        ),
        Value::Number(n) => (
            Some(Value::Number(n.clone())),
            Some("#,##0".to_string()),
            comma_number(n),
            Some("right".to_string()),
        ),
        Value::Null => (None, None, String::new(), None),
        _ => (Some(value.clone()), None, py_str(value), None),
    }
}

fn fmt_cell(cell: &mut GCell, value: &Value) {
    let (v, nf, text, ha) = disp(value);
    cell.value = v;
    cell.numfmt = nf;
    cell.text = text;
    if let Some(h) = ha {
        cell.halign = Some(h);
    }
}

fn lbl<'a>(grid: &'a mut Grid, r: usize, c: usize, label: Option<&Value>) -> &'a mut GCell {
    let cell = grid.cell(r, c);
    if let Some(v) = label {
        if !matches!(v, Value::Null) {
            let s = py_str(v);
            if !s.is_empty() {
                cell.value = Some(match v {
                    Value::String(_) => Value::String(s.clone()),
                    _ => Value::String(s.clone()),
                });
                cell.text = s;
            }
        }
    }
    cell
}

fn lbl_str<'a>(grid: &'a mut Grid, r: usize, c: usize, label: Option<&str>) -> &'a mut GCell {
    let cell = grid.cell(r, c);
    if let Some(s) = label {
        if !s.is_empty() {
            cell.value = Some(Value::String(s.to_string()));
            cell.text = s.to_string();
        }
    }
    cell
}

fn style(
    cell: &mut GCell,
    bold: bool,
    size: usize,
    color: &str,
    fill: Option<&str>,
    align: Option<Al>,
) {
    cell.bold = bold;
    cell.size = size;
    cell.color = color.to_string();
    if let Some(f) = fill {
        cell.fill = Some(f.to_string());
    }
    if let Some(a) = align {
        cell.halign = Some(a.h.to_string());
        cell.indent = a.ind;
        cell.wrap = a.wrap;
    }
}

fn merge(grid: &mut Grid, r1: usize, c1: usize, r2: usize, c2: usize) {
    grid.merges.push((r1, c1, r2, c2));
    grid.cell(r2, c2);
}

fn grid_borders(grid: &mut Grid, r1: usize, c1: usize, r2: usize, c2: usize) {
    for r in r1..=r2 {
        for c in c1..=c2 {
            let cell = grid.cell(r, c);
            cell.top = Some(THIN.to_string());
            cell.left = Some(THIN.to_string());
            cell.right = Some(THIN.to_string());
            cell.bottom = Some(THIN.to_string());
        }
    }
}

fn edges(
    cell: &mut GCell,
    top: Option<&str>,
    left: Option<&str>,
    right: Option<&str>,
    bottom: Option<&str>,
) {
    if let Some(x) = top {
        cell.top = Some(x.to_string());
    }
    if let Some(x) = left {
        cell.left = Some(x.to_string());
    }
    if let Some(x) = right {
        cell.right = Some(x.to_string());
    }
    if let Some(x) = bottom {
        cell.bottom = Some(x.to_string());
    }
}

fn box_edges(
    cell: &mut GCell,
    top: Option<&str>,
    left: Option<&str>,
    right: Option<&str>,
    bottom: Option<&str>,
) {
    cell.top = top.map(str::to_string);
    cell.left = left.map(str::to_string);
    cell.right = right.map(str::to_string);
    cell.bottom = bottom.map(str::to_string);
}

fn clear_left(cell: &mut GCell) {
    cell.left = None;
}

fn titles(grid: &mut Grid, m: &Value, mut r: usize, ncol: usize) -> usize {
    if let Some(v) = obj(m, "title") {
        if !matches!(v, Value::Null) {
            merge(grid, r, 1, r, ncol.max(2));
            let c = lbl(grid, r, 1, Some(v));
            style(c, true, 13, C_TEXT, None, Some(A_LEFT));
            grid.heights.insert(r, 22.0);
            r += 1;
        }
    }
    if let Some(v) = obj(m, "unit") {
        if !matches!(v, Value::Null) {
            merge(grid, r, 1, r, ncol.max(2));
            let c = lbl(grid, r, 1, Some(v));
            style(c, false, 9, C_MUTE, None, Some(A_RIGHT));
            r += 1;
        }
    }
    r
}

fn outer(grid: &mut Grid, r1: usize, c1: usize, r2: usize, c2: usize) {
    for c in c1..=c2 {
        edges(grid.cell(r1, c), Some(HEAVY), None, None, None);
        edges(grid.cell(r2, c), None, None, None, Some(HEAVY));
    }
    for r in r1..=r2 {
        edges(grid.cell(r, c1), None, Some(HEAVY), None, None);
        edges(grid.cell(r, c2), None, None, Some(HEAVY), None);
    }
}

fn vband(grid: &mut Grid, col: usize, b0: usize, b1: usize) {
    box_edges(grid.cell(b0, col), Some(HEAVY), Some(HEAVY), None, None);
    for rr in (b0 + 1)..=b1 {
        box_edges(
            grid.cell(rr, col),
            None,
            Some(HEAVY),
            Some(HEAVY),
            if rr == b1 { Some(HEAVY) } else { None },
        );
    }
}

fn hband(grid: &mut Grid, row: usize, c0: usize, c1: usize) {
    box_edges(grid.cell(row, c0), Some(HEAVY), Some(HEAVY), None, None);
    for cc in (c0 + 1)..=c1 {
        box_edges(
            grid.cell(row, cc),
            Some(HEAVY),
            None,
            if cc == c1 { Some(HEAVY) } else { None },
            Some(HEAVY),
        );
    }
}

fn interp(a: &str, b: &str, t: f64) -> String {
    let aa = a.trim_start_matches('#');
    let bb = b.trim_start_matches('#');
    let mut out = String::new();
    for i in [0, 2, 4] {
        let av = u8::from_str_radix(&aa[i..i + 2], 16).unwrap() as f64;
        let bv = u8::from_str_radix(&bb[i..i + 2], 16).unwrap() as f64;
        out.push_str(&format!("{:02X}", (av + (bv - av) * t).round() as u8));
    }
    out
}

fn level_fill(k: usize, d: usize) -> String {
    if d <= 1 {
        HEAD_FILL.to_string()
    } else {
        interp(HEAD_FILL, SUB_FILL, k.min(d - 1) as f64 / (d - 1) as f64)
    }
}

fn themed(mut grid: Grid, theme_name: &str) -> Result<Grid, Box<dyn Error>> {
    if theme_name == "color" {
        return Ok(grid);
    }
    let t = theme(theme_name)?;
    let d = grid.depth.unwrap_or(2).max(2);
    let mut fmap: HashMap<String, Option<String>> = HashMap::new();
    fmap.insert(C_NAVY.to_string(), t.navy.map(str::to_string));
    for k in 0..d {
        let dst = match (t.head, t.sub) {
            (Some(h), Some(s)) => Some(interp(h, s, k.min(d - 1) as f64 / (d - 1) as f64)),
            _ => None,
        };
        fmap.insert(level_fill(k, d), dst);
    }
    let cmap: HashMap<&str, &str> =
        HashMap::from([(C_WHITE, t.white), (C_TEXT, t.text), (C_MUTE, t.mute)]);
    for cell in grid.g.values_mut() {
        if let Some(fill) = &cell.fill {
            if let Some(new_fill) = fmap.get(fill) {
                cell.fill = new_fill.clone();
            }
        }
        if let Some(new_color) = cmap.get(cell.color.as_str()) {
            cell.color = (*new_color).to_string();
        }
    }
    Ok(grid)
}

fn build_row(grid: &mut Grid, m: &Value) -> Result<(), Box<dyn Error>> {
    let nodes = nodes_of(m)?;
    let attrs = arr(m, "attributes")?;
    let a_len = attrs.len();
    let depth = max_depth(&nodes);
    grid.depth = Some(depth);
    let levels = obj(m, "level_labels").and_then(|x| x.as_array());
    let c0 = 1;
    let c_at0 = 1 + depth;
    let ncol = depth + a_len;
    let r = titles(grid, m, 1, ncol);
    let hdr = r;

    if let Some(levels) = levels {
        for d in 0..depth {
            let s = levels.get(d).map(py_str).unwrap_or_default();
            let c = lbl_str(grid, hdr, c0 + d, Some(&s));
            style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_LEFT));
        }
    } else {
        let corner = obj(m, "corner_label")
            .map(py_str)
            .unwrap_or_default();
        let c = lbl_str(grid, hdr, c0, Some(&corner));
        style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_LEFT));
        for d in 1..depth {
            style(
                grid.cell(hdr, c0 + d),
                false,
                10,
                C_TEXT,
                Some(C_NAVY),
                None,
            );
        }
    }
    for (j, attr) in attrs.iter().enumerate() {
        let s = py_str(attr);
        let c = lbl_str(grid, hdr, c_at0 + j, Some(&s));
        style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_CENTER));
    }

    let mut cur_r = hdr + 1;
    let mut flat: Vec<(usize, usize, usize, bool)> = Vec::new();
    fn emit(
        grid: &mut Grid,
        node: &Node,
        level: usize,
        depth: usize,
        a_len: usize,
        c0: usize,
        c_at0: usize,
        cur_r: &mut usize,
        flat: &mut Vec<(usize, usize, usize, bool)>,
    ) {
        let lcol = c0 + level;
        let is_grp = !node.children.is_empty();
        let is_hdr = level < depth - 1;
        let r0 = *cur_r;
        let sfill = level_fill(level, depth);
        let lcolor = if level == 0 { C_TEXT } else { C_MUTE };
        for cc in c0..lcol {
            let f = level_fill(cc - c0, depth);
            style(grid.cell(r0, cc), false, 10, C_TEXT, Some(&f), None);
        }
        let align = if level == 0 {
            A_LEFT
        } else {
            Al {
                h: "left",
                ind: level,
                wrap: false,
            }
        };
        let c = lbl_str(grid, r0, lcol, Some(&node.label));
        style(c, is_hdr, 10, lcolor, Some(&sfill), Some(align));
        for cc in (lcol + 1)..c_at0 {
            style(grid.cell(r0, cc), is_hdr, 10, C_TEXT, Some(&sfill), None);
        }
        let vals = if is_grp { &node.summary } else { &node.values };
        for j in 0..a_len {
            let cell = grid.cell(r0, c_at0 + j);
            if let Some(vs) = vals {
                if let Some(v) = vs.get(j) {
                    if !matches!(v, Value::Null) {
                        fmt_cell(cell, v);
                    }
                }
            }
            style(cell, is_hdr, 10, C_TEXT, Some(&sfill), None);
        }
        *cur_r += 1;
        if is_grp {
            for ch in &node.children {
                emit(grid, ch, level + 1, depth, a_len, c0, c_at0, cur_r, flat);
            }
        }
        flat.push((level, r0, *cur_r - 1, is_grp));
    }
    for n in &nodes {
        emit(
            grid,
            n,
            0,
            depth,
            a_len,
            c0,
            c_at0,
            &mut cur_r,
            &mut flat,
        );
    }
    let last = cur_r - 1;

    grid_borders(grid, hdr, c0, last, ncol);
    outer(grid, hdr, c0, last, ncol);
    if levels.is_none() {
        grid.cell(hdr, c0).right = None;
        grid.cell(hdr, c0 + 1).left = None;
    }
    for (level, r0, _r1, is_grp) in &flat {
        if *level < depth - 1 {
            let lcol = c0 + *level;
            for c in lcol..=ncol {
                edges(grid.cell(*r0, c), Some(HEAVY), None, None, None);
            }
            let lo = if *is_grp { lcol + 1 } else { lcol };
            for c in lo..=ncol {
                edges(grid.cell(*r0, c), None, None, None, Some(HEAVY));
            }
        }
    }
    for (level, r0, r1, is_grp) in &flat {
        let lcol = c0 + *level;
        if *level < depth - 1 {
            if *is_grp {
                vband(grid, lcol, *r0, *r1);
            } else {
                edges(grid.cell(*r0, lcol), None, Some(HEAVY), None, None);
            }
            clear_left(grid.cell(*r0, lcol + 1));
        } else {
            edges(grid.cell(*r0, lcol), None, Some(HEAVY), None, None);
        }
        if *level >= 1 {
            grid.outline_rows.insert(*r0, *level);
        }
    }
    grid.summary_below = Some(false);
    grid.widths.insert(c0, 12.0);
    for d in 1..depth {
        grid.widths
            .insert(c0 + d, if d == depth - 1 { 14.0 } else { 12.0 });
    }
    for j in 0..a_len {
        grid.widths.insert(c_at0 + j, 12.0);
    }
    grid.freeze = Some((hdr + 1, c_at0));
    Ok(())
}

fn build_column(grid: &mut Grid, m: &Value) -> Result<(), Box<dyn Error>> {
    if obj(m, "nodes").is_some() {
        build_column_ndepth(grid, m)
    } else {
        build_column_2level(grid, m)
    }
}

fn build_column_ndepth(grid: &mut Grid, m: &Value) -> Result<(), Box<dyn Error>> {
    let nodes = nodes_of(m)?;
    let attrs = arr(m, "attributes")?;
    let a_len = attrs.len();
    let depth = max_depth(&nodes);
    grid.depth = Some(depth);
    let levels = obj(m, "level_labels").and_then(|x| x.as_array());
    let c_albl = 1;
    let ncol = 1 + tree_size(&nodes);
    let r = titles(grid, m, 1, ncol);
    let r0 = r;
    let r_data0 = r0 + depth;
    let last_r = r_data0 + a_len - 1;

    if let Some(levels) = levels {
        for k in 0..depth {
            let s = levels.get(k).map(py_str).unwrap_or_default();
            let c = lbl_str(grid, r0 + k, c_albl, Some(&s));
            style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_LEFT_IND));
        }
    } else {
        merge(grid, r0, c_albl, r0 + depth - 1, c_albl);
        let corner = obj(m, "corner_label")
            .map(py_str)
            .unwrap_or_default();
        let c = lbl_str(grid, r0, c_albl, Some(&corner));
        style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_CENTER));
    }
    for (i, attr) in attrs.iter().enumerate() {
        let s = py_str(attr);
        let c = lbl_str(grid, r_data0 + i, c_albl, Some(&s));
        style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_LEFT_IND));
    }

    let mut cur_c = 2;
    let mut flat: Vec<(usize, usize, usize, bool)> = Vec::new();
    fn emit(
        grid: &mut Grid,
        node: &Node,
        level: usize,
        depth: usize,
        a_len: usize,
        r0: usize,
        r_data0: usize,
        cur_c: &mut usize,
        flat: &mut Vec<(usize, usize, usize, bool)>,
    ) {
        let is_grp = !node.children.is_empty();
        let is_hdr = level < depth - 1;
        let c0 = *cur_c;
        let sfill = level_fill(level, depth);
        let lcolor = if level == 0 { C_TEXT } else { C_MUTE };
        for rr in r0..(r0 + level) {
            let f = level_fill(rr - r0, depth);
            style(grid.cell(rr, c0), false, 10, C_TEXT, Some(&f), None);
        }
        let c = lbl_str(grid, r0 + level, c0, Some(&node.label));
        style(
            c,
            is_hdr,
            10,
            lcolor,
            Some(&sfill),
            Some(if level == 0 { A_LEFT } else { A_CENTER }),
        );
        for rr in (r0 + level + 1)..(r0 + depth) {
            style(grid.cell(rr, c0), is_hdr, 10, C_TEXT, Some(&sfill), None);
        }
        let vals = if is_grp { &node.summary } else { &node.values };
        for i in 0..a_len {
            let cell = grid.cell(r_data0 + i, c0);
            if let Some(vs) = vals {
                if let Some(v) = vs.get(i) {
                    if !matches!(v, Value::Null) {
                        fmt_cell(cell, v);
                    }
                }
            }
            style(cell, is_hdr, 10, C_TEXT, Some(&sfill), None);
        }
        *cur_c += 1;
        if is_grp {
            for ch in &node.children {
                emit(grid, ch, level + 1, depth, a_len, r0, r_data0, cur_c, flat);
            }
        }
        let c1 = *cur_c - 1;
        if c1 > c0 {
            grid.soft_merges.push((r0 + level, c0, r0 + level, c1));
            for cc in (c0 + 1)..=c1 {
                style(
                    grid.cell(r0 + level, cc),
                    is_hdr,
                    10,
                    C_TEXT,
                    Some(&sfill),
                    None,
                );
            }
        }
        flat.push((level, c0, c1, is_grp));
    }
    for n in &nodes {
        emit(
            grid,
            n,
            0,
            depth,
            a_len,
            r0,
            r_data0,
            &mut cur_c,
            &mut flat,
        );
    }
    let last_c = cur_c - 1;

    grid_borders(grid, r0, c_albl, last_r, last_c);
    outer(grid, r0, c_albl, last_r, last_c);
    for cc in c_albl..=last_c {
        edges(
            grid.cell(r0 + depth - 1, cc),
            None,
            None,
            None,
            Some(HEAVY),
        );
    }
    for (level, c0, _c1, is_grp) in &flat {
        if *level < depth - 1 {
            let lrow = r0 + *level;
            for rr in lrow..=last_r {
                edges(grid.cell(rr, *c0), None, Some(HEAVY), None, None);
            }
            let lo = if *is_grp { lrow + 1 } else { lrow };
            for rr in lo..=last_r {
                edges(grid.cell(rr, *c0), None, None, Some(HEAVY), None);
            }
        }
    }
    for (level, c0, c1, is_grp) in &flat {
        let lrow = r0 + *level;
        if *level < depth - 1 {
            if *is_grp {
                hband(grid, lrow, *c0, *c1);
            } else {
                edges(grid.cell(lrow, *c0), Some(HEAVY), None, None, None);
            }
            grid.cell(lrow + 1, *c0).top = None;
        } else {
            edges(grid.cell(lrow, *c0), Some(HEAVY), None, None, None);
        }
        if *level >= 1 {
            grid.outline_cols.insert(*c0, *level);
        }
    }
    grid.summary_right = Some(false);
    grid.widths.insert(c_albl, 13.0);
    for cc in 2..=last_c {
        grid.widths.insert(cc, 9.0);
    }
    grid.freeze = Some((r_data0, 2));
    Ok(())
}

fn build_column_2level(grid: &mut Grid, m: &Value) -> Result<(), Box<dyn Error>> {
    let attrs = arr(m, "attributes")?;
    let a_len = attrs.len();
    let c_albl = 1;
    let mut total = 1;
    for it in arr(m, "items")? {
        let summary_count = if obj(it, "summary").is_some() { 1 } else { 0 };
        let details_count = obj(it, "details")
            .and_then(|x| x.as_array())
            .map(|x| x.len())
            .unwrap_or(0);
        total += summary_count + details_count;
    }
    let r = titles(grid, m, 1, total);
    let r_item = r;
    let r_sub = r + 1;
    let r_data0 = r + 2;

    merge(grid, r_item, c_albl, r_sub, c_albl);
    let corner = obj(m, "corner_label")
        .map(py_str)
        .unwrap_or_default();
    let c = lbl_str(grid, r_item, c_albl, Some(&corner));
    style(c, true, 10, C_WHITE, Some(C_NAVY), Some(A_CENTER));

    let mut c = 2;
    let mut blocks: Vec<(usize, usize, usize)> = Vec::new();
    for it in arr(m, "items")? {
        let c0 = c;
        let summ = opt_array(it, "summary");
        if let Some(summ_vals) = &summ {
            style(grid.cell(r_sub, c), false, 10, C_TEXT, Some(HEAD_FILL), None);
            for (i, v) in summ_vals.iter().enumerate() {
                let cell = grid.cell(r_data0 + i, c);
                fmt_cell(cell, v);
                style(cell, true, 10, C_TEXT, Some(HEAD_FILL), None);
            }
            c += 1;
        }
        let sub_first = c;
        if let Some(details) = obj(it, "details").and_then(|x| x.as_array()) {
            for d in details {
                let dlabel = label(d)?;
                let cell = lbl_str(grid, r_sub, c, Some(&dlabel));
                style(cell, false, 10, C_MUTE, Some(SUB_FILL), Some(A_CENTER));
                for (i, v) in opt_array(d, "values").unwrap_or_default().iter().enumerate() {
                    let cell = grid.cell(r_data0 + i, c);
                    fmt_cell(cell, v);
                    style(cell, false, 10, C_TEXT, Some(SUB_FILL), None);
                }
                c += 1;
            }
        }
        let c1 = c - 1;
        let it_label = label(it)?;
        for cc in c0..=c1 {
            let text = if cc == c0 { Some(it_label.as_str()) } else { None };
            let cell = lbl_str(grid, r_item, cc, text);
            style(
                cell,
                true,
                10,
                C_TEXT,
                Some(HEAD_FILL),
                if cc == c0 { Some(A_LEFT) } else { None },
            );
        }
        if c1 > c0 {
            grid.soft_merges.push((r_item, c0, r_item, c1));
        }
        blocks.push((c0, if summ.is_some() { sub_first } else { c0 }, c1));
    }

    let last_c = c - 1;
    for (i, attr) in attrs.iter().enumerate() {
        let s = py_str(attr);
        let cell = lbl_str(grid, r_data0 + i, c_albl, Some(&s));
        style(cell, true, 10, C_WHITE, Some(C_NAVY), Some(A_LEFT_IND));
    }
    let last_r = r_data0 + a_len - 1;

    grid_borders(grid, r_item, c_albl, last_r, last_c);
    outer(grid, r_item, c_albl, last_r, last_c);
    for cc in c_albl..=last_c {
        edges(grid.cell(r_sub, cc), None, None, None, Some(HEAVY));
    }
    for (c0, sub_first, c1) in &blocks {
        for cc in *c0..=*c1 {
            box_edges(
                grid.cell(r_item, cc),
                Some(HEAVY),
                if cc == *c0 { Some(HEAVY) } else { None },
                if cc == *c1 { Some(HEAVY) } else { None },
                None,
            );
        }
        for rr in r_item..=last_r {
            edges(grid.cell(rr, *c0), None, Some(HEAVY), None, None);
            edges(grid.cell(rr, *c1), None, None, Some(HEAVY), None);
        }
        if *sub_first > *c0 {
            for rr in r_sub..=last_r {
                edges(grid.cell(rr, *sub_first), None, Some(HEAVY), None, None);
            }
            for cc in *sub_first..=*c1 {
                edges(grid.cell(r_sub, cc), Some(HEAVY), None, None, None);
            }
            grid.cell(r_sub, *c0).top = None;
        }
    }
    for cc in 2..=last_c {
        if !blocks.iter().any(|(c0, _, _)| cc == *c0) {
            grid.outline_cols.insert(cc, 1);
        }
    }
    grid.summary_right = Some(false);
    grid.widths.insert(c_albl, 13.0);
    for cc in 2..=last_c {
        grid.widths.insert(cc, 9.0);
    }
    grid.freeze = Some((r_data0, 2));
    Ok(())
}

fn map_get<'a>(v: &'a Value, keys: &[&str]) -> Option<&'a Value> {
    let mut cur = v;
    for key in keys {
        cur = cur.as_object()?.get(*key)?;
    }
    Some(cur)
}

fn num_value(v: &Value) -> f64 {
    match v {
        Value::Number(n) => n.as_f64().unwrap_or(0.0),
        Value::String(s) => s.parse::<f64>().unwrap_or(0.0),
        Value::Bool(b) => {
            if *b {
                1.0
            } else {
                0.0
            }
        }
        _ => 0.0,
    }
}

fn sum_values(vals: impl Iterator<Item = Value>) -> Value {
    let sum: f64 = vals.map(|v| num_value(&v)).sum();
    if (sum.fract()).abs() < 1e-9 {
        Value::Number(Number::from(sum as i64))
    } else {
        value_number(sum)
    }
}

fn build_both(grid: &mut Grid, m: &Value) -> Result<(), Box<dyn Error>> {
    let rg = arr(m, "row_groups")?;
    let cg = arr(m, "col_groups")?;
    let data = req(m, "data")?;
    let mut cg_det: HashMap<String, Vec<String>> = HashMap::new();
    for c in cg {
        let lab = label(c)?;
        let details = arr(c, "details")?.iter().map(py_str).collect::<Vec<_>>();
        cg_det.insert(lab, details);
    }
    let c_grp = 1;
    let c_det = 2;
    let ncol = 2 + cg.iter().map(|c| 1 + arr(c, "details").map(|x| x.len()).unwrap_or(0)).sum::<usize>();
    let r = titles(grid, m, 1, ncol);
    let r_citem = r;
    let r_csub = r + 1;
    let r0 = r + 2;

    let mut ccols: Vec<(usize, String, Option<String>)> = Vec::new();
    let mut col_blocks: Vec<(usize, usize, usize)> = Vec::new();
    let mut c = 3;
    for cgrp in cg {
        let c0 = c;
        let c_label = label(cgrp)?;
        ccols.push((c, c_label.clone(), None));
        c += 1;
        for cd in arr(cgrp, "details")? {
            ccols.push((c, c_label.clone(), Some(py_str(cd))));
            c += 1;
        }
        col_blocks.push((c0, c0 + 1, c - 1));
    }
    let last_c = c - 1;

    merge(grid, r_citem, c_grp, r_csub, c_det);
    let corner = obj(m, "corner_label")
        .map(py_str)
        .unwrap_or_default();
    let cell = lbl_str(grid, r_citem, c_grp, Some(&corner));
    style(cell, true, 10, C_WHITE, Some(C_NAVY), Some(A_CENTER));

    for (cgrp, (c0, _, c1)) in cg.iter().zip(col_blocks.iter()) {
        let lab = label(cgrp)?;
        for cc in *c0..=*c1 {
            let text = if cc == *c0 { Some(lab.as_str()) } else { None };
            let cell = lbl_str(grid, r_citem, cc, text);
            style(cell, true, 10, C_TEXT, Some(HEAD_FILL), Some(A_LEFT));
        }
    }
    for (cidx, _ci, cd) in &ccols {
        if let Some(cd) = cd {
            let cell = lbl_str(grid, r_csub, *cidx, Some(cd));
            style(cell, false, 10, C_MUTE, Some(SUB_FILL), Some(A_CENTER));
        } else {
            style(
                grid.cell(r_csub, *cidx),
                false,
                10,
                C_TEXT,
                Some(HEAD_FILL),
                None,
            );
        }
    }

    let v_fine = |ri: &str, rd: &str, ci: &str, cd: &str| -> Value {
        map_get(data, &[ri, rd, ci, cd]).cloned().unwrap_or(Value::Null)
    };
    let v_rd = |ri: &str, rd: &str, ci: &str, cd: &Option<String>| -> Value {
        if let Some(cd) = cd {
            v_fine(ri, rd, ci, cd)
        } else {
            let iter = cg_det
                .get(ci)
                .into_iter()
                .flat_map(|xs| xs.iter())
                .map(|x| v_fine(ri, rd, ci, x));
            sum_values(iter)
        }
    };
    let v_head = |rgrp: &Value, ci: &str, cd: &Option<String>| -> Result<Value, Box<dyn Error>> {
        let ri = label(rgrp)?;
        let iter = arr(rgrp, "details")?
            .iter()
            .map(|rd| v_rd(&ri, &py_str(rd), ci, cd));
        Ok(sum_values(iter))
    };

    let mut r = r0;
    let mut row_blocks: Vec<(usize, usize, Vec<usize>)> = Vec::new();
    for rgrp in rg {
        let b0 = r;
        let rlabel = label(rgrp)?;
        let cell = lbl_str(grid, r, c_grp, Some(&rlabel));
        style(cell, true, 10, C_TEXT, Some(HEAD_FILL), Some(A_LEFT));
        style(
            grid.cell(r, c_det),
            true,
            10,
            C_TEXT,
            Some(HEAD_FILL),
            None,
        );
        for (cidx, ci, cd) in &ccols {
            let v = v_head(rgrp, ci, cd)?;
            let cell = grid.cell(r, *cidx);
            fmt_cell(cell, &v);
            style(cell, true, 10, C_TEXT, Some(HEAD_FILL), None);
        }
        r += 1;
        let mut det_rows = Vec::new();
        for rd in arr(rgrp, "details")? {
            let rd_s = py_str(rd);
            style(
                grid.cell(r, c_grp),
                false,
                10,
                C_TEXT,
                Some(HEAD_FILL),
                None,
            );
            let cell = lbl_str(grid, r, c_det, Some(&rd_s));
            style(cell, false, 10, C_MUTE, Some(SUB_FILL), Some(A_LEFT_IND));
            for (cidx, ci, cd) in &ccols {
                let v = v_rd(&rlabel, &rd_s, ci, cd);
                let cell = grid.cell(r, *cidx);
                fmt_cell(cell, &v);
                style(cell, false, 10, C_TEXT, Some(SUB_FILL), None);
            }
            det_rows.push(r);
            r += 1;
        }
        row_blocks.push((b0, r - 1, det_rows));
    }
    let last_r = r - 1;

    grid_borders(grid, r_citem, c_grp, last_r, last_c);
    outer(grid, r_citem, c_grp, last_r, last_c);
    for cc in c_grp..=last_c {
        edges(grid.cell(r_csub, cc), None, None, None, Some(HEAVY));
    }
    for (c0, sub_first, c1) in &col_blocks {
        for cc in *c0..=*c1 {
            box_edges(
                grid.cell(r_citem, cc),
                Some(HEAVY),
                if cc == *c0 { Some(HEAVY) } else { None },
                if cc == *c1 { Some(HEAVY) } else { None },
                None,
            );
        }
        for rr in r_citem..=last_r {
            edges(grid.cell(rr, *c0), None, Some(HEAVY), None, None);
        }
        for rr in r_csub..=last_r {
            edges(grid.cell(rr, *sub_first), None, Some(HEAVY), None, None);
        }
        for cc in *sub_first..=*c1 {
            edges(grid.cell(r_csub, cc), Some(HEAVY), None, None, None);
        }
        grid.cell(r_csub, *c0).top = None;
    }
    for (b0, b1, det_rows) in &row_blocks {
        for cc in c_grp..=last_c {
            edges(grid.cell(*b0, cc), Some(HEAVY), None, None, None);
        }
        if !det_rows.is_empty() {
            for cc in c_det..=last_c {
                edges(grid.cell(*b0, cc), None, None, None, Some(HEAVY));
            }
        }
        vband(grid, c_grp, *b0, *b1);
        clear_left(grid.cell(*b0, c_det));
        for rr in det_rows {
            edges(grid.cell(*rr, c_det), None, Some(HEAVY), None, None);
            grid.outline_rows.insert(*rr, 1);
        }
    }
    for (cidx, _ci, cd) in &ccols {
        if cd.is_some() {
            grid.outline_cols.insert(*cidx, 1);
        }
    }
    grid.summary_below = Some(false);
    grid.summary_right = Some(false);
    grid.widths.insert(c_grp, 12.0);
    grid.widths.insert(c_det, 13.0);
    for cc in 3..=last_c {
        grid.widths.insert(cc, 8.5);
    }
    grid.freeze = Some((r0, 3));
    Ok(())
}

fn layout(model: &Value, orientation: &str) -> Result<Grid, Box<dyn Error>> {
    let mut grid = Grid::new();
    match orientation {
        "row" => build_row(&mut grid, model)?,
        "column" => build_column(&mut grid, model)?,
        "both" => build_both(&mut grid, model)?,
        _ => return Err("orientation must be 'row', 'column' or 'both'".into()),
    }
    Ok(grid)
}

fn dump_ir(grid: &Grid) -> Value {
    let mut cells = Map::new();
    for ((r, c), cell) in &grid.g {
        cells.insert(format!("{r},{c}"), serde_json::to_value(cell).unwrap());
    }
    let map_usize = |m: &BTreeMap<usize, usize>| -> Map<String, Value> {
        m.iter()
            .map(|(k, v)| (k.to_string(), json!(v)))
            .collect::<Map<_, _>>()
    };
    let map_f64 = |m: &BTreeMap<usize, f64>| -> Map<String, Value> {
        m.iter()
            .map(|(k, v)| (k.to_string(), json!(v)))
            .collect::<Map<_, _>>()
    };
    json!({
        "nr": grid.nr,
        "nc": grid.nc,
        "depth": grid.depth,
        "freeze": grid.freeze.map(|(r,c)| vec![r,c]),
        "merges": grid.merges.iter().map(|(a,b,c,d)| vec![*a,*b,*c,*d]).collect::<Vec<_>>(),
        "soft_merges": grid.soft_merges.iter().map(|(a,b,c,d)| vec![*a,*b,*c,*d]).collect::<Vec<_>>(),
        "widths": map_f64(&grid.widths),
        "heights": map_f64(&grid.heights),
        "outline_rows": map_usize(&grid.outline_rows),
        "outline_cols": map_usize(&grid.outline_cols),
        "cells": cells,
    })
}

type RecMap = BTreeMap<(usize, usize), (Option<String>, Option<String>, Option<String>, Option<String>)>;

fn weight(v: &Option<String>) -> usize {
    match v.as_deref() {
        Some(THIN) => 1,
        Some(HEAVY) => 2,
        _ => 0,
    }
}

fn heaviest(a: &Option<String>, b: &Option<String>) -> Option<String> {
    if weight(a) >= weight(b) {
        a.clone()
    } else {
        b.clone()
    }
}

fn reconcile(grid: &Grid) -> RecMap {
    let mut rec = BTreeMap::new();
    for (&(r, c), cell) in &grid.g {
        let mut top = cell.top.clone();
        let mut left = cell.left.clone();
        let mut right = cell.right.clone();
        let mut bottom = cell.bottom.clone();
        if let Some(up) = r.checked_sub(1).and_then(|rr| grid.g.get(&(rr, c))) {
            top = heaviest(&top, &up.bottom);
        }
        if let Some(lf) = c.checked_sub(1).and_then(|cc| grid.g.get(&(r, cc))) {
            left = heaviest(&left, &lf.right);
        }
        if let Some(rt) = grid.g.get(&(r, c + 1)) {
            right = heaviest(&right, &rt.left);
        }
        if let Some(dn) = grid.g.get(&(r + 1, c)) {
            bottom = heaviest(&bottom, &dn.top);
        }
        rec.insert((r, c), (top, left, right, bottom));
    }
    rec
}

fn region_border(
    rec: &RecMap,
    r1: usize,
    c1: usize,
    r2: usize,
    c2: usize,
) -> (Option<String>, Option<String>, Option<String>, Option<String>) {
    let none = (None, None, None, None);
    let tl = rec.get(&(r1, c1)).unwrap_or(&none);
    let tr = rec.get(&(r1, c2)).unwrap_or(&none);
    let bl = rec.get(&(r2, c1)).unwrap_or(&none);
    (tl.0.clone(), tl.1.clone(), tr.2.clone(), bl.3.clone())
}

fn is_dark(hexcolor: &Option<String>) -> bool {
    let Some(h) = hexcolor else {
        return false;
    };
    if h.len() < 6 {
        return false;
    }
    let r = u8::from_str_radix(&h[0..2], 16).unwrap_or(255) as f64;
    let g = u8::from_str_radix(&h[2..4], 16).unwrap_or(255) as f64;
    let b = u8::from_str_radix(&h[4..6], 16).unwrap_or(255) as f64;
    0.299 * r + 0.587 * g + 0.114 * b < 110.0
}

fn load_data(path: &str) -> Result<Value, Box<dyn Error>> {
    let text = fs::read_to_string(path)?;
    if path.to_lowercase().ends_with(".yaml") || path.to_lowercase().ends_with(".yml") {
        Ok(serde_yaml::from_str(&text)?)
    } else {
        Ok(serde_json::from_str(&text)?)
    }
}

fn cell_display(cell: &GCell) -> String {
    if let Some(v) = &cell.value {
        match v {
            Value::String(s) => s.clone(),
            Value::Number(_) => cell.text.clone(),
            Value::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            Value::Null => String::new(),
            _ => cell.text.clone(),
        }
    } else {
        cell.text.clone()
    }
}

fn xl_rgb(hex: &str) -> rust_xlsxwriter::Color {
    rust_xlsxwriter::Color::RGB(u32::from_str_radix(hex.trim_start_matches('#'), 16).unwrap_or(0))
}

fn xl_cell_format(cell: &GCell) -> rust_xlsxwriter::Format {
    use rust_xlsxwriter::{Format, FormatAlign, FormatPattern};
    let mut f = Format::new()
        .set_font_name(FONT)
        .set_font_size(cell.size as f64)
        .set_font_color(xl_rgb(&cell.color));
    if cell.bold {
        f = f.set_bold();
    }
    if let Some(fill) = &cell.fill {
        f = f.set_background_color(xl_rgb(fill)).set_pattern(FormatPattern::Solid);
    }
    if let Some(nf) = &cell.numfmt {
        f = f.set_num_format(nf);
    }
    if let Some(h) = &cell.halign {
        f = match h.as_str() {
            "left" => f.set_align(FormatAlign::Left),
            "center" => f.set_align(FormatAlign::Center),
            "right" => f.set_align(FormatAlign::Right),
            _ => f,
        };
        f = f.set_align(FormatAlign::VerticalCenter);
        if cell.indent > 0 {
            f = f.set_indent(cell.indent as u8);
        }
        if cell.wrap {
            f = f.set_text_wrap();
        }
    }
    f
}

fn xl_with_borders(
    mut f: rust_xlsxwriter::Format,
    t: &Theme,
    top: &Option<String>,
    left: &Option<String>,
    right: &Option<String>,
    bottom: &Option<String>,
) -> rust_xlsxwriter::Format {
    use rust_xlsxwriter::FormatBorder;
    let kind = |s: &Option<String>| match s.as_deref() {
        Some("thin") => Some(FormatBorder::Thin),
        Some("medium") => Some(FormatBorder::Medium),
        _ => None,
    };
    let col = |s: &Option<String>| {
        if s.as_deref() == Some("medium") {
            xl_rgb(t.heavy)
        } else {
            xl_rgb(t.grid)
        }
    };
    if let Some(b) = kind(top) {
        f = f.set_border_top(b).set_border_top_color(col(top));
    }
    if let Some(b) = kind(left) {
        f = f.set_border_left(b).set_border_left_color(col(left));
    }
    if let Some(b) = kind(right) {
        f = f.set_border_right(b).set_border_right_color(col(right));
    }
    if let Some(b) = kind(bottom) {
        f = f.set_border_bottom(b).set_border_bottom_color(col(bottom));
    }
    f
}

fn render_xlsx(grid: &Grid, path: &str, t: &Theme) -> Result<(), Box<dyn Error>> {
    use rust_xlsxwriter::Workbook;
    let mut workbook = Workbook::new();
    let ws = workbook.add_worksheet();
    ws.set_screen_gridlines(false);

    for ((r, c), cell) in &grid.g {
        let f = xl_with_borders(xl_cell_format(cell), t, &cell.top, &cell.left, &cell.right, &cell.bottom);
        let (row, col) = ((*r - 1) as u32, (*c - 1) as u16);
        match &cell.value {
            Some(Value::Number(n)) => {
                ws.write_number_with_format(row, col, n.as_f64().unwrap_or(0.0), &f)?;
            }
            Some(Value::String(s)) => {
                ws.write_string_with_format(row, col, s, &f)?;
            }
            _ => {
                if !cell.text.is_empty() {
                    ws.write_string_with_format(row, col, &cell.text, &f)?;
                } else {
                    ws.write_blank(row, col, &f)?;
                }
            }
        }
    }

    // 병합: 영역 외곽 테두리(top-left.top/left, top-right.right, bottom-left.bottom) + top-left 서식
    for (r1, c1, r2, c2) in &grid.merges {
        let tl = grid.g.get(&(*r1, *c1)).cloned().unwrap_or_default();
        let right = grid.g.get(&(*r1, *c2)).map(|x| x.right.clone()).unwrap_or_else(|| tl.right.clone());
        let bottom = grid.g.get(&(*r2, *c1)).map(|x| x.bottom.clone()).unwrap_or_else(|| tl.bottom.clone());
        let f = xl_with_borders(xl_cell_format(&tl), t, &tl.top, &tl.left, &right, &bottom);
        let text = cell_display(&tl);
        ws.merge_range((*r1 - 1) as u32, (*c1 - 1) as u16, (*r2 - 1) as u32, (*c2 - 1) as u16, &text, &f)?;
    }
    for (c, w) in &grid.widths {
        ws.set_column_width((*c - 1) as u16, *w)?;
    }
    for (r, h) in &grid.heights {
        ws.set_row_height((*r - 1) as u32, *h)?;
    }
    if let Some((r, c)) = grid.freeze {
        ws.set_freeze_panes((r - 1) as u32, (c - 1) as u16)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn dx_borders(
    t: &Theme,
    top: &Option<String>,
    left: &Option<String>,
    right: &Option<String>,
    bottom: &Option<String>,
) -> docx_rs::TableCellBorders {
    use docx_rs::{BorderType, TableCellBorder, TableCellBorderPosition, TableCellBorders};
    let spec = |w: &Option<String>| -> Option<(usize, String)> {
        match w.as_deref() {
            Some(THIN) => Some((4, t.grid.to_string())),
            Some(HEAVY) => Some((12, t.heavy.to_string())),
            _ => None,
        }
    };
    let mut b = TableCellBorders::new();
    for (pos, w) in [
        (TableCellBorderPosition::Top, top),
        (TableCellBorderPosition::Left, left),
        (TableCellBorderPosition::Right, right),
        (TableCellBorderPosition::Bottom, bottom),
    ] {
        if let Some((sz, col)) = spec(w) {
            b = b.set(TableCellBorder::new(pos).border_type(BorderType::Single).size(sz).color(col));
        } else {
            b = b.set(TableCellBorder::new(pos).border_type(BorderType::Nil));
        }
    }
    b
}

fn render_docx(grid: &Grid, path: &str, t: &Theme) -> Result<(), Box<dyn Error>> {
    use docx_rs::{
        AlignmentType, Docx, Paragraph, Run, RunFonts, Shading, ShdType, Table, TableCell,
        TableLayoutType, TableRow, VAlignType, WidthType,
    };
    use std::fs::File;

    let rec = reconcile(grid);
    let merges: Vec<(usize, usize, usize, usize)> =
        grid.merges.iter().chain(grid.soft_merges.iter()).cloned().collect();

    // Word vMerge 좌측선 버그 회피: 세로 병합을 '행별 가로 병합'으로 분해(텍스트=첫행, 채움=전행 반복)
    let mut covered: BTreeSet<(usize, usize)> = BTreeSet::new();
    let mut lead_region: BTreeMap<(usize, usize), (usize, usize, usize, usize)> = BTreeMap::new();
    let mut fill_over: BTreeMap<(usize, usize), Option<String>> = BTreeMap::new();
    for &(r1, c1, r2, c2) in &merges {
        let anchor_fill = grid.g.get(&(r1, c1)).and_then(|c| c.fill.clone());
        for r in r1..=r2 {
            lead_region.insert((r, c1), (r1, c1, r2, c2));
            if r != r1 {
                fill_over.insert((r, c1), anchor_fill.clone());
            }
            for c in (c1 + 1)..=c2 {
                covered.insert((r, c));
            }
        }
    }

    let default = GCell::default();
    let mut rows = Vec::new();
    for r in 1..=grid.nr {
        let mut cells = Vec::new();
        for c in 1..=grid.nc {
            if covered.contains(&(r, c)) {
                continue;
            }
            let cell = grid.g.get(&(r, c)).unwrap_or(&default);
            let region = lead_region.get(&(r, c)).copied();
            let fill = fill_over.get(&(r, c)).cloned().unwrap_or_else(|| cell.fill.clone());

            let (top, left, right, bottom) = if let Some((br1, bc1, br2, bc2)) = region {
                let (t0, l0, r0, b0) = region_border(&rec, br1, bc1, br2, bc2);
                (if r == br1 { t0 } else { None }, l0, r0, if r == br2 { b0 } else { None })
            } else {
                let (mut t0, mut l0, mut r0, mut b0) = region_border(&rec, r, c, r, c);
                if is_dark(&cell.fill) {
                    let gf = |rr: usize, cc: usize| grid.g.get(&(rr, cc)).and_then(|x| x.fill.clone());
                    if l0.as_deref() == Some(HEAVY) && c > 1 && !is_dark(&gf(r, c - 1)) { l0 = None; }
                    if r0.as_deref() == Some(HEAVY) && c < grid.nc && !is_dark(&gf(r, c + 1)) { r0 = None; }
                    if t0.as_deref() == Some(HEAVY) && r > 1 && !is_dark(&gf(r - 1, c)) { t0 = None; }
                    if b0.as_deref() == Some(HEAVY) && r < grid.nr && !is_dark(&gf(r + 1, c)) { b0 = None; }
                }
                (t0, l0, r0, b0)
            };

            let show_text = region.map_or(true, |reg| r == reg.0);
            let mut para = Paragraph::new();
            if show_text && !cell.text.is_empty() {
                let mut run = Run::new()
                    .add_text(&cell.text)
                    .fonts(RunFonts::new().east_asia(FONT).ascii(FONT).hi_ansi(FONT))
                    .size((cell.size * 2) as usize)
                    .color(cell.color.clone());
                if cell.bold {
                    run = run.bold();
                }
                para = para.add_run(run);
            }
            if let Some(h) = &cell.halign {
                para = para.align(match h.as_str() {
                    "center" => AlignmentType::Center,
                    "right" => AlignmentType::Right,
                    _ => AlignmentType::Left,
                });
            }

            let mut tc = TableCell::new()
                .add_paragraph(para)
                .vertical_align(VAlignType::Center)
                .set_borders(dx_borders(t, &top, &left, &right, &bottom));
            if let Some(f) = &fill {
                tc = tc.shading(Shading::new().shd_type(ShdType::Clear).color("auto").fill(f.clone()));
            }
            if let Some((_, bc1, _, bc2)) = region {
                if bc2 > bc1 {
                    tc = tc.grid_span(bc2 - bc1 + 1);
                }
            }
            cells.push(tc);
        }
        rows.push(TableRow::new(cells));
    }

    // 열 너비(dxa) — 페이지 가용폭(약 6.5in)에 맞춰 스케일
    let usable_in = 6.5_f64;
    let raw: Vec<f64> = (1..=grid.nc).map(|c| grid.widths.get(&c).copied().unwrap_or(8.43) * 0.092).collect();
    let sum: f64 = raw.iter().sum();
    let scale = if sum > 0.0 { (usable_in / sum).min(1.0) } else { 1.0 };
    let grid_dxa: Vec<usize> = raw.iter().map(|w| (w * scale * 1440.0).round() as usize).collect();
    let total: usize = grid_dxa.iter().sum();

    let table = Table::new(rows).layout(TableLayoutType::Fixed).set_grid(grid_dxa).width(total, WidthType::Dxa);
    let doc = Docx::new().add_table(table);
    let file = File::create(path)?;
    doc.build().pack(file)?;
    Ok(())
}

fn generate(model: &Value, out: &str, orient: &str, theme_name: &str) -> Result<(), Box<dyn Error>> {
    let ext = Path::new(out)
        .extension()
        .and_then(|x| x.to_str())
        .unwrap_or("")
        .to_lowercase();
    let t = theme(theme_name)?;
    let grid = themed(layout(model, orient)?, theme_name)?;
    match ext.as_str() {
        "xlsx" => render_xlsx(&grid, out, &t),
        "docx" => render_docx(&grid, out, &t),
        "pptx" => Err("pptx is not supported in the Rust build; use the Python engine.".into()),
        _ => Err("output extension must be .xlsx or .docx".into()),
    }
}

fn usage() -> &'static str {
    "drilldown-table <data.{yaml,json}> <out.{xlsx,docx}> [row|column|both] [color|grey|mono]"
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut args: Vec<String> = env::args().skip(1).collect();
    if args.iter().any(|x| x == "-h" || x == "--help") || args.len() < 2 {
        eprintln!("{}", usage());
        if args.len() < 2 {
            std::process::exit(2);
        }
        return Ok(());
    }
    let dump_ir_flag = args.iter().position(|x| x == "--dump-ir");
    if let Some(i) = dump_ir_flag {
        args.remove(i);
    }
    let data_path = args.get(0).cloned().ok_or("missing data path")?;
    let out_path = args.get(1).cloned().ok_or("missing output path")?;
    let orient = args.get(2).map(String::as_str).unwrap_or("row");
    let theme_name = args.get(3).map(String::as_str).unwrap_or("color");
    let model = load_data(&data_path)?;
    if dump_ir_flag.is_some() {
        let grid = themed(layout(&model, orient)?, theme_name)?;
        println!("{}", serde_json::to_string_pretty(&dump_ir(&grid))?);
    } else {
        generate(&model, &out_path, orient, theme_name)?;
        println!("saved {}  (orient={}, theme={})", out_path, orient, theme_name);
    }
    Ok(())
}
