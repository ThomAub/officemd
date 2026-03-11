//! Rectangle-based table detection using union-find clustering.

use std::collections::HashMap;

use log::debug;

use crate::pdf_inspector::types::{PdfRect, TextItem};

use super::Table;

/// Disjoint-set (union-find) for clustering indices.
struct UnionFind {
    parent: Vec<usize>,
    rank: Vec<usize>,
}

impl UnionFind {
    fn new(n: usize) -> Self {
        Self {
            parent: (0..n).collect(),
            rank: vec![0; n],
        }
    }

    fn find(&mut self, x: usize) -> usize {
        if self.parent[x] != x {
            self.parent[x] = self.find(self.parent[x]);
        }
        self.parent[x]
    }

    fn union(&mut self, a: usize, b: usize) {
        let ra = self.find(a);
        let rb = self.find(b);
        if ra == rb {
            return;
        }
        if self.rank[ra] < self.rank[rb] {
            self.parent[ra] = rb;
        } else if self.rank[ra] > self.rank[rb] {
            self.parent[rb] = ra;
        } else {
            self.parent[rb] = ra;
            self.rank[ra] += 1;
        }
    }
}

/// Check if two rects overlap after expanding each by `tol` on all sides.
pub(crate) fn rects_overlap(a: &(f32, f32, f32, f32), b: &(f32, f32, f32, f32), tol: f32) -> bool {
    // a and b are (x, y, w, h) where (x,y) is bottom-left corner
    let (ax, ay, aw, ah) = *a;
    let (bx, by, bw, bh) = *b;
    // Expand each rect by tol
    let a_left = ax - tol;
    let a_right = ax + aw + tol;
    let a_bottom = ay - tol;
    let a_top = ay + ah + tol;
    let b_left = bx - tol;
    let b_right = bx + bw + tol;
    let b_bottom = by - tol;
    let b_top = by + bh + tol;
    // AABB overlap: NOT (separated)
    !(a_right < b_left || b_right < a_left || a_top < b_bottom || b_top < a_bottom)
}

/// Cluster rects by spatial overlap using union-find.
/// Returns groups of rect indices; only groups with ≥ `min_size` rects are returned.
pub(crate) fn cluster_rects(
    rects: &[(f32, f32, f32, f32)],
    tolerance: f32,
    min_size: usize,
) -> Vec<Vec<usize>> {
    let n = rects.len();
    let mut uf = UnionFind::new(n);

    for i in 0..n {
        for j in (i + 1)..n {
            if rects_overlap(&rects[i], &rects[j], tolerance) {
                uf.union(i, j);
            }
        }
    }

    // Group indices by root
    let mut groups: HashMap<usize, Vec<usize>> = HashMap::new();
    for i in 0..n {
        groups.entry(uf.find(i)).or_default().push(i);
    }

    // Sort by root index for deterministic output order
    let mut result: Vec<(usize, Vec<usize>)> = groups
        .into_iter()
        .filter(|(_, g)| g.len() >= min_size)
        .collect();
    result.sort_by_key(|(root, _)| *root);
    result.into_iter().map(|(_, g)| g).collect()
}

/// A bounding box hint from cell-border rects that failed full grid validation.
///
/// When a rect cluster contains cell-sized borders but they don't form a valid
/// grid (e.g. only horizontal row borders with no vertical column dividers),
/// the bounding box of those cell-sized rects can still be used to scope
/// heuristic table detection, preventing unrelated items (graph labels, etc.)
/// from being merged into the table.
#[derive(Debug, Clone)]
pub struct RectHintRegion {
    /// X coordinate of the left edge.
    pub x_left: f32,
    /// X coordinate of the right edge.
    pub x_right: f32,
    /// Y coordinate of the top edge (highest value in PDF space)
    pub y_top: f32,
    /// Y coordinate of the bottom edge (lowest value in PDF space)
    pub y_bottom: f32,
    /// Number of cell-sized rects contributing to this hint.
    pub rect_count: usize,
}

#[derive(Clone)]
struct RectClusterGroup {
    rects: Vec<(f32, f32, f32, f32)>,
    x_left: f32,
    x_right: f32,
    y_top: f32,
    y_bottom: f32,
}

impl RectClusterGroup {
    fn from_rects(rects: Vec<(f32, f32, f32, f32)>) -> Option<Self> {
        let x_left = rects.iter().map(|(x, _, _, _)| *x).reduce(f32::min)?;
        let x_right = rects.iter().map(|(x, _, w, _)| *x + *w).reduce(f32::max)?;
        let y_bottom = rects.iter().map(|(_, y, _, _)| *y).reduce(f32::min)?;
        let y_top = rects.iter().map(|(_, y, _, h)| *y + *h).reduce(f32::max)?;
        Some(Self {
            rects,
            x_left,
            x_right,
            y_top,
            y_bottom,
        })
    }

    fn merge(&mut self, other: Self) {
        self.rects.extend(other.rects);
        self.x_left = self.x_left.min(other.x_left);
        self.x_right = self.x_right.max(other.x_right);
        self.y_bottom = self.y_bottom.min(other.y_bottom);
        self.y_top = self.y_top.max(other.y_top);
    }
}

fn merge_rect_cluster_bands(groups: Vec<RectClusterGroup>) -> Vec<RectClusterGroup> {
    if groups.len() < 2 {
        return groups;
    }

    let mut sorted = groups;
    sorted.sort_by(|a, b| {
        b.y_top
            .partial_cmp(&a.y_top)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then_with(|| {
                a.x_left
                    .partial_cmp(&b.x_left)
                    .unwrap_or(std::cmp::Ordering::Equal)
            })
    });

    let mut merged: Vec<RectClusterGroup> = Vec::new();
    for group in sorted {
        let should_merge = merged.last_mut().is_some_and(|prev| {
            let overlap =
                (prev.y_top.min(group.y_top) - prev.y_bottom.max(group.y_bottom)).max(0.0);
            let prev_h = (prev.y_top - prev.y_bottom).max(1.0);
            let group_h = (group.y_top - group.y_bottom).max(1.0);
            let overlap_ratio = overlap / prev_h.min(group_h);
            let horizontal_gap = if prev.x_right < group.x_left {
                group.x_left - prev.x_right
            } else if group.x_right < prev.x_left {
                prev.x_left - group.x_right
            } else {
                0.0
            };
            let combined_width = prev.x_right.max(group.x_right) - prev.x_left.min(group.x_left);

            overlap_ratio >= 0.65 && horizontal_gap <= 90.0 && combined_width <= 620.0
        });

        if should_merge {
            merged.last_mut().unwrap().merge(group);
        } else {
            merged.push(group);
        }
    }

    merged
}

/// Detect tables from explicit rectangle (`re`) operators in the PDF.
///
/// Many PDFs draw cell borders using `re` (rectangle) operators.  Table pages
/// typically have 100-200+ rects while non-table pages have < 30.  This function
/// clusters spatially connected rectangles into groups, then identifies grids of
/// cell-sized rectangles within each cluster and assigns text items to cells.
///
/// Also returns hint regions: bounding boxes of cell-sized rects from clusters
/// that failed full grid validation.  These can be used to scope heuristic
/// detection and prevent unrelated items from being merged into tables.
pub fn detect_tables_from_rects(
    items: &[TextItem],
    rects: &[PdfRect],
    page: u32,
) -> (Vec<Table>, Vec<RectHintRegion>) {
    // Filter rects on this page; normalize negative widths/heights; skip tiny rects.
    let mut page_rects: Vec<(f32, f32, f32, f32)> = Vec::new(); // (x, y, w, h) normalized
    for r in rects {
        if r.page != page {
            continue;
        }
        let (mut x, mut y, mut w, mut h) = (r.x, r.y, r.width, r.height);
        if w < 0.0 {
            x += w;
            w = -w;
        }
        if h < 0.0 {
            y += h;
            h = -h;
        }
        // Skip tiny rects (borders, dots, decorations)
        if w < 5.0 || h < 5.0 {
            continue;
        }
        page_rects.push((x, y, w, h));
    }

    debug!(
        "page {}: {} rects after size filter",
        page,
        page_rects.len(),
    );

    let mut tables = Vec::new();
    let mut hint_regions = Vec::new();

    // Single clustering pass with min_size=4; large clusters (>= 6) attempt
    // full grid detection first, smaller ones produce hint regions only.
    if page_rects.len() >= 4 {
        let clusters = cluster_rects(&page_rects, 3.0, 4);
        debug!("page {}: {} clusters with >= 4 rects", page, clusters.len());

        let merged_groups = merge_rect_cluster_bands(
            clusters
                .iter()
                .filter_map(|cluster_indices| {
                    let group_rects: Vec<(f32, f32, f32, f32)> =
                        cluster_indices.iter().map(|&i| page_rects[i]).collect();
                    RectClusterGroup::from_rects(group_rects)
                })
                .collect(),
        );

        for group in merged_groups {
            let group_rects = group.rects;

            if group_rects.len() >= 6 {
                // Large cluster: attempt full grid detection
                if let Some(table) = detect_table_from_rect_group(items, &group_rects, page) {
                    tables.push(table);
                    continue;
                } else if let Some(table) = detect_row_stripe_table(items, &group_rects, page) {
                    tables.push(table);
                    continue;
                }
            }

            // Cluster too small for grid, or grid detection failed: extract hint region
            if let Some(hint) = extract_hint_region(&group_rects) {
                debug!(
                    "page {}: hint region x={:.1}..{:.1} y={:.1}..{:.1} ({} rects)",
                    page,
                    hint.x_left,
                    hint.x_right,
                    hint.y_bottom,
                    hint.y_top,
                    group_rects.len()
                );
                hint_regions.push(hint);
            }
        }
    }

    merge_hint_regions(&mut hint_regions);

    (tables, hint_regions)
}

/// Extract a hint region from a rect cluster that failed grid validation.
///
/// Only produces hints from small clusters (≤ 8 rects) where a few cell-border
/// rects define a table's row boundaries.  Large clusters (form-style decorative
/// rects) are not suitable for hint regions since they typically span the whole page.
///
/// Filters out oversized "bounding box" rects (height > 4× the median height),
/// then computes the Y bounding box of the remaining cell-sized rects.
fn extract_hint_region(group_rects: &[(f32, f32, f32, f32)]) -> Option<RectHintRegion> {
    // Ignore tiny clusters. Large failed clusters can still be useful when their
    // cell-sized rects bound a table region on a busy page.
    if group_rects.len() < 2 {
        return None;
    }

    // Compute median height to identify cell-sized rects
    let mut heights: Vec<f32> = group_rects.iter().map(|&(_, _, _, h)| h).collect();
    heights.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median_h = heights[heights.len() / 2];

    // Keep only cell-sized rects (height ≤ 4× median)
    let cell_rects: Vec<&(f32, f32, f32, f32)> = group_rects
        .iter()
        .filter(|(_, _, _, h)| *h <= median_h * 4.0)
        .collect();

    if cell_rects.len() < 2 {
        return None;
    }

    // Compute bounding box of cell-sized rects
    let x_left = cell_rects.iter().map(|(x, _, _, _)| *x).reduce(f32::min)?;
    let x_right = cell_rects
        .iter()
        .map(|(x, _, w, _)| *x + *w)
        .reduce(f32::max)?;
    let y_bottom = cell_rects.iter().map(|(_, y, _, _)| *y).reduce(f32::min)?;
    let y_top = cell_rects
        .iter()
        .map(|(_, y, _, h)| *y + *h)
        .reduce(f32::max)?;

    // The region must have meaningful size but not span an unreasonable area.
    let region_width = x_right - x_left;
    let region_height = y_top - y_bottom;
    if !(40.0..=560.0).contains(&region_width) || !(10.0..=500.0).contains(&region_height) {
        return None;
    }

    Some(RectHintRegion {
        x_left,
        x_right,
        y_top,
        y_bottom,
        rect_count: cell_rects.len(),
    })
}

fn looks_like_overflowed_sparse_table(
    cells: &[Vec<String>],
    rect_count: usize,
    long_cell_threshold: usize,
) -> bool {
    if rect_count >= 10 {
        return false;
    }

    let mut saw_non_empty = false;
    let mut saw_long = false;
    for cell in cells.iter().flat_map(|row| row.iter()) {
        if cell.trim().is_empty() {
            continue;
        }
        saw_non_empty = true;
        if cell.chars().count() >= long_cell_threshold {
            saw_long = true;
            break;
        }
    }

    saw_non_empty && saw_long
}

fn merge_hint_regions(hints: &mut Vec<RectHintRegion>) {
    if hints.len() < 2 {
        return;
    }

    hints.sort_by(|a, b| {
        b.y_top
            .partial_cmp(&a.y_top)
            .unwrap_or(std::cmp::Ordering::Equal)
    });

    let mut merged: Vec<RectHintRegion> = Vec::with_capacity(hints.len());
    for hint in hints.drain(..) {
        let overlaps = merged.last_mut().is_some_and(|prev| {
            hint.x_left <= prev.x_right + 12.0
                && hint.x_right >= prev.x_left - 12.0
                && hint.y_bottom <= prev.y_top + 12.0
                && hint.y_top >= prev.y_bottom - 12.0
        });

        if overlaps {
            let prev = merged.last_mut().unwrap();
            prev.x_left = prev.x_left.min(hint.x_left);
            prev.x_right = prev.x_right.max(hint.x_right);
            prev.y_bottom = prev.y_bottom.min(hint.y_bottom);
            prev.y_top = prev.y_top.max(hint.y_top);
            prev.rect_count += hint.rect_count;
        } else {
            merged.push(hint);
        }
    }

    *hints = merged;
}

fn finalize_rect_backed_table(
    mut cells: Vec<Vec<String>>,
    item_indices: Vec<usize>,
    col_edges: &[f32],
    row_edges: &[f32],
    rect_count: usize,
    debug_prefix: &str,
) -> Option<Table> {
    if item_indices.is_empty() {
        debug!("  {}rejected: no text items assigned to grid", debug_prefix);
        return None;
    }

    let mut trimmed_col_edges: Vec<f32> = col_edges.to_vec();
    while cells.first().is_some_and(|row| row.len() > 2)
        && cells
            .iter()
            .all(|row| row.first().is_some_and(|c| c.trim().is_empty()))
    {
        for row in &mut cells {
            row.remove(0);
        }
        trimmed_col_edges.remove(0);
    }
    while cells.first().is_some_and(|row| row.len() > 2)
        && cells
            .iter()
            .all(|row| row.last().is_some_and(|c| c.trim().is_empty()))
    {
        for row in &mut cells {
            row.pop();
        }
        trimmed_col_edges.pop();
    }

    let non_empty_rows = cells
        .iter()
        .filter(|row| row.iter().any(|c| !c.trim().is_empty()))
        .count();
    if non_empty_rows < 2 {
        debug!(
            "  {}rejected: only {} non-empty rows",
            debug_prefix, non_empty_rows
        );
        return None;
    }

    let total_cells = (cells.len() * cells.first().map_or(0, Vec::len)) as f32;
    let non_empty_cells = cells
        .iter()
        .flat_map(|row| row.iter())
        .filter(|c| !c.trim().is_empty())
        .count();
    let content_ratio = if total_cells > 0.0 {
        non_empty_cells as f32 / total_cells
    } else {
        0.0
    };
    if content_ratio < 0.25 {
        debug!(
            "  {}rejected: content ratio {:.2} < 0.25 ({} non-empty / {} total)",
            debug_prefix, content_ratio, non_empty_cells, total_cells as u32
        );
        return None;
    }

    for col in 0..cells.first().map_or(0, Vec::len) {
        let col_has_content = cells
            .iter()
            .any(|row| row.get(col).is_some_and(|c| !c.trim().is_empty()));
        if !col_has_content {
            debug!(
                "  {}rejected: column {} is completely empty",
                debug_prefix, col
            );
            return None;
        }
    }

    if looks_like_overflowed_sparse_table(&cells, rect_count, 80) {
        debug!(
            "  {}rejected: sparse rect table has overflowed cells",
            debug_prefix
        );
        return None;
    }

    let columns: Vec<f32> = trimmed_col_edges
        .windows(2)
        .map(|pair| (pair[0] + pair[1]) / 2.0)
        .collect();
    let rows: Vec<f32> = row_edges
        .windows(2)
        .map(|pair| (pair[0] + pair[1]) / 2.0)
        .collect();

    Some(Table {
        columns,
        rows,
        cells,
        item_indices,
    })
}

/// Detect a single table from a cluster of spatially connected rects.
///
/// Contains the grid-detection logic: snap edges, fill-ratio check,
/// assign items to grid, content density validation.
pub(crate) fn detect_table_from_rect_group(
    items: &[TextItem],
    group_rects: &[(f32, f32, f32, f32)],
    page: u32,
) -> Option<Table> {
    // Extract unique X and Y edges from all rects
    let mut x_edges: Vec<f32> = Vec::new();
    let mut y_edges: Vec<f32> = Vec::new();
    for &(x, y, w, h) in group_rects {
        x_edges.push(x);
        x_edges.push(x + w);
        y_edges.push(y);
        y_edges.push(y + h);
    }

    let x_edges = snap_edges(&x_edges, 6.0);
    let y_edges = snap_edges(&y_edges, 6.0);

    if x_edges.len() < 3 || y_edges.len() < 4 {
        debug!(
            "  rejected: {} x-edges, {} y-edges (need >=3, >=4)",
            x_edges.len(),
            y_edges.len()
        );
        return None;
    }

    // Sort column edges left-to-right, row edges top-to-bottom (highest Y first for PDF)
    let mut col_edges = x_edges;
    col_edges.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mut row_edges = y_edges;
    row_edges.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));

    let num_cols = col_edges.len() - 1;
    let num_rows = row_edges.len() - 1;

    if num_cols < 2 || num_rows < 2 {
        return None;
    }

    // Reject grids that are too large — real tables rarely exceed 12 columns.
    // Form-style PDFs with scattered field boxes produce huge sparse grids.
    if num_cols > 12 {
        return None;
    }

    // Verify that cell-sized rects actually fill the grid
    // Count how many grid cells have a matching rect
    let mut filled_cells = 0u32;
    for row in 0..num_rows {
        let y_top = row_edges[row];
        let y_bot = row_edges[row + 1];
        for col in 0..num_cols {
            let x_left = col_edges[col];
            let x_right = col_edges[col + 1];
            // Check if any rect approximately covers this cell
            let cell_covered = group_rects.iter().any(|&(rx, ry, rw, rh)| {
                let tol = 6.0;
                rx <= x_left + tol
                    && (rx + rw) >= x_right - tol
                    && ry <= y_top + tol
                    && (ry + rh) >= y_bot - tol
            });
            if cell_covered {
                filled_cells += 1;
            }
        }
    }

    let total_cells = (num_cols * num_rows) as f32;
    let fill_ratio = filled_cells as f32 / total_cells;

    debug!(
        "  grid: {}x{} = {} cells, {} filled, ratio={:.2}",
        num_rows, num_cols, total_cells as u32, filled_cells, fill_ratio
    );

    // Require at least 30% of cells to be backed by rects
    if fill_ratio < 0.3 {
        debug!("  rejected: fill ratio {:.2} < 0.30", fill_ratio);
        return None;
    }

    // Build table: assign text items to cells
    let (mut cells, item_indices) = assign_items_to_grid(items, &col_edges, &row_edges, page);

    // Consolidate vertically-merged cells: rects spanning multiple grid rows
    // should have their text collected into the first sub-row.
    propagate_merged_cells(&mut cells, &col_edges, &row_edges, group_rects);

    finalize_rect_backed_table(
        cells,
        item_indices,
        &col_edges,
        &row_edges,
        group_rects.len(),
        "",
    )
}

/// Deduplicate nearby edge values within a tolerance, returning sorted unique edges.
pub(crate) fn snap_edges(values: &[f32], tolerance: f32) -> Vec<f32> {
    let mut sorted: Vec<f32> = values.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let mut snapped: Vec<f32> = Vec::new();
    for &v in &sorted {
        if let Some(last) = snapped.last() {
            if (v - *last).abs() <= tolerance {
                continue; // Skip — too close to previous edge
            }
        }
        snapped.push(v);
    }
    snapped
}

/// Assign text items to grid cells defined by column/row edges.
///
/// Returns `(cells, item_indices)` where `cells[row][col]` is the cell text
/// and `item_indices` lists the original item indices that were consumed.
pub(crate) fn assign_items_to_grid(
    items: &[TextItem],
    col_edges: &[f32],
    row_edges: &[f32],
    page: u32,
) -> (Vec<Vec<String>>, Vec<usize>) {
    let num_cols = col_edges.len() - 1;
    let num_rows = row_edges.len() - 1;

    // Collect items per cell for proper sorting before joining
    let mut cell_items: Vec<Vec<Vec<(usize, &TextItem)>>> =
        vec![vec![Vec::new(); num_cols]; num_rows];
    let mut indices = Vec::new();

    for (idx, item) in items.iter().enumerate() {
        if item.page != page {
            continue;
        }
        if item.is_rotated {
            continue;
        }
        // Use item center for assignment
        let cx = item.x + item.width / 2.0;
        let cy = item.y;

        // Find column: cx must be between col_edges[c] and col_edges[c+1]
        let col = (0..num_cols).find(|&c| cx >= col_edges[c] - 2.0 && cx <= col_edges[c + 1] + 2.0);
        // Find row: cy must be between row_edges[r+1] (bottom) and row_edges[r] (top)
        let row = (0..num_rows).find(|&r| cy >= row_edges[r + 1] - 2.0 && cy <= row_edges[r] + 2.0);

        if let (Some(c), Some(r)) = (col, row) {
            cell_items[r][c].push((idx, item));
            indices.push(idx);
        }
    }

    // Build cell strings: sort items within each cell by Y descending then X ascending
    let mut cells: Vec<Vec<String>> = Vec::with_capacity(num_rows);
    for row_items in &mut cell_items {
        let mut row_cells = Vec::with_capacity(num_cols);
        for col_items in row_items.iter_mut() {
            col_items.sort_by(|a, b| {
                b.1.y
                    .partial_cmp(&a.1.y)
                    .unwrap_or(std::cmp::Ordering::Equal)
                    .then_with(|| {
                        a.1.x
                            .partial_cmp(&b.1.x)
                            .unwrap_or(std::cmp::Ordering::Equal)
                    })
            });
            let text: String = col_items
                .iter()
                .map(|(_, item)| item.text.trim())
                .filter(|t| !t.is_empty())
                .collect::<Vec<_>>()
                .join(" ");
            row_cells.push(text);
        }
        cells.push(row_cells);
    }

    (cells, indices)
}

/// Consolidate text in vertically-merged cells.
///
/// When a single rect spans multiple grid rows (e.g. a "Classification" label
/// covering several price sub-rows), text ends up in only one sub-row while the
/// others have an empty cell.  This function detects such spans and moves all
/// text into the first sub-row, clearing the rest so that downstream
/// continuation-merge in `clean_table_cells` collapses sub-rows correctly.
fn propagate_merged_cells(
    cells: &mut [Vec<String>],
    col_edges: &[f32],
    row_edges: &[f32],
    group_rects: &[(f32, f32, f32, f32)],
) {
    let num_cols = col_edges.len() - 1;
    let num_rows = row_edges.len() - 1;
    let tol = 6.0;

    for col in 0..num_cols {
        for rect in group_rects {
            let (rx, ry, rw, rh) = *rect;

            // Rect must cover this column
            if rx > col_edges[col] + tol || (rx + rw) < col_edges[col + 1] - tol {
                continue;
            }

            // Find first and last grid rows that the rect spans
            let first_row = (0..num_rows)
                .find(|&r| ry <= row_edges[r] + tol && (ry + rh) >= row_edges[r + 1] - tol);
            let last_row = (0..num_rows)
                .rfind(|&r| ry <= row_edges[r] + tol && (ry + rh) >= row_edges[r + 1] - tol);

            let (first, last) = match (first_row, last_row) {
                (Some(f), Some(l)) if l > f => (f, l),
                _ => continue, // Single row or no match — skip
            };

            // Collect all text from sub-rows within the merged range
            let mut combined = String::new();
            for row in cells.iter().take(last + 1).skip(first) {
                let text = row[col].trim();
                if !text.is_empty() {
                    if !combined.is_empty() {
                        combined.push(' ');
                    }
                    combined.push_str(text);
                }
            }

            // Place combined text in the first sub-row, clear the rest
            cells[first][col] = combined;
            for row in cells.iter_mut().take(last + 1).skip(first + 1) {
                row[col] = String::new();
            }
        }
    }
}

/// Check if rects form a row-stripe pattern (full-width horizontal bands).
///
/// Row-stripe shading uses rects that all share similar X position and width,
/// spanning the full table width. This produces only ~2 unique X-edges, which
/// makes normal grid detection fail (1-column grid).
fn is_row_stripe_pattern(rects: &[(f32, f32, f32, f32)]) -> bool {
    if rects.len() < 3 {
        return false;
    }

    let mut widths: Vec<f32> = rects.iter().map(|&(_, _, w, _)| w).collect();
    widths.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median_width = widths[widths.len() / 2];

    // Must be page-spanning (>200pt)
    if median_width <= 200.0 {
        return false;
    }

    // >75% of rects should have width within 10% of median
    let within_tolerance = rects
        .iter()
        .filter(|&&(_, _, w, _)| (w - median_width).abs() <= median_width * 0.10)
        .count();

    within_tolerance as f32 / rects.len() as f32 > 0.75
}

/// Detect a table from row-stripe rects by using rect Y-edges for rows
/// and text X-position clustering for columns.
fn detect_row_stripe_table(
    items: &[TextItem],
    group_rects: &[(f32, f32, f32, f32)],
    page: u32,
) -> Option<Table> {
    if !is_row_stripe_pattern(group_rects) {
        return None;
    }

    debug!(
        "  trying row-stripe detection ({} rects)",
        group_rects.len()
    );

    // Extract Y-edges from rects
    let mut y_edges: Vec<f32> = Vec::new();
    for &(_, y, _, h) in group_rects {
        y_edges.push(y);
        y_edges.push(y + h);
    }
    let y_edges = snap_edges(&y_edges, 6.0);

    if y_edges.len() < 4 {
        debug!("  row-stripe rejected: only {} y-edges", y_edges.len());
        return None;
    }

    // Sort row edges top-to-bottom (highest Y first for PDF)
    let mut row_edges = y_edges;
    row_edges.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));

    // Compute the bounding box of the stripe region for filtering items
    let y_top = row_edges[0];
    let y_bottom = *row_edges.last().unwrap();
    let x_left = group_rects
        .iter()
        .map(|&(x, _, _, _)| x)
        .reduce(f32::min)
        .unwrap();
    let x_right = group_rects
        .iter()
        .map(|&(x, _, w, _)| x + w)
        .reduce(f32::max)
        .unwrap();

    // Gather page items within the stripe region
    let page_items: Vec<(usize, &TextItem)> = items
        .iter()
        .enumerate()
        .filter(|(_, item)| {
            item.page == page
                && !item.is_rotated
                && item.y >= y_bottom - 2.0
                && item.y <= y_top + 2.0
                && item.x >= x_left - 5.0
                && item.x + item.width <= x_right + 5.0
        })
        .collect();

    if page_items.is_empty() {
        return None;
    }

    // Derive column boundaries from text X-position clustering.
    // Use a lower threshold than find_column_boundaries (which clamps at 25pt min)
    // since we already know this is a table from the rects and narrow columns
    // (e.g. row-number + date at 21pt gap) should stay separate.
    let columns = cluster_x_positions(&page_items, 15.0);

    if columns.len() < 2 {
        debug!(
            "  row-stripe rejected: only {} columns from text clustering",
            columns.len()
        );
        return None;
    }

    // Convert column centers to column edges (midpoints between adjacent, plus outer edges)
    let mut col_edges: Vec<f32> = Vec::with_capacity(columns.len() + 1);

    // Left edge: minimum item X minus small padding
    let min_x = page_items
        .iter()
        .map(|(_, i)| i.x)
        .reduce(f32::min)
        .unwrap();
    col_edges.push(min_x - 5.0);

    // Midpoints between adjacent column centers
    for pair in columns.windows(2) {
        col_edges.push((pair[0] + pair[1]) / 2.0);
    }

    // Right edge: maximum item right edge plus small padding
    let max_x_right = page_items
        .iter()
        .map(|(_, i)| i.x + i.width)
        .reduce(f32::max)
        .unwrap();
    col_edges.push(max_x_right + 5.0);

    let num_cols = col_edges.len() - 1;
    let num_rows = row_edges.len() - 1;

    debug!(
        "  row-stripe grid: {}x{} ({} col edges, {} row edges)",
        num_rows,
        num_cols,
        col_edges.len(),
        row_edges.len()
    );

    // Assign items to grid
    let (cells, item_indices) = assign_items_to_grid(items, &col_edges, &row_edges, page);

    if item_indices.is_empty() {
        debug!("  row-stripe rejected: no items assigned");
        return None;
    }

    finalize_rect_backed_table(
        cells,
        item_indices,
        &col_edges,
        &row_edges,
        group_rects.len(),
        "row-stripe ",
    )
}

/// Cluster text item X positions into column centers with a given minimum threshold.
///
/// Similar to `find_column_boundaries` in grid.rs but with a lower minimum threshold
/// suitable for rect-backed tables where we already know tabular structure exists
/// (no need for anti-paragraph safeguards).
fn cluster_x_positions(items: &[(usize, &TextItem)], min_threshold: f32) -> Vec<f32> {
    let mut x_positions: Vec<f32> = items.iter().map(|(_, i)| i.x).collect();
    x_positions.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    if x_positions.is_empty() {
        return vec![];
    }

    let x_range = x_positions.last().unwrap() - x_positions.first().unwrap();
    let avg_gap = if x_positions.len() > 1 {
        x_range / (x_positions.len() - 1) as f32
    } else {
        60.0
    };
    let cluster_threshold = avg_gap.clamp(min_threshold, 50.0);

    let mut columns = Vec::new();
    let mut cluster_sum = x_positions[0];
    let mut cluster_count = 1usize;

    for &x in &x_positions[1..] {
        let cluster_center = cluster_sum / cluster_count as f32;
        if x - cluster_center > cluster_threshold {
            columns.push(cluster_center);
            cluster_sum = x;
            cluster_count = 1;
        } else {
            cluster_sum += x;
            cluster_count += 1;
        }
    }
    columns.push(cluster_sum / cluster_count as f32);

    // Filter: each column needs multiple items
    let min_items_per_col = (items.len() / columns.len().max(1) / 4).max(2);
    columns
        .into_iter()
        .filter(|&col_x| {
            items
                .iter()
                .filter(|(_, i)| (i.x - col_x).abs() < cluster_threshold)
                .count()
                >= min_items_per_col
        })
        .collect()
}
