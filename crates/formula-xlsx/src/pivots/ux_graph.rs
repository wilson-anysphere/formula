use std::collections::BTreeMap;

use crate::{XlsxDocument, XlsxError, XlsxPackage};

use super::graph::PivotTableInstance;
use super::pivot_charts::PivotChartPart;
use super::slicers::{SlicerDefinition, TimelineDefinition};

/// A unified view of pivot-related UX artifacts (pivot tables/caches, pivot charts, slicers, and
/// timelines) plus their connections.
///
/// The UI can use this struct to quickly answer questions like:
/// - Which slicers/timelines are connected to a given pivot table?
/// - Which pivot charts are bound to a given pivot table?
/// - Which pivot tables does a slicer/timeline affect?
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxPivotUxGraph {
    pub pivot_tables: Vec<PivotTableInstance>,
    pub pivot_charts: Vec<PivotChartPart>,
    pub slicers: Vec<SlicerDefinition>,
    pub timelines: Vec<TimelineDefinition>,

    /// For each slicer (index into [`XlsxPivotUxGraph::slicers`]), the indices of connected pivot
    /// tables (indices into [`XlsxPivotUxGraph::pivot_tables`]).
    pub slicer_to_pivot_tables: Vec<Vec<usize>>,
    /// For each timeline (index into [`XlsxPivotUxGraph::timelines`]), the indices of connected
    /// pivot tables (indices into [`XlsxPivotUxGraph::pivot_tables`]).
    pub timeline_to_pivot_tables: Vec<Vec<usize>>,
    /// For each pivot chart (index into [`XlsxPivotUxGraph::pivot_charts`]), the index of the
    /// pivot table it targets (index into [`XlsxPivotUxGraph::pivot_tables`]).
    pub pivot_chart_to_pivot_table: Vec<Option<usize>>,

    /// For each pivot table (index into [`XlsxPivotUxGraph::pivot_tables`]), the slicer indices
    /// that are connected to it (indices into [`XlsxPivotUxGraph::slicers`]).
    pub pivot_table_to_slicers: Vec<Vec<usize>>,
    /// For each pivot table (index into [`XlsxPivotUxGraph::pivot_tables`]), the timeline indices
    /// that are connected to it (indices into [`XlsxPivotUxGraph::timelines`]).
    pub pivot_table_to_timelines: Vec<Vec<usize>>,
    /// For each pivot table (index into [`XlsxPivotUxGraph::pivot_tables`]), the pivot chart
    /// indices that target it (indices into [`XlsxPivotUxGraph::pivot_charts`]).
    pub pivot_table_to_pivot_charts: Vec<Vec<usize>>,
}

fn build_pivot_ux_graph(
    pivot_tables: Vec<PivotTableInstance>,
    pivot_charts: Vec<PivotChartPart>,
    slicers: Vec<SlicerDefinition>,
    timelines: Vec<TimelineDefinition>,
) -> XlsxPivotUxGraph {
    let mut pivot_table_index_by_part: BTreeMap<String, usize> = BTreeMap::new();
    for (idx, table) in pivot_tables.iter().enumerate() {
        // Deterministic: first wins.
        pivot_table_index_by_part
            .entry(table.pivot_table_part.clone())
            .or_insert(idx);
    }

    let mut slicer_to_pivot_tables: Vec<Vec<usize>> = Vec::with_capacity(slicers.len());
    let mut pivot_table_to_slicers: Vec<Vec<usize>> = vec![Vec::new(); pivot_tables.len()];
    for (slicer_idx, slicer) in slicers.iter().enumerate() {
        let mut connected = Vec::new();
        for part in &slicer.connected_pivot_tables {
            if let Some(&pivot_idx) = pivot_table_index_by_part.get(part) {
                connected.push(pivot_idx);
            }
        }
        connected.sort_unstable();
        connected.dedup();
        for &pivot_idx in &connected {
            if let Some(list) = pivot_table_to_slicers.get_mut(pivot_idx) {
                list.push(slicer_idx);
            }
        }
        slicer_to_pivot_tables.push(connected);
    }
    for list in &mut pivot_table_to_slicers {
        list.sort_unstable();
        list.dedup();
    }

    let mut timeline_to_pivot_tables: Vec<Vec<usize>> = Vec::with_capacity(timelines.len());
    let mut pivot_table_to_timelines: Vec<Vec<usize>> = vec![Vec::new(); pivot_tables.len()];
    for (timeline_idx, timeline) in timelines.iter().enumerate() {
        let mut connected = Vec::new();
        for part in &timeline.connected_pivot_tables {
            if let Some(&pivot_idx) = pivot_table_index_by_part.get(part) {
                connected.push(pivot_idx);
            }
        }
        connected.sort_unstable();
        connected.dedup();
        for &pivot_idx in &connected {
            if let Some(list) = pivot_table_to_timelines.get_mut(pivot_idx) {
                list.push(timeline_idx);
            }
        }
        timeline_to_pivot_tables.push(connected);
    }
    for list in &mut pivot_table_to_timelines {
        list.sort_unstable();
        list.dedup();
    }

    let mut pivot_chart_to_pivot_table: Vec<Option<usize>> = Vec::with_capacity(pivot_charts.len());
    let mut pivot_table_to_pivot_charts: Vec<Vec<usize>> = vec![Vec::new(); pivot_tables.len()];
    for (chart_idx, chart) in pivot_charts.iter().enumerate() {
        let target = chart
            .pivot_source_part
            .as_deref()
            .and_then(|part| pivot_table_index_by_part.get(part).copied());
        pivot_chart_to_pivot_table.push(target);
        if let Some(pivot_idx) = target {
            if let Some(list) = pivot_table_to_pivot_charts.get_mut(pivot_idx) {
                list.push(chart_idx);
            }
        }
    }
    for list in &mut pivot_table_to_pivot_charts {
        list.sort_unstable();
        list.dedup();
    }

    XlsxPivotUxGraph {
        pivot_tables,
        pivot_charts,
        slicers,
        timelines,
        slicer_to_pivot_tables,
        timeline_to_pivot_tables,
        pivot_chart_to_pivot_table,
        pivot_table_to_slicers,
        pivot_table_to_timelines,
        pivot_table_to_pivot_charts,
    }
}

impl XlsxPackage {
    /// Build a unified pivot UX graph by correlating:
    /// - [`XlsxPackage::pivot_graph`] (pivot tables + caches + sheet placement)
    /// - [`XlsxPackage::pivot_chart_parts`] (pivot charts)
    /// - [`XlsxPackage::pivot_slicer_parts`] (slicers + timelines)
    pub fn pivot_ux_graph(&self) -> Result<XlsxPivotUxGraph, XlsxError> {
        let pivot_graph = self.pivot_graph()?;
        let pivot_charts = self.pivot_chart_parts()?;
        let slicer_parts = self.pivot_slicer_parts()?;
        Ok(build_pivot_ux_graph(
            pivot_graph.pivot_tables,
            pivot_charts,
            slicer_parts.slicers,
            slicer_parts.timelines,
        ))
    }
}

impl XlsxDocument {
    /// Build a unified pivot UX graph using only the preserved parts in an [`XlsxDocument`].
    pub fn pivot_ux_graph(&self) -> Result<XlsxPivotUxGraph, XlsxError> {
        let pivot_graph = self.pivot_graph()?;
        let pivot_charts = self.pivot_chart_parts()?;
        let slicer_parts = self.pivot_slicer_parts()?;
        Ok(build_pivot_ux_graph(
            pivot_graph.pivot_tables,
            pivot_charts,
            slicer_parts.slicers,
            slicer_parts.timelines,
        ))
    }
}
