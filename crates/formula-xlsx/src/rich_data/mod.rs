//! Parsers/utilities for Excel "rich data" parts.
//!
//! Excel stores cell-level rich values (data types, images-in-cells, etc.) via:
//! - `xl/worksheets/sheet*.xml` `c/@vm` (value-metadata index)
//! - `xl/metadata.xml` `<valueMetadata>` + `<futureMetadata name="XLRICHVALUE">`
//! - `xl/richData/richValue.xml` (rich-value records, indexed by `rvb/@i`)
//!
//! This module currently only supports resolving `vm` indices to rich-value indices.

pub mod metadata;

