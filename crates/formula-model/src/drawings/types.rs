use std::collections::{BTreeMap, HashMap};

use serde::{Deserialize, Serialize};

use crate::CellRef;

/// English Metric Units (EMU) are the canonical unit used by DrawingML.
///
/// 1 inch = 914_400 EMU.
pub type Emu = i64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct EmuSize {
    pub cx: Emu,
    pub cy: Emu,
}

impl EmuSize {
    pub const fn new(cx: Emu, cy: Emu) -> Self {
        Self { cx, cy }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellOffset {
    pub x_emu: Emu,
    pub y_emu: Emu,
}

impl CellOffset {
    pub const fn new(x_emu: Emu, y_emu: Emu) -> Self {
        Self { x_emu, y_emu }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct AnchorPoint {
    pub cell: CellRef,
    pub offset: CellOffset,
}

impl AnchorPoint {
    pub const fn new(cell: CellRef, offset: CellOffset) -> Self {
        Self { cell, offset }
    }
}

/// Spreadsheet drawing anchors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Anchor {
    /// `xdr:oneCellAnchor` – anchored to a single cell with explicit size.
    OneCell { from: AnchorPoint, ext: EmuSize },
    /// `xdr:twoCellAnchor` – anchored between two cell corners.
    TwoCell { from: AnchorPoint, to: AnchorPoint },
    /// `xdr:absoluteAnchor` – absolute position from the sheet origin.
    Absolute { pos: CellOffset, ext: EmuSize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
pub struct DrawingObjectId(pub u32);

#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
pub struct ImageId(pub String);

impl ImageId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ImageData {
    pub bytes: Vec<u8>,
    /// Best-effort content type (e.g. `image/png`). Optional because
    /// callers may only have file extension information.
    pub content_type: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ImageStore {
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    images: BTreeMap<ImageId, ImageData>,
}

impl ImageStore {
    pub fn insert(&mut self, id: ImageId, image: ImageData) -> Option<ImageData> {
        self.images.insert(id, image)
    }

    pub fn get(&self, id: &ImageId) -> Option<&ImageData> {
        self.images.get(id)
    }

    pub fn ids(&self) -> impl Iterator<Item = &ImageId> {
        self.images.keys()
    }

    pub fn iter(&self) -> impl Iterator<Item = (&ImageId, &ImageData)> {
        self.images.iter()
    }

    pub fn is_empty(&self) -> bool {
        self.images.is_empty()
    }

    pub fn ensure_unique_name(&self, base: &str, ext: &str) -> ImageId {
        if !self
            .images
            .contains_key(&ImageId::new(format!("{base}.{ext}")))
        {
            return ImageId::new(format!("{base}.{ext}"));
        }

        let mut i: u64 = 1;
        loop {
            let candidate = ImageId::new(format!("{base}{i}.{ext}"));
            if !self.images.contains_key(&candidate) {
                return candidate;
            }
            i = i.wrapping_add(1);
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum DrawingObjectKind {
    Image {
        image_id: ImageId,
    },
    Shape {
        /// Raw DrawingML payload (e.g. `<xdr:sp>…</xdr:sp>`).
        raw_xml: String,
    },
    ChartPlaceholder {
        /// Relationship ID inside the drawing part (e.g. `rId5`).
        rel_id: String,
        raw_xml: String,
    },
    Unknown {
        raw_xml: String,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DrawingObject {
    pub id: DrawingObjectId,
    pub kind: DrawingObjectKind,
    pub anchor: Anchor,
    /// z-order from back (lower) to front (higher).
    pub z_order: i32,
    /// Best-effort size in EMU. For `Anchor::OneCell` and `Anchor::Absolute` this
    /// matches the anchor ext. For `Anchor::TwoCell` it is extracted from the
    /// object's transform if present.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<EmuSize>,
    /// Per-object preserved metadata for format compatibility layers.
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    pub preserved: HashMap<String, String>,
}
