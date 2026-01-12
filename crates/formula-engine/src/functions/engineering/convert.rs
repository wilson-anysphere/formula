use std::collections::HashMap;
use std::sync::OnceLock;

use crate::value::ErrorKind;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Dimension {
    Length,
    Mass,
    Time,
    Pressure,
    Force,
    Energy,
    Power,
    Temperature,
    Area,
    Volume,
    Speed,
    Information,
    MagneticFluxDensity,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PrefixSupport {
    None,
    Metric,
    Binary,
}

#[derive(Debug, Clone, Copy)]
struct UnitDef {
    dim: Dimension,
    /// Convert a value in this unit to the dimension base unit via:
    ///
    /// `base = value * scale + offset`
    ///
    /// For most units, `offset` is 0. Temperature units use non-zero offsets.
    scale: f64,
    offset: f64,
    prefix: PrefixSupport,
}

/// CONVERT(number, from_unit, to_unit)
///
/// Excel-compatible unit conversions.
///
/// Error semantics:
/// - Invalid unit name => `#N/A`
/// - Incompatible unit dimensions => `#N/A`
/// - Non-finite numeric input/output => `#NUM!`
pub fn convert(number: f64, from_unit: &str, to_unit: &str) -> Result<f64, ErrorKind> {
    if !number.is_finite() {
        return Err(ErrorKind::Num);
    }

    let from = parse_unit(from_unit).ok_or(ErrorKind::NA)?;
    let to = parse_unit(to_unit).ok_or(ErrorKind::NA)?;
    if from.dim != to.dim {
        return Err(ErrorKind::NA);
    }

    let base = number * from.scale + from.offset;
    if !base.is_finite() {
        return Err(ErrorKind::Num);
    }

    let out = (base - to.offset) / to.scale;
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

fn parse_unit(unit: &str) -> Option<UnitDef> {
    // Excel unit names are case-sensitive; match exact names first.
    if let Some(def) = unit_table().get(unit).copied() {
        return Some(def);
    }

    // Information units use binary prefixes in Excel (k/M/G/...).
    for (prefix, factor) in BINARY_PREFIXES {
        if let Some(rest) = unit.strip_prefix(prefix) {
            if rest.is_empty() {
                continue;
            }
            if let Some(base) = unit_table().get(rest).copied() {
                if base.prefix != PrefixSupport::Binary {
                    continue;
                }
                return Some(UnitDef {
                    dim: base.dim,
                    scale: base.scale * factor,
                    offset: base.offset,
                    prefix: base.prefix,
                });
            }
        }
    }

    // Metric prefixes apply to a subset of units (SI and a few others).
    for (prefix, factor) in METRIC_PREFIXES {
        if let Some(rest) = unit.strip_prefix(prefix) {
            if rest.is_empty() {
                continue;
            }
            if let Some(base) = unit_table().get(rest).copied() {
                if base.prefix != PrefixSupport::Metric {
                    continue;
                }
                let exp = metric_prefix_exponent(rest);
                return Some(UnitDef {
                    dim: base.dim,
                    scale: base.scale * factor.powi(exp),
                    offset: base.offset,
                    prefix: base.prefix,
                });
            }
        }
    }

    None
}

fn metric_prefix_exponent(base_unit: &str) -> i32 {
    match base_unit.chars().last() {
        Some('2') => 2,
        Some('3') => 3,
        _ => 1,
    }
}

const METRIC_PREFIXES: &[(&str, f64)] = &[
    // Largest prefixes first; `da` must come before `d`.
    ("Y", 1e24),
    ("Z", 1e21),
    ("E", 1e18),
    ("P", 1e15),
    ("T", 1e12),
    ("G", 1e9),
    ("M", 1e6),
    ("k", 1e3),
    ("h", 1e2),
    ("da", 1e1),
    ("d", 1e-1),
    ("c", 1e-2),
    ("m", 1e-3),
    // Excel historically uses `u` for micro, but accepts `µ` in modern versions.
    ("u", 1e-6),
    ("µ", 1e-6),
    ("n", 1e-9),
    ("p", 1e-12),
    ("f", 1e-15),
    ("a", 1e-18),
    ("z", 1e-21),
    ("y", 1e-24),
];

const BINARY_PREFIXES: &[(&str, f64)] = &[
    // 2^(10*n) prefixes for information units.
    ("Y", 1_208_925_819_614_629_174_706_176.0), // 2^80
    ("Z", 1_180_591_620_717_411_303_424.0),     // 2^70
    ("E", 1_152_921_504_606_846_976.0),         // 2^60
    ("P", 1_125_899_906_842_624.0),             // 2^50
    ("T", 1_099_511_627_776.0),                 // 2^40
    ("G", 1_073_741_824.0),                     // 2^30
    ("M", 1_048_576.0),                         // 2^20
    ("k", 1_024.0),                             // 2^10
];

fn unit_table() -> &'static HashMap<&'static str, UnitDef> {
    static TABLE: OnceLock<HashMap<&'static str, UnitDef>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let mut units: HashMap<&'static str, UnitDef> = HashMap::new();

        // Length (base: meter).
        units.insert(
            "m",
            UnitDef {
                dim: Dimension::Length,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "in",
            UnitDef {
                dim: Dimension::Length,
                scale: 0.0254,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ft",
            UnitDef {
                dim: Dimension::Length,
                scale: 0.3048,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "yd",
            UnitDef {
                dim: Dimension::Length,
                scale: 0.9144,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "mi",
            UnitDef {
                dim: Dimension::Length,
                scale: 1609.344,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "Nmi",
            UnitDef {
                dim: Dimension::Length,
                scale: 1852.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ang",
            UnitDef {
                dim: Dimension::Length,
                scale: 1e-10,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        // Typographic pica: 1 pica = 1/6 inch.
        units.insert(
            "Pica",
            UnitDef {
                dim: Dimension::Length,
                scale: 0.0254 / 6.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Area (base: square meter).
        units.insert(
            "m2",
            UnitDef {
                dim: Dimension::Area,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "in2",
            UnitDef {
                dim: Dimension::Area,
                scale: 0.0254_f64.powi(2),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ft2",
            UnitDef {
                dim: Dimension::Area,
                scale: 0.3048_f64.powi(2),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "yd2",
            UnitDef {
                dim: Dimension::Area,
                scale: 0.9144_f64.powi(2),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "acre",
            UnitDef {
                dim: Dimension::Area,
                scale: 4046.8564224,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ha",
            UnitDef {
                dim: Dimension::Area,
                scale: 10000.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Volume (base: cubic meter).
        units.insert(
            "m3",
            UnitDef {
                dim: Dimension::Volume,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "in3",
            UnitDef {
                dim: Dimension::Volume,
                scale: 0.0254_f64.powi(3),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ft3",
            UnitDef {
                dim: Dimension::Volume,
                scale: 0.3048_f64.powi(3),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "yd3",
            UnitDef {
                dim: Dimension::Volume,
                scale: 0.9144_f64.powi(3),
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        // Liter.
        let liter = UnitDef {
            dim: Dimension::Volume,
            scale: 0.001,
            offset: 0.0,
            prefix: PrefixSupport::Metric,
        };
        units.insert("L", liter);
        units.insert("l", liter);

        units.insert(
            "gal",
            UnitDef {
                dim: Dimension::Volume,
                // US liquid gallon.
                scale: 0.003_785_411_784,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Mass (base: gram).
        units.insert(
            "g",
            UnitDef {
                dim: Dimension::Mass,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "lbm",
            UnitDef {
                dim: Dimension::Mass,
                scale: 453.59237,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "ozm",
            UnitDef {
                dim: Dimension::Mass,
                scale: 453.59237 / 16.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        // Slug.
        units.insert(
            "sg",
            UnitDef {
                dim: Dimension::Mass,
                // 1 slug = 32.17404855643044 lbm.
                scale: 32.174_048_556_430_44 * 453.59237,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Time (base: second).
        let second = UnitDef {
            dim: Dimension::Time,
            scale: 1.0,
            offset: 0.0,
            prefix: PrefixSupport::Metric,
        };
        units.insert("sec", second);
        units.insert("s", second);
        units.insert(
            "mn",
            UnitDef {
                dim: Dimension::Time,
                scale: 60.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "hr",
            UnitDef {
                dim: Dimension::Time,
                scale: 3600.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "day",
            UnitDef {
                dim: Dimension::Time,
                scale: 86400.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "yr",
            UnitDef {
                dim: Dimension::Time,
                // Julian year.
                scale: 31_557_600.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Pressure (base: pascal).
        units.insert(
            "Pa",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "atm",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 101_325.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "psi",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 6_894.757_293_168_361,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "Torr",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 101_325.0 / 760.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "mmHg",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 101_325.0 / 760.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "bar",
            UnitDef {
                dim: Dimension::Pressure,
                scale: 100_000.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );

        // Force (base: newton).
        units.insert(
            "N",
            UnitDef {
                dim: Dimension::Force,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "dyn",
            UnitDef {
                dim: Dimension::Force,
                scale: 1e-5,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "lbf",
            UnitDef {
                dim: Dimension::Force,
                scale: 4.448_221_615_260_5,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Energy (base: joule).
        units.insert(
            "J",
            UnitDef {
                dim: Dimension::Energy,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "e",
            UnitDef {
                dim: Dimension::Energy,
                // erg
                scale: 1e-7,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "cal",
            UnitDef {
                dim: Dimension::Energy,
                // International table calorie.
                scale: 4.1868,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "Wh",
            UnitDef {
                dim: Dimension::Energy,
                scale: 3600.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "BTU",
            UnitDef {
                dim: Dimension::Energy,
                // British thermal unit (IT).
                scale: 1055.055_852_62,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Power (base: watt).
        units.insert(
            "W",
            UnitDef {
                dim: Dimension::Power,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "hp",
            UnitDef {
                dim: Dimension::Power,
                // Mechanical horsepower.
                scale: 745.699_871_582_270_2,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Temperature (base: kelvin).
        units.insert(
            "K",
            UnitDef {
                dim: Dimension::Temperature,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "C",
            UnitDef {
                dim: Dimension::Temperature,
                scale: 1.0,
                offset: 273.15,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "F",
            UnitDef {
                dim: Dimension::Temperature,
                scale: 5.0 / 9.0,
                offset: 273.15 - 32.0 * (5.0 / 9.0),
                prefix: PrefixSupport::None,
            },
        );

        // Speed (base: meter/second).
        units.insert(
            "m/s",
            UnitDef {
                dim: Dimension::Speed,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "m/hr",
            UnitDef {
                dim: Dimension::Speed,
                scale: 1.0 / 3600.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "mph",
            UnitDef {
                dim: Dimension::Speed,
                scale: 1609.344 / 3600.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );
        units.insert(
            "kn",
            UnitDef {
                dim: Dimension::Speed,
                scale: 1852.0 / 3600.0,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        // Information (base: bit).
        units.insert(
            "bit",
            UnitDef {
                dim: Dimension::Information,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Binary,
            },
        );
        units.insert(
            "byte",
            UnitDef {
                dim: Dimension::Information,
                scale: 8.0,
                offset: 0.0,
                prefix: PrefixSupport::Binary,
            },
        );

        // Magnetism (base: tesla).
        units.insert(
            "T",
            UnitDef {
                dim: Dimension::MagneticFluxDensity,
                scale: 1.0,
                offset: 0.0,
                prefix: PrefixSupport::Metric,
            },
        );
        units.insert(
            "ga",
            UnitDef {
                dim: Dimension::MagneticFluxDensity,
                // 1 gauss = 1e-4 tesla.
                scale: 1e-4,
                offset: 0.0,
                prefix: PrefixSupport::None,
            },
        );

        units
    })
}
