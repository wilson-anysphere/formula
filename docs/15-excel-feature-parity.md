# Excel Feature Parity Checklist

## Overview

This document tracks every Excel feature and our implementation status. Features are categorized by priority:

- **P0**: Required for launch (breaks compatibility if missing)
- **P1**: Required for power users (significant gap if missing)
- **P2**: Nice to have (can add post-launch)
- **P3**: Rare/legacy (low priority)

---

## Formula Functions

### Math & Trigonometry (P0)

| Function | Status | Notes |
|----------|--------|-------|
| ABS | ⬜ | |
| ACOS | ⬜ | |
| ACOSH | ⬜ | |
| ACOT | ⬜ | |
| ACOTH | ⬜ | |
| AGGREGATE | ⬜ | Complex, 19 subtypes |
| ARABIC | ⬜ | |
| ASIN | ⬜ | |
| ASINH | ⬜ | |
| ATAN | ⬜ | |
| ATAN2 | ⬜ | |
| ATANH | ⬜ | |
| BASE | ⬜ | |
| CEILING | ⬜ | Multiple modes |
| CEILING.MATH | ⬜ | |
| CEILING.PRECISE | ⬜ | |
| COMBIN | ⬜ | |
| COMBINA | ⬜ | |
| COS | ⬜ | |
| COSH | ⬜ | |
| COT | ⬜ | |
| COTH | ⬜ | |
| CSC | ⬜ | |
| CSCH | ⬜ | |
| DECIMAL | ⬜ | |
| DEGREES | ⬜ | |
| EVEN | ⬜ | |
| EXP | ⬜ | |
| FACT | ⬜ | |
| FACTDOUBLE | ⬜ | |
| FLOOR | ⬜ | Multiple modes |
| FLOOR.MATH | ⬜ | |
| FLOOR.PRECISE | ⬜ | |
| GCD | ⬜ | |
| INT | ⬜ | |
| ISO.CEILING | ⬜ | |
| LCM | ⬜ | |
| LET | ⬜ | Important for readability |
| LN | ⬜ | |
| LOG | ⬜ | |
| LOG10 | ⬜ | |
| MDETERM | ⬜ | |
| MINVERSE | ⬜ | |
| MMULT | ⬜ | |
| MOD | ⬜ | |
| MROUND | ⬜ | |
| MULTINOMIAL | ⬜ | |
| MUNIT | ⬜ | |
| ODD | ⬜ | |
| PI | ⬜ | |
| POWER | ⬜ | |
| PRODUCT | ⬜ | |
| QUOTIENT | ⬜ | |
| RADIANS | ⬜ | |
| RAND | ⬜ | Volatile |
| RANDBETWEEN | ⬜ | Volatile |
| RANDARRAY | ⬜ | Dynamic array |
| ROMAN | ⬜ | |
| ROUND | ⬜ | |
| ROUNDDOWN | ⬜ | |
| ROUNDUP | ⬜ | |
| SEC | ⬜ | |
| SECH | ⬜ | |
| SEQUENCE | ⬜ | Dynamic array |
| SERIESSUM | ⬜ | |
| SIGN | ⬜ | |
| SIN | ⬜ | |
| SINH | ⬜ | |
| SQRT | ⬜ | |
| SQRTPI | ⬜ | |
| SUBTOTAL | ⬜ | 11 function types |
| SUM | ✅ | Core function |
| SUMIF | ⬜ | |
| SUMIFS | ⬜ | |
| SUMPRODUCT | ⬜ | |
| SUMSQ | ✅ | |
| SUMX2MY2 | ⬜ | |
| SUMX2PY2 | ⬜ | |
| SUMXMY2 | ⬜ | |
| TAN | ⬜ | |
| TANH | ⬜ | |
| TRUNC | ⬜ | |

### Lookup & Reference (P0)

| Function | Status | Notes |
|----------|--------|-------|
| ADDRESS | ⬜ | |
| AREAS | ⬜ | |
| CHOOSE | ✅ | Supports arrays and range unions |
| CHOOSECOLS | ✅ | Dynamic array |
| CHOOSEROWS | ✅ | Dynamic array |
| COLUMN | ⬜ | |
| COLUMNS | ⬜ | |
| DROP | ✅ | Dynamic array |
| EXPAND | ✅ | Dynamic array |
| FILTER | ⬜ | Dynamic array, critical |
| FORMULATEXT | ⬜ | |
| GETPIVOTDATA | ✅ | MVP: supports tabular pivot outputs (limited layouts) |
| HLOOKUP | ✅ | Includes wildcard exact match + approximate modes |
| HSTACK | ⬜ | Dynamic array |
| HYPERLINK | ⬜ | |
| INDEX | ✅ | |
| INDIRECT | ⬜ | Volatile, complex |
| LOOKUP | ⬜ | Legacy |
| MATCH | ✅ | Includes wildcard exact match + approximate modes |
| OFFSET | ⬜ | Volatile |
| ROW | ⬜ | |
| ROWS | ⬜ | |
| RTD | ⬜ | Real-time data |
| SORT | ⬜ | Dynamic array |
| SORTBY | ⬜ | Dynamic array |
| TAKE | ✅ | Dynamic array |
| TOCOL | ⬜ | Dynamic array |
| TOROW | ⬜ | Dynamic array |
| TRANSPOSE | ⬜ | |
| UNIQUE | ⬜ | Dynamic array, critical |
| VLOOKUP | ✅ | Includes wildcard exact match + approximate modes |
| VSTACK | ⬜ | Dynamic array |
| WRAPCOLS | ⬜ | Dynamic array |
| WRAPROWS | ⬜ | Dynamic array |
| XLOOKUP | ✅ | Supports match_mode/search_mode + 2D return_array spilling |
| XMATCH | ✅ | Supports match_mode/search_mode + binary search modes |

### Text Functions (P0)

| Function | Status | Notes |
|----------|--------|-------|
| ASC | ⬜ | |
| BAHTTEXT | ⬜ | |
| CHAR | ⬜ | |
| CLEAN | ⬜ | |
| CODE | ⬜ | |
| CONCAT | ⬜ | |
| CONCATENATE | ⬜ | Legacy, still used |
| DBCS | ⬜ | |
| DOLLAR | ⬜ | |
| EXACT | ⬜ | |
| FIND | ⬜ | Case-sensitive |
| FINDB | ⬜ | Byte-based |
| FIXED | ⬜ | |
| LEFT | ⬜ | |
| LEFTB | ⬜ | Byte-based |
| LEN | ⬜ | |
| LENB | ⬜ | Byte-based |
| LOWER | ⬜ | |
| MID | ⬜ | |
| MIDB | ⬜ | Byte-based |
| NUMBERVALUE | ⬜ | |
| PHONETIC | ⬜ | |
| PROPER | ⬜ | |
| REPLACE | ⬜ | |
| REPLACEB | ⬜ | Byte-based |
| REPT | ⬜ | |
| RIGHT | ⬜ | |
| RIGHTB | ⬜ | Byte-based |
| SEARCH | ⬜ | Case-insensitive |
| SEARCHB | ⬜ | Byte-based |
| SUBSTITUTE | ⬜ | |
| T | ⬜ | |
| TEXT | ⬜ | Complex formatting |
| TEXTAFTER | ⬜ | |
| TEXTBEFORE | ⬜ | |
| TEXTJOIN | ⬜ | |
| TEXTSPLIT | ⬜ | Dynamic array |
| TRIM | ⬜ | |
| UNICHAR | ⬜ | |
| UNICODE | ⬜ | |
| UPPER | ⬜ | |
| VALUE | ⬜ | |
| VALUETOTEXT | ⬜ | |

### Logical Functions (P0)

| Function | Status | Notes |
|----------|--------|-------|
| AND | ⬜ | |
| BYCOL | ⬜ | Lambda helper |
| BYROW | ⬜ | Lambda helper |
| FALSE | ⬜ | |
| IF | ✅ | Most used function |
| IFERROR | ✅ | |
| IFNA | ⬜ | |
| IFS | ⬜ | Multiple conditions |
| LAMBDA | ⬜ | User-defined functions |
| LET | ⬜ | Variable binding |
| MAKEARRAY | ⬜ | Lambda helper |
| MAP | ⬜ | Lambda helper |
| NOT | ⬜ | |
| OR | ⬜ | |
| REDUCE | ⬜ | Lambda helper |
| SCAN | ⬜ | Lambda helper |
| SWITCH | ⬜ | |
| TRUE | ⬜ | |
| XOR | ⬜ | |

### Date & Time (P0)

| Function | Status | Notes |
|----------|--------|-------|
| DATE | ⬜ | |
| DATEDIF | ⬜ | Undocumented but used |
| DATEVALUE | ⬜ | |
| DAY | ⬜ | |
| DAYS | ⬜ | |
| DAYS360 | ⬜ | |
| EDATE | ⬜ | |
| EOMONTH | ⬜ | |
| HOUR | ⬜ | |
| ISOWEEKNUM | ⬜ | |
| MINUTE | ⬜ | |
| MONTH | ⬜ | |
| NETWORKDAYS | ⬜ | |
| NETWORKDAYS.INTL | ⬜ | |
| NOW | ⬜ | Volatile |
| SECOND | ⬜ | |
| TIME | ⬜ | |
| TIMEVALUE | ⬜ | |
| TODAY | ⬜ | Volatile |
| WEEKDAY | ⬜ | |
| WEEKNUM | ⬜ | |
| WORKDAY | ⬜ | |
| WORKDAY.INTL | ⬜ | |
| YEAR | ⬜ | |
| YEARFRAC | ⬜ | |

### Statistical Functions (P1)

| Function | Status | Notes |
|----------|--------|-------|
| AVEDEV | ✅ | |
| AVERAGE | ✅ | P0 |
| AVERAGEA | ✅ | |
| AVERAGEIF | ✅ | P0 |
| AVERAGEIFS | ✅ | P0 |
| BETA.DIST | ⬜ | |
| BETA.INV | ⬜ | |
| BINOM.DIST | ⬜ | |
| BINOM.DIST.RANGE | ⬜ | |
| BINOM.INV | ⬜ | |
| CHISQ.DIST | ⬜ | |
| CHISQ.DIST.RT | ⬜ | |
| CHISQ.INV | ⬜ | |
| CHISQ.INV.RT | ⬜ | |
| CHISQ.TEST | ⬜ | |
| CONFIDENCE.NORM | ⬜ | |
| CONFIDENCE.T | ⬜ | |
| CORREL | ✅ | |
| COUNT | ✅ | P0 |
| COUNTA | ✅ | P0 |
| COUNTBLANK | ✅ | P0 |
| COUNTIF | ✅ | P0 |
| COUNTIFS | ✅ | P0 |
| COVARIANCE.P | ✅ | |
| COVARIANCE.S | ✅ | |
| DEVSQ | ✅ | |
| EXPON.DIST | ⬜ | |
| F.DIST | ⬜ | |
| F.DIST.RT | ⬜ | |
| F.INV | ⬜ | |
| F.INV.RT | ⬜ | |
| F.TEST | ⬜ | |
| FISHER | ⬜ | |
| FISHERINV | ⬜ | |
| FORECAST | ✅ | |
| FORECAST.ETS | ⬜ | |
| FORECAST.ETS.CONFINT | ⬜ | |
| FORECAST.ETS.SEASONALITY | ⬜ | |
| FORECAST.ETS.STAT | ⬜ | |
| FORECAST.LINEAR | ✅ | |
| FREQUENCY | ⬜ | Array function |
| GAMMA | ⬜ | |
| GAMMA.DIST | ⬜ | |
| GAMMA.INV | ⬜ | |
| GAMMALN | ⬜ | |
| GAMMALN.PRECISE | ⬜ | |
| GAUSS | ⬜ | |
| GEOMEAN | ✅ | |
| GROWTH | ⬜ | Array function |
| HARMEAN | ✅ | |
| HYPGEOM.DIST | ⬜ | |
| INTERCEPT | ✅ | |
| KURT | ⬜ | |
| LARGE | ✅ | P0 |
| LINEST | ⬜ | Array function |
| LOGEST | ⬜ | Array function |
| LOGNORM.DIST | ⬜ | |
| LOGNORM.INV | ⬜ | |
| MAX | ✅ | P0 |
| MAXA | ✅ | |
| MAXIFS | ✅ | P0 |
| MEDIAN | ✅ | P0 |
| MIN | ✅ | P0 |
| MINA | ✅ | |
| MINIFS | ✅ | P0 |
| MODE.MULT | ✅ | Array function |
| MODE.SNGL | ✅ | |
| NEGBINOM.DIST | ⬜ | |
| NORM.DIST | ⬜ | |
| NORM.INV | ⬜ | |
| NORM.S.DIST | ⬜ | |
| NORM.S.INV | ⬜ | |
| PEARSON | ✅ | |
| PERCENTILE.EXC | ✅ | |
| PERCENTILE.INC | ✅ | |
| PERCENTRANK.EXC | ✅ | |
| PERCENTRANK.INC | ✅ | |
| PERMUT | ⬜ | |
| PERMUTATIONA | ⬜ | |
| PHI | ⬜ | |
| POISSON.DIST | ⬜ | |
| PROB | ⬜ | |
| QUARTILE.EXC | ✅ | |
| QUARTILE.INC | ✅ | |
| RANK.AVG | ✅ | |
| RANK.EQ | ✅ | |
| RSQ | ✅ | |
| SKEW | ⬜ | |
| SKEW.P | ⬜ | |
| SLOPE | ✅ | |
| SMALL | ✅ | P0 |
| STANDARDIZE | ✅ | |
| STDEV.P | ✅ | |
| STDEV.S | ✅ | |
| STDEVA | ✅ | |
| STDEVPA | ✅ | |
| STEYX | ✅ | |
| T.DIST | ⬜ | |
| T.DIST.2T | ⬜ | |
| T.DIST.RT | ⬜ | |
| T.INV | ⬜ | |
| T.INV.2T | ⬜ | |
| T.TEST | ⬜ | |
| TREND | ⬜ | Array function |
| TRIMMEAN | ✅ | |
| VAR.P | ✅ | |
| VAR.S | ✅ | |
| VARA | ✅ | |
| VARPA | ✅ | |
| WEIBULL.DIST | ⬜ | |
| Z.TEST | ⬜ | |

### Financial Functions (P1)

| Function | Status | Notes |
|----------|--------|-------|
| ACCRINT | ⬜ | |
| ACCRINTM | ⬜ | |
| AMORDEGRC | ⬜ | |
| AMORLINC | ⬜ | |
| COUPDAYBS | ⬜ | |
| COUPDAYS | ⬜ | |
| COUPDAYSNC | ⬜ | |
| COUPNCD | ⬜ | |
| COUPNUM | ⬜ | |
| COUPPCD | ⬜ | |
| CUMIPMT | ⬜ | |
| CUMPRINC | ⬜ | |
| DB | ⬜ | |
| DDB | ✅ | |
| DISC | ⬜ | |
| DOLLARDE | ⬜ | |
| DOLLARFR | ⬜ | |
| DURATION | ⬜ | |
| EFFECT | ⬜ | |
| FV | ✅ | P0 |
| FVSCHEDULE | ⬜ | |
| INTRATE | ⬜ | |
| IPMT | ✅ | |
| IRR | ✅ | P0 |
| ISPMT | ⬜ | |
| MDURATION | ⬜ | |
| MIRR | ✅ | |
| NOMINAL | ⬜ | |
| NPER | ✅ | |
| NPV | ✅ | P0 |
| ODDFPRICE | ⬜ | |
| ODDFYIELD | ⬜ | |
| ODDLPRICE | ⬜ | |
| ODDLYIELD | ⬜ | |
| PDURATION | ⬜ | |
| PMT | ✅ | P0 |
| PPMT | ✅ | |
| PRICE | ⬜ | |
| PRICEDISC | ⬜ | |
| PRICEMAT | ⬜ | |
| PV | ✅ | P0 |
| RATE | ✅ | |
| RECEIVED | ⬜ | |
| RRI | ⬜ | |
| SLN | ✅ | |
| SYD | ✅ | |
| TBILLEQ | ⬜ | |
| TBILLPRICE | ⬜ | |
| TBILLYIELD | ⬜ | |
| VDB | ⬜ | |
| XIRR | ✅ | P0 |
| XNPV | ✅ | P0 |
| YIELD | ⬜ | |
| YIELDDISC | ⬜ | |
| YIELDMAT | ⬜ | |

---

## File Format Features

### Worksheets (P0)

| Feature | Status | Notes |
|---------|--------|-------|
| Multiple sheets | ⬜ | |
| Sheet naming | ⬜ | |
| Sheet ordering | ⬜ | |
| Sheet color tabs | ⬜ | |
| Hidden sheets | ⬜ | |
| Very hidden sheets | ⬜ | |
| Sheet protection | ⬜ | |

### Cell Content (P0)

| Feature | Status | Notes |
|---------|--------|-------|
| Numbers | ⬜ | |
| Text | ⬜ | |
| Formulas | ⬜ | |
| Dates | ⬜ | |
| Times | ⬜ | |
| Booleans | ⬜ | |
| Errors | ⬜ | |
| Rich text | ⬜ | |
| Hyperlinks | ⬜ | |
| Comments | ⬜ | |
| Notes | ⬜ | |
| Images in cells | ⬜ | |

### Formatting (P0)

| Feature | Status | Notes |
|---------|--------|-------|
| Font family | ⬜ | |
| Font size | ⬜ | |
| Font color | ⬜ | |
| Bold | ⬜ | |
| Italic | ⬜ | |
| Underline | ⬜ | |
| Strikethrough | ⬜ | |
| Background color | ⬜ | |
| Number formats | ⬜ | Complex |
| Date formats | ⬜ | |
| Custom formats | ⬜ | |
| Borders | ⬜ | |
| Alignment | ⬜ | |
| Text wrap | ⬜ | |
| Merge cells | ⬜ | |
| Cell rotation | ⬜ | |

### Conditional Formatting (P0)

| Feature | Status | Notes |
|---------|--------|-------|
| Cell value rules | ⬜ | |
| Formula rules | ⬜ | |
| Data bars | ⬜ | |
| Color scales | ⬜ | |
| Icon sets | ⬜ | |
| Top/bottom rules | ⬜ | |
| Unique/duplicate | ⬜ | |

### Data Features (P0)

| Feature | Status | Notes |
|---------|--------|-------|
| Tables | ⬜ | |
| Structured references | ⬜ | |
| Data validation | ⬜ | |
| Dropdown lists | ⬜ | |
| Named ranges | ⬜ | |
| AutoFilter | ⬜ | |
| Sort | ⬜ | |
| Group/outline | ⬜ | |

### Charts (P1)

| Feature | Status | Notes |
|---------|--------|-------|
| Column charts (clustered/stacked/100% stacked) | ⬜ | DrawingML `c:barChart` + `c:grouping` |
| Bar charts (clustered/stacked/100% stacked) | ⬜ | DrawingML `c:barChart` + `c:barDir="bar"` |
| Line charts | ⬜ | |
| Pie charts | ⬜ | |
| Area charts (stacked/100% stacked) | ⬜ | |
| Scatter plots | ⬜ | |
| Bubble charts | ⬜ | |
| Radar charts | ⬜ | |
| Waterfall charts | ⬜ | Often ChartEx (Excel 2016+) |
| Histogram charts | ⬜ | Often ChartEx (Excel 2016+) |
| Pareto charts | ⬜ | Often ChartEx (Excel 2016+) |
| Box & whisker charts | ⬜ | Often ChartEx (Excel 2016+) |
| Treemap charts | ⬜ | Often ChartEx (Excel 2016+) |
| Sunburst charts | ⬜ | Often ChartEx (Excel 2016+) |
| Funnel charts | ⬜ | Often ChartEx (Excel 2016+) |
| Stock charts (OHLC) | ⬜ | DrawingML `c:stockChart` |
| Combo charts | ⬜ | |
| Map charts (preserve/placeholder) | ⬜ | Preserve XML and show placeholder if unsupported |
| Sparklines | ⬜ | |
| Axis formatting fidelity | ⬜ | Ticks, gridlines, number formats, scaling |
| Titles, legends, and data labels | ⬜ | Rich text + layout |
| Theme-based colors | ⬜ | Depends on theme fidelity (Task 109) |
| Series formatting + markers | ⬜ | Line/fill/marker, per-point overrides |
| Layout anchored to sheet (EMUs) | ⬜ | Drawing anchors must match Excel |
| Lossless round-trip (unknown chart types/parts) | ⬜ | Preserve chart-related parts byte-for-byte when unedited |

### Pivot Tables (P1)

| Feature | Status | Notes |
|---------|--------|-------|
| Basic pivot | ⬜ | |
| Multiple value fields | ⬜ | |
| Calculated fields | ⬜ | |
| Pivot charts | ⬜ | |
| Slicers | ⬜ | |
| Timelines | ⬜ | |

### Advanced (P2)

| Feature | Status | Notes |
|---------|--------|-------|
| Power Query | ⬜ | |
| Data Model | ⬜ | |
| Power Pivot | ⬜ | |
| Cube functions | ⬜ | |
| Solver | ⬜ | |
| Goal Seek | ⬜ | |
| Scenarios | ⬜ | |

### VBA (P2)

| Feature | Status | Notes |
|---------|--------|-------|
| Preserve vbaProject.bin | ⬜ | P0 actually |
| Parse VBA code | ⬜ | |
| Display VBA | ⬜ | |
| Execute VBA | ⬜ | |

---

## Progress Summary

| Category | Total | Done | % |
|----------|-------|------|---|
| Math Functions | 65 | 0 | 0% |
| Lookup Functions | 35 | 0 | 0% |
| Text Functions | 50 | 0 | 0% |
| Logical Functions | 20 | 0 | 0% |
| Date Functions | 25 | 0 | 0% |
| Statistical Functions | 100 | 0 | 0% |
| Financial Functions | 55 | 0 | 0% |
| File Features | 100+ | 0 | 0% |
| **Total** | **~500** | **0** | **0%** |

---

## Implementation Order

### Phase 1: Core Functions (Months 1-3)
1. Basic math: SUM, AVERAGE, MIN, MAX, COUNT
2. Lookup: VLOOKUP, INDEX, MATCH, XLOOKUP
3. Logical: IF, AND, OR, IFERROR
4. Text: LEFT, RIGHT, MID, CONCATENATE, TRIM
5. Date: DATE, TODAY, NOW, YEAR, MONTH, DAY

### Phase 2: Extended Functions (Months 3-6)
1. All remaining P0 functions
2. Dynamic array functions (FILTER, SORT, UNIQUE)
3. LAMBDA and helper functions
4. Statistical functions
5. Financial functions

### Phase 3: Full Parity (Months 6-12)
1. P1 functions
2. P2 functions
3. Edge cases and compatibility fixes
4. Engineering functions
5. Cube functions
