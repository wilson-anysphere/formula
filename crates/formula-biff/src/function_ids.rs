#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FunctionSpec {
    pub id: u16,
    pub name: &'static str,
    pub min_args: u8,
    pub max_args: u8,
}

// NOTE: Function IDs are BIFF built-in function indices (the `iftab` values used by
// `PtgFunc`/`PtgFuncVar`). These are shared across BIFF8/BIFF12 for "classic" Excel
// functions.
//
// The full BIFF12 id <-> name mapping lives in [`crate::ftab`]. This module stores
// argument-count metadata for *worksheet* functions so we can:
// - decode `PtgFunc` (fixed-arity calls where argc is implicit)
// - encode functions by validating argument counts and selecting `PtgFunc` vs
//   `PtgFuncVar`
//
// Many FTAB entries correspond to legacy Excel 4.0 (XLM) macro sheet functions or
// UI/command helpers. Those are intentionally left as `None` until we need them.
pub(crate) const FTAB_ARG_RANGES: [Option<(u8, u8)>; 485] = [
    Some((0, 255)), //   0 COUNT
    Some((2, 3)),   //   1 IF
    Some((1, 1)),   //   2 ISNA
    Some((1, 1)),   //   3 ISERROR
    Some((0, 255)), //   4 SUM
    Some((1, 255)), //   5 AVERAGE
    Some((1, 255)), //   6 MIN
    Some((1, 255)), //   7 MAX
    Some((0, 1)),   //   8 ROW
    Some((0, 1)),   //   9 COLUMN
    Some((0, 0)),   //  10 NA
    Some((2, 255)), //  11 NPV
    Some((1, 255)), //  12 STDEV
    Some((1, 2)),   //  13 DOLLAR
    Some((1, 3)),   //  14 FIXED
    Some((1, 1)),   //  15 SIN
    Some((1, 1)),   //  16 COS
    Some((1, 1)),   //  17 TAN
    Some((1, 1)),   //  18 ATAN
    Some((0, 0)),   //  19 PI
    Some((1, 1)),   //  20 SQRT
    Some((1, 1)),   //  21 EXP
    Some((1, 1)),   //  22 LN
    Some((1, 1)),   //  23 LOG10
    Some((1, 1)),   //  24 ABS
    Some((1, 1)),   //  25 INT
    Some((1, 1)),   //  26 SIGN
    Some((2, 2)),   //  27 ROUND
    Some((2, 3)),   //  28 LOOKUP
    Some((2, 4)),   //  29 INDEX
    Some((2, 2)),   //  30 REPT
    Some((3, 3)),   //  31 MID
    Some((1, 1)),   //  32 LEN
    Some((1, 1)),   //  33 VALUE
    Some((0, 0)),   //  34 TRUE
    Some((0, 0)),   //  35 FALSE
    Some((1, 255)), //  36 AND
    Some((1, 255)), //  37 OR
    Some((1, 1)),   //  38 NOT
    Some((2, 2)),   //  39 MOD
    Some((3, 3)),   //  40 DCOUNT
    Some((3, 3)),   //  41 DSUM
    Some((3, 3)),   //  42 DAVERAGE
    Some((3, 3)),   //  43 DMIN
    Some((3, 3)),   //  44 DMAX
    Some((3, 3)),   //  45 DSTDEV
    Some((1, 255)), //  46 VAR
    Some((3, 3)),   //  47 DVAR
    Some((2, 2)),   //  48 TEXT
    Some((1, 4)),   //  49 LINEST
    Some((1, 4)),   //  50 TREND
    Some((1, 4)),   //  51 LOGEST
    Some((1, 4)),   //  52 GROWTH
    None,           //  53 GOTO
    None,           //  54 HALT
    None,           //  55 RETURN
    Some((3, 5)),   //  56 PV
    Some((3, 5)),   //  57 FV
    Some((3, 5)),   //  58 NPER
    Some((3, 5)),   //  59 PMT
    Some((3, 6)),   //  60 RATE
    Some((3, 3)),   //  61 MIRR
    Some((1, 2)),   //  62 IRR
    Some((0, 0)),   //  63 RAND
    Some((2, 3)),   //  64 MATCH
    Some((3, 3)),   //  65 DATE
    Some((3, 3)),   //  66 TIME
    Some((1, 1)),   //  67 DAY
    Some((1, 1)),   //  68 MONTH
    Some((1, 1)),   //  69 YEAR
    Some((1, 2)),   //  70 WEEKDAY
    Some((1, 1)),   //  71 HOUR
    Some((1, 1)),   //  72 MINUTE
    Some((1, 1)),   //  73 SECOND
    Some((0, 0)),   //  74 NOW
    Some((1, 1)),   //  75 AREAS
    Some((1, 1)),   //  76 ROWS
    Some((1, 1)),   //  77 COLUMNS
    Some((3, 5)),   //  78 OFFSET
    None,           //  79 ABSREF
    None,           //  80 RELREF
    None,           //  81 ARGUMENT
    Some((2, 3)),   //  82 SEARCH
    Some((1, 1)),   //  83 TRANSPOSE
    None,           //  84 ERROR
    None,           //  85 STEP
    Some((1, 1)),   //  86 TYPE
    None,           //  87 ECHO
    None,           //  88 SET.NAME
    Some((0, 0)),   //  89 CALLER
    None,           //  90 DEREF
    None,           //  91 WINDOWS
    Some((4, 4)),   //  92 SERIES
    None,           //  93 DOCUMENTS
    None,           //  94 ACTIVE.CELL
    None,           //  95 SELECTION
    None,           //  96 RESULT
    Some((2, 2)),   //  97 ATAN2
    Some((1, 1)),   //  98 ASIN
    Some((1, 1)),   //  99 ACOS
    Some((2, 255)), // 100 CHOOSE
    Some((3, 4)),   // 101 HLOOKUP
    Some((3, 4)),   // 102 VLOOKUP
    None,           // 103 LINKS
    None,           // 104 INPUT
    Some((1, 1)),   // 105 ISREF
    None,           // 106 GET.FORMULA
    None,           // 107 GET.NAME
    None,           // 108 SET.VALUE
    Some((1, 2)),   // 109 LOG
    None,           // 110 EXEC
    Some((1, 1)),   // 111 CHAR
    Some((1, 1)),   // 112 LOWER
    Some((1, 1)),   // 113 UPPER
    Some((1, 1)),   // 114 PROPER
    Some((1, 2)),   // 115 LEFT
    Some((1, 2)),   // 116 RIGHT
    Some((2, 2)),   // 117 EXACT
    Some((1, 1)),   // 118 TRIM
    Some((4, 4)),   // 119 REPLACE
    Some((3, 4)),   // 120 SUBSTITUTE
    Some((1, 1)),   // 121 CODE
    None,           // 122 NAMES
    None,           // 123 DIRECTORY
    Some((2, 3)),   // 124 FIND
    Some((1, 2)),   // 125 CELL
    Some((1, 1)),   // 126 ISERR
    Some((1, 1)),   // 127 ISTEXT
    Some((1, 1)),   // 128 ISNUMBER
    Some((1, 1)),   // 129 ISBLANK
    Some((1, 1)),   // 130 T
    Some((1, 1)),   // 131 N
    None,           // 132 FOPEN
    None,           // 133 FCLOSE
    None,           // 134 FSIZE
    None,           // 135 FREADLN
    None,           // 136 FREAD
    None,           // 137 FWRITELN
    None,           // 138 FWRITE
    None,           // 139 FPOS
    Some((1, 1)),   // 140 DATEVALUE
    Some((1, 1)),   // 141 TIMEVALUE
    Some((3, 3)),   // 142 SLN
    Some((4, 4)),   // 143 SYD
    Some((4, 5)),   // 144 DDB
    None,           // 145 GET.DEF
    None,           // 146 REFTEXT
    None,           // 147 TEXTREF
    Some((1, 2)),   // 148 INDIRECT
    None,           // 149 REGISTER
    None,           // 150 CALL
    None,           // 151 ADD.BAR
    None,           // 152 ADD.MENU
    None,           // 153 ADD.COMMAND
    None,           // 154 ENABLE.COMMAND
    None,           // 155 CHECK.COMMAND
    None,           // 156 RENAME.COMMAND
    None,           // 157 SHOW.BAR
    None,           // 158 DELETE.MENU
    None,           // 159 DELETE.COMMAND
    None,           // 160 GET.CHART.ITEM
    None,           // 161 DIALOG.BOX
    Some((1, 1)),   // 162 CLEAN
    Some((1, 1)),   // 163 MDETERM
    Some((1, 1)),   // 164 MINVERSE
    Some((2, 2)),   // 165 MMULT
    None,           // 166 FILES
    Some((4, 6)),   // 167 IPMT
    Some((4, 6)),   // 168 PPMT
    Some((0, 255)), // 169 COUNTA
    None,           // 170 CANCEL.KEY
    None,           // 171 FOR
    None,           // 172 WHILE
    None,           // 173 BREAK
    None,           // 174 NEXT
    None,           // 175 INITIATE
    None,           // 176 REQUEST
    None,           // 177 POKE
    None,           // 178 EXECUTE
    None,           // 179 TERMINATE
    None,           // 180 RESTART
    None,           // 181 HELP
    None,           // 182 GET.BAR
    Some((1, 255)), // 183 PRODUCT
    Some((1, 1)),   // 184 FACT
    Some((2, 2)),   // 185 GET.CELL
    Some((1, 1)),   // 186 GET.WORKSPACE
    Some((1, 1)),   // 187 GET.WINDOW
    Some((1, 1)),   // 188 GET.DOCUMENT
    Some((3, 3)),   // 189 DPRODUCT
    Some((1, 1)),   // 190 ISNONTEXT
    None,           // 191 GET.NOTE
    None,           // 192 NOTE
    Some((1, 255)), // 193 STDEVP
    Some((1, 255)), // 194 VARP
    Some((3, 3)),   // 195 DSTDEVP
    Some((3, 3)),   // 196 DVARP
    Some((1, 2)),   // 197 TRUNC
    Some((1, 1)),   // 198 ISLOGICAL
    Some((3, 3)),   // 199 DCOUNTA
    None,           // 200 DELETE.BAR
    None,           // 201 UNREGISTER
    None,           // 202 <reserved>
    None,           // 203 <reserved>
    Some((1, 2)),   // 204 USDOLLAR
    Some((2, 3)),   // 205 FINDB
    Some((2, 3)),   // 206 SEARCHB
    Some((4, 4)),   // 207 REPLACEB
    Some((1, 2)),   // 208 LEFTB
    Some((1, 2)),   // 209 RIGHTB
    Some((3, 3)),   // 210 MIDB
    Some((1, 1)),   // 211 LENB
    Some((2, 2)),   // 212 ROUNDUP
    Some((2, 2)),   // 213 ROUNDDOWN
    Some((1, 1)),   // 214 ASC
    Some((1, 1)),   // 215 DBCS
    Some((2, 3)),   // 216 RANK
    None,           // 217 <reserved>
    None,           // 218 <reserved>
    Some((2, 5)),   // 219 ADDRESS
    Some((2, 3)),   // 220 DAYS360
    Some((0, 0)),   // 221 TODAY
    Some((5, 7)),   // 222 VDB
    None,           // 223 ELSE
    None,           // 224 ELSE.IF
    None,           // 225 END.IF
    None,           // 226 FOR.CELL
    Some((1, 255)), // 227 MEDIAN
    Some((1, 255)), // 228 SUMPRODUCT
    Some((1, 1)),   // 229 SINH
    Some((1, 1)),   // 230 COSH
    Some((1, 1)),   // 231 TANH
    Some((1, 1)),   // 232 ASINH
    Some((1, 1)),   // 233 ACOSH
    Some((1, 1)),   // 234 ATANH
    Some((3, 3)),   // 235 DGET
    None,           // 236 CREATE.OBJECT
    None,           // 237 VOLATILE
    None,           // 238 LAST.ERROR
    None,           // 239 CUSTOM.UNDO
    None,           // 240 CUSTOM.REPEAT
    None,           // 241 FORMULA.CONVERT
    None,           // 242 GET.LINK.INFO
    None,           // 243 TEXT.BOX
    Some((1, 1)),   // 244 INFO
    None,           // 245 GROUP
    None,           // 246 GET.OBJECT
    Some((4, 5)),   // 247 DB
    None,           // 248 PAUSE
    None,           // 249 <reserved>
    None,           // 250 <reserved>
    None,           // 251 RESUME
    Some((2, 2)),   // 252 FREQUENCY
    None,           // 253 ADD.TOOLBAR
    None,           // 254 DELETE.TOOLBAR
    None,           // 255 USER
    None,           // 256 RESET.TOOLBAR
    Some((1, 1)),   // 257 EVALUATE
    None,           // 258 GET.TOOLBAR
    None,           // 259 GET.TOOL
    None,           // 260 SPELLING.CHECK
    Some((1, 1)),   // 261 ERROR.TYPE
    None,           // 262 APP.TITLE
    None,           // 263 WINDOW.TITLE
    None,           // 264 SAVE.TOOLBAR
    None,           // 265 ENABLE.TOOL
    None,           // 266 PRESS.TOOL
    None,           // 267 REGISTER.ID
    Some((1, 1)),   // 268 GET.WORKBOOK
    Some((1, 255)), // 269 AVEDEV
    Some((3, 5)),   // 270 BETADIST
    Some((1, 1)),   // 271 GAMMALN
    Some((3, 5)),   // 272 BETAINV
    Some((4, 4)),   // 273 BINOMDIST
    Some((2, 2)),   // 274 CHIDIST
    Some((2, 2)),   // 275 CHIINV
    Some((2, 2)),   // 276 COMBIN
    Some((3, 3)),   // 277 CONFIDENCE
    Some((3, 3)),   // 278 CRITBINOM
    Some((1, 1)),   // 279 EVEN
    Some((3, 3)),   // 280 EXPONDIST
    Some((3, 3)),   // 281 FDIST
    Some((3, 3)),   // 282 FINV
    Some((1, 1)),   // 283 FISHER
    Some((1, 1)),   // 284 FISHERINV
    Some((2, 2)),   // 285 FLOOR
    Some((4, 4)),   // 286 GAMMADIST
    Some((3, 3)),   // 287 GAMMAINV
    Some((2, 2)),   // 288 CEILING
    Some((4, 4)),   // 289 HYPGEOMDIST
    Some((3, 3)),   // 290 LOGNORMDIST
    Some((3, 3)),   // 291 LOGINV
    Some((3, 3)),   // 292 NEGBINOMDIST
    Some((4, 4)),   // 293 NORMDIST
    Some((1, 1)),   // 294 NORMSDIST
    Some((3, 3)),   // 295 NORMINV
    Some((1, 1)),   // 296 NORMSINV
    Some((3, 3)),   // 297 STANDARDIZE
    Some((1, 1)),   // 298 ODD
    Some((2, 2)),   // 299 PERMUT
    Some((3, 3)),   // 300 POISSON
    Some((3, 3)),   // 301 TDIST
    Some((4, 4)),   // 302 WEIBULL
    Some((2, 2)),   // 303 SUMXMY2
    Some((2, 2)),   // 304 SUMX2MY2
    Some((2, 2)),   // 305 SUMX2PY2
    Some((2, 2)),   // 306 CHITEST
    Some((2, 2)),   // 307 CORREL
    Some((2, 2)),   // 308 COVAR
    Some((3, 3)),   // 309 FORECAST
    Some((2, 2)),   // 310 FTEST
    Some((2, 2)),   // 311 INTERCEPT
    Some((2, 2)),   // 312 PEARSON
    Some((2, 2)),   // 313 RSQ
    Some((2, 2)),   // 314 STEYX
    Some((2, 2)),   // 315 SLOPE
    Some((4, 4)),   // 316 TTEST
    Some((3, 4)),   // 317 PROB
    Some((1, 255)), // 318 DEVSQ
    Some((1, 255)), // 319 GEOMEAN
    Some((1, 255)), // 320 HARMEAN
    Some((1, 255)), // 321 SUMSQ
    Some((1, 255)), // 322 KURT
    Some((1, 255)), // 323 SKEW
    Some((2, 3)),   // 324 ZTEST
    Some((2, 2)),   // 325 LARGE
    Some((2, 2)),   // 326 SMALL
    Some((2, 2)),   // 327 QUARTILE
    Some((2, 2)),   // 328 PERCENTILE
    Some((2, 3)),   // 329 PERCENTRANK
    Some((1, 255)), // 330 MODE
    Some((2, 2)),   // 331 TRIMMEAN
    Some((2, 2)),   // 332 TINV
    None,           // 333 <reserved>
    None,           // 334 MOVIE.COMMAND
    None,           // 335 GET.MOVIE
    Some((1, 255)), // 336 CONCATENATE
    Some((2, 2)),   // 337 POWER
    None,           // 338 PIVOT.ADD.DATA
    None,           // 339 GET.PIVOT.TABLE
    None,           // 340 GET.PIVOT.FIELD
    None,           // 341 GET.PIVOT.ITEM
    Some((1, 1)),   // 342 RADIANS
    Some((1, 1)),   // 343 DEGREES
    Some((2, 255)), // 344 SUBTOTAL
    Some((2, 3)),   // 345 SUMIF
    Some((2, 2)),   // 346 COUNTIF
    Some((1, 1)),   // 347 COUNTBLANK
    None,           // 348 SCENARIO.GET
    None,           // 349 OPTIONS.LISTS.GET
    Some((4, 4)),   // 350 ISPMT
    Some((3, 3)),   // 351 DATEDIF
    None,           // 352 DATESTRING
    None,           // 353 NUMBERSTRING
    Some((1, 2)),   // 354 ROMAN
    None,           // 355 OPEN.DIALOG
    None,           // 356 SAVE.DIALOG
    None,           // 357 VIEW.GET
    Some((2, 255)), // 358 GETPIVOTDATA
    Some((1, 2)),   // 359 HYPERLINK
    Some((1, 1)),   // 360 PHONETIC
    Some((1, 255)), // 361 AVERAGEA
    Some((1, 255)), // 362 MAXA
    Some((1, 255)), // 363 MINA
    Some((1, 255)), // 364 STDEVPA
    Some((1, 255)), // 365 VARPA
    Some((1, 255)), // 366 STDEVA
    Some((1, 255)), // 367 VARA
    Some((1, 1)),   // 368 BAHTTEXT
    Some((1, 1)),   // 369 THAIDAYOFWEEK
    Some((1, 1)),   // 370 THAIDIGIT
    Some((1, 1)),   // 371 THAIMONTHOFYEAR
    Some((1, 1)),   // 372 THAINUMSOUND
    Some((1, 1)),   // 373 THAINUMSTRING
    Some((1, 1)),   // 374 THAISTRINGLENGTH
    Some((1, 1)),   // 375 ISTHAIDIGIT
    Some((1, 1)),   // 376 ROUNDBAHTDOWN
    Some((1, 1)),   // 377 ROUNDBAHTUP
    Some((1, 1)),   // 378 THAIYEAR
    Some((3, 255)), // 379 RTD
    Some((2, 255)), // 380 CUBEVALUE
    Some((2, 3)),   // 381 CUBEMEMBER
    Some((3, 3)),   // 382 CUBEMEMBERPROPERTY
    Some((3, 4)),   // 383 CUBERANKEDMEMBER
    Some((1, 2)),   // 384 HEX2BIN
    Some((1, 1)),   // 385 HEX2DEC
    Some((1, 2)),   // 386 HEX2OCT
    Some((1, 2)),   // 387 DEC2BIN
    Some((1, 2)),   // 388 DEC2HEX
    Some((1, 2)),   // 389 DEC2OCT
    Some((1, 2)),   // 390 OCT2BIN
    Some((1, 2)),   // 391 OCT2HEX
    Some((1, 1)),   // 392 OCT2DEC
    Some((1, 1)),   // 393 BIN2DEC
    Some((1, 2)),   // 394 BIN2OCT
    Some((1, 2)),   // 395 BIN2HEX
    Some((2, 2)),   // 396 IMSUB
    Some((2, 2)),   // 397 IMDIV
    Some((2, 2)),   // 398 IMPOWER
    Some((1, 1)),   // 399 IMABS
    Some((1, 1)),   // 400 IMSQRT
    Some((1, 1)),   // 401 IMLN
    Some((1, 1)),   // 402 IMLOG2
    Some((1, 1)),   // 403 IMLOG10
    Some((1, 1)),   // 404 IMSIN
    Some((1, 1)),   // 405 IMCOS
    Some((1, 1)),   // 406 IMEXP
    Some((1, 1)),   // 407 IMARGUMENT
    Some((1, 1)),   // 408 IMCONJUGATE
    Some((1, 1)),   // 409 IMAGINARY
    Some((1, 1)),   // 410 IMREAL
    Some((2, 3)),   // 411 COMPLEX
    Some((1, 255)), // 412 IMSUM
    Some((1, 255)), // 413 IMPRODUCT
    Some((4, 4)),   // 414 SERIESSUM
    Some((1, 1)),   // 415 FACTDOUBLE
    Some((1, 1)),   // 416 SQRTPI
    Some((2, 2)),   // 417 QUOTIENT
    Some((1, 2)),   // 418 DELTA
    Some((1, 2)),   // 419 GESTEP
    Some((1, 1)),   // 420 ISEVEN
    Some((1, 1)),   // 421 ISODD
    Some((2, 2)),   // 422 MROUND
    Some((1, 2)),   // 423 ERF
    Some((1, 1)),   // 424 ERFC
    Some((2, 2)),   // 425 BESSELJ
    Some((2, 2)),   // 426 BESSELK
    Some((2, 2)),   // 427 BESSELY
    Some((2, 2)),   // 428 BESSELI
    Some((2, 3)),   // 429 XIRR
    Some((3, 3)),   // 430 XNPV
    Some((5, 6)),   // 431 PRICEMAT
    Some((5, 6)),   // 432 YIELDMAT
    Some((4, 5)),   // 433 INTRATE
    Some((4, 5)),   // 434 RECEIVED
    Some((4, 5)),   // 435 DISC
    Some((4, 5)),   // 436 PRICEDISC
    Some((4, 5)),   // 437 YIELDDISC
    Some((3, 3)),   // 438 TBILLEQ
    Some((3, 3)),   // 439 TBILLPRICE
    Some((3, 3)),   // 440 TBILLYIELD
    Some((6, 7)),   // 441 PRICE
    Some((6, 7)),   // 442 YIELD
    Some((2, 2)),   // 443 DOLLARDE
    Some((2, 2)),   // 444 DOLLARFR
    Some((2, 2)),   // 445 NOMINAL
    Some((2, 2)),   // 446 EFFECT
    Some((6, 6)),   // 447 CUMPRINC
    Some((6, 6)),   // 448 CUMIPMT
    Some((2, 2)),   // 449 EDATE
    Some((2, 2)),   // 450 EOMONTH
    Some((2, 3)),   // 451 YEARFRAC
    Some((3, 4)),   // 452 COUPDAYBS
    Some((3, 4)),   // 453 COUPDAYS
    Some((3, 4)),   // 454 COUPDAYSNC
    Some((3, 4)),   // 455 COUPNCD
    Some((3, 4)),   // 456 COUPNUM
    Some((3, 4)),   // 457 COUPPCD
    Some((5, 6)),   // 458 DURATION
    Some((5, 6)),   // 459 MDURATION
    Some((7, 8)),   // 460 ODDLPRICE
    Some((7, 8)),   // 461 ODDLYIELD
    Some((8, 9)),   // 462 ODDFPRICE
    Some((8, 9)),   // 463 ODDFYIELD
    Some((2, 2)),   // 464 RANDBETWEEN
    Some((1, 2)),   // 465 WEEKNUM
    Some((6, 7)),   // 466 AMORDEGRC
    Some((6, 7)),   // 467 AMORLINC
    Some((3, 3)),   // 468 CONVERT
    Some((6, 8)),   // 469 ACCRINT
    Some((4, 5)),   // 470 ACCRINTM
    Some((2, 3)),   // 471 WORKDAY
    Some((2, 3)),   // 472 NETWORKDAYS
    Some((1, 255)), // 473 GCD
    Some((1, 255)), // 474 MULTINOMIAL
    Some((1, 255)), // 475 LCM
    Some((2, 2)),   // 476 FVSCHEDULE
    Some((3, 4)),   // 477 CUBEKPIMEMBER
    Some((2, 5)),   // 478 CUBESET
    Some((1, 1)),   // 479 CUBESETCOUNT
    Some((2, 2)),   // 480 IFERROR
    Some((2, 254)), // 481 COUNTIFS
    Some((3, 255)), // 482 SUMIFS
    Some((2, 3)),   // 483 AVERAGEIF
    Some((3, 255)), // 484 AVERAGEIFS
];

pub fn function_name_to_id(name: &str) -> Option<u16> {
    crate::ftab::function_id_from_name(name)
}

/// Like [`function_name_to_id`], but assumes `name` is already ASCII-uppercase.
///
/// This exists to avoid redundant `to_ascii_uppercase()` allocations in hot paths where the
/// caller already maintains an uppercase function name.
pub fn function_name_to_id_uppercase(name: &str) -> Option<u16> {
    crate::ftab::function_id_from_uppercase_name(name)
}

pub fn function_id_to_name(id: u16) -> Option<&'static str> {
    crate::ftab::function_name_from_id(id)
}

pub fn function_spec_from_id(id: u16) -> Option<FunctionSpec> {
    let name = crate::ftab::function_name_from_id(id)?;
    let (min_args, max_args) = FTAB_ARG_RANGES
        .get(id as usize)
        .copied()
        .flatten()?;
    Some(FunctionSpec {
        id,
        name,
        min_args,
        max_args,
    })
}

#[cfg(feature = "encode")]
pub(crate) fn function_spec_from_name(name: &str) -> Option<FunctionSpec> {
    let name = name.trim();
    if name.is_empty() {
        return None;
    }

    let mut buf = [0u8; 64];
    let upper_owned: String;
    let upper: &str = if name.len() <= buf.len() {
        for (dst, src) in buf[..name.len()].iter_mut().zip(name.as_bytes()) {
            *dst = src.to_ascii_uppercase();
        }
        std::str::from_utf8(&buf[..name.len()]).expect("ASCII uppercasing preserves UTF-8")
    } else {
        upper_owned = name.to_ascii_uppercase();
        &upper_owned
    };

    // Encode supports only true built-in FTAB entries. `function_id_from_name` also
    // returns `0x00FF` (USER) for forward-compat `_xlfn.` functions; reject those by
    // ensuring the canonical FTAB name matches the requested (possibly `_xlfn.`-stripped)
    // name.
    let normalized = upper.strip_prefix("_XLFN.").unwrap_or(upper);
    let id = crate::ftab::function_id_from_uppercase_name(upper)?;
    let canonical = crate::ftab::function_name_from_id(id)?;
    if canonical != normalized {
        return None;
    }

    function_spec_from_id(id)
}
