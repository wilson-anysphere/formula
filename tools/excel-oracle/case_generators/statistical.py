from __future__ import annotations

import itertools
from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
    excel_serial_1900,
) -> None:
    # ------------------------------------------------------------------
    # Aggregates over ranges
    # ------------------------------------------------------------------
    agg_vals = [0, 1, -1, 2, -2, 0.5]
    # Keep range sizes small to make Excel automation fast, and cap the number of
    # permutations so the overall corpus stays ~1k.
    for a, b, c in itertools.islice(itertools.product(agg_vals, repeat=3), 30):
        inputs = [
            CellInput("A1", a),
            CellInput("A2", b),
            CellInput("A3", c),
        ]
        add_case(cases, prefix="sum_r", tags=["agg", "SUM"], formula="=SUM(A1:A3)", inputs=inputs)
        add_case(cases, prefix="avg_r", tags=["agg", "AVERAGE"], formula="=AVERAGE(A1:A3)", inputs=inputs)
        add_case(cases, prefix="min_r", tags=["agg", "MIN"], formula="=MIN(A1:A3)", inputs=inputs)
        add_case(cases, prefix="max_r", tags=["agg", "MAX"], formula="=MAX(A1:A3)", inputs=inputs)

    # Type coercion edges (SUM ignores text in ranges, but direct args coerce differently).
    add_case(
        cases,
        prefix="sum_args",
        tags=["agg", "SUM", "coercion"],
        formula='=SUM("5",3)',
        inputs=[],
    )
    add_case(
        cases,
        prefix="sum_args",
        tags=["agg", "SUM", "coercion"],
        formula='=SUM("abc",3)',
        inputs=[],
    )
    add_case(
        cases,
        prefix="sum_rng_text",
        tags=["agg", "SUM", "coercion"],
        formula="=SUM(A1:A3)",
        inputs=[CellInput("A1", 1), CellInput("A2", "text"), CellInput("A3", 3)],
    )

    # COUNT / COUNTA / COUNTBLANK
    count_range_inputs = [
        CellInput("A1", 1),
        CellInput("A2", "x"),
        CellInput("A3", True),
        CellInput("A4", None),  # blank
        CellInput("A5", ""),  # empty string counts as blank for COUNTBLANK
        CellInput("A6", formula="=1/0"),  # errors are ignored by COUNT* in Excel
    ]
    add_case(cases, prefix="count", tags=["agg", "COUNT"], formula="=COUNT(A1:A6)", inputs=count_range_inputs)
    add_case(cases, prefix="counta", tags=["agg", "COUNTA"], formula="=COUNTA(A1:A6)", inputs=count_range_inputs)
    add_case(
        cases,
        prefix="countblank",
        tags=["agg", "COUNTBLANK"],
        formula="=COUNTBLANK(A1:A6)",
        inputs=count_range_inputs,
    )

    # COUNTIF (include numeric, text + wildcards, blanks)
    countif_num_inputs = [CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3), CellInput("A4", 4)]
    add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A4,">2")',
        inputs=countif_num_inputs,
    )
    countif_text_inputs = [
        CellInput("A1", "apple"),
        CellInput("A2", "banana"),
        CellInput("A3", "apricot"),
        CellInput("A4", None),
        CellInput("A5", ""),
    ]
    add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A5,"ap*")',
        inputs=countif_text_inputs,
    )
    add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A5,"")',
        inputs=countif_text_inputs,
        description="Blank criteria matches truly blank cells and empty-string cells",
    )

    # SUMPRODUCT
    sumproduct_inputs = [
        CellInput("A1", 1),
        CellInput("A2", 2),
        CellInput("A3", 3),
        CellInput("B1", 4),
        CellInput("B2", 5),
        CellInput("B3", 6),
    ]
    add_case(
        cases,
        prefix="sumproduct",
        tags=["agg", "SUMPRODUCT"],
        formula="=SUMPRODUCT(A1:A3,B1:B3)",
        inputs=sumproduct_inputs,
    )
    add_case(
        cases,
        prefix="sumproduct",
        tags=["agg", "SUMPRODUCT", "error"],
        formula="=SUMPRODUCT(A1:A2,B1:B2)",
        inputs=[
            CellInput("A1", 1),
            CellInput("A2", formula="=1/0"),
            CellInput("B1", 2),
            CellInput("B2", 3),
        ],
        description="SUMPRODUCT propagates errors from any element",
    )

    subtotal_inputs = [CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3)]
    add_case(
        cases,
        prefix="subtotal",
        tags=["agg", "SUBTOTAL"],
        formula="=SUBTOTAL(9,A1:A3)",
        inputs=subtotal_inputs,
    )
    add_case(
        cases,
        prefix="aggregate",
        tags=["agg", "AGGREGATE"],
        formula="=AGGREGATE(9,4,A1:A3)",
        inputs=subtotal_inputs,
    )

    # ------------------------------------------------------------------
    # Criteria-string semantics (COUNTIF/SUMIF/AVERAGEIF and *IFS variants)
    # ------------------------------------------------------------------
    # These functions have a large surface area of Excel-compat behaviors:
    # operator parsing, wildcards/escapes, blank/error/boolean matching, and
    # numeric-vs-text coercion.

    # Shared text range for wildcard / escape / blank / error criteria probing.
    criteria_text_inputs = [
        CellInput("A1", "apple"),
        CellInput("A2", "apricot"),
        CellInput("A3", "apex"),
        CellInput("A4", "fax"),
        CellInput("A5", "xylophone"),
        CellInput("A6", "max"),
        CellInput("A7", "*"),
        CellInput("A8", "?"),
        CellInput("A9", "~"),
        CellInput("A10", "~a"),
        CellInput("A11", "a"),
        CellInput("A12", ""),
        CellInput("A13", None),
        CellInput("A14", formula="=1/0"),
        CellInput("A15", formula="=NA()"),
    ]

    for crit in ["ap*", "*x*", "?a?"]:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "wildcards"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["~*", "~?", "~~", "~a"]:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "escapes"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["", "=", "<>"]:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "blanks"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["#DIV/0!", "#N/A"]:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "errors"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    # Numeric-vs-text and operator parsing.
    criteria_num_inputs = [
        CellInput("B1", 5),
        CellInput("B2", "5"),
        CellInput("B3", 6),
        CellInput("B4", "6"),
        CellInput("B5", 0),
        CellInput("B6", "0"),
        CellInput("B7", 3),
        CellInput("B8", "3"),
        CellInput("B9", None),
        CellInput("B10", ""),
    ]

    for crit_expr in ['">5"', '"<=0"', '"<>3"']:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "operators"],
            formula=f"=COUNTIF(B1:B10,{crit_expr})",
            inputs=criteria_num_inputs,
        )

    for crit_expr in ['"5"', "5", '"5*"']:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "numeric-vs-text"],
            formula=f"=COUNTIF(B1:B10,{crit_expr})",
            inputs=criteria_num_inputs,
        )

    # Boolean criteria. Mix booleans, numbers, and text booleans.
    criteria_bool_inputs = [
        CellInput("C1", True),
        CellInput("C2", False),
        CellInput("C3", 1),
        CellInput("C4", 0),
        CellInput("C5", "TRUE"),
        CellInput("C6", "FALSE"),
        CellInput("C7", None),
    ]

    for crit_expr in ["TRUE", "FALSE", '"TRUE"', '"FALSE"']:
        add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "booleans"],
            formula=f"=COUNTIF(C1:C7,{crit_expr})",
            inputs=criteria_bool_inputs,
            output_cell="D1",
        )

    # Date criteria against date serial numbers.
    #
    # Avoid locale-dependent date parsing (e.g. "1/1/2020") by building criteria
    # from DATE(...) so the criterion is numeric and stable across locales.
    date_serials = [
        excel_serial_1900(2019, 12, 31),
        excel_serial_1900(2020, 1, 1),
        excel_serial_1900(2020, 2, 1),
        excel_serial_1900(2021, 1, 1),
    ]
    criteria_date_inputs = [CellInput(f"D{i+1}", v) for i, v in enumerate(date_serials)]
    add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula='=COUNTIF(D1:D4,">"&DATE(2020,1,1))',
        inputs=criteria_date_inputs,
    )
    add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula='=COUNTIF(D1:D4,"<="&DATE(2020,1,1))',
        inputs=criteria_date_inputs,
    )
    add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula="=COUNTIF(D1:D4,DATE(2020,1,1))",
        inputs=criteria_date_inputs,
    )

    # Multi-criteria fixtures (shared by *IFS variants).
    ifs_inputs = [
        CellInput("E1", "apple"),
        CellInput("E2", "apex"),
        CellInput("E3", "banana"),
        CellInput("E4", "fax"),
        CellInput("E5", "max"),
        CellInput("E6", "*"),
        CellInput("E7", None),
        CellInput("F1", 1),
        CellInput("F2", 2),
        CellInput("F3", 3),
        CellInput("F4", 4),
        CellInput("F5", 5),
        CellInput("F6", 6),
        CellInput("F7", 7),
        CellInput("G1", 10),
        CellInput("G2", 20),
        CellInput("G3", 30),
        CellInput("G4", 40),
        CellInput("G5", 50),
        CellInput("G6", 60),
        CellInput("G7", 70),
    ]

    # COUNTIFS basics + wildcards/escapes + blanks.
    add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "wildcards"],
        formula='=COUNTIFS(E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "escapes"],
        formula='=COUNTIFS(E1:E7,"~*",F1:F7,">=0")',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "blanks"],
        formula='=COUNTIFS(E1:E7,"",F1:F7,">=0")',
        inputs=ifs_inputs,
    )

    # COUNTIFS invalid-arity and shape mismatch should error (#VALUE).
    add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "arg-count"],
        formula='=COUNTIFS(E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "shape-mismatch"],
        formula='=COUNTIFS(E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    # SUMIF: wildcards/escapes and operator parsing.
    add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "wildcards"],
        formula='=SUMIF(E1:E7,"*x*",G1:G7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "escapes"],
        formula='=SUMIF(E1:E7,"~*",G1:G7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "operators"],
        formula='=SUMIF(F1:F7,">5",G1:G7)',
        inputs=ifs_inputs,
    )

    # SUMIF: errors in the summed range only matter if the row is included.
    sumif_err_inputs = [
        CellInput("I1", 1),
        CellInput("I2", 2),
        CellInput("I3", 3),
        CellInput("J1", 10),
        CellInput("J2", formula="=1/0"),
        CellInput("J3", 30),
    ]
    add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "sum-range-errors"],
        formula='=SUMIF(I1:I3,">2",J1:J3)',
        inputs=sumif_err_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "sum-range-errors"],
        formula='=SUMIF(I1:I3,">1",J1:J3)',
        inputs=sumif_err_inputs,
    )

    # SUMIFS: multi-criteria, invalid-arity, and shape mismatch should error (#VALUE).
    add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "wildcards"],
        formula='=SUMIFS(G1:G7,E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "arg-count"],
        formula='=SUMIFS(G1:G7,E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "shape-mismatch"],
        formula='=SUMIFS(G1:G7,E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    sumifs_err_inputs = [
        CellInput("K1", 1),
        CellInput("K2", 2),
        CellInput("K3", 3),
        CellInput("K4", 4),
        CellInput("L1", "x"),
        CellInput("L2", "x"),
        CellInput("L3", "y"),
        CellInput("L4", "y"),
        CellInput("M1", 10),
        CellInput("M2", formula="=1/0"),
        CellInput("M3", 30),
        CellInput("M4", 40),
    ]
    add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "sum-range-errors"],
        formula='=SUMIFS(M1:M4,K1:K4,">=3",L1:L4,"y")',
        inputs=sumifs_err_inputs,
    )
    add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "sum-range-errors"],
        formula='=SUMIFS(M1:M4,K1:K4,">=2",L1:L4,"x")',
        inputs=sumifs_err_inputs,
    )

    # AVERAGEIF: #DIV/0! when no numeric values are included.
    add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "wildcards"],
        formula='=AVERAGEIF(E1:E7,"*x*",G1:G7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "operators"],
        formula='=AVERAGEIF(F1:F7,">5",G1:G7)',
        inputs=ifs_inputs,
    )
    avgif_no_numeric_inputs = [
        CellInput("N1", 1),
        CellInput("N2", 2),
        CellInput("N3", 3),
        CellInput("O1", "a"),
        CellInput("O2", "b"),
        CellInput("O3", "c"),
    ]
    add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "no-numeric"],
        formula='=AVERAGEIF(N1:N3,">0",O1:O3)',
        inputs=avgif_no_numeric_inputs,
    )
    add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "no-numeric"],
        formula='=AVERAGEIF(N1:N3,">10",N1:N3)',
        inputs=avgif_no_numeric_inputs,
    )

    # AVERAGEIFS: multi-criteria, invalid-arity/shape mismatch (#VALUE), and #DIV/0! on no matches.
    add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "wildcards"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "arg-count"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "shape-mismatch"],
        formula='=AVERAGEIFS(G1:G7,E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )
    add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "no-numeric"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"no_match",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    # ------------------------------------------------------------------
    # Statistical / regression functions (function-catalog backfill)
    # ------------------------------------------------------------------
    # Prefer array literals to keep the corpus compact and deterministic.
    add_case(cases, prefix="avedev", tags=["stat", "AVEDEV"], formula="=AVEDEV({1,2,3,4})")
    add_case(cases, prefix="averagea", tags=["stat", "AVERAGEA"], formula='=AVERAGEA({1,"x",TRUE})')
    add_case(cases, prefix="maxa", tags=["stat", "MAXA"], formula='=MAXA({1,"x",TRUE})')
    add_case(cases, prefix="mina", tags=["stat", "MINA"], formula='=MINA({1,"x",FALSE})')
    add_case(cases, prefix="median", tags=["stat", "MEDIAN"], formula="=MEDIAN({1,2,3,4})")
    add_case(cases, prefix="mode", tags=["stat", "MODE"], formula="=MODE({1,1,2,3})")
    add_case(cases, prefix="mode_sngl", tags=["stat", "MODE.SNGL"], formula="=MODE.SNGL({1,1,2,3})")
    add_case(cases, prefix="mode_mult", tags=["stat", "MODE.MULT"], formula="=MODE.MULT({1,1,2,2,3})", output_cell="C1")

    add_case(cases, prefix="devsq", tags=["stat", "DEVSQ"], formula="=DEVSQ({1,2,3})")
    add_case(cases, prefix="geomean", tags=["stat", "GEOMEAN"], formula="=GEOMEAN({1,2,3,4})")
    add_case(cases, prefix="harmean", tags=["stat", "HARMEAN"], formula="=HARMEAN({1,2,4})")

    add_case(cases, prefix="kurt", tags=["stat", "KURT"], formula="=KURT({3,4,5,2,3,4,5,6,4,7})")
    add_case(cases, prefix="skew", tags=["stat", "SKEW"], formula="=SKEW({3,4,5,2,3,4,5,6,4,7})")
    add_case(cases, prefix="skew_p", tags=["stat", "SKEW.P"], formula="=SKEW.P({3,4,5,2,3,4,5,6,4,7})")

    # Normal distribution helpers (CDF/PDF/INV + transforms + legacy aliases).
    add_case(cases, prefix="norm_dist", tags=["stat", "NORM.DIST"], formula="=NORM.DIST(0,0,1,TRUE)")
    add_case(cases, prefix="norm_s_dist", tags=["stat", "NORM.S.DIST"], formula="=NORM.S.DIST(1,TRUE)")
    add_case(cases, prefix="norm_inv", tags=["stat", "NORM.INV"], formula="=NORM.INV(0.5,1,2)")
    add_case(cases, prefix="norm_s_inv", tags=["stat", "NORM.S.INV"], formula="=NORM.S.INV(0.975)")
    add_case(cases, prefix="normdist", tags=["stat", "NORMDIST"], formula="=NORMDIST(0,0,1,TRUE)")
    add_case(cases, prefix="normsdist", tags=["stat", "NORMSDIST"], formula="=NORMSDIST(1)")
    add_case(cases, prefix="norminv", tags=["stat", "NORMINV"], formula="=NORMINV(0.5,1,2)")
    add_case(cases, prefix="normsinv", tags=["stat", "NORMSINV"], formula="=NORMSINV(0.975)")
    add_case(cases, prefix="phi", tags=["stat", "PHI"], formula="=PHI(0)")
    add_case(cases, prefix="gauss", tags=["stat", "GAUSS"], formula="=GAUSS(1)")

    # Discrete distributions + probability functions (CDF/PDF/INV + legacy aliases).
    add_case(cases, prefix="binom_dist", tags=["stat", "BINOM.DIST"], formula="=BINOM.DIST(2,5,0.5,FALSE)")
    add_case(
        cases,
        prefix="binom_dist_range",
        tags=["stat", "BINOM.DIST.RANGE"],
        formula="=BINOM.DIST.RANGE(5,0.5,1,3)",
    )
    add_case(cases, prefix="binom_inv", tags=["stat", "BINOM.INV"], formula="=BINOM.INV(10,0.5,0.5)")
    add_case(cases, prefix="binomdist", tags=["stat", "BINOMDIST"], formula="=BINOMDIST(2,5,0.5,FALSE)")
    add_case(cases, prefix="critbinom", tags=["stat", "CRITBINOM"], formula="=CRITBINOM(10,0.5,0.5)")

    add_case(cases, prefix="poisson_dist", tags=["stat", "POISSON.DIST"], formula="=POISSON.DIST(2,3,FALSE)")
    add_case(cases, prefix="poisson", tags=["stat", "POISSON"], formula="=POISSON(2,3,TRUE)")

    add_case(
        cases,
        prefix="negbinom_dist",
        tags=["stat", "NEGBINOM.DIST"],
        formula="=NEGBINOM.DIST(3,2,0.5,FALSE)",
    )
    add_case(cases, prefix="negbinomdist", tags=["stat", "NEGBINOMDIST"], formula="=NEGBINOMDIST(3,2,0.5)")

    add_case(
        cases,
        prefix="hypgeom_dist",
        tags=["stat", "HYPGEOM.DIST"],
        formula="=HYPGEOM.DIST(2,5,5,10,FALSE)",
    )
    add_case(cases, prefix="hypgeomdist", tags=["stat", "HYPGEOMDIST"], formula="=HYPGEOMDIST(2,5,5,10)")

    add_case(cases, prefix="prob", tags=["stat", "PROB"], formula="=PROB({0,1,2},{0.2,0.5,0.3},0,1)")

    # Hypothesis tests + legacy aliases.
    add_case(cases, prefix="z_test", tags=["stat", "Z.TEST"], formula="=Z.TEST({1,2,3,4},2)")
    add_case(cases, prefix="ztest", tags=["stat", "ZTEST"], formula="=ZTEST({1,2,3,4},2)")

    add_case(cases, prefix="t_test", tags=["stat", "T.TEST"], formula="=T.TEST({1,2,3},{3,2,1},2,1)")
    add_case(cases, prefix="ttest", tags=["stat", "TTEST"], formula="=TTEST({1,2,3},{3,2,1},2,1)")

    add_case(cases, prefix="f_test", tags=["stat", "F.TEST"], formula="=F.TEST({1,2,3},{1,2,3})")
    add_case(cases, prefix="ftest", tags=["stat", "FTEST"], formula="=FTEST({1,2,3},{1,2,3})")

    add_case(
        cases,
        prefix="chisq_test",
        tags=["stat", "CHISQ.TEST"],
        formula="=CHISQ.TEST({10,20;30,40},{12,18;28,42})",
    )
    add_case(cases, prefix="chitest", tags=["stat", "CHITEST"], formula="=CHITEST({10,20;30,40},{12,18;28,42})")

    add_case(
        cases,
        prefix="frequency",
        tags=["stat", "FREQUENCY"],
        formula="=FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89})",
        output_cell="C1",
    )

    add_case(cases, prefix="large", tags=["stat", "LARGE"], formula="=LARGE({1,2,3,4},2)")
    add_case(cases, prefix="small", tags=["stat", "SMALL"], formula="=SMALL({1,2,3,4},2)")

    add_case(cases, prefix="percentile", tags=["stat", "PERCENTILE"], formula="=PERCENTILE({1,2,3,4},0.25)")
    add_case(
        cases,
        prefix="percentile_inc",
        tags=["stat", "PERCENTILE.INC"],
        formula="=PERCENTILE.INC({1,2,3,4},0.25)",
    )
    add_case(
        cases,
        prefix="percentile_exc",
        tags=["stat", "PERCENTILE.EXC"],
        formula="=PERCENTILE.EXC({1,2,3,4},0.25)",
    )

    # PERCENTRANK and variants: use a 1..9 set so both inclusive and exclusive variants
    # yield simple finite decimals without needing explicit `significance`.
    add_case(
        cases,
        prefix="percentrank",
        tags=["stat", "PERCENTRANK"],
        formula="=PERCENTRANK({1,2,3,4,5,6,7,8,9},2)",
    )
    add_case(
        cases,
        prefix="percentrank_inc",
        tags=["stat", "PERCENTRANK.INC"],
        formula="=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9},2)",
    )
    add_case(
        cases,
        prefix="percentrank_exc",
        tags=["stat", "PERCENTRANK.EXC"],
        formula="=PERCENTRANK.EXC({1,2,3,4,5,6,7,8,9},2)",
    )

    add_case(cases, prefix="quartile", tags=["stat", "QUARTILE"], formula="=QUARTILE({1,2,3,4},1)")
    add_case(cases, prefix="quartile_inc", tags=["stat", "QUARTILE.INC"], formula="=QUARTILE.INC({1,2,3,4},1)")
    add_case(cases, prefix="quartile_exc", tags=["stat", "QUARTILE.EXC"], formula="=QUARTILE.EXC({1,2,3,4},1)")

    add_case(cases, prefix="rank", tags=["stat", "RANK"], formula="=RANK(2,{1,2,2,3})")
    add_case(cases, prefix="rank_eq", tags=["stat", "RANK.EQ"], formula="=RANK.EQ(2,{1,2,2,3})")
    add_case(cases, prefix="rank_avg", tags=["stat", "RANK.AVG"], formula="=RANK.AVG(2,{1,2,2,3})")

    add_case(cases, prefix="stdev", tags=["stat", "STDEV"], formula="=STDEV({1,2,3,4})")
    add_case(cases, prefix="stdev_s", tags=["stat", "STDEV.S"], formula="=STDEV.S({1,2,3,4})")
    add_case(cases, prefix="stdev_p", tags=["stat", "STDEV.P"], formula="=STDEV.P({1,2,3,4})")
    add_case(cases, prefix="stdeva", tags=["stat", "STDEVA"], formula="=STDEVA({1,2,3,TRUE})")
    add_case(cases, prefix="stdevp", tags=["stat", "STDEVP"], formula="=STDEVP({1,2,3,TRUE})")
    add_case(cases, prefix="stdevpa", tags=["stat", "STDEVPA"], formula="=STDEVPA({1,2,3,TRUE})")

    add_case(cases, prefix="var", tags=["stat", "VAR"], formula="=VAR({1,2,3,4})")
    add_case(cases, prefix="var_s", tags=["stat", "VAR.S"], formula="=VAR.S({1,2,3,4})")
    add_case(cases, prefix="var_p", tags=["stat", "VAR.P"], formula="=VAR.P({1,2,3,4})")
    add_case(cases, prefix="vara", tags=["stat", "VARA"], formula="=VARA({1,2,3,TRUE})")
    add_case(cases, prefix="varp", tags=["stat", "VARP"], formula="=VARP({1,2,3,TRUE})")
    add_case(cases, prefix="varpa", tags=["stat", "VARPA"], formula="=VARPA({1,2,3,TRUE})")

    add_case(cases, prefix="trimmean", tags=["stat", "TRIMMEAN"], formula="=TRIMMEAN({1,2,3,100},0.5)")

    add_case(cases, prefix="standardize", tags=["stat", "STANDARDIZE"], formula="=STANDARDIZE(1,3,2)")

    add_case(cases, prefix="correl", tags=["stat", "CORREL"], formula="=CORREL({1,2,3},{1,5,7})")
    add_case(cases, prefix="pearson", tags=["stat", "PEARSON"], formula="=PEARSON({1,2,3},{1,5,7})")
    add_case(cases, prefix="covar", tags=["stat", "COVAR"], formula="=COVAR({1,2,3},{1,5,7})")
    add_case(cases, prefix="cov_p", tags=["stat", "COVARIANCE.P"], formula="=COVARIANCE.P({1,2,3},{1,5,7})")
    add_case(cases, prefix="cov_s", tags=["stat", "COVARIANCE.S"], formula="=COVARIANCE.S({1,2,3},{1,5,7})")

    add_case(cases, prefix="rsq", tags=["stat", "RSQ"], formula="=RSQ({1,2,3},{1,2,3})")
    add_case(cases, prefix="slope", tags=["stat", "SLOPE"], formula="=SLOPE({1,2,3},{1,2,3})")
    add_case(cases, prefix="intercept", tags=["stat", "INTERCEPT"], formula="=INTERCEPT({1,2,3},{1,2,3})")
    add_case(cases, prefix="forecast", tags=["stat", "FORECAST"], formula="=FORECAST(4,{1,2,3},{1,2,3})")
    add_case(
        cases,
        prefix="forecast_linear",
        tags=["stat", "FORECAST.LINEAR"],
        formula="=FORECAST.LINEAR(4,{1,2,3},{1,2,3})",
    )
    add_case(
        cases,
        prefix="forecast_ets",
        tags=["stat", "FORECAST.ETS"],
        formula="=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},1)",
    )
    # Date timelines in Excel often use monthly/quarterly/yearly serials, which are not evenly
    # spaced in days (28/29/30/31 and 365/366). Include deterministic monthly date timeline cases
    # to ensure engines handle this common pattern.
    add_case(
        cases,
        prefix="forecast_ets_monthly",
        tags=["stat", "FORECAST.ETS", "dates"],
        formula="=FORECAST.ETS(43952,{10,10,10,10},{43831,43862,43891,43922},1)",
        description="Monthly date serial timeline (2020-01-01..2020-04-01) forecasting 2020-05-01",
    )
    add_case(
        cases,
        prefix="forecast_ets_month_end",
        tags=["stat", "FORECAST.ETS", "dates", "eom"],
        formula="=FORECAST.ETS(44012,{10,10,10,10},{43890,43921,43951,43982},1)",
        description="Month-end date serial timeline (2020-02-29..2020-05-31) forecasting 2020-06-30",
    )
    add_case(
        cases,
        prefix="forecast_ets_confint",
        tags=["stat", "FORECAST.ETS.CONFINT"],
        formula="=FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0.95,1)",
    )
    add_case(
        cases,
        prefix="forecast_ets_confint_monthly",
        tags=["stat", "FORECAST.ETS.CONFINT", "dates"],
        formula="=FORECAST.ETS.CONFINT(43952,{10,10,10,10},{43831,43862,43891,43922},0.95,1)",
        description="Confidence interval for monthly date serial timeline (perfect fit -> 0)",
    )
    add_case(
        cases,
        prefix="forecast_ets_confint_month_end",
        tags=["stat", "FORECAST.ETS.CONFINT", "dates", "eom"],
        formula="=FORECAST.ETS.CONFINT(44012,{10,10,10,10},{43890,43921,43951,43982},0.95,1)",
        description="Confidence interval for month-end date serial timeline (perfect fit -> 0)",
    )
    add_case(
        cases,
        prefix="forecast_ets_seasonality",
        tags=["stat", "FORECAST.ETS.SEASONALITY"],
        formula="=FORECAST.ETS.SEASONALITY({10,20,10,20,10,20,10,20},{1,2,3,4,5,6,7,8})",
    )
    add_case(
        cases,
        prefix="forecast_ets_seasonality_monthly",
        tags=["stat", "FORECAST.ETS.SEASONALITY", "dates"],
        formula="=FORECAST.ETS.SEASONALITY({10,10,10,10},{43831,43862,43891,43922})",
        description="Seasonality detection on a monthly date serial timeline (constant series)",
    )
    add_case(
        cases,
        prefix="forecast_ets_seasonality_month_end",
        tags=["stat", "FORECAST.ETS.SEASONALITY", "dates", "eom"],
        formula="=FORECAST.ETS.SEASONALITY({10,10,10,10},{43890,43921,43951,43982})",
        description="Seasonality detection on a month-end date serial timeline (constant series)",
    )
    add_case(
        cases,
        prefix="forecast_ets_stat",
        tags=["stat", "FORECAST.ETS.STAT"],
        formula="=FORECAST.ETS.STAT({1,2,3,4,5,6},{1,2,3,4,5,6},1,1,1,8)",
    )
    add_case(
        cases,
        prefix="forecast_ets_stat_monthly",
        tags=["stat", "FORECAST.ETS.STAT", "dates"],
        formula="=FORECAST.ETS.STAT({10,10,10,10},{43831,43862,43891,43922},1,1,1,8)",
        description="RMSE for monthly date serial timeline (perfect fit -> 0)",
    )
    add_case(
        cases,
        prefix="forecast_ets_stat_month_end",
        tags=["stat", "FORECAST.ETS.STAT", "dates", "eom"],
        formula="=FORECAST.ETS.STAT({10,10,10,10},{43890,43921,43951,43982},1,1,1,8)",
        description="RMSE for month-end date serial timeline (perfect fit -> 0)",
    )

    # STEYX should be 0 for a perfectly linear relationship (y = 2x + 1).
    add_case(cases, prefix="steyx", tags=["stat", "STEYX"], formula="=STEYX({3,5,7,9,11},{1,2,3,4,5})")

    # MAXIFS / MINIFS (criteria-based aggregates)
    maxifs_inputs = [
        CellInput("A1", 10),
        CellInput("A2", 20),
        CellInput("A3", 30),
        CellInput("B1", "A"),
        CellInput("B2", "B"),
        CellInput("B3", "A"),
    ]
    add_case(
        cases,
        prefix="maxifs",
        tags=["agg", "MAXIFS"],
        formula='=MAXIFS(A1:A3,B1:B3,"A")',
        inputs=maxifs_inputs,
    )
    add_case(
        cases,
        prefix="minifs",
        tags=["agg", "MINIFS"],
        formula='=MINIFS(A1:A3,B1:B3,"A")',
        inputs=maxifs_inputs,
    )
