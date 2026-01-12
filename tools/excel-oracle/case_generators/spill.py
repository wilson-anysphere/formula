from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Dynamic arrays / spilling
    # ------------------------------------------------------------------
    add_case(
        cases,
        prefix="spill_range",
        tags=["spill", "range"],
        formula="=A1:A3",
        inputs=[CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3)],
        output_cell="C1",
        description="Reference spill",
    )
    add_case(
        cases,
        prefix="spill_transpose",
        tags=["spill", "TRANSPOSE"],
        formula="=TRANSPOSE(A1:C1)",
        inputs=[CellInput("A1", 1), CellInput("B1", 2), CellInput("C1", 3)],
        output_cell="E1",
        description="Function spill",
    )
    add_case(
        cases,
        prefix="spill_minverse",
        tags=["spill", "MINVERSE", "matrix"],
        formula="=MINVERSE({1,2;3,5})",
        output_cell="C1",
        description="Matrix inverse spill",
    )
    add_case(
        cases,
        prefix="spill_mmult",
        tags=["spill", "MMULT", "matrix"],
        formula="=MMULT({1,2;3,4},{5;6})",
        output_cell="C1",
        description="Matrix multiply spill",
    )
    add_case(
        cases,
        prefix="spill_munit",
        tags=["spill", "MUNIT", "matrix"],
        formula="=MUNIT(3)",
        output_cell="C1",
        description="Identity matrix spill (Excel 365+)",
    )
    add_case(
        cases,
        prefix="spill_sequence",
        tags=["spill", "SEQUENCE", "dynarr"],
        formula="=SEQUENCE(2,2,1,1)",
        inputs=[],
        output_cell="C1",
        description="Dynamic array function (Excel 365+)",
    )
    add_case(
        cases,
        prefix="spill_textsplit",
        tags=["spill", "TEXTSPLIT", "dynarr"],
        formula='=TEXTSPLIT("a,b,c",",")',
        description="Dynamic array function (Excel 365+)",
    )
    add_case(
        cases,
        prefix="spill_frequency",
        tags=["spill", "FREQUENCY"],
        formula="=FREQUENCY({1,2,3,4,5},{2,4})",
        output_cell="C1",
        description="Array/spill-producing histogram bucket counts",
    )

    # FILTER / SORT / UNIQUE (simple spill cases)
    filter_inputs = [CellInput(f"A{i}", i) for i in range(1, 6)]
    add_case(
        cases,
        prefix="spill_filter",
        tags=["spill", "FILTER", "dynarr"],
        formula="=FILTER(A1:A5,A1:A5>2)",
        inputs=filter_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="spill_filter",
        tags=["spill", "FILTER", "dynarr"],
        formula='=FILTER(A1:A5,A1:A5>10,"none")',
        inputs=filter_inputs,
        output_cell="C1",
        description="if_empty fallback (no matches)",
    )

    sort_inputs = [CellInput("A1", 3), CellInput("A2", 1), CellInput("A3", 2)]
    add_case(
        cases,
        prefix="spill_sort",
        tags=["spill", "SORT", "dynarr"],
        formula="=SORT(A1:A3)",
        inputs=sort_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="spill_sort",
        tags=["spill", "SORT", "dynarr"],
        formula="=SORT(A1:A3,1,-1)",
        inputs=sort_inputs,
        output_cell="C1",
    )

    unique_inputs = [
        CellInput("A1", 1),
        CellInput("A2", 1),
        CellInput("A3", 2),
        CellInput("A4", 3),
        CellInput("A5", 3),
        CellInput("A6", 3),
    ]
    add_case(
        cases,
        prefix="spill_unique",
        tags=["spill", "UNIQUE", "dynarr"],
        formula="=UNIQUE(A1:A6)",
        inputs=unique_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="spill_unique",
        tags=["spill", "UNIQUE", "dynarr"],
        formula="=UNIQUE(A1:A6,FALSE,TRUE)",
        inputs=unique_inputs,
        output_cell="C1",
        description="Return values that occur exactly once",
    )

    # Dynamic array helpers / shape functions (function-catalog backfill)
    add_case(cases, prefix="choose", tags=["lookup", "CHOOSE"], formula='=CHOOSE(2,"one","two","three")')
    add_case(
        cases,
        prefix="choosecols",
        tags=["spill", "CHOOSECOLS", "dynarr"],
        formula="=CHOOSECOLS({1,2,3;4,5,6},1,3)",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="chooserows",
        tags=["spill", "CHOOSEROWS", "dynarr"],
        formula="=CHOOSEROWS({1,2;3,4;5,6},1,3)",
        output_cell="C1",
    )
    add_case(cases, prefix="hstack", tags=["spill", "HSTACK", "dynarr"], formula="=HSTACK({1,2},{3,4})", output_cell="C1")
    add_case(cases, prefix="vstack", tags=["spill", "VSTACK", "dynarr"], formula="=VSTACK({1,2},{3,4})", output_cell="C1")
    add_case(cases, prefix="take", tags=["spill", "TAKE", "dynarr"], formula="=TAKE({1,2,3;4,5,6},1,2)", output_cell="C1")
    add_case(cases, prefix="drop", tags=["spill", "DROP", "dynarr"], formula="=DROP({1,2,3;4,5,6},1,1)", output_cell="C1")
    add_case(cases, prefix="tocol", tags=["spill", "TOCOL", "dynarr"], formula="=TOCOL({1,2;3,4})", output_cell="C1")
    add_case(cases, prefix="torow", tags=["spill", "TOROW", "dynarr"], formula="=TOROW({1,2;3,4})", output_cell="C1")
    add_case(cases, prefix="wraprows", tags=["spill", "WRAPROWS", "dynarr"], formula="=WRAPROWS({1,2,3,4,5,6},2)", output_cell="C1")
    add_case(cases, prefix="wrapcols", tags=["spill", "WRAPCOLS", "dynarr"], formula="=WRAPCOLS({1,2,3,4,5,6},2)", output_cell="C1")
    add_case(
        cases,
        prefix="expand",
        tags=["spill", "EXPAND", "dynarr"],
        formula="=EXPAND({1,2;3,4},3,3,0)",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="sortby",
        tags=["spill", "SORTBY", "dynarr"],
        formula='=SORTBY({"b";"a";"c"},{2;1;3})',
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="makearray",
        tags=["spill", "MAKEARRAY", "LAMBDA", "dynarr"],
        formula="=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="map",
        tags=["spill", "MAP", "LAMBDA", "dynarr"],
        formula="=MAP({1,2,3},LAMBDA(x,x*2))",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="reduce",
        tags=["spill", "REDUCE", "LAMBDA", "dynarr"],
        formula="=REDUCE(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="scan",
        tags=["spill", "SCAN", "LAMBDA", "dynarr"],
        formula="=SCAN(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="byrow",
        tags=["spill", "BYROW", "LAMBDA", "dynarr"],
        formula="=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))",
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="bycol",
        tags=["spill", "BYCOL", "LAMBDA", "dynarr"],
        formula="=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))",
        output_cell="C1",
    )
    add_case(cases, prefix="let", tags=["lambda", "LET"], formula="=LET(a,2,b,a*3,c,b+1,c)")
    add_case(cases, prefix="lambda", tags=["lambda", "LAMBDA"], formula="=LAMBDA(x,x+1)(2)")
    add_case(
        cases,
        prefix="isomitted",
        tags=["lambda", "LAMBDA", "ISOMITTED"],
        formula="=LAMBDA(x,y,ISOMITTED(y))(1)",
        description="Missing LAMBDA arguments bind as blank and are detectable via ISOMITTED",
    )
