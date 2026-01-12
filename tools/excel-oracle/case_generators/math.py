from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Basic math functions
    # ------------------------------------------------------------------
    math_vals = [0, 1, -1, 2, -2, 0.5, -0.5, 10.25, -10.25, 1e-9, -1e-9, 1e9, -1e9]

    for v in math_vals:
        add_case(cases, prefix="abs", tags=["math", "ABS"], formula="=ABS(A1)", inputs=[CellInput("A1", v)])
        add_case(cases, prefix="sign", tags=["math", "SIGN"], formula="=SIGN(A1)", inputs=[CellInput("A1", v)])
        add_case(cases, prefix="int", tags=["math", "INT"], formula="=INT(A1)", inputs=[CellInput("A1", v)])

    mod_divisors = [1, 2, 3, -2, -3, 10]
    for a in [0, 1, -1, 2, -2, 10, -10, 10.5, -10.5]:
        for b in mod_divisors:
            add_case(
                cases,
                prefix="mod",
                tags=["math", "MOD"],
                formula="=MOD(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    # Keep ROUND* combinations capped so the corpus stays <2k cases total.
    round_vals = [0, 0.1, 0.5, 1.5, 2.5, 10.25, -10.25, 1234.5678, -1234.5678, 1e-7]
    round_digits = [-2, -1, 0, 1, 2]
    for func in ["ROUND", "ROUNDUP", "ROUNDDOWN"]:
        for v in round_vals:
            for d in round_digits:
                add_case(
                    cases,
                    prefix=func.lower(),
                    tags=["math", func],
                    formula=f"={func}(A1,B1)",
                    inputs=[CellInput("A1", v), CellInput("B1", d)],
                )

    # ------------------------------------------------------------------
    # Extended math functions (function-catalog backfill)
    # ------------------------------------------------------------------
    add_case(cases, prefix="pi", tags=["math", "PI"], formula="=PI()")
    add_case(cases, prefix="sin", tags=["math", "SIN", "PI"], formula="=SIN(PI())")
    add_case(cases, prefix="cos", tags=["math", "COS"], formula="=COS(0)")
    add_case(cases, prefix="tan", tags=["math", "TAN"], formula="=TAN(0)")
    add_case(cases, prefix="acos", tags=["math", "ACOS"], formula="=ACOS(1)")
    add_case(cases, prefix="asin", tags=["math", "ASIN"], formula="=ASIN(0)")
    add_case(cases, prefix="atan", tags=["math", "ATAN"], formula="=ATAN(1)")
    add_case(cases, prefix="atan2", tags=["math", "ATAN2"], formula="=ATAN2(1,1)")
    add_case(cases, prefix="exp_ln", tags=["math", "EXP", "LN"], formula="=LN(EXP(1))")
    add_case(cases, prefix="log", tags=["math", "LOG"], formula="=LOG(100,10)")
    add_case(cases, prefix="log10", tags=["math", "LOG10"], formula="=LOG10(100)")
    add_case(cases, prefix="power", tags=["math", "POWER"], formula="=POWER(2,3)")
    add_case(cases, prefix="sqrt", tags=["math", "SQRT"], formula="=SQRT(4)")
    add_case(cases, prefix="product", tags=["math", "PRODUCT"], formula="=PRODUCT(1,2,3)")
    add_case(cases, prefix="sumsq", tags=["math", "SUMSQ"], formula="=SUMSQ(1,2,3)")
    add_case(cases, prefix="trunc", tags=["math", "TRUNC"], formula="=TRUNC(3.14159,2)")
    add_case(cases, prefix="ceiling", tags=["math", "CEILING"], formula="=CEILING(1.2,1)")
    add_case(cases, prefix="ceiling_math", tags=["math", "CEILING.MATH"], formula="=CEILING.MATH(1.2,1)")
    add_case(cases, prefix="ceiling_precise", tags=["math", "CEILING.PRECISE"], formula="=CEILING.PRECISE(1.2,1)")
    add_case(cases, prefix="iso_ceiling", tags=["math", "ISO.CEILING"], formula="=ISO.CEILING(1.2,1)")
    add_case(cases, prefix="floor", tags=["math", "FLOOR"], formula="=FLOOR(1.2,1)")
    add_case(cases, prefix="floor_math", tags=["math", "FLOOR.MATH"], formula="=FLOOR.MATH(1.2,1)")
    add_case(cases, prefix="floor_precise", tags=["math", "FLOOR.PRECISE"], formula="=FLOOR.PRECISE(1.2,1)")
    add_case(cases, prefix="roman", tags=["math", "ROMAN"], formula="=ROMAN(499,0)")
    add_case(cases, prefix="arabic", tags=["math", "ARABIC"], formula='=ARABIC("MCMXCIX")')
    add_case(cases, prefix="mdeterm", tags=["math", "MDETERM", "matrix"], formula="=MDETERM({1,2;3,4})")

    # Math & trig backfill (more).
    add_case(cases, prefix="radians", tags=["math", "RADIANS"], formula="=RADIANS(180)")
    add_case(cases, prefix="degrees", tags=["math", "DEGREES"], formula="=DEGREES(PI())")
    add_case(cases, prefix="sinh", tags=["math", "SINH"], formula="=SINH(0)")
    add_case(cases, prefix="cosh", tags=["math", "COSH"], formula="=COSH(0)")
    add_case(cases, prefix="tanh", tags=["math", "TANH"], formula="=TANH(0)")
    add_case(cases, prefix="asinh", tags=["math", "ASINH"], formula="=ASINH(0)")
    add_case(cases, prefix="acosh", tags=["math", "ACOSH"], formula="=ACOSH(1)")
    add_case(cases, prefix="atanh", tags=["math", "ATANH"], formula="=ATANH(0)")
    add_case(cases, prefix="cot", tags=["math", "COT"], formula="=COT(1)")
    add_case(cases, prefix="csc", tags=["math", "CSC"], formula="=CSC(1)")
    add_case(cases, prefix="sec", tags=["math", "SEC"], formula="=SEC(1)")
    add_case(cases, prefix="acot", tags=["math", "ACOT"], formula="=ACOT(1)")
    add_case(cases, prefix="coth", tags=["math", "COTH"], formula="=COTH(1)")
    add_case(cases, prefix="csch", tags=["math", "CSCH"], formula="=CSCH(1)")
    add_case(cases, prefix="sech", tags=["math", "SECH"], formula="=SECH(1)")
    add_case(cases, prefix="acoth", tags=["math", "ACOTH"], formula="=ACOTH(2)")
    add_case(cases, prefix="fact", tags=["math", "FACT"], formula="=FACT(5)")
    add_case(cases, prefix="factdouble", tags=["math", "FACTDOUBLE"], formula="=FACTDOUBLE(6)")
    add_case(cases, prefix="combin", tags=["math", "COMBIN"], formula="=COMBIN(5,2)")
    add_case(cases, prefix="combina", tags=["math", "COMBINA"], formula="=COMBINA(3,2)")
    add_case(cases, prefix="permut", tags=["math", "PERMUT"], formula="=PERMUT(5,2)")
    add_case(cases, prefix="permutationa", tags=["math", "PERMUTATIONA"], formula="=PERMUTATIONA(3,2)")
    add_case(cases, prefix="gcd", tags=["math", "GCD"], formula="=GCD(24,36)")
    add_case(cases, prefix="lcm", tags=["math", "LCM"], formula="=LCM(4,6)")
    add_case(cases, prefix="multinomial", tags=["math", "MULTINOMIAL"], formula="=MULTINOMIAL(1,2,3)")
    add_case(cases, prefix="mround", tags=["math", "MROUND"], formula="=MROUND(10,3)")
    add_case(cases, prefix="even", tags=["math", "EVEN"], formula="=EVEN(1)")
    add_case(cases, prefix="odd", tags=["math", "ODD"], formula="=ODD(2)")
    add_case(cases, prefix="iseven", tags=["math", "ISEVEN"], formula="=ISEVEN(2)")
    add_case(cases, prefix="isodd", tags=["math", "ISODD"], formula="=ISODD(3)")
    add_case(cases, prefix="quotient", tags=["math", "QUOTIENT"], formula="=QUOTIENT(5,2)")
    add_case(cases, prefix="sqrtpi", tags=["math", "SQRTPI"], formula="=SQRTPI(1)")
    add_case(cases, prefix="delta", tags=["math", "DELTA"], formula="=DELTA(0)")
    add_case(cases, prefix="gestep", tags=["math", "GESTEP"], formula="=GESTEP(1,0)")
    add_case(cases, prefix="seriessum", tags=["math", "SERIESSUM"], formula="=SERIESSUM(2,0,1,{1,2,3})")
    add_case(cases, prefix="sumxmy2", tags=["math", "SUMXMY2"], formula="=SUMXMY2({1,2},{3,4})")
    add_case(cases, prefix="sumx2my2", tags=["math", "SUMX2MY2"], formula="=SUMX2MY2({1,2},{3,4})")
    add_case(cases, prefix="sumx2py2", tags=["math", "SUMX2PY2"], formula="=SUMX2PY2({1,2},{3,4})")

    # Error/coercion semantics backfill.
    add_case(cases, prefix="acosh", tags=["math", "ACOSH", "error"], formula="=ACOSH(0.5)")
    add_case(cases, prefix="atanh", tags=["math", "ATANH", "error"], formula="=ATANH(1)")
    add_case(cases, prefix="acoth", tags=["math", "ACOTH", "error"], formula="=ACOTH(1)")
    add_case(cases, prefix="coth", tags=["math", "COTH", "error"], formula="=COTH(0)")
    add_case(cases, prefix="csch", tags=["math", "CSCH", "error"], formula="=CSCH(0)")
    add_case(cases, prefix="cot", tags=["math", "COT", "error"], formula="=COT(0)")
    add_case(cases, prefix="csc", tags=["math", "CSC", "error"], formula="=CSC(0)")
    add_case(cases, prefix="fact", tags=["math", "FACT", "error"], formula="=FACT(-1)")
    add_case(cases, prefix="combin", tags=["math", "COMBIN", "error"], formula="=COMBIN(5,7)")
    add_case(cases, prefix="permut", tags=["math", "PERMUT", "error"], formula="=PERMUT(5,7)")
    add_case(cases, prefix="gcd", tags=["math", "GCD", "error"], formula="=GCD(-2,4)")
    add_case(cases, prefix="mround", tags=["math", "MROUND", "error"], formula="=MROUND(-10,3)")
    add_case(cases, prefix="quotient", tags=["math", "QUOTIENT", "error"], formula="=QUOTIENT(5,0)")
    add_case(cases, prefix="sqrtpi", tags=["math", "SQRTPI", "error"], formula="=SQRTPI(-1)")
    add_case(cases, prefix="seriessum", tags=["math", "SERIESSUM", "error"], formula="=SERIESSUM(0,-1,1,{1})")
    add_case(cases, prefix="sumxmy2", tags=["math", "SUMXMY2", "error"], formula="=SUMXMY2({1,2},{3})")
    add_case(
        cases,
        prefix="sumxmy2",
        tags=["math", "SUMXMY2", "coercion"],
        formula='=SUMXMY2({1,"x",3},{1,2,3})',
        description="SUMX* treats text as 0 within arrays",
    )
