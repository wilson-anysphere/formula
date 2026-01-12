from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Engineering functions (complex + special math)
    # ------------------------------------------------------------------
    add_case(cases, prefix="complex", tags=["engineering", "COMPLEX"], formula="=COMPLEX(3,4)")
    add_case(cases, prefix="imabs", tags=["engineering", "IMABS"], formula='=IMABS("3+4i")')
    add_case(
        cases,
        prefix="imaginary",
        tags=["engineering", "IMAGINARY"],
        formula='=IMAGINARY("3+4i")',
    )
    add_case(cases, prefix="imreal", tags=["engineering", "IMREAL"], formula='=IMREAL("3+4i")')
    add_case(
        cases,
        prefix="imargument",
        tags=["engineering", "IMARGUMENT"],
        formula='=IMARGUMENT("1+i")',
    )
    add_case(
        cases,
        prefix="imconjugate",
        tags=["engineering", "IMCONJUGATE"],
        formula='=IMCONJUGATE("3+4i")',
    )
    add_case(cases, prefix="imsum", tags=["engineering", "IMSUM"], formula='=IMSUM("1+i","1-i")')
    add_case(
        cases,
        prefix="improduct",
        tags=["engineering", "IMPRODUCT"],
        formula='=IMPRODUCT("1+i","1-i")',
    )
    add_case(cases, prefix="imsub", tags=["engineering", "IMSUB"], formula='=IMSUB("3+4i","1+2i")')
    add_case(cases, prefix="imdiv", tags=["engineering", "IMDIV"], formula='=IMDIV("1+i","1-i")')
    add_case(cases, prefix="impower", tags=["engineering", "IMPOWER"], formula='=IMPOWER("i",2)')
    add_case(cases, prefix="imsqrt", tags=["engineering", "IMSQRT"], formula='=IMSQRT("-1")')
    add_case(cases, prefix="imln", tags=["engineering", "IMLN"], formula='=IMLN("1")')
    add_case(cases, prefix="imlog2", tags=["engineering", "IMLOG2"], formula='=IMLOG2("2")')
    add_case(cases, prefix="imlog10", tags=["engineering", "IMLOG10"], formula='=IMLOG10("100")')
    add_case(cases, prefix="imsin", tags=["engineering", "IMSIN"], formula='=IMSIN("0")')
    add_case(cases, prefix="imcos", tags=["engineering", "IMCOS"], formula='=IMCOS("0")')
    add_case(cases, prefix="imexp", tags=["engineering", "IMEXP"], formula='=IMEXP("0")')

    add_case(cases, prefix="erf", tags=["engineering", "ERF"], formula="=ERF(1)")
    add_case(cases, prefix="erfc", tags=["engineering", "ERFC"], formula="=ERFC(1)")
    add_case(cases, prefix="besselj", tags=["engineering", "BESSELJ"], formula="=BESSELJ(1,0)")
    add_case(cases, prefix="bessely", tags=["engineering", "BESSELY"], formula="=BESSELY(1,0)")
    add_case(cases, prefix="besseli", tags=["engineering", "BESSELI"], formula="=BESSELI(1,0)")
    add_case(cases, prefix="besselk", tags=["engineering", "BESSELK"], formula="=BESSELK(1,0)")

    add_case(
        cases,
        prefix="convert",
        tags=["engineering", "CONVERT"],
        formula='=CONVERT(1,"m","ft")',
    )
