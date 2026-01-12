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

    # ------------------------------------------------------------------
    # Engineering functions (base conversions + bitwise)
    # ------------------------------------------------------------------
    add_case(cases, prefix="bin2dec", tags=["engineering", "BIN2DEC"], formula='=BIN2DEC("1010")')
    add_case(cases, prefix="bin2oct", tags=["engineering", "BIN2OCT"], formula='=BIN2OCT("1010")')
    add_case(cases, prefix="bin2hex", tags=["engineering", "BIN2HEX"], formula='=BIN2HEX("1010")')

    add_case(cases, prefix="oct2dec", tags=["engineering", "OCT2DEC"], formula='=OCT2DEC("17")')
    add_case(cases, prefix="oct2bin", tags=["engineering", "OCT2BIN"], formula='=OCT2BIN("17")')
    add_case(cases, prefix="oct2hex", tags=["engineering", "OCT2HEX"], formula='=OCT2HEX("17")')

    add_case(cases, prefix="hex2dec", tags=["engineering", "HEX2DEC"], formula='=HEX2DEC("FF")')
    add_case(cases, prefix="hex2bin", tags=["engineering", "HEX2BIN"], formula='=HEX2BIN("FF")')
    add_case(cases, prefix="hex2oct", tags=["engineering", "HEX2OCT"], formula='=HEX2OCT("FF")')

    add_case(cases, prefix="dec2bin", tags=["engineering", "DEC2BIN"], formula="=DEC2BIN(10)")
    add_case(cases, prefix="dec2oct", tags=["engineering", "DEC2OCT"], formula="=DEC2OCT(10)")
    add_case(cases, prefix="dec2hex", tags=["engineering", "DEC2HEX"], formula="=DEC2HEX(10)")

    add_case(cases, prefix="base", tags=["engineering", "BASE"], formula="=BASE(15,16,4)")
    add_case(cases, prefix="decimal", tags=["engineering", "DECIMAL"], formula='=DECIMAL("FF",16)')

    add_case(cases, prefix="bitand", tags=["engineering", "BITAND"], formula="=BITAND(5,3)")
    add_case(cases, prefix="bitor", tags=["engineering", "BITOR"], formula="=BITOR(5,3)")
    add_case(cases, prefix="bitxor", tags=["engineering", "BITXOR"], formula="=BITXOR(5,3)")
    add_case(
        cases, prefix="bitlshift", tags=["engineering", "BITLSHIFT"], formula="=BITLSHIFT(1,3)"
    )
    add_case(
        cases, prefix="bitrshift", tags=["engineering", "BITRSHIFT"], formula="=BITRSHIFT(8,3)"
    )
