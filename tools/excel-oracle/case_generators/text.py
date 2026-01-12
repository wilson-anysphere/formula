from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
    excel_serial_1900,
) -> None:
    # ------------------------------------------------------------------
    # Text functions (keep cases deterministic; avoid locale-dependent parsing where possible)
    # ------------------------------------------------------------------
    strings = ["", "a", "foo", "Hello", "12345", "a b c", "This is a test", "こんにちは"]
    num_chars = [0, 1, 2, 3, 5]
    for s in strings:
        add_case(cases, prefix="len", tags=["text", "LEN"], formula="=LEN(A1)", inputs=[CellInput("A1", s)])
        for n in num_chars:
            add_case(
                cases,
                prefix="left",
                tags=["text", "LEFT"],
                formula="=LEFT(A1,B1)",
                inputs=[CellInput("A1", s), CellInput("B1", n)],
            )
            add_case(
                cases,
                prefix="right",
                tags=["text", "RIGHT"],
                formula="=RIGHT(A1,B1)",
                inputs=[CellInput("A1", s), CellInput("B1", n)],
            )

    mid_starts = [1, 2, 3]
    mid_lens = [0, 1, 2]
    for s in strings:
        for start in mid_starts:
            for ln in mid_lens:
                add_case(
                    cases,
                    prefix="mid",
                    tags=["text", "MID"],
                    formula="=MID(A1,B1,C1)",
                    inputs=[CellInput("A1", s), CellInput("B1", start), CellInput("C1", ln)],
                    output_cell="D1",
                )

    # CONCATENATE is legacy but widely used; CONCAT exists in newer Excel.
    for a in strings:
        for b in ["", "X", "123"]:
            add_case(
                cases,
                prefix="concat",
                tags=["text", "CONCATENATE"],
                formula="=CONCATENATE(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    # FIND / SEARCH differences: FIND is case-sensitive, SEARCH is not.
    find_haystacks = ["foobar", "FooBar", "abcabc"]
    find_needles = ["foo", "Foo", "bar", "z"]
    for needle in find_needles:
        for hay in find_haystacks:
            add_case(
                cases,
                prefix="find",
                tags=["text", "FIND"],
                formula="=FIND(A1,B1)",
                inputs=[CellInput("A1", needle), CellInput("B1", hay)],
            )
            add_case(
                cases,
                prefix="search",
                tags=["text", "SEARCH"],
                formula="=SEARCH(A1,B1)",
                inputs=[CellInput("A1", needle), CellInput("B1", hay)],
            )

    # SUBSTITUTE
    for s in ["foo bar foo", "aaaa", "123123", ""]:
        add_case(
            cases,
            prefix="substitute",
            tags=["text", "SUBSTITUTE"],
            formula='=SUBSTITUTE(A1,"foo","x")',
            inputs=[CellInput("A1", s)],
        )

    # Additional text functions
    add_case(
        cases,
        prefix="clean",
        tags=["text", "CLEAN"],
        formula="=CLEAN(A1)",
        inputs=[CellInput("A1", "a\u0000\u0009b\u001Fc\u007Fd")],
        description="CLEAN strips non-printable ASCII control codes",
    )
    add_case(
        cases,
        prefix="trim",
        tags=["text", "TRIM"],
        formula="=TRIM(A1)",
        inputs=[CellInput("A1", "  a   b  ")],
    )
    add_case(
        cases,
        prefix="trim",
        tags=["text", "TRIM"],
        formula="=TRIM(A1)",
        inputs=[CellInput("A1", "\ta  b")],
        description="TRIM collapses spaces but preserves tabs",
    )
    add_case(cases, prefix="upper", tags=["text", "UPPER"], formula="=UPPER(A1)", inputs=[CellInput("A1", "Abc")])
    add_case(cases, prefix="lower", tags=["text", "LOWER"], formula="=LOWER(A1)", inputs=[CellInput("A1", "AbC")])
    add_case(
        cases,
        prefix="proper",
        tags=["text", "PROPER"],
        formula="=PROPER(A1)",
        inputs=[CellInput("A1", "hELLO wORLD")],
    )
    add_case(
        cases,
        prefix="exact",
        tags=["text", "EXACT"],
        formula="=EXACT(A1,B1)",
        inputs=[CellInput("A1", "Hello"), CellInput("B1", "hello")],
    )
    add_case(
        cases,
        prefix="exact",
        tags=["text", "EXACT"],
        formula="=EXACT(A1,B1)",
        inputs=[CellInput("A1", "Hello"), CellInput("B1", "Hello")],
    )
    add_case(cases, prefix="replace", tags=["text", "REPLACE"], formula='=REPLACE("abcdef",2,3,"X")')
    add_case(cases, prefix="replace", tags=["text", "REPLACE"], formula='=REPLACE("abc",5,1,"X")')

    # Legacy DBCS / byte-count text functions.
    #
    # In single-byte locales (en-US), these behave identically to their non-`B`
    # equivalents. In DBCS locales they become locale/codepage-dependent.
    add_case(cases, prefix="lenb", tags=["text", "LENB"], formula='=LENB("abc")')
    add_case(cases, prefix="leftb", tags=["text", "LEFTB"], formula='=LEFTB("abc",2)')
    add_case(cases, prefix="rightb", tags=["text", "RIGHTB"], formula='=RIGHTB("abc",2)')
    add_case(cases, prefix="midb", tags=["text", "MIDB"], formula='=MIDB("abc",2,2)')
    add_case(cases, prefix="findb", tags=["text", "FINDB"], formula='=FINDB("b","abc")')
    add_case(cases, prefix="searchb", tags=["text", "SEARCHB"], formula='=SEARCHB("B","abc")')
    add_case(cases, prefix="replaceb", tags=["text", "REPLACEB"], formula='=REPLACEB("abcdef",2,3,"X")')
    add_case(cases, prefix="asc", tags=["text", "ASC"], formula='=ASC("ABC")')
    add_case(cases, prefix="dbcs", tags=["text", "DBCS"], formula='=DBCS("ABC")')
    add_case(
        cases,
        prefix="phonetic",
        tags=["text", "PHONETIC"],
        formula="=PHONETIC(A1)",
        inputs=[CellInput("A1", "abc")],
    )

    # Thai localization functions (deterministic, locale-independent).
    add_case(cases, prefix="thai", tags=["thai", "BAHTTEXT"], formula="=BAHTTEXT(1234.5)")
    add_case(cases, prefix="thai", tags=["thai", "BAHTTEXT"], formula="=BAHTTEXT(0)")
    add_case(cases, prefix="thai", tags=["thai", "BAHTTEXT"], formula="=BAHTTEXT(-11.25)")
    add_case(cases, prefix="thai", tags=["thai", "THAIDIGIT"], formula='=THAIDIGIT("123")')
    # Use an integer input to keep this locale-independent across decimal separator differences.
    add_case(cases, prefix="thai", tags=["thai", "THAIDIGIT"], formula="=THAIDIGIT(1234)")
    add_case(
        cases,
        prefix="thai",
        tags=["thai", "ISTHAIDIGIT", "THAIDIGIT"],
        formula='=ISTHAIDIGIT(THAIDIGIT("123"))',
    )
    add_case(cases, prefix="thai", tags=["thai", "ISTHAIDIGIT"], formula='=ISTHAIDIGIT("๑๒๓")')
    add_case(cases, prefix="thai", tags=["thai", "THAINUMSTRING"], formula="=THAINUMSTRING(1234.5)")
    add_case(cases, prefix="thai", tags=["thai", "THAINUMSTRING"], formula="=THAINUMSTRING(-1234.5)")
    add_case(cases, prefix="thai", tags=["thai", "THAINUMSOUND"], formula="=THAINUMSOUND(1234.5)")
    add_case(cases, prefix="thai", tags=["thai", "THAINUMSOUND"], formula="=THAINUMSOUND(-1234.5)")
    add_case(cases, prefix="thai", tags=["thai", "THAISTRINGLENGTH"], formula='=THAISTRINGLENGTH("เก้า")')
    add_case(cases, prefix="thai", tags=["thai", "ROUNDBAHTDOWN"], formula="=ROUNDBAHTDOWN(1.26)")
    add_case(cases, prefix="thai", tags=["thai", "ROUNDBAHTUP"], formula="=ROUNDBAHTUP(1.26)")
    add_case(cases, prefix="thai", tags=["thai", "ROUNDBAHTDOWN"], formula="=ROUNDBAHTDOWN(-1.26)")
    add_case(cases, prefix="thai", tags=["thai", "ROUNDBAHTUP"], formula="=ROUNDBAHTUP(-1.26)")
    add_case(
        cases,
        prefix="thai",
        tags=["thai", "THAIDAYOFWEEK", "DATE"],
        formula="=THAIDAYOFWEEK(DATE(2020,1,1))",
    )
    add_case(
        cases,
        prefix="thai",
        tags=["thai", "THAIDAYOFWEEK", "DATE"],
        formula="=THAIDAYOFWEEK(DATE(2020,1,5))",
    )
    add_case(
        cases,
        prefix="thai",
        tags=["thai", "THAIMONTHOFYEAR", "DATE"],
        formula="=THAIMONTHOFYEAR(DATE(2020,1,1))",
    )
    add_case(
        cases,
        prefix="thai",
        tags=["thai", "THAIMONTHOFYEAR", "DATE"],
        formula="=THAIMONTHOFYEAR(DATE(2020,12,31))",
    )
    add_case(cases, prefix="thai", tags=["thai", "THAIYEAR", "DATE"], formula="=THAIYEAR(DATE(2020,1,1))")
    add_case(cases, prefix="thai", tags=["thai", "THAIYEAR", "DATE"], formula="=THAIYEAR(DATE(1900,1,1))")

    # CONCAT (unlike CONCATENATE, CONCAT flattens ranges)
    add_case(
        cases,
        prefix="concat_new",
        tags=["text", "CONCAT"],
        formula='=CONCAT(A1:A2,"c")',
        inputs=[CellInput("A1", "a"), CellInput("A2", "b")],
    )

    # TEXTJOIN
    textjoin_inputs = [
        CellInput("A1", "a"),
        CellInput("A2", None),
        CellInput("A3", ""),
        CellInput("A4", 1),
    ]
    add_case(
        cases,
        prefix="textjoin",
        tags=["text", "TEXTJOIN"],
        formula='=TEXTJOIN(",",TRUE,A1:A4)',
        inputs=textjoin_inputs,
    )
    add_case(
        cases,
        prefix="textjoin",
        tags=["text", "TEXTJOIN"],
        formula='=TEXTJOIN(",",FALSE,A1:A4)',
        inputs=textjoin_inputs,
    )

    # TEXTSPLIT (dynamic array)
    add_case(
        cases,
        prefix="textsplit_basic",
        tags=["text", "TEXTSPLIT"],
        formula='=TEXTSPLIT("a,b,c",",")',
    )

    # TEXT / VALUE / NUMBERVALUE / DOLLAR
    # TEXT formatting: keep these cases locale-independent by avoiding thousand/decimal separators
    # and currency symbols in the *result* string.
    add_case(cases, prefix="text_fmt", tags=["text", "TEXT"], formula='=TEXT(1234.567,"0")')
    add_case(cases, prefix="text_pct", tags=["text", "TEXT"], formula='=TEXT(1.23,"0%")')
    add_case(cases, prefix="text_int", tags=["text", "TEXT"], formula='=TEXT(-1,"0")')
    add_case(cases, prefix="value", tags=["text", "VALUE", "coercion"], formula='=VALUE("1234")')
    add_case(cases, prefix="value", tags=["text", "VALUE"], formula='=VALUE("(1000)")')
    add_case(cases, prefix="value", tags=["text", "VALUE"], formula='=VALUE("10%")')
    add_case(cases, prefix="value", tags=["text", "VALUE", "error"], formula='=VALUE("nope")')
    add_case(
        cases,
        prefix="numbervalue",
        tags=["text", "NUMBERVALUE", "coercion"],
        formula='=NUMBERVALUE("1.234,5", ",", ".")',
    )
    add_case(
        cases,
        prefix="numbervalue",
        tags=["text", "NUMBERVALUE", "error"],
        formula='=NUMBERVALUE("1,23", ",", ",")',
    )
    # DOLLAR returns a localized currency string. Wrap in N(...) so the case result is
    # locale-independent while still exercising the function.
    add_case(cases, prefix="dollar", tags=["text", "DOLLAR"], formula="=N(DOLLAR(1234.567,2))")
    add_case(cases, prefix="dollar", tags=["text", "DOLLAR"], formula="=N(DOLLAR(-1234.567,2))")

    # TEXT: Excel number format codes (dates/sections/conditions). These are
    # locale-sensitive, but high-signal for Excel compatibility.
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format"],
        formula=r'=TEXT(A1,"0")',
        inputs=[CellInput("A1", 1234.567)],
    )
    multi_section = '"0.00;(0.00);""zero"";""text:""@"'
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", 1.2)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", -1.2)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", 0)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", "hi")],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "conditions"],
        formula=r'=TEXT(A1,"[<0]""neg"";""pos""")',
        inputs=[CellInput("A1", -1)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "conditions"],
        formula=r'=TEXT(A1,"[<0]""neg"";""pos""")',
        inputs=[CellInput("A1", 1)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "date"],
        formula=r'=TEXT(A1,"yyyy-mm-dd")',
        inputs=[CellInput("A1", excel_serial_1900(2024, 1, 10))],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "date", "time"],
        formula=r'=TEXT(A1,"yyyy-mm-dd hh:mm")',
        inputs=[CellInput("A1", excel_serial_1900(2024, 1, 10) + 0.5)],
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "locale"],
        formula=r'=TEXT(A1,"[$€-407]0")',
        inputs=[CellInput("A1", 1)],
        description="Currency symbol bracket token + locale code (LCID 0x0407 de-DE)",
    )
    add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "invalid"],
        formula=r'=TEXT(A1,"")',
        inputs=[CellInput("A1", 1234.5)],
        description="Empty format_text should fall back to General (Excel behavior)",
    )
