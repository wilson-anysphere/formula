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
    # Financial functions (P1)
    # ------------------------------------------------------------------
    # Keep this section relatively small; it's mainly intended to cover
    # algorithmic edge cases and common parameter combinations.
    add_case(cases, prefix="pv", tags=["financial", "PV"], formula="=PV(0, 3, -10)")
    add_case(cases, prefix="pv", tags=["financial", "PV"], formula="=PV(0.05, 10, -100)")
    add_case(cases, prefix="fv", tags=["financial", "FV"], formula="=FV(0, 5, -10)")
    add_case(cases, prefix="fv", tags=["financial", "FV"], formula="=FV(0.05, 10, -100)")
    add_case(
        cases,
        prefix="fvschedule",
        tags=["financial", "FVSCHEDULE"],
        formula="=FVSCHEDULE(100,A1:A3)",
        inputs=[CellInput("A1", 0.1), CellInput("A2", 0.2), CellInput("A3", 0.3)],
        output_cell="C1",
    )
    add_case(cases, prefix="pmt", tags=["financial", "PMT"], formula="=PMT(0, 2, 10)")
    add_case(cases, prefix="pmt", tags=["financial", "PMT"], formula="=PMT(0.05, 10, 1000)")
    add_case(cases, prefix="ipmt", tags=["financial", "IPMT"], formula="=IPMT(0.05, 1, 10, 1000)")
    add_case(cases, prefix="ipmt", tags=["financial", "IPMT"], formula="=IPMT(0.05, 10, 10, 1000, 0, 1)")
    add_case(cases, prefix="ppmt", tags=["financial", "PPMT"], formula="=PPMT(0.05, 1, 10, 1000)")
    add_case(cases, prefix="ppmt", tags=["financial", "PPMT"], formula="=PPMT(0.05, 10, 10, 1000, 0, 1)")
    add_case(
        cases,
        prefix="cumipmt",
        tags=["financial", "CUMIPMT"],
        formula="=CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0)",
        description="Excel docs amortization accumulation example",
    )
    add_case(
        cases,
        prefix="cumprinc",
        tags=["financial", "CUMPRINC"],
        formula="=CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0)",
        description="Excel docs amortization accumulation example",
    )
    add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0, -10, 100)")
    add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0.05, -100, 1000)")
    add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(10, -100, 1000)")
    add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(12, -50, 500)")
    add_case(cases, prefix="effect", tags=["financial", "EFFECT"], formula="=EFFECT(0.1,12)")
    # Integer-coercion variants for the `npery` argument (same shape as bond `frequency`).
    add_case(
        cases,
        prefix="effect",
        tags=["financial", "EFFECT", "coercion"],
        formula="=EFFECT(0.1,2.9)",
        description="npery=2.9",
    )
    add_case(
        cases,
        prefix="effect",
        tags=["financial", "EFFECT", "coercion"],
        formula="=EFFECT(0.1,1.1)",
        description="npery=1.1",
    )
    add_case(
        cases,
        prefix="effect",
        tags=["financial", "EFFECT", "coercion"],
        formula="=EFFECT(0.1,1.999999999)",
        description="npery=1.999999999",
    )
    add_case(cases, prefix="nominal", tags=["financial", "NOMINAL"], formula="=NOMINAL(0.1,12)")
    add_case(
        cases,
        prefix="nominal",
        tags=["financial", "NOMINAL", "coercion"],
        formula="=NOMINAL(0.1,2.9)",
        description="npery=2.9",
    )
    add_case(
        cases,
        prefix="nominal",
        tags=["financial", "NOMINAL", "coercion"],
        formula="=NOMINAL(0.1,1.1)",
        description="npery=1.1",
    )
    add_case(cases, prefix="rri", tags=["financial", "RRI"], formula="=RRI(10,-100,200)")
    add_case(cases, prefix="pduration", tags=["financial", "PDURATION"], formula="=PDURATION(0.025,2000,2200)")
    coup_settlement = excel_serial_1900(2024, 6, 15)
    coup_maturity = excel_serial_1900(2025, 1, 1)
    coup_inputs = [CellInput("A1", coup_settlement), CellInput("A2", coup_maturity)]
    add_case(
        cases,
        prefix="coupdaybs",
        tags=["financial", "COUPDAYBS"],
        formula="=COUPDAYBS(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    add_case(
        cases,
        prefix="coupdays",
        tags=["financial", "COUPDAYS"],
        formula="=COUPDAYS(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    add_case(
        cases,
        prefix="coupdaysnc",
        tags=["financial", "COUPDAYSNC"],
        formula="=COUPDAYSNC(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    add_case(
        cases,
        prefix="coupncd",
        tags=["financial", "COUPNCD"],
        formula="=COUPNCD(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    add_case(
        cases,
        prefix="coupnum",
        tags=["financial", "COUPNUM"],
        formula="=COUPNUM(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    add_case(
        cases,
        prefix="couppcd",
        tags=["financial", "COUPPCD"],
        formula="=COUPPCD(A1,A2,2,0)",
        inputs=coup_inputs,
    )
    # Integer-coercion variants for `frequency` and `basis` (truncate toward zero).
    add_case(
        cases,
        prefix="coupdaybs",
        tags=["financial", "COUPDAYBS", "coercion"],
        formula="=COUPDAYBS(A1,A2,2.9,0)",
        inputs=coup_inputs,
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="coupncd",
        tags=["financial", "COUPNCD", "coercion"],
        formula="=COUPNCD(A1,A2,2,0.9)",
        inputs=coup_inputs,
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="fvschedule",
        tags=["financial", "FVSCHEDULE"],
        formula="=FVSCHEDULE(100,{0.1,0.2})",
    )
    add_case(
        cases,
        prefix="fvschedule",
        tags=["financial", "FVSCHEDULE"],
        formula="=FVSCHEDULE(100,A1:A2)",
        inputs=[CellInput("A1", 0.1), CellInput("A2", 0.2)],
    )
    add_case(
        cases,
        prefix="fvschedule",
        tags=["financial", "FVSCHEDULE"],
        formula="=FVSCHEDULE(100,(A1:A2,B1:B2))",
        inputs=[
            CellInput("A1", 0.1),
            CellInput("A2", 0.2),
            CellInput("B1", 0.3),
            CellInput("B2", 0.4),
        ],
    )
    add_case(cases, prefix="sln", tags=["financial", "SLN"], formula="=SLN(30, 0, 3)")
    add_case(cases, prefix="syd", tags=["financial", "SYD"], formula="=SYD(30, 0, 3, 1)")
    add_case(cases, prefix="ddb", tags=["financial", "DDB"], formula="=DDB(1000, 100, 5, 1)")
    add_case(cases, prefix="ispmt", tags=["financial", "ISPMT"], formula="=ISPMT(0.1, 1, 3, 300)")
    add_case(cases, prefix="dollarde", tags=["financial", "DOLLARDE"], formula="=DOLLARDE(1.02, 16)")
    add_case(cases, prefix="dollarfr", tags=["financial", "DOLLARFR"], formula="=DOLLARFR(1.125, 16)")
    add_case(cases, prefix="db", tags=["financial", "DB"], formula="=DB(10000, 1000, 5, 1)")
    add_case(cases, prefix="db", tags=["financial", "DB"], formula="=DB(10000, 1000, 5, 1, 7)")
    add_case(cases, prefix="db", tags=["financial", "DB"], formula="=DB(10000, 1000, 5, 6, 7)")
    add_case(cases, prefix="vdb", tags=["financial", "VDB"], formula="=VDB(2400, 300, 10, 0, 1)")
    add_case(cases, prefix="vdb", tags=["financial", "VDB"], formula="=VDB(2400, 300, 10, 0, 0.5)")
    add_case(cases, prefix="vdb", tags=["financial", "VDB"], formula="=VDB(2400, 0, 10, 6, 10, 2, FALSE)")
    add_case(cases, prefix="vdb", tags=["financial", "VDB"], formula="=VDB(2400, 0, 10, 6, 10, 2, TRUE)")

    # Coupon date helper functions (COUP*).
    add_case(
        cases,
        prefix="coupdaybs",
        tags=["financial", "COUPDAYBS"],
        formula="=COUPDAYBS(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )
    add_case(
        cases,
        prefix="coupdays",
        tags=["financial", "COUPDAYS"],
        formula="=COUPDAYS(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )
    add_case(
        cases,
        prefix="coupdaysnc",
        tags=["financial", "COUPDAYSNC"],
        formula="=COUPDAYSNC(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )
    add_case(
        cases,
        prefix="coupncd",
        tags=["financial", "COUPNCD"],
        formula="=COUPNCD(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )
    add_case(
        cases,
        prefix="coupnum",
        tags=["financial", "COUPNUM"],
        formula="=COUPNUM(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )
    add_case(
        cases,
        prefix="couppcd",
        tags=["financial", "COUPPCD"],
        formula="=COUPPCD(DATE(2020,3,1),DATE(2025,1,15),2,0)",
    )

    # Coupon schedule edge cases around month-end dates that are not the 31st.
    #
    # Excel's coupon schedule rules are subtle for maturities like Apr 30 or Feb 28/29; these cases
    # are intended to disambiguate whether Excel treats such maturities as an explicit end-of-month
    # schedule (pinned to month end) or simply as an EDATE-style day-of-month schedule.
    #
    # These are especially important for basis=1 where `COUPDAYS` depends on the computed PCD/NCD
    # dates (actual day count between coupon dates).
    add_case(
        cases,
        prefix="couppcd_eom_apr30",
        tags=["financial", "COUPPCD", "coupon_schedule", "eom_edge"],
        formula="=COUPPCD(DATE(2020,2,15),DATE(2020,4,30),4,1)",
        description="COUPPCD with maturity=2020-04-30 (month-end but not 31st), basis=1",
    )
    add_case(
        cases,
        prefix="coupdays_eom_apr30",
        tags=["financial", "COUPDAYS", "coupon_schedule", "eom_edge"],
        formula="=COUPDAYS(DATE(2020,2,15),DATE(2020,4,30),4,1)",
        description="COUPDAYS with maturity=2020-04-30 (month-end but not 31st), basis=1",
    )
    add_case(
        cases,
        prefix="couppcd_eom_feb28",
        tags=["financial", "COUPPCD", "coupon_schedule", "eom_edge"],
        formula="=COUPPCD(DATE(2020,11,15),DATE(2021,2,28),2,1)",
        description="COUPPCD with maturity=2021-02-28 (month-end), basis=1",
    )
    add_case(
        cases,
        prefix="coupdays_eom_feb28",
        tags=["financial", "COUPDAYS", "coupon_schedule", "eom_edge"],
        formula="=COUPDAYS(DATE(2020,11,15),DATE(2021,2,28),2,1)",
        description="COUPDAYS with maturity=2021-02-28 (month-end), basis=1",
    )

    # Basis=4 (European 30E/360) day-count edge cases.
    #
    # For basis=4, Excel uses European DAYS360 for day counts like COUPDAYBS, but models the coupon
    # period length `E` used by COUPDAYS as the fixed `360/frequency` value. Excel then computes
    # COUPDAYSNC as the remaining portion of the modeled period: `DSC = E - A`.
    #
    # This can differ from European DAYS360 between coupon dates and/or between settlement and NCD,
    # especially for end-of-month schedules involving February.
    add_case(
        cases,
        prefix="coupdaybs_b4_eom_feb28",
        tags=["financial", "COUPDAYBS", "coupon_schedule", "basis4"],
        formula="=COUPDAYBS(DATE(2020,11,15),DATE(2021,2,28),2,4)",
        description="COUPDAYBS basis=4 (European 30E/360) for an end-of-month schedule involving February",
    )
    add_case(
        cases,
        prefix="coupdays_b4_eom_feb28",
        tags=["financial", "COUPDAYS", "coupon_schedule", "basis4"],
        formula="=COUPDAYS(DATE(2020,11,15),DATE(2021,2,28),2,4)",
        description="COUPDAYS basis=4 uses fixed 360/frequency even when European DAYS360 between coupon dates differs",
    )
    add_case(
        cases,
        prefix="coupdaysnc_b4_eom_feb28",
        tags=["financial", "COUPDAYSNC", "coupon_schedule", "basis4"],
        formula="=COUPDAYSNC(DATE(2020,11,15),DATE(2021,2,28),2,4)",
        description="COUPDAYSNC basis=4 is computed as E - A (preserves additivity even when DAYS360 is not additive)",
    )

    # Discount securities / T-bills.
    add_case(
        cases,
        prefix="disc",
        tags=["financial", "DISC"],
        formula="=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
    )
    add_case(
        cases,
        prefix="disc",
        tags=["financial", "DISC"],
        formula="=DISC(45292.9,45475.1,97,100,1)",
        description="Fractional serial date inputs should be floored (2024-01-01..2024-07-02, basis=1)",
    )
    add_case(
        cases,
        prefix="pricedisc",
        tags=["financial", "PRICEDISC"],
        formula="=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100)",
    )
    add_case(
        cases,
        prefix="pricedisc",
        tags=["financial", "PRICEDISC"],
        formula="=PRICEDISC(DATE(2024,1,1),DATE(2024,7,2),0.05,100,1)",
        description="PRICEDISC with leap-year interval (183 days) using basis=1",
    )
    add_case(
        cases,
        prefix="yielddisc",
        tags=["financial", "YIELDDISC"],
        formula="=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
    )
    add_case(
        cases,
        prefix="yielddisc",
        tags=["financial", "YIELDDISC"],
        formula="=YIELDDISC(DATE(2024,1,1),DATE(2024,7,2),97,100,1)",
        description="YIELDDISC with leap-year interval (183 days) using basis=1",
    )
    add_case(
        cases,
        prefix="intrate",
        tags=["financial", "INTRATE"],
        formula="=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100)",
    )
    add_case(
        cases,
        prefix="intrate",
        tags=["financial", "INTRATE"],
        formula="=INTRATE(DATE(2024,1,1),DATE(2024,7,2),97,100,1)",
        description="INTRATE with leap-year interval (183 days) using basis=1",
    )
    add_case(
        cases,
        prefix="received",
        tags=["financial", "RECEIVED"],
        formula="=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05)",
    )
    add_case(
        cases,
        prefix="received",
        tags=["financial", "RECEIVED"],
        formula='=RECEIVED("2024-01-01","2024-07-02",95,0.05,1)',
        description="RECEIVED with ISO date text inputs using basis=1",
    )
    add_case(
        cases,
        prefix="pricemat",
        tags=["financial", "PRICEMAT"],
        formula="=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
    )
    add_case(
        cases,
        prefix="pricemat",
        tags=["financial", "PRICEMAT"],
        formula="=PRICEMAT(DATE(2024,1,1),DATE(2024,7,2),DATE(2024,1,1),0.05,0.04,1)",
        description="PRICEMAT with issue=settlement and leap-year interval using basis=1",
    )
    add_case(
        cases,
        prefix="yieldmat",
        tags=["financial", "YIELDMAT"],
        formula="=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077)",
    )
    add_case(
        cases,
        prefix="yieldmat",
        tags=["financial", "YIELDMAT"],
        formula="=YIELDMAT(DATE(2024,1,1),DATE(2024,7,2),DATE(2024,1,1),0.05,99,1)",
        description="YIELDMAT with issue=settlement and leap-year interval using basis=1",
    )
    add_case(
        cases,
        prefix="tbillprice",
        tags=["financial", "TBILLPRICE"],
        formula="=TBILLPRICE(DATE(2020,1,1),DATE(2020,7,1),0.05)",
    )
    add_case(
        cases,
        prefix="tbillyield",
        tags=["financial", "TBILLYIELD"],
        formula="=TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),97.47222222222223)",
    )
    add_case(
        cases,
        prefix="tbilleq",
        tags=["financial", "TBILLEQ"],
        formula="=TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),0.05)",
    )

    # Odd-coupon bond functions (`ODDF*` / `ODDL*`).
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,0)",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0)",
    )

    # Integer-coercion variants for `frequency` and `basis`.
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2.9,0)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1.1,0)",
        description="frequency=1.1",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1.999999999,0)",
        description="frequency=1.999999999",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE", "coercion"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2.9,0)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD", "coercion"],
        formula="=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2.9,0)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1.9)",
        description="basis=1.9",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE", "coercion"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,-0.1)",
        description="basis=-0.1",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE", "coercion"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD", "coercion"],
        formula="=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0.9)",
        description="basis=0.9",
    )

    # Long odd-stub period cases (DFC/E > 1 or DSM/E > 1).
    add_case(
        cases,
        prefix="oddfprice_long_stub",
        tags=["financial", "ODDFPRICE", "odd_coupon", "long_stub", "basis0"],
        formula="=ODDFPRICE(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),0.0785,0.0625,100,2,0)",
        description="Long odd-first coupon period (DFC/E > 1), basis=0",
    )
    add_case(
        cases,
        prefix="oddfyield_long_stub",
        tags=["financial", "ODDFYIELD", "odd_coupon", "long_stub", "basis0"],
        formula="=ODDFYIELD(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),0.0785,98,100,2,0)",
        description="Long odd-first coupon period (DFC/E > 1), basis=0",
    )
    add_case(
        cases,
        prefix="oddfprice_long_stub",
        tags=["financial", "ODDFPRICE", "odd_coupon", "long_stub", "basis1"],
        formula="=ODDFPRICE(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),0.0785,0.0625,100,2,1)",
        description="Long odd-first coupon period (DFC/E > 1), basis=1 (crosses leap day)",
    )
    add_case(
        cases,
        prefix="oddfyield_long_stub",
        tags=["financial", "ODDFYIELD", "odd_coupon", "long_stub", "basis1"],
        formula="=ODDFYIELD(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),0.0785,98,100,2,1)",
        description="Long odd-first coupon period (DFC/E > 1), basis=1 (crosses leap day)",
    )
    add_case(
        cases,
        prefix="oddlprice_long_stub",
        tags=["financial", "ODDLPRICE", "odd_coupon", "long_stub", "basis0"],
        formula="=ODDLPRICE(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)",
        description="Long odd-last coupon period (DSM/E > 1), basis=0",
    )
    add_case(
        cases,
        prefix="oddlyield_long_stub",
        tags=["financial", "ODDLYIELD", "odd_coupon", "long_stub", "basis0"],
        formula="=ODDLYIELD(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),0.0785,98,100,2,0)",
        description="Long odd-last coupon period (DSM/E > 1), basis=0",
    )
    add_case(
        cases,
        prefix="oddlprice_long_stub",
        tags=["financial", "ODDLPRICE", "odd_coupon", "long_stub", "basis1"],
        formula="=ODDLPRICE(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1)",
        description="Long odd-last coupon period (DSM/E > 1), basis=1",
    )
    add_case(
        cases,
        prefix="oddlyield_long_stub",
        tags=["financial", "ODDLYIELD", "odd_coupon", "long_stub", "basis1"],
        formula="=ODDLYIELD(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),0.0785,98,100,2,1)",
        description="Long odd-last coupon period (DSM/E > 1), basis=1",
    )

    # French accounting depreciation functions (AMOR*)
    amor_date_purchased = excel_serial_1900(2008, 8, 19)
    amor_first_period = excel_serial_1900(2008, 12, 31)
    amor_inputs = [CellInput("A1", amor_date_purchased), CellInput("A2", amor_first_period)]
    add_case(
        cases,
        prefix="amorlinc",
        tags=["financial", "AMORLINC"],
        formula="=AMORLINC(2400,A1,A2,300,1,0.15,1)",
        inputs=amor_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="amordegrec",
        tags=["financial", "AMORDEGRC"],
        formula="=AMORDEGRC(2400,A1,A2,300,1,0.15,1)",
        inputs=amor_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="price",
        tags=["financial", "PRICE"],
        formula="=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
    )
    add_case(
        cases,
        prefix="price",
        tags=["financial", "PRICE", "coercion"],
        formula="=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2.9,0)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="price",
        tags=["financial", "PRICE", "coercion"],
        formula="=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="yield",
        tags=["financial", "YIELD"],
        formula="=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)",
    )
    add_case(
        cases,
        prefix="yield",
        tags=["financial", "YIELD", "coercion"],
        formula="=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2.9,0)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="yield",
        tags=["financial", "YIELD", "coercion"],
        formula="=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="duration",
        tags=["financial", "DURATION"],
        formula="=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    )
    add_case(
        cases,
        prefix="duration",
        tags=["financial", "DURATION", "coercion"],
        formula="=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2.9,1)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="duration",
        tags=["financial", "DURATION", "coercion"],
        formula="=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="mduration",
        tags=["financial", "MDURATION"],
        formula="=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    )
    add_case(
        cases,
        prefix="mduration",
        tags=["financial", "MDURATION", "coercion"],
        formula="=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2.9,1)",
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="mduration",
        tags=["financial", "MDURATION", "coercion"],
        formula="=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0.9)",
        description="basis=0.9",
    )

    # Odd-coupon bond functions (ODDF*/ODDL*) frequency variants + non-par redemption scaling.
    #
    # These are intentionally "roundtrip" cases: pick a yield, compute price via ODDF/ODDLPRICE,
    # then recover the yield via ODDF/ODDLYIELD.
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,105,1,0)",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0)",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,105,1,0),105,1,0)",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0),100,4,0)",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,105,1,0)",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0)",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,105,1,0),105,1,0)",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0),100,4,0)",
    )

    # Accrued interest.
    add_case(
        cases,
        prefix="accrintm",
        tags=["financial", "ACCRINTM"],
        formula="=ACCRINTM(A1,A2,0.1,1000,0)",
        inputs=[CellInput("A1", excel_serial_1900(2020, 1, 1)), CellInput("A2", excel_serial_1900(2020, 7, 1))],
    )
    add_case(
        cases,
        prefix="accrintm",
        tags=["financial", "ACCRINTM", "coercion"],
        formula="=ACCRINTM(A1,A2,0.1,1000,0.9)",
        inputs=[CellInput("A1", excel_serial_1900(2020, 1, 1)), CellInput("A2", excel_serial_1900(2020, 7, 1))],
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="accrint",
        tags=["financial", "ACCRINT"],
        formula="=ACCRINT(A1,A2,A3,0.1,1000,2,0,FALSE)",
        inputs=[
            CellInput("A1", excel_serial_1900(2020, 2, 15)),
            CellInput("A2", excel_serial_1900(2020, 5, 15)),
            CellInput("A3", excel_serial_1900(2020, 8, 15)),
        ],
    )
    add_case(
        cases,
        prefix="accrint",
        tags=["financial", "ACCRINT", "coercion"],
        formula="=ACCRINT(A1,A2,A3,0.1,1000,2.9,0,FALSE)",
        inputs=[
            CellInput("A1", excel_serial_1900(2020, 2, 15)),
            CellInput("A2", excel_serial_1900(2020, 5, 15)),
            CellInput("A3", excel_serial_1900(2020, 8, 15)),
        ],
        description="frequency=2.9",
    )
    add_case(
        cases,
        prefix="accrint",
        tags=["financial", "ACCRINT", "coercion"],
        formula="=ACCRINT(A1,A2,A3,0.1,1000,2,0.9,FALSE)",
        inputs=[
            CellInput("A1", excel_serial_1900(2020, 2, 15)),
            CellInput("A2", excel_serial_1900(2020, 5, 15)),
            CellInput("A3", excel_serial_1900(2020, 8, 15)),
        ],
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="accrint_stub_default",
        tags=["financial", "ACCRINT"],
        formula="=ACCRINT(A1,A2,A3,0.1,1000,2,0)",
        inputs=[
            CellInput("A1", excel_serial_1900(2020, 2, 15)),
            CellInput("A2", excel_serial_1900(2020, 5, 15)),
            CellInput("A3", excel_serial_1900(2020, 4, 15)),
        ],
        description="Settlement before first_interest; exercises default calc_method behavior.",
    )
    add_case(
        cases,
        prefix="accrint_stub_true",
        tags=["financial", "ACCRINT"],
        formula="=ACCRINT(A1,A2,A3,0.1,1000,2,0,TRUE)",
        inputs=[
            CellInput("A1", excel_serial_1900(2020, 2, 15)),
            CellInput("A2", excel_serial_1900(2020, 5, 15)),
            CellInput("A3", excel_serial_1900(2020, 4, 15)),
        ],
        description="Settlement before first_interest; calc_method TRUE accrues from the start of the regular coupon period.",
    )

    # EOM schedule behavior for ACCRINT: when `first_interest` is month-end (even if not the 31st),
    # Excel pins coupon dates to month-end. This impacts basis=1 (Actual/Actual) because `E` is the
    # actual number of days between coupon dates.
    add_case(
        cases,
        prefix="accrint_eom_first_interest",
        tags=["financial", "ACCRINT", "coupon_schedule", "eom_edge"],
        formula="=ACCRINT(DATE(2019,12,15),DATE(2020,4,30),DATE(2020,8,15),0.12,1000,4,1)",
        description="ACCRINT with first_interest=2020-04-30 (month-end not 31st), basis=1",
    )
    add_case(
        cases,
        prefix="accrint_eom_first_interest_calc_method_false",
        tags=["financial", "ACCRINT", "coupon_schedule", "eom_edge"],
        formula="=ACCRINT(DATE(2020,1,15),DATE(2020,4,30),DATE(2020,2,15),0.12,1000,4,1,FALSE)",
        description="ACCRINT with first_interest=2020-04-30 (month-end not 31st), basis=1, calc_method=FALSE",
    )
    add_case(
        cases,
        prefix="accrint_eom_first_interest_calc_method_true",
        tags=["financial", "ACCRINT", "coupon_schedule", "eom_edge"],
        formula="=ACCRINT(DATE(2020,1,15),DATE(2020,4,30),DATE(2020,2,15),0.12,1000,4,1,TRUE)",
        description="ACCRINT with first_interest=2020-04-30 (month-end not 31st), basis=1, calc_method=TRUE",
    )

    # Odd-coupon bond functions (ODDF*/ODDL*)
    #
    # Keep these cases small + focused on:
    # - input validation semantics (negative yields / negative coupon rates)
    # - date coercion (ISO-like date text should be accepted and coerced to a date serial)
    # - default/boundary parameters (basis omitted should match basis=0; yld=0 should be finite)
    #
    # Also include basis-coverage cases across 0..4. These functions use coupon-period day-count
    # ratios (DSC/E, etc) rather than YEARFRAC-based coupon sizing, which is parity-sensitive for
    # actual-day bases (1/2/3) when regular coupon periods have different day lengths.

    # Default basis (omitted) should match basis=0.
    # Also include yld=0, which is a common edge where discount factors become 1.0.
    add_case(
        cases,
        prefix="oddfprice_basis_omitted",
        tags=["financial", "odd_coupon", "ODDFPRICE", "basis_omitted", "basis0"],
        formula=(
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),"
            "0.0785,0.0625,100,2)"
        ),
        description="ODDFPRICE with basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddfyield_basis_omitted",
        tags=["financial", "odd_coupon", "ODDFYIELD", "basis_omitted", "basis0"],
        formula=(
            "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),"
            "0.0785,98,100,2)"
        ),
        description="ODDFYIELD with basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddlprice_basis_omitted",
        tags=["financial", "odd_coupon", "ODDLPRICE", "basis_omitted", "basis0"],
        formula=(
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),"
            "0.0785,0.0625,100,2)"
        ),
        description="ODDLPRICE with basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddlyield_basis_omitted",
        tags=["financial", "odd_coupon", "ODDLYIELD", "basis_omitted", "basis0"],
        formula=(
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),"
            "0.0785,98,100,2)"
        ),
        description="ODDLYIELD with basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddfprice_basis_omitted_yld0",
        tags=["financial", "odd_coupon", "ODDFPRICE", "basis_omitted", "basis0", "yld0"],
        formula=(
            "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),"
            "0.0785,0,100,2)"
        ),
        description="ODDFPRICE with yld=0 and basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddfyield_basis_omitted_from_yld0_price",
        tags=["financial", "odd_coupon", "ODDFYIELD", "basis_omitted", "basis0", "yld0"],
        formula=(
            "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),"
            "0.0785,A1,100,2)"
        ),
        inputs=[
            CellInput(
                "A1",
                formula=(
                    "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),"
                    "0.0785,0,100,2)"
                ),
            )
        ],
        description="ODDFYIELD with basis omitted, using a price computed by ODDFPRICE with yld=0 (yield should roundtrip to 0)",
    )
    add_case(
        cases,
        prefix="oddlprice_basis_omitted_yld0",
        tags=["financial", "odd_coupon", "ODDLPRICE", "basis_omitted", "basis0", "yld0"],
        formula=(
            "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),"
            "0.0785,0,100,2)"
        ),
        description="ODDLPRICE with yld=0 and basis omitted (defaults to 0)",
    )
    add_case(
        cases,
        prefix="oddlyield_basis_omitted_from_yld0_price",
        tags=["financial", "odd_coupon", "ODDLYIELD", "basis_omitted", "basis0", "yld0"],
        formula=(
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),"
            "0.0785,A1,100,2)"
        ),
        inputs=[
            CellInput(
                "A1",
                formula=(
                    "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),"
                    "0.0785,0,100,2)"
                ),
            )
        ],
        description="ODDLYIELD with basis omitted, using a price computed by ODDLPRICE with yld=0 (yield should roundtrip to 0)",
    )

    for basis in [0, 1, 2, 3, 4]:
        add_case(
            cases,
            prefix=f"oddfprice_basis{basis}",
            tags=["financial", "odd_coupon", "ODDFPRICE", f"basis{basis}"],
            formula=(
                "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,29),"
                f"0.08,0.075,100,2,{basis})"
            ),
            description=f"ODDFPRICE basis={basis} (semiannual; clamped Feb coupon)",
        )
        add_case(
            cases,
            prefix=f"oddfyield_basis{basis}",
            tags=["financial", "odd_coupon", "ODDFYIELD", f"basis{basis}"],
            formula=(
                "=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,29),"
                f"0.08,98,100,2,{basis})"
            ),
            description=f"ODDFYIELD basis={basis} (semiannual; fixed price)",
        )
        add_case(
            cases,
            prefix=f"oddlprice_basis{basis}",
            tags=["financial", "odd_coupon", "ODDLPRICE", f"basis{basis}"],
            formula=(
                "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),"
                f"0.06,0.055,100,4,{basis})"
            ),
            description=f"ODDLPRICE basis={basis} (quarterly; short odd last)",
        )
        add_case(
            cases,
            prefix=f"oddlyield_basis{basis}",
            tags=["financial", "odd_coupon", "ODDLYIELD", f"basis{basis}"],
            formula=(
                "=ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),"
                f"0.06,98,100,4,{basis})"
            ),
            description=f"ODDLYIELD basis={basis} (quarterly; fixed price)",
        )

    # Month-end schedule disambiguation for ODDL*: last_interest on Apr 30 (month-end but not 31st).
    # For basis=1, the computed regular period length `E` depends on whether the schedule pins to
    # month-end (Jan 31 -> Apr 30) vs a day-of-month schedule (Jan 30 -> Apr 30).
    add_case(
        cases,
        prefix="oddlprice_eom_apr30_basis1",
        tags=["financial", "odd_coupon", "ODDLPRICE", "coupon_schedule", "eom_edge", "basis1"],
        formula="=ODDLPRICE(DATE(2020,5,15),DATE(2020,7,31),DATE(2020,4,30),0.06,0.055,100,4,1)",
        description="ODDLPRICE basis=1 with last_interest=2020-04-30 (month-end but not 31st)",
    )
    add_case(
        cases,
        prefix="oddlyield_eom_apr30_basis1",
        tags=["financial", "odd_coupon", "ODDLYIELD", "coupon_schedule", "eom_edge", "basis1"],
        formula="=ODDLYIELD(DATE(2020,5,15),DATE(2020,7,31),DATE(2020,4,30),0.06,98,100,4,1)",
        description="ODDLYIELD basis=1 with last_interest=2020-04-30 (month-end but not 31st)",
    )

    # Schedule alignment / misalignment cases (ODDF*/ODDL*).
    #
    # These scenarios are intended to pin Excel's subtle schedule validation rules around:
    # - `first_coupon` / `maturity` alignment for ODDF* (maturity-anchored coupon stepping)
    # - `last_interest` / `maturity` alignment for ODDL*
    #
    # Note: these cases are tagged `invalid_schedule` for historical reasons, but the slice includes
    # both true `#NUM!` invalid-schedule errors *and* "misaligned but chronologically valid" inputs
    # that the engine currently accepts/pins to numeric results. Validate against real Excel before
    # tightening schedule constraints.
    #
    # NOTE: The pinned dataset in CI is currently a synthetic baseline generated from the engine.
    # Treat these as regression tests for current engine behavior until we can patch with real Excel
    # results (Task 486).
    add_case(
        cases,
        prefix="oddfprice_invalid_schedule_dom_mismatch",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFPRICE"],
        formula=(
            "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,28),"
            "0.08,0.075,100,2,0)"
        ),
        description=(
            "ODDFPRICE invalid schedule: first_coupon is not reachable from maturity by stepping 6 months "
            "(day-of-month mismatch; maturity-anchored EDATE schedule hits 2020-02-29, not 2020-02-28)"
        ),
    )
    add_case(
        cases,
        prefix="oddfyield_invalid_schedule_dom_mismatch",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFYIELD"],
        formula=(
            "=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,28),"
            "0.08,98,100,2,0)"
        ),
        description=(
            "ODDFYIELD invalid schedule: first_coupon is not reachable from maturity by stepping 6 months "
            "(day-of-month mismatch; expected #NUM!)"
        ),
    )

    # EOM mismatch: maturity is end-of-month but first_coupon is not (schedule is pinned to month-end).
    add_case(
        cases,
        prefix="oddfprice_invalid_schedule_maturity_eom_first_not",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFPRICE"],
        formula=(
            "=ODDFPRICE(DATE(2022,12,20),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,30),"
            "0.05,0.06,100,2,0)"
        ),
        description=(
            "ODDFPRICE invalid schedule: maturity is EOM (month-end schedule) but first_coupon is not; "
            "maturity-anchored EOMONTH stepping never hits first_coupon"
        ),
    )
    add_case(
        cases,
        prefix="oddfyield_invalid_schedule_maturity_eom_first_not",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFYIELD"],
        formula=(
            "=ODDFYIELD(DATE(2022,12,20),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,30),"
            "0.05,98,100,2,0)"
        ),
        description=(
            "ODDFYIELD invalid schedule: maturity is EOM but first_coupon is not (EOM schedule stepping)"
        ),
    )

    # EOM mismatch: maturity is not end-of-month but first_coupon is (schedule is day-of-month based).
    add_case(
        cases,
        prefix="oddfprice_invalid_schedule_maturity_not_eom_first_is",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFPRICE"],
        formula=(
            "=ODDFPRICE(DATE(2022,12,20),DATE(2024,7,30),DATE(2022,12,15),DATE(2023,1,31),"
            "0.05,0.06,100,2,0)"
        ),
        description=(
            "ODDFPRICE invalid schedule: maturity is not EOM (EDATE schedule) but first_coupon is EOM; "
            "maturity-anchored EDATE stepping never hits first_coupon"
        ),
    )
    add_case(
        cases,
        prefix="oddfyield_invalid_schedule_maturity_not_eom_first_is",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFYIELD"],
        formula=(
            "=ODDFYIELD(DATE(2022,12,20),DATE(2024,7,30),DATE(2022,12,15),DATE(2023,1,31),"
            "0.05,98,100,2,0)"
        ),
        description=(
            "ODDFYIELD invalid schedule: maturity is not EOM but first_coupon is EOM (schedule stepping)"
        ),
    )

    # ODDL*: basic schedule-alignment invalidity (last_interest must be before maturity).
    #
    # We also include a "misaligned but chronologically valid" variant where `last_interest` is not
    # reachable from `maturity` by stepping whole coupon periods under the maturity-anchored EOM
    # schedule (Excel behavior for this is subtle; pin via oracle).
    add_case(
        cases,
        prefix="oddlprice_invalid_schedule_misaligned_last_interest",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2024,8,1),DATE(2025,1,31),DATE(2024,7,30),0.05,0.04,100,2,0)",
        description=(
            "ODDLPRICE invalid schedule: last_interest is not on the maturity-anchored coupon schedule "
            "implied by maturity+frequency (maturity is EOM but last_interest is not)"
        ),
    )
    add_case(
        cases,
        prefix="oddlyield_invalid_schedule_misaligned_last_interest",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2024,8,1),DATE(2025,1,31),DATE(2024,7,30),0.05,99,100,2,0)",
        description=(
            "ODDLYIELD invalid schedule: last_interest is not on the maturity-anchored coupon schedule "
            "implied by maturity+frequency (maturity is EOM but last_interest is not)"
        ),
    )

    # Extra EOM mismatch cases for ODDL* (basis=1 so coupon-period `E` depends on the actual
    # previous regular coupon dates).
    add_case(
        cases,
        prefix="oddlprice_invalid_schedule_maturity_eom_last_not_basis1",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLPRICE", "basis1"],
        formula="=ODDLPRICE(DATE(2024,8,15),DATE(2025,1,31),DATE(2024,7,30),0.05,0.04,100,2,1)",
        description=(
            "ODDLPRICE invalid schedule: maturity is EOM but last_interest is not (basis=1; EOM schedule mismatch)"
        ),
    )
    add_case(
        cases,
        prefix="oddlyield_invalid_schedule_maturity_eom_last_not_basis1",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLYIELD", "basis1"],
        formula="=ODDLYIELD(DATE(2024,8,15),DATE(2025,1,31),DATE(2024,7,30),0.05,99,100,2,1)",
        description=(
            "ODDLYIELD invalid schedule: maturity is EOM but last_interest is not (basis=1; EOM schedule mismatch)"
        ),
    )
    add_case(
        cases,
        prefix="oddlprice_invalid_schedule_maturity_not_eom_last_is_basis1",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLPRICE", "basis1"],
        formula="=ODDLPRICE(DATE(2024,8,15),DATE(2025,1,30),DATE(2024,7,31),0.05,0.04,100,2,1)",
        description=(
            "ODDLPRICE invalid schedule: maturity is not EOM but last_interest is EOM (basis=1; EOM schedule mismatch)"
        ),
    )
    add_case(
        cases,
        prefix="oddlyield_invalid_schedule_maturity_not_eom_last_is_basis1",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLYIELD", "basis1"],
        formula="=ODDLYIELD(DATE(2024,8,15),DATE(2025,1,30),DATE(2024,7,31),0.05,99,100,2,1)",
        description=(
            "ODDLYIELD invalid schedule: maturity is not EOM but last_interest is EOM (basis=1; EOM schedule mismatch)"
        ),
    )

    add_case(
        cases,
        prefix="oddlprice_invalid_schedule_last_after_maturity",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2025,1,2),0.05,0.04,100,2,0)",
        description=(
            "ODDLPRICE invalid schedule: last_interest is after maturity (schedule alignment / chronology)"
        ),
    )
    add_case(
        cases,
        prefix="oddlyield_invalid_schedule_last_after_maturity",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2024,7,1),DATE(2025,1,1),DATE(2025,1,2),0.05,99,100,2,0)",
        description=(
            "ODDLYIELD invalid schedule: last_interest is after maturity (schedule alignment / chronology)"
        ),
    )

    # Minimal ordering invalidity that is specifically about odd-coupon schedule semantics (not just
    # settlement >= maturity): ODDF* requires settlement <= first_coupon.
    add_case(
        cases,
        prefix="oddfprice_invalid_schedule_settlement_after_first",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFPRICE"],
        formula=(
            "=ODDFPRICE(DATE(2020,8,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),"
            "0.05,0.04,100,2,0)"
        ),
        description="ODDFPRICE invalid schedule: settlement is after first_coupon (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_invalid_schedule_settlement_after_first",
        tags=["financial", "odd_coupon", "invalid_schedule", "ODDFYIELD"],
        formula=(
            "=ODDFYIELD(DATE(2020,8,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),"
            "0.05,98,100,2,0)"
        ),
        description="ODDFYIELD invalid schedule: settlement is after first_coupon (expected #NUM!)",
    )

    add_case(
        cases,
        prefix="oddfprice_neg_yld",
        tags=["financial", "odd_coupon", "ODDFPRICE", "odd_coupon_validation", "neg_yld"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-0.01,100,2,0)",
        description="ODDFPRICE with negative yld (confirm whether Excel returns #NUM! or a price)",
    )
    add_case(
        cases,
        prefix="oddlprice_neg_yld",
        tags=["financial", "odd_coupon", "ODDLPRICE", "odd_coupon_validation", "neg_yld"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-0.01,100,2,0)",
        description="ODDLPRICE with negative yld (confirm whether Excel returns #NUM! or a price)",
    )
    add_case(
        cases,
        prefix="oddfprice_neg_yld_below_minus1",
        tags=["financial", "odd_coupon", "ODDFPRICE", "odd_coupon_validation", "neg_yld", "yld_below_minus1"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-1.5,100,2,0)",
        description="ODDFPRICE with yld=-1.5 (below -1 but still within the per-period domain yld > -frequency when frequency=2)",
    )
    add_case(
        cases,
        prefix="oddlprice_neg_yld_below_minus1",
        tags=["financial", "odd_coupon", "ODDLPRICE", "odd_coupon_validation", "neg_yld", "yld_below_minus1"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0)",
        description="ODDLPRICE with yld=-1.5 (below -1 but still within the per-period domain yld > -frequency when frequency=2)",
    )
    add_case(
        cases,
        prefix="oddlprice_neg_yld_below_minus1_settlement_before_last_interest",
        tags=[
            "financial",
            "odd_coupon",
            "ODDLPRICE",
            "odd_coupon_validation",
            "neg_yld",
            "yld_below_minus1",
            "settlement_before_last_interest",
        ],
        formula="=ODDLPRICE(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0)",
        description="ODDLPRICE with yld=-1.5 and settlement before last_interest (covers the settlement < last_interest pricing path)",
    )
    add_case(
        cases,
        prefix="oddfprice_yld_eq_neg_frequency",
        tags=[
            "financial",
            "odd_coupon",
            "ODDFPRICE",
            "odd_coupon_validation",
            "yld_eq_neg_frequency",
        ],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-2,100,2,0)",
        description="ODDFPRICE with yld == -frequency (discount base hits 0; confirm Excel returns #DIV/0!)",
    )
    add_case(
        cases,
        prefix="oddlprice_yld_eq_neg_frequency",
        tags=[
            "financial",
            "odd_coupon",
            "ODDLPRICE",
            "odd_coupon_validation",
            "yld_eq_neg_frequency",
        ],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-2,100,2,0)",
        description="ODDLPRICE with yld == -frequency (discount base hits 0; confirm Excel returns #DIV/0!)",
    )
    add_case(
        cases,
        prefix="oddfprice_yld_below_neg_frequency",
        tags=[
            "financial",
            "odd_coupon",
            "ODDFPRICE",
            "odd_coupon_validation",
            "yld_lt_neg_frequency",
        ],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-2.5,100,2,0)",
        description="ODDFPRICE with yld < -frequency (confirm Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlprice_yld_below_neg_frequency",
        tags=[
            "financial",
            "odd_coupon",
            "ODDLPRICE",
            "odd_coupon_validation",
            "yld_lt_neg_frequency",
        ],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-2.5,100,2,0)",
        description="ODDLPRICE with yld < -frequency (confirm Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfprice_neg_rate",
        tags=["financial", "odd_coupon", "ODDFPRICE", "odd_coupon_validation", "neg_rate"],
        formula="=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,0.0625,100,2,0)",
        description="ODDFPRICE with negative coupon rate (confirm whether Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlprice_neg_rate",
        tags=["financial", "odd_coupon", "ODDLPRICE", "odd_coupon_validation", "neg_rate"],
        formula="=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,0.0625,100,2,0)",
        description="ODDLPRICE with negative coupon rate (confirm whether Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_neg_rate",
        tags=["financial", "odd_coupon", "ODDFYIELD", "odd_coupon_validation", "neg_rate"],
        formula="=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,98,100,2,0)",
        description="ODDFYIELD with negative coupon rate (confirm whether Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlyield_neg_rate",
        tags=["financial", "odd_coupon", "ODDLYIELD", "odd_coupon_validation", "neg_rate"],
        formula="=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,98,100,2,0)",
        description="ODDLYIELD with negative coupon rate (confirm whether Excel returns #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_high_price",
        tags=["financial", "odd_coupon", "ODDFYIELD", "odd_coupon_validation", "implied_neg_yld"],
        formula="=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,300,100,2,0)",
        description="ODDFYIELD with a price that implies a negative yield if allowed (confirm whether Excel returns a negative yield or #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlyield_high_price",
        tags=["financial", "odd_coupon", "ODDLYIELD", "odd_coupon_validation", "implied_neg_yld"],
        formula="=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,300,100,2,0)",
        description="ODDLYIELD with a price that implies a negative yield if allowed (confirm whether Excel returns a negative yield or #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_roundtrip_neg_yld_below_minus1",
        tags=[
            "financial",
            "odd_coupon",
            "ODDFYIELD",
            "odd_coupon_validation",
            "neg_yld",
            "yld_below_minus1",
        ],
        formula=(
            "=ODDFYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),DATE(2021,3,1),0.0785,"
            "ODDFPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),DATE(2021,3,1),0.0785,-1.5,100,2,0),"
            "100,2,0)"
        ),
        description="ODDFYIELD roundtrip: yield from ODDFPRICE at yld=-1.5 (below -1) to confirm Excel allows negative yields in both ODDFPRICE and ODDFYIELD",
    )
    add_case(
        cases,
        prefix="oddlyield_roundtrip_neg_yld_below_minus1",
        tags=[
            "financial",
            "odd_coupon",
            "ODDLYIELD",
            "odd_coupon_validation",
            "neg_yld",
            "yld_below_minus1",
        ],
        formula=(
            "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,"
            "ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0),"
            "100,2,0)"
        ),
        description="ODDLYIELD roundtrip: yield from ODDLPRICE at yld=-1.5 (below -1) to confirm Excel allows negative yields in both ODDLPRICE and ODDLYIELD",
    )
    add_case(
        cases,
        prefix="oddlyield_roundtrip_neg_yld_below_minus1_settlement_before_last_interest",
        tags=[
            "financial",
            "odd_coupon",
            "ODDLYIELD",
            "odd_coupon_validation",
            "neg_yld",
            "yld_below_minus1",
            "settlement_before_last_interest",
        ],
        formula=(
            "=ODDLYIELD(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,"
            "ODDLPRICE(DATE(2020,8,1),DATE(2021,3,1),DATE(2020,10,15),0.0785,-1.5,100,2,0),"
            "100,2,0)"
        ),
        description="ODDLYIELD roundtrip: yield from ODDLPRICE at yld=-1.5 with settlement before last_interest (covers the settlement < last_interest path)",
    )
    add_case(
        cases,
        prefix="financial_oddfprice_zero_coupon",
        tags=["financial", "ODDFPRICE", "zero-coupon"],
        formula="=ODDFPRICE(43831,44378,43739,44013,0,0.1,100,2,0)",
        output_cell="A1",
        description="ODDFPRICE zero-coupon: discounted redemption with odd first coupon schedule",
    )
    add_case(
        cases,
        prefix="financial_oddfyield_zero_coupon",
        tags=["financial", "ODDFYIELD", "zero-coupon"],
        formula="=ODDFYIELD(43831,44378,43739,44013,0,86.3837598531476,100,2,0)",
        output_cell="A1",
        description="ODDFYIELD inverts ODDFPRICE for a zero-coupon bond",
    )
    add_case(
        cases,
        prefix="financial_oddlprice_zero_coupon",
        tags=["financial", "ODDLPRICE", "zero-coupon"],
        formula="=ODDLPRICE(44228,44317,44197,0,0.1,100,2,0)",
        output_cell="A1",
        description="ODDLPRICE zero-coupon: discounted redemption with odd last coupon schedule",
    )
    add_case(
        cases,
        prefix="financial_oddlyield_zero_coupon",
        tags=["financial", "ODDLYIELD", "zero-coupon"],
        formula="=ODDLYIELD(44228,44317,44197,0,97.59000729485331,100,2,0)",
        output_cell="A1",
        description="ODDLYIELD inverts ODDLPRICE for a zero-coupon bond",
    )

    add_case(
        cases,
        prefix="oddfprice_date_text",
        tags=["financial", "odd_coupon", "ODDFPRICE"],
        formula='=ODDFPRICE("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,0.0625,100,2,0)',
        description="ODDFPRICE with ISO date text arguments (date coercion)",
    )
    add_case(
        cases,
        prefix="oddlprice_date_text",
        tags=["financial", "odd_coupon", "ODDLPRICE"],
        formula='=ODDLPRICE("2020-11-11","2021-03-01","2020-10-15",0.0785,0.0625,100,2,0)',
        description="ODDLPRICE with ISO date text arguments (date coercion)",
    )
    add_case(
        cases,
        prefix="oddfyield_date_text",
        tags=["financial", "odd_coupon", "ODDFYIELD"],
        formula='=ODDFYIELD("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,A1,100,2,0)',
        inputs=[
            CellInput(
                "A1",
                formula='=ODDFPRICE("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,0.0625,100,2,0)',
            )
        ],
        description="ODDFYIELD with ISO date text arguments (date coercion + roundtrip via ODDFPRICE)",
    )
    add_case(
        cases,
        prefix="oddlyield_date_text",
        tags=["financial", "odd_coupon", "ODDLYIELD"],
        formula='=ODDLYIELD("2020-11-11","2021-03-01","2020-10-15",0.0785,A1,100,2,0)',
        inputs=[
            CellInput(
                "A1",
                formula='=ODDLPRICE("2020-11-11","2021-03-01","2020-10-15",0.0785,0.0625,100,2,0)',
            )
        ],
        description="ODDLYIELD with ISO date text arguments (date coercion + roundtrip via ODDLPRICE)",
    )

    # Settlement before `last_interest` (ODDL* supports settlement anywhere before maturity).
    add_case(
        cases,
        prefix="oddlprice_settlement_before_last_interest",
        tags=["financial", "odd_coupon", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2023,10,15),DATE(2025,3,1),DATE(2024,7,1),0.06,0.05,100,2,0)",
        description="ODDLPRICE with settlement before last_interest (multiple remaining coupons)",
    )
    add_case(
        cases,
        prefix="oddlyield_settlement_before_last_interest",
        tags=["financial", "odd_coupon", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2023,10,15),DATE(2025,3,1),DATE(2024,7,1),0.06,ODDLPRICE(DATE(2023,10,15),DATE(2025,3,1),DATE(2024,7,1),0.06,0.05,100,2,0),100,2,0)",
        description="ODDLYIELD inverts ODDLPRICE when settlement is before last_interest",
    )
    add_case(
        cases,
        prefix="oddlprice_settlement_before_last_interest_eom",
        tags=["financial", "odd_coupon", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2023,10,15),DATE(2025,2,15),DATE(2024,7,31),0.06,0.05,100,2,0)",
        description="ODDLPRICE with settlement before last_interest on an end-of-month coupon schedule",
    )
    add_case(
        cases,
        prefix="oddlyield_settlement_before_last_interest_eom",
        tags=["financial", "odd_coupon", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2023,10,15),DATE(2025,2,15),DATE(2024,7,31),0.06,ODDLPRICE(DATE(2023,10,15),DATE(2025,2,15),DATE(2024,7,31),0.06,0.05,100,2,0),100,2,0)",
        description="ODDLYIELD inverts ODDLPRICE for an end-of-month coupon schedule when settlement is before last_interest",
    )

    # Boundary-date equality behavior (Excel quirks).
    #
    # These cases are intentionally simple and vary only one boundary at a time. Today they are
    # primarily used to pin current engine behavior in CI; verify real Excel strictness/leniency by
    # generating a real Excel dataset via tools/excel-oracle/run-excel-oracle.ps1 (Task 393).
    #
    # Locked by:
    # - `crates/formula-engine/tests/functions/financial_odd_coupon.rs`
    # - `crates/formula-engine/tests/functions/financial_oddcoupons.rs`
    add_case(
        cases,
        prefix="oddfprice_issue_eq_settlement",
        tags=["financial", "odd_coupon", "boundary", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        description="ODDFPRICE boundary: issue == settlement",
    )
    add_case(
        cases,
        prefix="oddfyield_issue_eq_settlement",
        tags=["financial", "odd_coupon", "boundary", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        description="ODDFYIELD boundary: issue == settlement",
    )
    add_case(
        cases,
        prefix="oddfprice_settlement_eq_first_coupon",
        tags=["financial", "odd_coupon", "boundary", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        description="ODDFPRICE boundary: settlement == first_coupon",
    )
    add_case(
        cases,
        prefix="oddfyield_settlement_eq_first_coupon",
        tags=["financial", "odd_coupon", "boundary", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        description="ODDFYIELD boundary: settlement == first_coupon",
    )
    add_case(
        cases,
        prefix="oddfprice_first_coupon_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        description="ODDFPRICE boundary: first_coupon == maturity",
    )
    add_case(
        cases,
        prefix="oddfyield_first_coupon_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        description="ODDFYIELD boundary: first_coupon == maturity",
    )
    add_case(
        cases,
        prefix="oddfprice_settlement_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDFPRICE", "error"],
        formula="=ODDFPRICE(DATE(2020,7,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        description="ODDFPRICE boundary: settlement == maturity (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_settlement_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDFYIELD", "error"],
        formula="=ODDFYIELD(DATE(2020,7,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,99,100,2,0)",
        description="ODDFYIELD boundary: settlement == maturity (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfprice_issue_eq_first_coupon",
        tags=["financial", "odd_coupon", "boundary", "ODDFPRICE", "error"],
        formula="=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        description="ODDFPRICE boundary: issue == first_coupon (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddfyield_issue_eq_first_coupon",
        tags=["financial", "odd_coupon", "boundary", "ODDFYIELD", "error"],
        formula="=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,99,100,2,0)",
        description="ODDFYIELD boundary: issue == first_coupon (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlprice_settlement_eq_last_interest",
        tags=["financial", "odd_coupon", "boundary", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0)",
        description="ODDLPRICE boundary: settlement == last_interest",
    )
    add_case(
        cases,
        prefix="oddlyield_settlement_eq_last_interest",
        tags=["financial", "odd_coupon", "boundary", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0),100,2,0)",
        description="ODDLYIELD boundary: settlement == last_interest",
    )
    add_case(
        cases,
        prefix="oddlprice_last_interest_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDLPRICE", "error"],
        formula="=ODDLPRICE(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,0.04,100,2,0)",
        description="ODDLPRICE boundary: last_interest == maturity (expected #NUM!)",
    )
    add_case(
        cases,
        prefix="oddlyield_last_interest_eq_maturity",
        tags=["financial", "odd_coupon", "boundary", "ODDLYIELD", "error"],
        formula="=ODDLYIELD(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,99,100,2,0)",
        description="ODDLYIELD boundary: last_interest == maturity (expected #NUM!)",
    )

    # Odd-coupon bond functions: deterministic coercion/error cases (avoid NaN/Inf which can be
    # awkward for the Excel oracle harness).
    add_case(
        cases,
        prefix="oddfprice_value",
        tags=["financial", "odd_coupon", "ODDFPRICE", "error"],
        formula='=ODDFPRICE("nope",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,0.05,100,2)',
        description="Unparseable settlement date should return #VALUE!",
    )
    add_case(
        cases,
        prefix="oddfyield_value",
        tags=["financial", "odd_coupon", "ODDFYIELD", "error"],
        formula='=ODDFYIELD("nope",DATE(2025,1,1),DATE(2019,1,1),DATE(2020,7,1),0.05,95,100,2)',
        description="Unparseable settlement date should return #VALUE!",
    )
    add_case(
        cases,
        prefix="oddlprice_value",
        tags=["financial", "odd_coupon", "ODDLPRICE", "error"],
        formula='=ODDLPRICE(DATE(2020,1,1),"nope",DATE(2024,7,1),0.05,0.05,100,2)',
        description="Unparseable maturity date should return #VALUE!",
    )
    add_case(
        cases,
        prefix="oddlyield_value",
        tags=["financial", "odd_coupon", "ODDLYIELD", "error"],
        formula='=ODDLYIELD(DATE(2020,1,1),"nope",DATE(2024,7,1),0.05,95,100,2)',
        description="Unparseable maturity date should return #VALUE!",
    )

    # Basis backfill (ODDF*/ODDL*).
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,0.06,100,2,0)",
        description="ODDFPRICE basis=0",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,98,100,2,0)",
        description="ODDFYIELD basis=0",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,0.06,100,2,0)",
        description="ODDLPRICE basis=0",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,98,100,2,0)",
        description="ODDLYIELD basis=0",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,0.06,100,2,1)",
        description="ODDFPRICE basis=1",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,98,100,2,1)",
        description="ODDFYIELD basis=1",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,0.06,100,2,1)",
        description="ODDLPRICE basis=1",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,98,100,2,1)",
        description="ODDLYIELD basis=1",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,0.06,100,2,2)",
        description="ODDFPRICE basis=2",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,98,100,2,2)",
        description="ODDFYIELD basis=2",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,0.06,100,2,2)",
        description="ODDLPRICE basis=2",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,98,100,2,2)",
        description="ODDLYIELD basis=2",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,0.06,100,2,3)",
        description="ODDFPRICE basis=3",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,98,100,2,3)",
        description="ODDFYIELD basis=3",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,0.06,100,2,3)",
        description="ODDLPRICE basis=3",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,98,100,2,3)",
        description="ODDLYIELD basis=3",
    )
    add_case(
        cases,
        prefix="oddfprice",
        tags=["financial", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,0.06,100,2,4)",
        description="ODDFPRICE basis=4",
    )
    add_case(
        cases,
        prefix="oddfyield",
        tags=["financial", "ODDFYIELD"],
        formula="=ODDFYIELD(DATE(2019,2,28),DATE(2019,9,30),DATE(2019,1,31),DATE(2019,3,31),0.05,98,100,2,4)",
        description="ODDFYIELD basis=4",
    )
    add_case(
        cases,
        prefix="oddlprice",
        tags=["financial", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,0.06,100,2,4)",
        description="ODDLPRICE basis=4",
    )
    add_case(
        cases,
        prefix="oddlyield",
        tags=["financial", "ODDLYIELD"],
        formula="=ODDLYIELD(DATE(2019,3,15),DATE(2019,3,31),DATE(2019,2,28),0.05,98,100,2,4)",
        description="ODDLYIELD basis=4",
    )

    # Range-based cashflow functions.
    cashflows = [-100.0, 30.0, 40.0, 50.0]
    cf_inputs = [CellInput(f"A{i+1}", v) for i, v in enumerate(cashflows)]
    add_case(
        cases,
        prefix="npv",
        tags=["financial", "NPV"],
        formula="=NPV(0.1,A1:A4)",
        inputs=cf_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="irr",
        tags=["financial", "IRR"],
        formula="=IRR(A1:A4)",
        inputs=cf_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="mirr",
        tags=["financial", "MIRR"],
        formula="=MIRR(A1:A4,0.1,0.12)",
        inputs=cf_inputs,
        output_cell="C1",
    )
    add_case(
        cases,
        prefix="irr_num",
        tags=["financial", "IRR", "error"],
        formula="=IRR(A1:A3)",
        inputs=[CellInput("A1", 10), CellInput("A2", 20), CellInput("A3", 30)],
        output_cell="C1",
        description="IRR requires at least one positive and one negative cashflow",
    )

    # XNPV/XIRR with explicit date serials (Excel 1900 system with Lotus bug).
    x_values = [-10000.0, 2000.0, 3000.0, 4000.0, 5000.0]
    x_dates = [
        excel_serial_1900(2020, 1, 1),
        excel_serial_1900(2020, 7, 1),
        excel_serial_1900(2021, 1, 1),
        excel_serial_1900(2021, 7, 1),
        excel_serial_1900(2022, 1, 1),
    ]
    x_inputs = []
    for i, (v, d) in enumerate(zip(x_values, x_dates), start=1):
        x_inputs.append(CellInput(f"A{i}", v))
        x_inputs.append(CellInput(f"B{i}", d))
    add_case(
        cases,
        prefix="xnpv",
        tags=["financial", "XNPV"],
        formula="=XNPV(0.1,A1:A5,B1:B5)",
        inputs=x_inputs,
        output_cell="D1",
    )
    add_case(
        cases,
        prefix="xirr",
        tags=["financial", "XIRR"],
        formula="=XIRR(A1:A5,B1:B5)",
        inputs=x_inputs,
        output_cell="D1",
    )
