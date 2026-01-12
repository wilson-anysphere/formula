from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Coupon date helper functions (COUP*): end-of-month schedule behavior
    # ------------------------------------------------------------------
    #
    # These cases historically lived only in the committed `cases.json` file,
    # which meant running `tools/excel-oracle/generate_cases.py` would drop them.
    # Keep them here so the corpus remains fully regeneratable.
    add_case(
        cases,
        prefix="financial_coupncd_eom",
        tags=["financial", "COUPNCD", "eom"],
        formula="=DAY(COUPNCD(DATE(2021,3,1),DATE(2021,8,31),2,0))",
        output_cell="A1",
        description="COUPNCD end-of-month schedule restores month-end after February",
    )
    add_case(
        cases,
        prefix="financial_couppcd_eom",
        tags=["financial", "COUPPCD", "eom"],
        formula="=DAY(COUPPCD(DATE(2020,12,1),DATE(2021,8,31),2,0))",
        output_cell="A1",
        description="COUPPCD end-of-month schedule restores month-end across February",
    )
    add_case(
        cases,
        prefix="financial_coupnum_eom",
        tags=["financial", "COUPNUM", "eom"],
        formula="=COUPNUM(DATE(2020,12,1),DATE(2021,8,31),2,0)",
        output_cell="A1",
        description="COUPNUM counts remaining coupons on an end-of-month schedule",
    )
    add_case(
        cases,
        prefix="financial_coupdaybs_eom",
        tags=["financial", "COUPDAYBS", "eom"],
        formula="=COUPDAYBS(DATE(2021,3,1),DATE(2021,8,31),2,0)",
        output_cell="A1",
        description="COUPDAYBS day-count in the current coupon period (EOM schedule)",
    )
    add_case(
        cases,
        prefix="financial_coupdaysnc_eom",
        tags=["financial", "COUPDAYSNC", "eom"],
        formula="=COUPDAYSNC(DATE(2021,3,1),DATE(2021,8,31),2,0)",
        output_cell="A1",
        description="COUPDAYSNC days from settlement to next coupon date (EOM schedule)",
    )
    add_case(
        cases,
        prefix="financial_coupdays_eom",
        tags=["financial", "COUPDAYS", "eom"],
        formula="=COUPDAYS(DATE(2021,3,1),DATE(2021,8,31),2,0)",
        output_cell="A1",
        description="COUPDAYS days in coupon period (EOM schedule)",
    )
    add_case(
        cases,
        prefix="financial_couppcd_eom_maturity_feb",
        tags=["financial", "COUPPCD", "eom"],
        formula="=DAY(COUPPCD(DATE(2030,9,1),DATE(2031,2,28),2,0))",
        output_cell="A1",
        description="COUPPCD EOM schedule restores month-end when maturity is Feb month-end (28th -> 31st)",
    )

    # ------------------------------------------------------------------
    # Odd-coupon bond functions (ODDF* / ODDL*): end-of-month schedules
    # ------------------------------------------------------------------
    add_case(
        cases,
        prefix="oddfprice_eom",
        tags=["financial", "odd_coupon", "ODDFPRICE"],
        formula="=ODDFPRICE(DATE(2020,2,15),DATE(2020,12,31),DATE(2020,1,31),DATE(2020,6,30),0.05,0.04,100,2,0)",
        description="Odd first coupon with EOM schedule (30-Jun -> 31-Dec) to exercise end-of-month stepping",
    )
    add_case(
        cases,
        prefix="oddfyield_eom",
        tags=["financial", "odd_coupon", "ODDFYIELD", "ODDFPRICE"],
        formula="=ODDFYIELD(DATE(2020,2,15),DATE(2020,12,31),DATE(2020,1,31),DATE(2020,6,30),0.05,ODDFPRICE(DATE(2020,2,15),DATE(2020,12,31),DATE(2020,1,31),DATE(2020,6,30),0.05,0.04,100,2,0),100,2,0)",
        description="ODDFYIELD roundtrip against ODDFPRICE for an EOM-aligned schedule",
    )
    add_case(
        cases,
        prefix="oddlprice_eom",
        tags=["financial", "odd_coupon", "ODDLPRICE"],
        formula="=ODDLPRICE(DATE(2020,7,15),DATE(2020,12,31),DATE(2020,6,30),0.05,0.04,100,2,1)",
        description="Odd last coupon with EOM last_interest/maturity (basis=1) to exercise end-of-month stepping",
    )
    add_case(
        cases,
        prefix="oddlyield_eom",
        tags=["financial", "odd_coupon", "ODDLYIELD", "ODDLPRICE"],
        formula="=ODDLYIELD(DATE(2020,7,15),DATE(2020,12,31),DATE(2020,6,30),0.05,ODDLPRICE(DATE(2020,7,15),DATE(2020,12,31),DATE(2020,6,30),0.05,0.04,100,2,1),100,2,1)",
        description="ODDLYIELD roundtrip against ODDLPRICE for an EOM-aligned schedule",
    )
