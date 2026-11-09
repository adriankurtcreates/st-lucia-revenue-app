import numpy as np


def calendarise_growth(cy_year1_rate, cy_year2_rate, weight=0.75):
    """
    Converts Calendar Year (CY) forecast paths into April-March Fiscal Year (FY) trends.
    Formula: FY growth = w * CY(t) + (1 - w) * CY(t+1)
    """
    # FIX: validate weight so a bad parameter can't silently corrupt all projections
    if not 0.0 <= weight <= 1.0:
        raise ValueError(f"weight must be in [0, 1], got {weight}")
    return (weight * cy_year1_rate) + ((1.0 - weight) * cy_year2_rate)


def calculate_nominal_growth(real_gdp_growth, cpi_inflation):
    """
    Computes true nominal growth rate combining volume and price factors.
    Formula: (1 + real_gdp_growth) * (1 + cpi_inflation) - 1
    """
    return ((1.0 + real_gdp_growth) * (1.0 + cpi_inflation)) - 1.0


def project_current_tax_head(baseline_val, driver_growth, buoyancy=1.0, policy_measure=0.0):
    """
    Core Revenue Equation:
    Revenue(t) = Revenue(t-1) * (1 + driver_growth)^buoyancy + policy_measure(t)
    """
    # FIX: guard against non-finite inputs so NaN can't propagate through projections
    if not np.isfinite(baseline_val):
        baseline_val = 0.0
    if not np.isfinite(driver_growth):
        driver_growth = 0.0
    base_projected = baseline_val * np.power((1.0 + driver_growth), buoyancy)
    return base_projected + policy_measure


def run_arrears_stock_flow(opening_stock, base_collection_rate, write_off_rate,
                           accrual_rate, current_tax_collections, rate_adjustment=0.0):
    """
    Executes the IMF-aligned Arrears Tank Dynamics:
    1. Adjusted Collection Rate = Base Collection Rate + Policy/Amnesty Adjustment
    2. Collections = Opening Stock * Adjusted Collection Rate
    3. Write-offs = Opening Stock * Write-off Rate
    4. New Arrears = Accrual Rate * Current-Year Collections of the underlying tax head
    5. Closing Stock = Opening Stock + New Arrears - Collections - Write-offs
    """
    # FIX: sanitise all inputs so NaN / negative current collections cannot corrupt the tank
    if not np.isfinite(opening_stock) or opening_stock < 0:
        opening_stock = 0.0
    if not np.isfinite(current_tax_collections) or current_tax_collections < 0:
        current_tax_collections = 0.0

    effective_collection_rate = max(0.0, base_collection_rate + rate_adjustment)

    collections = opening_stock * effective_collection_rate
    write_offs = opening_stock * write_off_rate
    new_arrears = accrual_rate * current_tax_collections

    closing_stock = opening_stock + new_arrears - collections - write_offs

    # FIX: explicit NaN-safe floor (max(0.0, nan) can return nan depending on argument order)
    closing_stock = 0.0 if not np.isfinite(closing_stock) else max(0.0, closing_stock)

    return {
        "collections": collections,
        "write_offs": write_offs,
        "new_arrears": new_arrears,
        "closing_stock": closing_stock,
    }
