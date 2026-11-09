def calculate_net_vat_bridge(gross_vat_collections, refund_claim_rate=0.125, audit_efficiency_rate=0.15):
    """
    IMF-Approved VAT Bridge:
    Isolates input tax refund volatility from baseline macro domestic economic metrics.

    Formula:
      Gross Claims = Gross VAT Collections * Refund Claim Rate
      Disallowed/Corrected Claims = Gross Claims * Audit Efficiency Rate
      Approved Refunds = Gross Claims - Disallowed/Corrected Claims
      Net VAT Yield = Gross VAT Collections - Approved Refunds

    NOTE: baseline VAT values supplied to the model must be GROSS collections.
    Defaults now match the app slider defaults (12.5% claims, 15% audit recovery).
    """
    # FIX: clamp rates into [0, 1] so an errant caller cannot produce negative net VAT
    refund_claim_rate = min(1.0, max(0.0, float(refund_claim_rate)))
    audit_efficiency_rate = min(1.0, max(0.0, float(audit_efficiency_rate)))

    gross_claims = gross_vat_collections * refund_claim_rate
    disallowed_claims = gross_claims * audit_efficiency_rate
    approved_refunds = gross_claims - disallowed_claims
    net_vat_yield = gross_vat_collections - approved_refunds

    return {
        "gross_vat": gross_vat_collections,
        "gross_claims": gross_claims,
        "approved_refunds": approved_refunds,
        "net_vat": net_vat_yield,
    }


def adjust_compliance_gap(base_revenue, baseline_tax_gap=0.30, policy_efficiency_gain=0.02):
    """
    Models targeted compliance intervention shocks (e.g., modernizing single window systems).

    Parameters:
      - base_revenue: Projections generated directly by macro drivers.
      - baseline_tax_gap: The fraction of total potential tax uncollected due to non-compliance.
      - policy_efficiency_gain: Pct points recovered via administrative intervention.

    NOTE: Because the potential pool and current-efficiency factor share the (1 - gap)
    term, this function is mathematically equivalent to a one-time multiplier of
    base_revenue * (1 + gain / (1 - gap)). It is designed to be applied ONCE to the
    baseline level, NOT compounded year-over-year. See app.py for correct usage.
    """
    # FIX: guard against gap = 1.0 (ZeroDivisionError) and out-of-range inputs
    if not 0.0 <= baseline_tax_gap < 1.0:
        raise ValueError("baseline_tax_gap must be in [0, 1)")
    policy_efficiency_gain = max(0.0, float(policy_efficiency_gain))

    theoretical_potential_pool = base_revenue / (1.0 - baseline_tax_gap)
    new_collection_efficiency = (1.0 - baseline_tax_gap) + policy_efficiency_gain
    new_collection_efficiency = min(1.0, new_collection_efficiency)

    adjusted_revenue = theoretical_potential_pool * new_collection_efficiency
    return adjusted_revenue
