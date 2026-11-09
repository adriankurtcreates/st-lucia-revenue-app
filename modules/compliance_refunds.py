def calculate_net_vat_bridge(gross_vat_collections, refund_claim_rate=0.15, audit_efficiency_rate=0.20):
    """
    IMF-Approved VAT Bridge:
    Isolates input tax refund volatility from baseline macro domestic economic metrics.
    
    Formula:
      Gross Claims = Gross VAT Collections * Refund Claim Rate
      Disallowed/Corrected Claims = Gross Claims * Audit Efficiency Rate
      Approved Refunds = Gross Claims - Disallowed/Corrected Claims
      Net VAT Yield = Gross VAT Collections - Approved Refunds
    """
    gross_claims = gross_vat_collections * refund_claim_rate
    disallowed_claims = gross_claims * audit_efficiency_rate
    approved_refunds = gross_claims - disallowed_claims
    net_vat_yield = gross_vat_collections - approved_refunds
    
    return {
        "gross_vat": gross_vat_collections,
        "gross_claims": gross_claims,
        "approved_refunds": approved_refunds,
        "net_vat": net_vat_yield
    }

def adjust_compliance_gap(base_revenue, baseline_tax_gap=0.30, policy_efficiency_gain=0.02):
    """
    Models targeted compliance intervention shocks (e.g., modernizing single window systems).
    
    Parameters:
      - base_revenue: Projections generated directly by macro drivers.
      - baseline_tax_gap: The fraction of total potential tax uncollected due to non-compliance.
      - policy_efficiency_gain: Pct points recovered via administrative intervention.
    """
    # Step 1: Reverse engineer the theoretical maximum potential tax pool
    theoretical_potential_pool = base_revenue / (1.0 - baseline_tax_gap)
    
    # Step 2: Calculate target expansion rate by tightening collection controls
    new_collection_efficiency = (1.0 - baseline_tax_gap) + policy_efficiency_gain
    # Cap total structural efficiency strictly at 100% of theoretical limits
    new_collection_efficiency = min(1.0, new_collection_efficiency)
    
    adjusted_revenue = theoretical_potential_pool * new_collection_efficiency
    return adjusted_revenue
