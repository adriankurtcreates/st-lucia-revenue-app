def calculate_net_vat(gross_vat_projected, refund_claim_rate, audit_efficiency_rate):
    """
    IMF Method: Net VAT = Gross VAT - (Gross Claims * (1 - Audit Correction Rate))
    """
    gross_claims = gross_vat_projected * refund_claim_rate
    approved_claims = gross_claims * (1.0 - audit_efficiency_rate)
    net_vat = gross_vat_projected - approved_claims
    return net_vat

def apply_compliance_shock(base_revenue, tax_gap_pct, collection_efficiency):
    """
    Models administrative performance gains (e.g., post-amnesty or audit push)
    """
    potential_revenue = base_revenue / (1 - tax_gap_pct)
    realised_revenue = potential_revenue * collection_efficiency
    return realised_revenue
