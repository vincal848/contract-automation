"""Mirrors CalculateMetrics in src/ReportAutomation.bas, row for row.

VBA cannot run in CI, so this is the part of the macro's logic that can be
pinned with ordinary tests. Keep it in step with CalculateMetrics by hand --
there is no code generation linking the two.

ponytail: hand-mirrored, so a VBA edit that is not copied here (or vice versa)
leaves these tests green while the macro is wrong. Ceiling: one reviewer's
discipline; upgrade path is exporting CalculatedMetrics from Excel and diffing it.
"""


def calculate_metrics(revenues, expenses):
    """Returns one dict per month: net_profit, revenue_growth_pct,
    profit_margin_pct, total_revenue -- the same columns A-D that
    CalculateMetrics writes to CalculatedMetrics.

    revenue_growth_pct is None (VBA writes "N/A") for the first month and for
    any month whose previous revenue was zero, instead of dividing by it.
    profit_margin_pct is None (VBA writes "N/A") for a month with zero revenue.
    """
    if len(revenues) != len(expenses):
        raise ValueError("revenues and expenses must have the same length")

    rows = []
    for i in range(len(revenues)):
        revenue = revenues[i]
        expense = expenses[i]
        net_profit = revenue - expense

        if i > 0 and revenues[i - 1] != 0:
            revenue_growth_pct = (revenue - revenues[i - 1]) / revenues[i - 1] * 100
        else:
            revenue_growth_pct = None

        if revenue != 0:
            profit_margin_pct = (net_profit / revenue) * 100
        else:
            profit_margin_pct = None

        rows.append({
            "net_profit": net_profit,
            "revenue_growth_pct": revenue_growth_pct,
            "profit_margin_pct": profit_margin_pct,
            "total_revenue": revenue,
        })
    return rows
