"""Flagship Deck builder example: a consulting-style business review.

Builds the six-slide deck showcased in the 0.10.0 release using one fluent
chain — no coordinates, no template file. The corporate theme drives the whole
look: composed title slide, KPI stat tiles with deltas, an emphasized chart
with a message headline, comparison cards, kickers, and footers with page
numbers.

Run with: uv run python examples/deck/business_review.py
"""

import pandas as pd

from easypptx import Deck

df = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [1200, 1350, 1480, 1610],
    "Expenses": [900, 940, 1010, 1080],
})

(
    Deck(theme="corporate", footer="Acme Corp — Q3 Business Review")
    .title_slide(
        "Q3 Business Review",
        subtitle="Finance & Strategy · October 2026",
        notes="60 seconds max",
    )
    .slide("The quarter in numbers", kicker="Q3 FINANCIALS")
    .stats([
        ("+12%", "Revenue growth", "+3pt vs Q2"),
        ("$1.61M", "Q4 revenue run-rate"),
        ("67%", "Gross margin", "+1.5pt"),
        ("38", "New logos", "-4 vs plan"),
    ])
    .bullets([
        "Fourth consecutive quarter of accelerating growth",
        ("EMEA and APAC each grew >15%", 1),
        "Cost discipline held: expense growth under 7%",
    ])
    .slide("Revenue is pulling away from costs", kicker="Q3 FINANCIALS")
    .chart(
        df,
        value_columns=["Revenue", "Expenses"],
        emphasize="Revenue",
        headline="Revenue accelerated every quarter while expenses stayed flat",
        show_values=True,
        number_format="#,##0",
        y_title="USD (k)",
    )
    .section("The path forward")
    .slide("Two ways to run Q4", kicker="DECISION")
    .compare(
        (
            "Hold the course",
            [
                "Keep current headcount",
                "Focus on EMEA expansion",
                ("Lower risk, ~8% growth", 1),
            ],
        ),
        (
            "Accelerate",
            [
                "Add 6 AEs in APAC",
                "Launch partner program",
                ("Higher risk, ~15% growth", 1),
            ],
        ),
    )
    .slide("Recommendation", kicker="DECISION", notes="pause here for discussion")
    .bullets([
        "Accelerate: the pipeline supports it",
        ("APAC win-rate is 8pt above global average", 1),
        ("Partner-sourced deals close 2x faster", 1),
        "Hold hiring in G&A; all growth headcount to sales",
    ])
    .save("business_review.pptx")
)
