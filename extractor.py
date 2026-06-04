"""
extractor.py
------------
Sends a PDF (as bytes) to the Claude API and returns structured financial data.
Works for any company's financial statements — income statement, cash flow, balance sheet.
"""

import anthropic
import base64
import json
import re
from typing import Optional

EXTRACTION_PROMPT = """You are a senior financial data extraction expert.

Carefully read this financial statement PDF and extract ALL financial data you can find.

Return ONLY a valid JSON object with these exact fields. Use null for any value you cannot find or are uncertain about. NEVER guess or infer numbers.

{
  "company_name": "string — legal name of the company",
  "reporting_period_label": "string — e.g. 'Year Ended December 31, 2023' or 'Three Months Ended March 31, 2024'",
  "period": "string — standardized: YYYY_FY for annual, YYYY_Q1/Q2/Q3/Q4 for quarterly",
  "currency": "string — e.g. 'CAD', 'USD'",
  "unit": "string — e.g. '$', 'thousands', 'millions'",
  "audit_status": "string — 'Audited', 'Unaudited', or 'Review Engagement'",

  "income_statement": {
    "revenue": number or null,
    "cogs": number or null,
    "gross_profit": number or null,
    "gross_margin_pct": number or null,
    "operating_expenses": number or null,
    "depreciation_amortization": number or null,
    "rent_occupancy": number or null,
    "distribution_expense": number or null,
    "total_opex": number or null,
    "operating_income": number or null,
    "ebit": number or null,
    "interest_expense": number or null,
    "other_income": number or null,
    "government_assistance": number or null,
    "income_before_tax": number or null,
    "income_tax_current": number or null,
    "income_tax_deferred": number or null,
    "net_income": number or null,
    "ebitda": number or null
  },

  "cash_flow": {
    "net_income": number or null,
    "depreciation_amortization": number or null,
    "stock_based_compensation": number or null,
    "working_capital_change": number or null,
    "operating_cash_flow": number or null,
    "capex": number or null,
    "proceeds_asset_disposals": number or null,
    "investing_cash_flow": number or null,
    "financing_cash_flow": number or null,
    "free_cash_flow": number or null,
    "net_change_in_cash": number or null
  },

  "balance_sheet": {
    "cash": number or null,
    "accounts_receivable": number or null,
    "inventories": number or null,
    "prepaid_other_current": number or null,
    "total_current_assets": number or null,
    "ppe_net": number or null,
    "total_assets": number or null,
    "accounts_payable": number or null,
    "short_term_debt": number or null,
    "deferred_revenue": number or null,
    "total_current_liabilities": number or null,
    "long_term_debt": number or null,
    "total_liabilities": number or null,
    "retained_earnings": number or null,
    "total_equity": number or null
  },

  "segment_data": {
    "segments": []
  },

  "additional_metrics": {
    "notes": "string — any important notes, one-time items, restatements, or caveats"
  }
}

Rules:
- All monetary values must be raw numbers (e.g. 27069791 not "$27.1M")
- If the statement shows values in thousands, convert to actual dollars (multiply by 1000)
- If the statement shows values in millions, convert to actual dollars (multiply by 1000000)
- Preserve signs: expenses are positive, losses are negative
- gross_margin_pct should be a decimal (e.g. 0.207 for 20.7%)
- Return ONLY the JSON object — no markdown, no explanation"""


def extract_financials(pdf_bytes: bytes, api_key: str, model: str = "claude-sonnet-4-6") -> dict:
    """
    Extract financial data from a PDF using the Claude API.

    Args:
        pdf_bytes: Raw PDF file bytes
        api_key: Anthropic API key
        model: Claude model to use

    Returns:
        dict with extracted financial data, plus '_extraction_error' key if failed
    """
    client = anthropic.Anthropic(api_key=api_key)

    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    try:
        response = client.messages.create(
            model=model,
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": EXTRACTION_PROMPT,
                        },
                    ],
                }
            ],
        )

        raw_text = response.content[0].text.strip()

        # Strip markdown code fences if present
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)

        data = json.loads(raw_text)
        data["_raw_response"] = raw_text
        return data

    except json.JSONDecodeError as e:
        return {
            "_extraction_error": f"Failed to parse JSON from Claude response: {e}",
            "_raw_response": raw_text if "raw_text" in dir() else "",
        }
    except anthropic.APIError as e:
        return {"_extraction_error": f"Claude API error: {e}"}
    except Exception as e:
        return {"_extraction_error": f"Unexpected error: {e}"}


def normalize_extracted_data(raw: dict) -> dict:
    """
    Flatten and normalize the nested extraction output into the flat dict
    format expected by excel_builder.py.
    """
    if "_extraction_error" in raw:
        return raw

    inc = raw.get("income_statement", {}) or {}
    cf  = raw.get("cash_flow", {}) or {}
    bs  = raw.get("balance_sheet", {}) or {}
    add = raw.get("additional_metrics", {}) or {}

    def n(val, default=None):
        """Return val if not None, else default."""
        return val if val is not None else default

    return {
        # Identity
        "company_name":           n(raw.get("company_name"), "Unknown Company"),
        "reporting_period_label": n(raw.get("reporting_period_label"), ""),
        "period":                 n(raw.get("period"), "UNKNOWN"),
        "currency":               n(raw.get("currency"), "USD"),
        "unit":                   n(raw.get("unit"), "$"),
        "audit_status":           n(raw.get("audit_status"), "Unaudited"),
        "source_file":            raw.get("_source_filename", ""),

        # Income Statement
        "revenue":               n(inc.get("revenue")),
        "cogs":                  n(inc.get("cogs")),
        "gross_profit":          n(inc.get("gross_profit")),
        "gross_margin_pct":      n(inc.get("gross_margin_pct")),
        "opex_operating":        n(inc.get("operating_expenses"), 0),
        "opex_amortization":     n(inc.get("depreciation_amortization"), 0),
        "opex_rent":             n(inc.get("rent_occupancy"), 0),
        "opex_distribution":     n(inc.get("distribution_expense"), 0),
        "total_opex":            n(inc.get("total_opex")),
        "operating_income":      n(inc.get("operating_income"), n(inc.get("ebit"))),
        "interest_expense":      n(inc.get("interest_expense"), 0),
        "other_income":          n(inc.get("other_income"), 0),
        "government_assistance": n(inc.get("government_assistance"), 0),
        "income_before_tax":     n(inc.get("income_before_tax")),
        "current_tax":           n(inc.get("income_tax_current"), 0),
        "future_tax":            n(inc.get("income_tax_deferred"), 0),
        "net_income":            n(inc.get("net_income")),
        "ebitda":                n(inc.get("ebitda")),

        # Cash Flow
        "ocf":         n(cf.get("operating_cash_flow")),
        "capex":       n(cf.get("capex"), 0),
        "investing_cf":n(cf.get("investing_cash_flow")),
        "financing_cf":n(cf.get("financing_cash_flow")),
        "amortization":n(cf.get("depreciation_amortization"), n(inc.get("depreciation_amortization"), 0)),
        "sbc":         n(cf.get("stock_based_compensation"), 0),

        # Balance Sheet
        "cash":               n(bs.get("cash"), 0),
        "accounts_receivable":n(bs.get("accounts_receivable")),
        "inventories":        n(bs.get("inventories")),
        "total_current_assets":n(bs.get("total_current_assets")),
        "capital_assets":     n(bs.get("ppe_net")),
        "total_assets":       n(bs.get("total_assets")),
        "accounts_payable":   n(bs.get("accounts_payable")),
        "short_term_debt":    n(bs.get("short_term_debt"), 0),
        "lt_debt_current":    0,
        "lt_debt_noncurrent": n(bs.get("long_term_debt"), 0),
        "total_debt":         n(bs.get("long_term_debt"), 0),
        "total_current_liabilities": n(bs.get("total_current_liabilities")),
        "total_liabilities":  n(bs.get("total_liabilities")),
        "retained_earnings":  n(bs.get("retained_earnings")),
        "total_equity":       n(bs.get("total_equity")),

        # Misc
        "notes": n(add.get("notes"), ""),
        "segment_data": raw.get("segment_data", {}),
    }
