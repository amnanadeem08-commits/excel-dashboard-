"""Rule-based marketplace copy generation for Excel Dashboard Studio."""

from pathlib import Path
import re

OUTPUT_DIR = Path(__file__).resolve().parents[1] / "outputs" / "marketplace_copy"

TEMPLATE_COPY = {
    "Small Business Bookkeeping / Profit & Loss Dashboard": {
        "title": "Small Business Bookkeeping and Profit & Loss Excel Dashboard",
        "short": "A beginner-friendly Excel bookkeeping dashboard for tracking income, expenses, profit, tax estimates, and monthly business performance.",
        "long": "This editable Excel dashboard helps small business owners organize income and expenses, review profit and loss, estimate tax/VAT/GST, and understand monthly performance without complicated accounting software. It includes sample data, lookup-ready categories, clean dashboard KPIs, and instruction notes designed for non-technical buyers.",
        "features": ["Income and expense trackers", "Profit and loss dashboard", "Monthly summary view", "Tax/VAT/GST estimate", "Category lookup structure", "Sample data included", "Printable dashboard layout"],
        "use_cases": ["Small business bookkeeping", "Freelancer income tracking", "Monthly profit review", "Simple tax estimate preparation", "Digital product marketplace template"],
        "price": "$12 - $29",
        "tags": ["bookkeeping", "profit loss", "excel dashboard", "small business", "expense tracker", "income tracker", "business finance", "excel template", "tax tracker", "monthly budget", "finance planner", "spreadsheet", "digital download"],
        "fiverr": "I will create a small business bookkeeping Excel dashboard template",
        "gumroad": "A ready-to-use Excel bookkeeping and profit dashboard for small businesses, freelancers, and service providers. Track income, expenses, profit, monthly trends, and tax estimates in one editable workbook.",
    },
    "Sales Dashboard": {
        "title": "Sales Dashboard Excel Template with Customers, Products, and Monthly Summary",
        "short": "A clean Excel sales dashboard for tracking orders, customers, products, sales trends, profit, and top performers.",
        "long": "This sales dashboard template gives business owners and sales teams a simple way to track orders, products, customers, profit, and monthly sales performance. It uses lookup-friendly customer and product sheets, sample sales data, KPI cards, and chart-ready summaries so buyers can start quickly and customize it for their own business.",
        "features": ["Sales data tracker", "Customer and product lookup sheets", "Monthly sales summary", "Profit tracking", "Top product and customer KPIs", "Chart-ready data", "Editable sample records"],
        "use_cases": ["Sales reporting", "Product performance tracking", "Customer sales analysis", "Monthly revenue review", "Small sales team dashboard"],
        "price": "$15 - $35",
        "tags": ["sales dashboard", "excel sales", "sales tracker", "customer tracker", "product sales", "revenue dashboard", "sales report", "excel template", "monthly sales", "profit tracker", "business dashboard", "spreadsheet", "digital download"],
        "fiverr": "I will create a professional Excel sales dashboard template",
        "gumroad": "A practical Excel sales dashboard for monitoring revenue, orders, customers, products, profit, and monthly performance with editable lookup tables and sample data.",
    },
    "Inventory Management Dashboard": {
        "title": "Inventory Management Excel Dashboard with Stock In, Stock Out, and Reorder Report",
        "short": "An editable Excel inventory dashboard for tracking products, suppliers, stock movement, reorder levels, and stock value.",
        "long": "This inventory management workbook helps small businesses monitor products, suppliers, stock in, stock out, current stock, reorder alerts, and stock value. It is designed to be easy to edit, with sample data, lookup-ready product IDs, conditional reorder flags, and a dashboard layout that works well for shops, warehouses, and online sellers.",
        "features": ["Product master sheet", "Supplier lookup sheet", "Stock in and stock out logs", "Current stock formulas", "Reorder report", "Low stock indicators", "Stock value tracking"],
        "use_cases": ["Retail inventory tracking", "Warehouse stock control", "Ecommerce product monitoring", "Supplier stock review", "Reorder planning"],
        "price": "$18 - $39",
        "tags": ["inventory tracker", "stock dashboard", "excel inventory", "reorder report", "stock management", "supplier tracker", "warehouse", "retail inventory", "excel template", "product tracker", "small business", "spreadsheet", "digital download"],
        "fiverr": "I will create an Excel inventory management dashboard template",
        "gumroad": "A local-first Excel inventory dashboard for tracking products, stock movement, suppliers, reorder levels, and current stock value with editable sample data.",
    },
    "Invoice & Receipt Template with Logo": {
        "title": "Invoice and Receipt Excel Template with Logo, Tax, and Payment Tracker",
        "short": "A printable Excel invoice and receipt template with logo area, customer lookup, product table, tax, totals, and payment tracking.",
        "long": "This Excel invoice and receipt workbook is built for small businesses that need a simple, editable, and printable billing system. It includes invoice and receipt layouts, customer and product lookup sheets, tax/VAT/GST fields, line item totals, invoice log, payment tracker, and clear buyer instructions.",
        "features": ["Printable invoice area", "Printable receipt area", "Logo placeholder", "Customer lookup", "Product/item table", "Discount and tax fields", "Invoice log", "Payment tracker"],
        "use_cases": ["Small business invoicing", "Freelance billing", "Receipt generation", "Payment tracking", "Printable client documents"],
        "price": "$10 - $25",
        "tags": ["invoice template", "receipt template", "excel invoice", "payment tracker", "billing", "tax invoice", "small business", "freelancer", "printable invoice", "excel template", "digital download", "client invoice", "spreadsheet"],
        "fiverr": "I will create a branded Excel invoice and receipt template",
        "gumroad": "A printable Excel invoice and receipt template with logo area, customer lookup, product line items, discount, tax/VAT/GST, payment status, and invoice tracking.",
    },
    "CRM Customer Tracker": {
        "title": "CRM Customer Tracker Excel Dashboard with Leads, Deals, and Follow-ups",
        "short": "A simple Excel CRM tracker for managing customers, leads, deals, follow-ups, pipeline value, and conversion performance.",
        "long": "This CRM customer tracker gives small teams and solo business owners a simple way to organize customers, leads, deals, follow-ups, and sales pipeline activity. It includes sample data, lookup formulas, status dropdowns, KPI cards, and chart-ready summaries for pipeline review and customer follow-up planning.",
        "features": ["Customer tracker", "Lead tracker", "Deal pipeline", "Follow-up log", "Pipeline summary", "Conversion KPIs", "Status dropdowns", "Sample CRM data"],
        "use_cases": ["Small business CRM", "Lead follow-up tracking", "Sales pipeline review", "Client relationship management", "Freelancer prospect tracking"],
        "price": "$15 - $35",
        "tags": ["crm tracker", "customer tracker", "lead tracker", "sales pipeline", "follow up tracker", "excel crm", "deal tracker", "client tracker", "excel dashboard", "small business", "spreadsheet", "digital download", "sales tracker"],
        "fiverr": "I will create an Excel CRM customer tracker dashboard",
        "gumroad": "An editable Excel CRM dashboard for tracking customers, leads, deals, follow-ups, pipeline value, and conversion rate with beginner-friendly sample data.",
    },
}


def generate_marketplace_copy(template_name):
    """Generate marketplace copy markdown for a template and save it locally."""
    if template_name not in TEMPLATE_COPY:
        raise ValueError(f"Unsupported template name: {template_name}")

    copy = TEMPLATE_COPY[template_name]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_DIR / f"{_slugify(template_name)}_marketplace_copy.md"
    path.write_text(_render_markdown(template_name, copy), encoding="utf-8")
    return path


def generate_all_marketplace_copy():
    """Generate marketplace copy files for every supported template."""
    return [generate_marketplace_copy(template_name) for template_name in TEMPLATE_COPY]


def _render_markdown(template_name, copy):
    return f"""# {copy['title']}

## Product Title
{copy['title']}

## Short Product Description
{copy['short']}

## Long Product Description
{copy['long']}

## Feature List
{_bullet_list(copy['features'])}

## Buyer Instructions
- Download the Excel workbook after purchase.
- Open the file in Microsoft Excel or a compatible spreadsheet app.
- Replace the sample data with your own business data.
- Update business details, logo, colors, currency, and tax settings before using the file with customers or reports.
- Keep lookup IDs and formula columns intact to avoid breaking linked sheets.
- Export printable sheets or dashboards to PDF when needed.

## Use Cases
{_bullet_list(copy['use_cases'])}

## Suggested Price Range
{copy['price']}

## Screenshot Checklist
- Dashboard or main printable sheet
- Data input sheet with sample rows
- Lookup/reference sheet
- Summary or tracker sheet
- Instructions sheet
- Close-up of KPI cards
- Example of dropdowns or formulas

## Etsy Tags
{', '.join(copy['tags'])}

## Fiverr Gig Title
{copy['fiverr']}

## Gumroad Product Description
{copy['gumroad']}

---
Generated locally by Excel Dashboard Studio for: {template_name}
"""


def _bullet_list(items):
    return "\n".join(f"- {item}" for item in items)


def _slugify(value):
    return re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")
