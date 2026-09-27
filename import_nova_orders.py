import csv
from datetime import datetime
from pathlib import Path

csv_path = Path(r"C:\Users\PC\Downloads\odoo-20.0\odoo-20.0\nova_custom\Nova_Candles_Orders_10.csv")
order_model = env["sale.order"]
partner_model = env["res.partner"]
product_model = env["product.product"]

product = product_model.search([("default_code", "=", "NOVA-IMPORTED-ORDER")], limit=1)
if not product:
    product = product_model.create({
        "name": "Imported order total",
        "default_code": "NOVA-IMPORTED-ORDER",
        "type": "service",
        "sale_ok": True,
        "purchase_ok": False,
        "list_price": 0.0,
        "taxes_id": [(6, 0, [])],
    })

created = 0
skipped = 0
for row in csv.DictReader(csv_path.open(encoding="utf-8-sig", newline="")):
    reference = row["Order Reference"].strip()
    existing = order_model.search([("client_order_ref", "=", reference)], limit=1)
    if existing:
        skipped += 1
        continue

    customer_name = row["Customer"].strip()
    partner = partner_model.search([("name", "=", customer_name)], limit=1)
    if not partner:
        partner = partner_model.create({
            "name": customer_name,
            "customer_rank": 1,
        })

    amount = float(row["Total Amount"])
    order = order_model.create({
        "partner_id": partner.id,
        "client_order_ref": reference,
        "date_order": datetime.strptime(row["Order Date"], "%Y-%m-%d"),
        "order_line": [(0, 0, {
            "product_id": product.id,
            "name": f"Imported total for {reference}",
            "product_uom_qty": 1.0,
            "price_unit": amount,
        })],
        "note": f"Imported from Nova_Candles_Orders_10.csv. Payment status: {row['Payment Status'].strip()}.",
    })

    if row["Status"].strip() == "Locked":
        order.action_confirm()
    elif row["Status"].strip() == "Quotation Sent":
        order.action_quotation_sent()
    created += 1
    print(f"CREATED {reference} {order.name} {order.state} {order.amount_total:.2f}")

env.cr.commit()
print(f"SUMMARY created={created} skipped={skipped} product_id={product.id}")
