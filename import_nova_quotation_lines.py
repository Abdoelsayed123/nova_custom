import csv
from collections import defaultdict
from pathlib import Path

csv_path = Path(r"C:\Users\PC\Downloads\odoo-20.0\odoo-20.0\nova_custom\Nova_Candles_Quotations_Lines.csv")
order_model = env["sale.order"]
product_model = env["product.product"]
rows_by_reference = defaultdict(list)

with csv_path.open(encoding="utf-8-sig", newline="") as csv_file:
    for row in csv.DictReader(csv_file):
        rows_by_reference[row["Order Reference"].strip()].append(row)

updated = 0
skipped = 0
for reference, rows in rows_by_reference.items():
    order = order_model.search([("client_order_ref", "=", reference)], limit=1)
    if not order:
        raise ValueError(f"Sales order not found: {reference}")

    markers = [f"Imported quotation line: {row['Product'].strip()}" for row in rows]
    if all(order.order_line.filtered(lambda line, marker=marker: line.name == marker) for marker in markers):
        skipped += 1
        continue

    line_values = []
    for row in rows:
        product_name = row["Product"].strip()
        product = product_model.search([("name", "=", product_name)], limit=1)
        if not product:
            product = product_model.create({
                "name": product_name,
                "default_code": f"NOVA-CSV-{len(product_model.search([('default_code', 'like', 'NOVA-CSV-')])) + 1:03d}",
                "type": "consu",
                "sale_ok": True,
                "purchase_ok": False,
                "list_price": float(row["Unit Price"]),
                "taxes_id": [(6, 0, [])],
            })

        quantity = float(row["Quantity"])
        subtotal = float(row["Subtotal"])
        line_values.append({
            "product_id": product.id,
            "name": f"Imported quotation line: {product_name}",
            "product_uom_qty": quantity,
            "price_unit": subtotal / quantity,
        })

    generic_line = order.order_line.filtered(
        lambda line: line.product_id.default_code == "NOVA-IMPORTED-ORDER"
    )[:1]
    if generic_line:
        generic_line.write(line_values[0])
        line_values = line_values[1:]

    if line_values:
        order.write({"order_line": [(0, 0, values) for values in line_values]})
    updated += 1
    print(f"UPDATED {reference} lines={len(rows)} total={order.amount_total:.2f}")

env.cr.commit()
print(f"SUMMARY updated={updated} skipped={skipped}")
