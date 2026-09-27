# NOVA Custom Addons

Place each installable NOVA module in its own folder directly under this directory.
Keep company-specific models, views, reports, and assets here instead of editing Odoo core addons.

The starter module is `nova_base`. The `nova_candles` module adds wick type and burn time to product forms.
After starting Odoo with `-c odoo.conf`, update the Apps list to discover the modules, then install **NOVA Candle Products**.

## Enterprise Demo Dataset

`nova_erp_enterprise_dataset.py` generates a synthetic Egyptian-market workbook with 1,000 products, 500 partners, 1,000 sales transactions, and 300 employees. Names, financial values, and transaction histories are generated for testing; the `SIM` phone values are intentionally non-dialable. Do not treat these rows as actual company records.

Install the dependency and generate the workbook:

```powershell
python -m pip install -r requirements.txt
python nova_erp_enterprise_dataset.py --output Nova_Candles_Enterprise_Demo.xlsx
```

The optional XML-RPC integration upserts products by internal reference into `product.template` and partners by reference into `res.partner`. It does not upload the sales or employee sheets, post invoices, set inventory quantities, or create payroll records. Those require separate Odoo models, modules, and permissions. Wholesale price and credit limit are sent only when compatible writable custom fields exist; otherwise the values remain in the workbook. Opening balance quantities also remain in Excel so the script cannot silently change stock.

For a separate import workbook, the product sheet can be named `Products` or `المنتجات`; the partner sheet can be `Partners`, `Customers`, `Distributors`, or `العملاء والموزعين`. Product headers accept `SKU`, `اسم المنتج`, `التكلفة`, `سعر البيع`, and `المخزون`. Partner headers accept `اسم العميل`, `المدينة`, `رقم الهاتف`, and `التصنيف`. Customer reference codes are optional; when absent, the script derives a stable reference from name, city, and phone so repeat syncs update the same partner.

Each row is validated independently. Errors are written to stderr with the sheet, Excel row number, and key. Creates are sent in batches of 100; if a batch fails, its rows are retried individually so one bad record does not stop the rest. The final report shows total, created, updated, and failed counts. Product stock is deliberately not written to `product.template`; use Odoo inventory adjustments for opening quantities.

Set credentials in environment variables, never in the script:

```powershell
$env:ODOO_URL = "https://your-odoo.example.com"
$env:ODOO_DB = "your_database"
$env:ODOO_USERNAME = "integration-user@example.com"
$env:ODOO_API_KEY = "your-api-key"
```

Preview the supported uploads first. Preview is the default and performs no database writes:

```powershell
python nova_erp_enterprise_dataset.py --input-workbook Nova_Candles_Enterprise_Demo.xlsx --upload-products --upload-partners
```

Only after reviewing the preview, add `--apply` to create or update records. The integration never deletes records. Use a dedicated Odoo user with only the required product and contact permissions, and test against a backup or staging database before production.