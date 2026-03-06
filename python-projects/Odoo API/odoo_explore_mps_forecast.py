import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    # --- Explore mrp.production.schedule.forecast fields ---
    print("=== mrp.production.schedule.forecast fields ===\n")
    try:
        fields = models.execute_kw(cfg.DB, uid, cfg.API_KEY,
            'mrp.production.schedule.forecast', 'fields_get',
            [], {'attributes': ['string', 'type']}
        )
        for fname, finfo in sorted(fields.items()):
            print(f"  {fname:<45} | {finfo['type']:<15} | {finfo['string']}")
    except Exception as e:
        print(f"  ERROR: {e}")

    # --- Get all MPS forecast records ---
    print("\n=== All MPS Forecast Records ===\n")
    try:
        forecasts = models.execute_kw(cfg.DB, uid, cfg.API_KEY,
            'mrp.production.schedule.forecast', 'search_read',
            [[]], {'fields': []}  # fetch all fields
        )
        print(f"  Total forecast records: {len(forecasts)}")
        for f in forecasts:
            print(f"  {f}")
    except Exception as e:
        print(f"  ERROR: {e}")

    # --- Check open MOs (manufacturing orders) for C-100 and C2 ---
    print("\n=== Open Manufacturing Orders ===\n")
    c2_products = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['default_code', 'in', cfg.C2_PRODUCT_REFS + ['101336']]]],
        {'fields': ['id', 'name', 'default_code']}
    )
    prod_ids = [p['id'] for p in c2_products]
    print(f"  Planning products: {[(p['default_code'], p['name']) for p in c2_products]}")

    mos = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.production', 'search_read',
        [[['product_id', 'in', prod_ids],
          ['state', 'in', ['draft', 'confirmed', 'progress']]]],
        {'fields': ['name', 'product_id', 'product_qty', 'date_planned_start', 'state']}
    )
    print(f"  Open MOs: {len(mos)}")
    for mo in mos:
        print(f"  {mo['name']} | {mo['product_id'][1]} | Qty: {mo['product_qty']} | "
              f"Planned: {mo['date_planned_start']} | State: {mo['state']}")

    # --- Check open POs for components ---
    print("\n=== Open Purchase Orders for key products ===\n")
    po_lines = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'purchase.order.line', 'search_read',
        [[['order_id.state', 'in', ['purchase', 'draft']],
          ['product_id', 'in', prod_ids]]],
        {'fields': ['product_id', 'product_qty', 'qty_received',
                    'date_planned', 'order_id']}
    )
    print(f"  Open PO lines for key products: {len(po_lines)}")
    for l in po_lines:
        remaining = l['product_qty'] - l['qty_received']
        print(f"  {l['product_id'][1]:<50} | Ordered: {l['product_qty']} | "
              f"Received: {l['qty_received']} | Remaining: {remaining} | "
              f"Expected: {l['date_planned']}")

    # --- Current stock for key products ---
    print("\n=== Current Stock (Montreal) ===\n")
    quants = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.quant', 'search_read',
        [[['product_id', 'in', prod_ids],
          ['location_id.usage', '=', 'internal'],
          ['location_id.complete_name', 'ilike', 'Montreal']]],
        {'fields': ['product_id', 'quantity', 'reserved_quantity', 'location_id']}
    )
    from collections import defaultdict
    stock_by_product = defaultdict(float)
    for q in quants:
        stock_by_product[q['product_id'][1]] += q['quantity'] - q['reserved_quantity']
    for prod, qty in stock_by_product.items():
        print(f"  {prod:<60} | Available: {qty}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
