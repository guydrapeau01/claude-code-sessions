import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    # --- MPS model fields ---
    print("=== MPS Fields (mrp.production.schedule) ===\n")
    fields = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.production.schedule', 'fields_get',
        [], {'attributes': ['string', 'type']}
    )
    for fname, finfo in sorted(fields.items()):
        print(f"  {fname:<45} | {finfo['type']:<15} | {finfo['string']}")

    print("\n=== Current MPS Entries (safe read) ===\n")
    mps_ids = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.production.schedule', 'search', [[]])
    print(f"  Total MPS entries: {len(mps_ids)}")
    if mps_ids:
        mps = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.production.schedule', 'read',
            [mps_ids[:5]], {'fields': list(fields.keys())[:10]}
        )
        for m in mps:
            print(f"  {m}")

    print("\n=== Warehouses ===\n")
    warehouses = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.warehouse', 'search_read',
        [[]], {'fields': ['id', 'name', 'code']}
    )
    for w in warehouses:
        print(f"  ID: {w['id']} | Name: {w['name']} | Code: {w['code']}")

    print("\n=== Reordering Rules fields ===\n")
    rr_fields = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.warehouse.orderpoint', 'fields_get',
        [], {'attributes': ['string', 'type']}
    )
    for fname, finfo in sorted(rr_fields.items()):
        print(f"  {fname:<45} | {finfo['type']:<15} | {finfo['string']}")

    print("\n=== Reordering Rules sample ===\n")
    rules = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.warehouse.orderpoint', 'search_read',
        [[]], {'fields': ['product_id', 'product_min_qty', 'product_max_qty',
                          'qty_on_hand', 'qty_forecast', 'location_id'], 'limit': 10}
    )
    print(f"  Total rules: {len(rules)}")
    for r in rules:
        print(f"  {r['product_id'][1] if r['product_id'] else 'N/A':<50} | "
              f"Min: {r['product_min_qty']} | Max: {r['product_max_qty']} | "
              f"On Hand: {r['qty_on_hand']} | Forecast: {r['qty_forecast']}")

    print("\n=== Open Sales Orders for key products ===\n")
    so_lines = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'sale.order.line', 'search_read',
        [[['order_id.state', 'in', ['sale', 'done']]]],
        {'fields': ['product_id', 'product_uom_qty', 'qty_delivered'], 'limit': 10}
    )
    print(f"  Sample SO lines: {len(so_lines)}")
    for l in so_lines:
        remaining = l['product_uom_qty'] - l['qty_delivered']
        print(f"  {l['product_id'][1] if l['product_id'] else 'N/A':<50} | "
              f"Ordered: {l['product_uom_qty']} | Remaining: {remaining}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
