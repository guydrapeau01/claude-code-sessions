import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    for ref in ['101306', '102175']:
        prod = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
            [[['default_code', '=', ref]]], {'fields': ['id', 'name', 'product_tmpl_id']}
        )
        if not prod:
            print(f"{ref}: NOT FOUND\n")
            continue
        p = prod[0]
        tid = p['product_tmpl_id'][0]
        sups = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.supplierinfo', 'search_read',
            [[['product_tmpl_id', '=', tid]]],
            {'fields': ['name', 'delay', 'min_qty', 'price', 'product_code']}
        )
        print(f"{ref} — {p['name']}")
        if sups:
            for s in sups:
                print(f"  Supplier: {s['name'][1] if s.get('name') else 'N/A'} | Lead: {s['delay']}d | Code: {s.get('product_code','')}")
        else:
            print("  NO SUPPLIER configured")
        print()

except Exception as e:
    import traceback
    traceback.print_exc()

input("\nPress Enter to close...")
