import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    print("=== Searching C2 products by default_code ===\n")
    c2_products = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['default_code', 'in', cfg.C2_PRODUCT_REFS]]],
        {'fields': ['id', 'name', 'default_code']}
    )
    if c2_products:
        for p in c2_products:
            print(f"  DB ID: {p['id']}  |  Internal Ref: {p['default_code']}  |  Name: {p['name']}")
    else:
        print("  ❌ No products found! Trying product.template instead...\n")
        c2_tmpl = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.template', 'search_read',
            [[['default_code', 'in', cfg.C2_PRODUCT_REFS]]],
            {'fields': ['id', 'name', 'default_code']}
        )
        for p in c2_tmpl:
            print(f"  Template DB ID: {p['id']}  |  Internal Ref: {p['default_code']}  |  Name: {p['name']}")

    print("\n=== Sample tickets with product assigned ===\n")
    tickets = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'helpdesk.ticket', 'search_read',
        [[['product_id', '!=', False]]],
        {'fields': ['id', 'product_id'], 'limit': 5}
    )
    for t in tickets:
        print(f"  Ticket #{t['id']}  |  product_id field: {t['product_id']}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
