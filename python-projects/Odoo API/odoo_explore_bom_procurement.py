import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    MONTHLY_PLAN = {'101336': 150, '101769': 20, '101711': 40, '102237': 20}

    # --- Resolve product info ---
    products = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['default_code', 'in', list(MONTHLY_PLAN.keys())]]],
        {'fields': ['id', 'name', 'default_code', 'product_tmpl_id']}
    )
    tmpl_ids    = [p['product_tmpl_id'][0] for p in products]
    variant_ids = [p['id'] for p in products]
    print(f"Variant IDs:  {variant_ids}")
    print(f"Template IDs: {tmpl_ids}\n")

    # --- Count all BOMs ---
    total_boms = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_count', [[]])
    print(f"Total BOMs in system: {total_boms}\n")

    # --- Show first 20 BOMs raw ---
    print("=== First 20 BOMs (raw) ===\n")
    all_boms = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_read',
        [[]], {'fields': ['id', 'product_id', 'product_tmpl_id', 'product_qty', 'type'], 'limit': 20}
    )
    for b in all_boms:
        prod_name = b['product_id'][1] if b.get('product_id') and b['product_id'] else '(no variant)'
        tmpl_name = b['product_tmpl_id'][1] if b.get('product_tmpl_id') else '(no template)'
        tmpl_id   = b['product_tmpl_id'][0] if b.get('product_tmpl_id') else None
        print(f"  BOM {b['id']:>4} | tmpl_id: {str(tmpl_id):<6} | {tmpl_name[:50]:<50} | variant: {prod_name[:30]}")

    # --- Try searching by name keywords ---
    print("\n=== Searching BOMs containing '101336' or 'C-100' or 'C2' in product name ===\n")
    for keyword in ['tremoFlo', 'tremoflo', 'C-100', 'C2 Device']:
        found = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_read',
            [[['product_tmpl_id.name', 'ilike', keyword]]],
            {'fields': ['id', 'product_tmpl_id', 'product_id'], 'limit': 5}
        )
        if found:
            print(f"  Keyword '{keyword}': {len(found)} BOMs found")
            for b in found:
                print(f"    BOM {b['id']} | {b['product_tmpl_id'][1] if b.get('product_tmpl_id') else 'N/A'}")

    # --- Check if BOMs use default_code in template ---
    print("\n=== Searching by default_code on template ===\n")
    for ref in MONTHLY_PLAN.keys():
        found = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_read',
            [[['product_tmpl_id.default_code', '=', ref]]],
            {'fields': ['id', 'product_tmpl_id', 'product_id']}
        )
        print(f"  Ref {ref}: {len(found)} BOMs")
        for b in found:
            print(f"    BOM {b['id']} | {b['product_tmpl_id'][1] if b.get('product_tmpl_id') else 'N/A'}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
