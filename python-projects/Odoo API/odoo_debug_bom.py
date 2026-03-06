import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    for ref in ['101104']:
        prod = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
            [[['default_code', '=', ref]]], {'fields': ['id', 'name', 'product_tmpl_id']}
        )
        if not prod:
            print(f"{ref}: NOT FOUND\n")
            continue
        p   = prod[0]
        tid = p['product_tmpl_id'][0]

        sups = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.supplierinfo', 'search_read',
            [[['product_tmpl_id', '=', tid]]],
            {'fields': ['name', 'delay', 'product_code']}
        )
        boms = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_read',
            [[['product_tmpl_id.default_code', '=', ref]]],
            {'fields': ['id', 'product_qty', 'type']}
        )

        print(f"{ref} — {p['name']}")
        print(f"  Suppliers: {[s['name'][1] for s in sups] if sups else 'NONE'}")
        bom_summary = [(f"ID:{b['id']} type:{b['type']}") for b in boms]
        print(f"  Has BOM:   {bool(boms)} {bom_summary}")

        if boms:
            lines = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom.line', 'search_read',
                [[['bom_id', '=', boms[0]['id']]]],
                {'fields': ['product_id', 'product_qty', 'child_bom_id']}
            )
            print(f"  BOM components:")
            for l in lines:
                child = f" → has child BOM" if l.get('child_bom_id') else ""
                print(f"    {l['product_id'][1][:50]}  qty:{l['product_qty']}{child}")
        print()

except Exception as e:
    import traceback
    traceback.print_exc()

input("\nPress Enter to close...")
