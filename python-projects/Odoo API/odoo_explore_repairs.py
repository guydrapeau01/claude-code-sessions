import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    print("=== Repair Order Fields ===\n")
    fields = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'repair.order', 'fields_get',
        [], {'attributes': ['string', 'type']}
    )
    for fname, finfo in sorted(fields.items()):
        print(f"  {fname:<45} | {finfo['type']:<15} | {finfo['string']}")

    print("\n=== Repair Order Stages/States ===\n")
    sample = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'repair.order', 'search_read',
        [[]], {'fields': ['id', 'name', 'state', 'product_id', 'lot_id',
                          'partner_id', 'create_date', 'user_id', 'company_id'],
               'limit': 3}
    )
    for r in sample:
        print(f"  Repair: {r['name']}  |  State: {r['state']}  |  Product: {r['product_id']}")

    print("\n=== Distinct States ===\n")
    all_repairs = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'repair.order', 'search_read',
        [[]], {'fields': ['state']}
    )
    states = set(r['state'] for r in all_repairs)
    for s in sorted(states):
        cnt = sum(1 for r in all_repairs if r['state'] == s)
        print(f"  State: {s:<20} | Count: {cnt}")

    print(f"\n  Total repairs: {len(all_repairs)}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
