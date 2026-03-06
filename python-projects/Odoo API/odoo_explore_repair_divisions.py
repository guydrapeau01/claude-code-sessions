import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    print("=== Distinct Divisions on Repair Orders ===\n")
    repairs = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'repair.order', 'search_read',
        [[['state', 'in', ['draft', 'confirmed', 'under_repair']]]],
        {'fields': ['id', 'name', 'division_id', 'state']}
    )

    divisions = {}
    no_division = 0
    for r in repairs:
        if r.get('division_id'):
            div_id   = r['division_id'][0]
            div_name = r['division_id'][1]
            divisions[div_id] = div_name
        else:
            no_division += 1

    print(f"  Found {len(divisions)} distinct divisions:\n")
    for did, dname in sorted(divisions.items(), key=lambda x: x[1]):
        cnt = sum(1 for r in repairs if r.get('division_id') and r['division_id'][0] == did)
        print(f"  ID: {did:<6} | Name: {dname:<40} | Active Repairs: {cnt}")

    print(f"\n  Repairs with NO division: {no_division}")
    print(f"  Total active repairs: {len(repairs)}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
