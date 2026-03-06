import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    # Fetch active repairs with division and partner info
    repairs = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'repair.order', 'search_read',
        [[['state', 'in', ['draft', 'confirmed', 'under_repair']]]],
        {'fields': ['id', 'name', 'state', 'division_id', 'partner_id']}
    )
    print(f"Total active repairs (before exclusions): {len(repairs)}\n")

    print("=== Division breakdown (raw) ===")
    from collections import Counter
    div_counts = Counter()
    for r in repairs:
        div_name = r['division_id'][1] if r.get('division_id') else 'NO DIVISION'
        div_counts[div_name] += 1
    for name, cnt in sorted(div_counts.items()):
        print(f"  '{name}': {cnt}")

    print("\n=== Checking cfg.REPAIR_DIVISIONS match ===")
    for div in cfg.REPAIR_DIVISIONS:
        cnt = sum(1 for r in repairs if r.get('division_id') and r['division_id'][1] == div)
        print(f"  cfg: '{div}' → {cnt} repairs")

    print("\n=== Sample Americas repairs ===")
    americas = [r for r in repairs if r.get('division_id') and 'Americas' in r['division_id'][1]]
    for r in americas[:5]:
        partner = r['partner_id'][1] if r.get('partner_id') else 'No partner'
        print(f"  {r['name']} | State: {r['state']} | Division: {r['division_id'][1]} | Partner: {partner}")

    print("\n=== Excluded customers check ===")
    excluded = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'res.partner', 'search_read',
        [[['name', 'in', cfg.REPAIR_EXCLUDED_CUSTOMERS]]],
        {'fields': ['id', 'name']}
    )
    print(f"  Excluded partner IDs: {[(p['name'], p['id']) for p in excluded]}")
    excluded_ids = [p['id'] for p in excluded]
    excluded_count = sum(1 for r in americas if r.get('partner_id') and r['partner_id'][0] in excluded_ids)
    print(f"  Americas repairs that would be excluded: {excluded_count}")
    print(f"  Americas repairs after exclusion: {len(americas) - excluded_count}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
