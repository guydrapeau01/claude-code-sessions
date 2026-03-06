import xmlrpc.client
import odoo_config as cfg
from collections import defaultdict

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    PLAN_PRODUCTS = cfg.ALL_PLAN_PRODUCTS
    MONTHLY_PLAN  = cfg.MONTHLY_PRODUCTION_PLAN

    # Resolve BOMs
    fin_prods = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['default_code', 'in', PLAN_PRODUCTS]]],
        {'fields': ['id', 'name', 'default_code', 'product_tmpl_id']}
    )
    bom_by_ref = {}
    for p in fin_prods:
        boms = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'search_read',
            [[['product_tmpl_id.default_code', '=', p['default_code']]]],
            {'fields': ['id', 'product_qty']}
        )
        if boms:
            bom_by_ref[p['default_code']] = boms[0]

    # Collect all component IDs from BOMs
    all_comp_ids = set()
    def get_comp_ids(bom_id, visited=None):
        if visited is None: visited = set()
        if bom_id in visited: return
        visited.add(bom_id)
        lines = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom.line', 'search_read',
            [[['bom_id', '=', bom_id]]],
            {'fields': ['product_id', 'child_bom_id']}
        )
        for l in lines:
            all_comp_ids.add(l['product_id'][0])
            if l.get('child_bom_id'):
                get_comp_ids(l['child_bom_id'][0], visited)

    for ref, bom in bom_by_ref.items():
        get_comp_ids(bom['id'])

    all_comp_ids = list(all_comp_ids)
    print(f"Total BOM components: {len(all_comp_ids)}\n")

    # Get stock
    quants = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.quant', 'search_read',
        [[['product_id', 'in', all_comp_ids], ['location_id.usage', '=', 'internal']]],
        {'fields': ['product_id', 'quantity', 'reserved_quantity']}
    )
    from collections import defaultdict
    _raw = defaultdict(lambda: [0.0, 0.0])
    for q in quants:
        _raw[q['product_id'][0]][0] += q['quantity']
        _raw[q['product_id'][0]][1] += q['reserved_quantity']
    stock_by_comp = {pid: max(0, v[0]-v[1]) for pid, v in _raw.items()}

    # Get open incoming
    moves = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.move', 'search_read',
        [[['product_id', 'in', all_comp_ids],
          ['state', 'in', ['waiting','confirmed','assigned','partially_available']],
          ['picking_type_id.code', '=', 'incoming']]],
        {'fields': ['product_id', 'product_qty', 'quantity_done', 'picking_id']}
    )
    pick_ids = list({m['picking_id'][0] for m in moves if m.get('picking_id')})
    pick_states = {}
    if pick_ids:
        picks = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'stock.picking', 'read',
            [pick_ids], {'fields': ['id', 'state']}
        )
        pick_states = {p['id']: p['state'] for p in picks}
    incoming_by_comp = defaultdict(float)
    for m in moves:
        pick_state = pick_states.get(m['picking_id'][0] if m.get('picking_id') else None, '')
        ordered    = m.get('product_qty') or 0
        done       = m.get('quantity_done') or 0
        remaining  = ordered if pick_state != 'done' else max(0, ordered - done)
        if remaining > 0:
            incoming_by_comp[m['product_id'][0]] += remaining

    # Find zero stock AND zero incoming parts
    zero_parts = [cid for cid in all_comp_ids
                  if stock_by_comp.get(cid, 0) == 0 and incoming_by_comp.get(cid, 0) == 0]
    print(f"Parts with 0 stock AND 0 incoming: {len(zero_parts)}\n")

    # Get product and supplier info for zero parts
    comp_prods = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['id', 'in', zero_parts]]],
        {'fields': ['id', 'name', 'default_code', 'product_tmpl_id']}
    )
    tmpl_ids = [p['product_tmpl_id'][0] for p in comp_prods]
    comp_by_id = {p['id']: p for p in comp_prods}

    suppliers = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.supplierinfo', 'search_read',
        [[['product_tmpl_id', 'in', tmpl_ids]]],
        {'fields': ['product_tmpl_id', 'name', 'delay', 'sequence']}
    )
    sup_by_tmpl = {}
    for s in sorted(suppliers, key=lambda x: x.get('sequence', 99)):
        tid = s['product_tmpl_id'][0]
        if tid not in sup_by_tmpl:
            sup_by_tmpl[tid] = s

    # Group by supplier
    by_supplier = defaultdict(list)
    for p in comp_prods:
        tid = p['product_tmpl_id'][0]
        sup = sup_by_tmpl.get(tid, {})
        sup_name = sup['name'][1] if sup.get('name') else 'NO SUPPLIER'
        by_supplier[sup_name].append({
            'ref':  p.get('default_code', '') or '',
            'name': p['name'],
            'lead': sup.get('delay', 0) or 0
        })

    print("=== Zero-stock parts grouped by supplier ===\n")
    print(f"{'Supplier':<35} | {'# Parts':>7} | {'Avg Lead':>9} | Sample parts")
    print("-" * 100)
    for sup_name in sorted(by_supplier.keys(), key=lambda x: -len(by_supplier[x])):
        parts = by_supplier[sup_name]
        avg_lead = sum(p['lead'] for p in parts) / len(parts)
        samples = ', '.join(p['ref'] for p in parts[:5])
        print(f"  {sup_name:<33} | {len(parts):>7} | {avg_lead:>8.0f}d | {samples}")

    print(f"\nTotal zero-stock parts: {len(zero_parts)}")
    print(f"Suppliers involved: {len(by_supplier)}")

except Exception as e:
    import traceback
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
