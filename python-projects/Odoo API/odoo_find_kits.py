import xmlrpc.client
import odoo_config as cfg

try:
    common = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/common")
    uid    = common.authenticate(cfg.DB, cfg.USERNAME, cfg.API_KEY, {})
    models = xmlrpc.client.ServerProxy(f"{cfg.URL}/xmlrpc/2/object")

    DEVICE_REFS = [
        '101336', '101711', '101769', '102237',
        '101490', '101759', '101760', '102240'
    ]

    devices = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.product', 'search_read',
        [[['default_code', 'in', DEVICE_REFS]]],
        {'fields': ['id', 'name', 'default_code', 'product_tmpl_id']}
    )
    device_by_id    = {d['id']: d for d in devices}
    device_tmpl_ids = [d['product_tmpl_id'][0] for d in devices]

    bom_lines = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom.line', 'search_read',
        [[['product_id.product_tmpl_id', 'in', device_tmpl_ids]]],
        {'fields': ['bom_id', 'product_id', 'product_qty']}
    )

    parent_bom_ids = list({l['bom_id'][0] for l in bom_lines})
    if not parent_bom_ids:
        print("No parent BOMs found.")
    else:
        parent_boms = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'mrp.bom', 'read',
            [parent_bom_ids], {'fields': ['id', 'product_tmpl_id', 'type']}
        )
        bom_map = {b['id']: b for b in parent_boms}

        kit_tmpl_ids = list({b['product_tmpl_id'][0] for b in parent_boms if b.get('product_tmpl_id')})
        kit_prods = models.execute_kw(cfg.DB, uid, cfg.API_KEY, 'product.template', 'read',
            [kit_tmpl_ids], {'fields': ['id', 'name', 'default_code']}
        )
        kit_by_tmpl = {p['id']: p for p in kit_prods}

        # Deduplicate: kit_ref -> {name, type, devices used}
        seen = {}
        for line in bom_lines:
            bom      = bom_map.get(line['bom_id'][0], {})
            tid      = bom.get('product_tmpl_id', [None])[0] if bom.get('product_tmpl_id') else None
            kit      = kit_by_tmpl.get(tid, {})
            kit_ref  = kit.get('default_code', '')
            kit_name = kit.get('name', '')
            btype    = bom.get('type', '')
            dev      = device_by_id.get(line['product_id'][0], {})
            dev_ref  = dev.get('default_code', '')
            if not kit_ref:
                continue
            if kit_ref not in seen:
                seen[kit_ref] = {'name': kit_name, 'type': btype, 'devices': []}
            if dev_ref and dev_ref not in seen[kit_ref]['devices']:
                seen[kit_ref]['devices'].append(dev_ref)

        lines = []
        lines.append("# ================================================================")
        lines.append("# KITS — copy/paste into odoo_config.py, fill in monthly qty")
        lines.append(f"# Found {len(seen)} unique kits containing your devices")
        lines.append("# ================================================================")
        lines.append("KIT_PRODUCTS = {")
        for ref, info in sorted(seen.items()):
            btype_comment = "kit" if info['type'] == 'phantom' else info['type']
            devices_str   = ', '.join(sorted(info['devices']))
            lines.append(f"    '{ref}': 0,   # [{btype_comment}]  devices: {devices_str}")
            lines.append(f"               # {info['name']}")
        lines.append("}")

        output = "\n".join(lines)
        print(output)

        out_file = "kits_found.txt"
        with open(out_file, 'w', encoding='utf-8') as f:
            f.write(output + "\n")
        print(f"\n✅ Saved to: {out_file}")

except Exception as e:
    import traceback
    traceback.print_exc()

input("\nPress Enter to close...")
