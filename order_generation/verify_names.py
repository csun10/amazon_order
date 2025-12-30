import json
from pathlib import Path

prods = []
template_dir = Path(__file__).parent / 'json_template'
for f in template_dir.glob('*.json'):
    try:
        data = json.load(open(f, encoding='utf-8'))
        for p in data.get('products', []):
            prods.append({'sku': p['产品编号'], 'name': p['产品名称']})
    except:
        pass

print(f'Total products: {len(prods)}')
named = [p for p in prods if p['name']]
print(f'Products with names: {len(named)}')
print(f'Products without names: {len(prods) - len(named)}')
print('\nSample products with Chinese names:')
for p in named[:15]:
    if any('\u4e00' <= c <= '\u9fff' for c in p['name']):
        print(f"  {p['sku']}: {p['name']}")
