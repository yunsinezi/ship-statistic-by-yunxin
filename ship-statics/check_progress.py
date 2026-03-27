import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('F:/ship-statics/templates/index.html', encoding='utf-8') as f:
    content = f.read()

print('Step 4 found:', 'Step 4' in content)
print('Step 5 found:', 'Step 5' in content)
print('Step 6 found:', 'Step 6' in content)

for fn in ['runGZ(', 'exportGZ(', 'setGZStatus(', 'renderGZTable(', 'gzData']:
    pos = content.find(fn)
    print(f'  JS {fn}:', 'YES' if pos > 0 else 'MISSING')

for div in ['gzPanel', 'gzTableCard', 'criteriaPanel']:
    pos = content.find(f'id="{div}"')
    print(f'  div #{div}:', 'YES' if pos > 0 else 'MISSING')

print('Total HTML size:', len(content))
