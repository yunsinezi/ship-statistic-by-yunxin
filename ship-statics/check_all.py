import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'F:/ship-statics')

# 检查前端缺失项
with open('F:/ship-statics/templates/index.html', encoding='utf-8') as f:
    html = f.read()

issues = []
for tag in ['Step 6', 'runGZ(', 'exportGZ(', 'setGZStatus(', 'renderGZTable(', 'gzPanel', 'gzData']:
    if tag not in html:
        issues.append(f'前端缺少: {tag}')

# 检查 plotter_gz.py 的 import 是否能工作
try:
    from core.plotter import COLORS, A3_W, A3_H, _add_border, fig_to_base64
    issues.append('plotter_gz import: OK')
except Exception as e:
    issues.append(f'plotter_gz import FAIL: {e}')

# 检查 app.py 路由
with open('F:/ship-statics/app.py', encoding='utf-8') as f:
    app = f.read()
for route in ['/calc_gz', '/export_gz', '/preview_gz']:
    if route in app:
        issues.append(f'路由 {route}: OK')
    else:
        issues.append(f'路由 {route}: MISSING')

for i in issues:
    print(i)
