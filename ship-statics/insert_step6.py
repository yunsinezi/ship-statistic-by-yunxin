import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('F:/ship-statics/templates/index.html', encoding='utf-8') as f:
    content = f.read()

# Step6 HTML 块
step6_html = '''    <!-- Step 6：GZ 曲线计算（阶段5新增） -->
    <div class="card-ship" style="border-color:#059669; border-width:2px;">
        <div class="card-header-ship" style="background:linear-gradient(90deg,#f0fff4,#dcfce7);">
            <span class="step-num" style="background:#059669;">6</span>
            <span style="color:#065f46; font-weight:bold;">稳性横截曲线（阶段5）</span>
            <span class="ms-auto badge" style="background:#059669; font-size:0.72rem;">GZ 曲线</span>
        </div>
        <div class="card-body-ship">
            <div class="row g-2 mb-2">
                <div class="col-6">
                    <label class="form-label-ship" style="font-size:0.8rem;">重心高度 KG (m)</label>
                    <input type="number" id="inputKG" class="form-control form-control-sm form-control-ship"
                           placeholder="例：4.5" step="0.01" min="0">
                </div>
                <div class="col-6">
                    <label class="form-label-ship" style="font-size:0.8rem;">横倾角步长 (°)</label>
                    <input type="number" id="thetaStep" class="form-control form-control-sm form-control-ship"
                           value="5" step="1" min="1" max="15">
                </div>
            </div>
            <div class="row g-2 mb-2">
                <div class="col-6">
                    <label class="form-label-ship" style="font-size:0.8rem;">水密度 ρ (t/m³)</label>
                    <select id="rhoSelect" class="form-select form-select-sm form-control-ship">
                        <option value="1.025" selected>海水 1.025</option>
                        <option value="1.000">淡水 1.000</option>
                    </select>
                </div>
                <div class="col-6">
                    <label class="form-label-ship" style="font-size:0.8rem;">最大横倾角 (°)</label>
                    <input type="number" id="thetaMax" class="form-control form-control-sm form-control-ship"
                           value="90" step="5" min="30" max="90">
                </div>
            </div>
            <div class="mb-2">
                <label class="form-label-ship" style="font-size:0.8rem;">
                    多排水量吃水列表（可选，逗号分隔）
                </label>
                <input type="text" id="gzDrafts" class="form-control form-control-sm form-control-ship"
                       placeholder="例：3.0,4.0,5.0,6.0（留空则只算单吃水）">
            </div>
            <div class="d-flex flex-column gap-2">
                <button class="btn-calc w-100" onclick="runGZ()" id="gzBtn"
                        style="background:#059669;">
                    <i class="bi bi-activity me-2"></i>计算 GZ 曲线
                </button>
                <button class="btn-export w-100" onclick="exportGZ()" id="exportGZBtn"
                        style="background:#065f46; border-color:#065f46;" disabled>
                    <i class="bi bi-file-zip me-2"></i>导出 GZ 曲线图+数据表
                </button>
            </div>
            <div class="status-bar status-idle mt-2" id="gzStatus">
                <i class="bi bi-info-circle"></i>
                <span>请先输入 KG 值，然后点击计算</span>
            </div>
        </div>
    </div>

'''

# 插入到 Step5 注释之前
marker = '    <!-- Step 5：曲线绘制（阶段4新增） -->'
idx = content.find(marker)
if idx < 0:
    print('ERROR: marker not found')
else:
    new_content = content[:idx] + step6_html + content[idx:]
    with open('F:/ship-statics/templates/index.html', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('Step6 inserted before Step5, OK')
    print('New file size:', len(new_content))
