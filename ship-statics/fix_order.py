import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('F:/ship-statics/templates/index.html', encoding='utf-8') as f:
    content = f.read()

# Step5 块：从 "<!-- Step 5" 到 Step4 div 开始之前（不含换行）
step5_start = '    <!-- Step 5：曲线绘制'
# Step4 块：从 div border-color:#2e7dd1 到下一个 </div></div> 结束
# 找 Step4 块结束：在 Step4 开始后，找到连续两个 </div> 后的位置

idx_s5_start = content.find(step5_start)

# Step4 div 特征
step4_div = '    <div class="card-ship" style="border-color:#2e7dd1;'
idx_s4_start = content.find(step4_div)

# Step4 块结束：找到 Step4 div 对应的结束（需要匹配嵌套）
# Step4 之后应该是 </div>\n\n    <!-- 或者 </div>\n    <!-- ══ 等标记
# 简单方法：从 idx_s4_start 开始，找到下一个 '</div>\n    <!--' 或 '</div>\n\n'
search_end_start = idx_s4_start + 100

# 找 Step4 块结束标记
end_markers = [
    '\n    <!-- ══ 阶段3',  # 静水力表格面板开始
    '\n    <!-- ══',       # 任意阶段标记
]
idx_s4_end = -1
for m in end_markers:
    pos = content.find(m, search_end_start)
    if pos > 0:
        idx_s4_end = pos
        break

# 如果找不到标记，就用 Step4 块的最后一个 </div>
# 计算 div 嵌套层级
if idx_s4_end < 0:
    # 从 idx_s4_start 开始，计数 <div 和 </div，当平衡时结束
    depth = 0
    i = idx_s4_start
    while i < len(content):
        if content[i:i+5] == '<div ':
            depth += 1
        elif content[i:i+6] == '</div>':
            depth -= 1
            if depth == 0:
                idx_s4_end = i + 6
                break
        i += 1

print('Step5 start:', idx_s5_start)
print('Step4 start:', idx_s4_start)
print('Step4 end:', idx_s4_end)

# 提取 Step5 和 Step4 块
step5_block = content[idx_s5_start:idx_s4_start]
step4_block = content[idx_s4_start:idx_s4_end]

print('Step5 block len:', len(step5_block))
print('Step4 block len:', len(step4_block))

# 重新组装：Step4 在前，Step5 在后
new_content = (
    content[:idx_s5_start] +
    step4_block +
    '\n\n' +
    step5_block +
    content[idx_s4_end:]
)

# 写回文件
with open('F:/ship-statics/templates/index.html', 'w', encoding='utf-8') as f:
    f.write(new_content)

print('Order fixed: Step4 -> Step5')
