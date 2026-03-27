# ============================================================
# core/plotter_gz_enhanced.py — 增强的 GZ 曲线绘制模块
#
# 阶段6 第二优先级
#
# 功能：
#   1. 绘制高质量的静稳性曲线
#   2. 标注规范要求的关键特征点
#   3. 支持多种导出格式（PNG、SVG）
#   4. 符合课程设计要求
#
# ============================================================

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Circle
import io
import base64
from typing import Dict, Tuple, Optional

# ══════════════════════════════════════════════════════════
# 增强的 GZ 曲线绘制
# ══════════════════════════════════════════════════════════

class EnhancedGzPlotter:
    """增强的 GZ 曲线绘制类"""
    
    def __init__(self, figsize: Tuple[float, float] = (12, 8), dpi: int = 300):
        """
        初始化绘图器
        
        参数：
            figsize: 图幅大小（英寸）
            dpi: 分辨率（点/英寸）
        """
        self.figsize = figsize
        self.dpi = dpi
        
        # 设置中文字体
        plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False
    
    def plot_gz_curve_with_annotations(self, gz_data: Dict, 
                                      condition_name: str = '',
                                      theta_f: Optional[float] = None) -> Tuple:
        """
        绘制带标注的 GZ 曲线
        
        参数：
            gz_data: GZ 曲线数据
            condition_name: 工况名称
            theta_f: 进水角（可选）
        
        返回：
            (fig, ax) 图形对象
        """
        # 创建图形
        fig, ax = plt.subplots(figsize=self.figsize, dpi=self.dpi)
        
        # 提取数据
        rows = gz_data.get('rows', [])
        if not rows:
            return fig, ax
        
        thetas = np.array([r['theta'] for r in rows])
        gzs = np.array([r['GZ'] for r in rows])
        
        # 绘制 GZ 曲线
        ax.plot(thetas, gzs, 'b-', linewidth=2.5, label='GZ 曲线（精确值）', zorder=3)
        
        # 绘制网格
        ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)
        ax.set_axisbelow(True)
        
        # 标注关键点
        self._annotate_key_points(ax, gz_data, thetas, gzs, theta_f)
        
        # 标注规范要求的衡准线
        self._annotate_criteria_lines(ax, gz_data)
        
        # 设置坐标轴
        ax.set_xlabel('横倾角 θ (°)', fontsize=12, fontweight='bold')
        ax.set_ylabel('复原力臂 GZ (m)', fontsize=12, fontweight='bold')
        ax.set_title(f'静稳性曲线 — {condition_name}', fontsize=14, fontweight='bold', pad=20)
        
        # 设置坐标轴范围
        ax.set_xlim(0, 90)
        ax.set_ylim(min(0, np.min(gzs) - 0.05), np.max(gzs) + 0.1)
        
        # 添加图例
        ax.legend(loc='upper right', fontsize=10, framealpha=0.95)
        
        # 添加规范信息文本框
        self._add_criteria_textbox(ax, gz_data)
        
        # 调整布局
        fig.tight_layout()
        
        return fig, ax
    
    def _annotate_key_points(self, ax, gz_data: Dict, thetas: np.ndarray, 
                            gzs: np.ndarray, theta_f: Optional[float] = None):
        """标注关键特征点"""
        
        # 1. GZ_max 点
        gz_max = gz_data.get('GZ_max', 0)
        theta_max_gz = gz_data.get('theta_max_gz', 0)
        
        if gz_max > 0:
            ax.plot(theta_max_gz, gz_max, 'ro', markersize=10, zorder=5, label='GZ_max')
            ax.annotate(
                f'GZ_max = {gz_max:.3f}m\nθ = {theta_max_gz:.1f}°',
                xy=(theta_max_gz, gz_max),
                xytext=(theta_max_gz + 10, gz_max + 0.05),
                fontsize=10,
                bbox=dict(boxstyle='round,pad=0.5', facecolor='yellow', alpha=0.7),
                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0', color='red', lw=1.5)
            )
        
        # 2. 稳性消失角 θ_vanish
        theta_vanish = gz_data.get('theta_vanish', 0)
        
        if theta_vanish > 0 and theta_vanish <= 90:
            # 找到 θ_vanish 对应的 GZ 值（应该接近 0）
            idx_vanish = np.argmin(np.abs(thetas - theta_vanish))
            gz_vanish = gzs[idx_vanish]
            
            ax.plot(theta_vanish, gz_vanish, 'gs', markersize=10, zorder=5, label='θ_vanish')
            ax.annotate(
                f'θ_vanish = {theta_vanish:.1f}°',
                xy=(theta_vanish, gz_vanish),
                xytext=(theta_vanish - 15, gz_vanish + 0.05),
                fontsize=10,
                bbox=dict(boxstyle='round,pad=0.5', facecolor='lightgreen', alpha=0.7),
                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0', color='green', lw=1.5)
            )
        
        # 3. 进水角 θ_f（如果有）
        if theta_f is not None and theta_f > 0:
            ax.axvline(x=theta_f, color='purple', linestyle='--', linewidth=2, alpha=0.7, label=f'进水角 θ_f = {theta_f:.1f}°')
    
    def _annotate_criteria_lines(self, ax, gz_data: Dict):
        """标注规范要求的衡准线"""
        
        # GZ_max 最小值衡准线
        gz_max_min = 0.20
        ax.axhline(y=gz_max_min, color='orange', linestyle=':', linewidth=1.5, alpha=0.6)
        ax.text(2, gz_max_min + 0.01, f'GZ_max ≥ {gz_max_min}m（规范要求）', 
               fontsize=9, color='orange', fontweight='bold')
    
    def _add_criteria_textbox(self, ax, gz_data: Dict):
        """添加规范衡准信息文本框"""
        
        gm = gz_data.get('GM', 0)
        gz_max = gz_data.get('GZ_max', 0)
        theta_max_gz = gz_data.get('theta_max_gz', 0)
        theta_vanish = gz_data.get('theta_vanish', 0)
        k_value = gz_data.get('K', 0)
        
        # 判定结果
        criteria = gz_data.get('criteria', {})
        overall = criteria.get('overall', '未判定')
        
        textstr = f'''稳性指标：
GM = {gm:.4f}m
GZ_max = {gz_max:.4f}m
θ_max_gz = {theta_max_gz:.1f}°
θ_vanish = {theta_vanish:.1f}°
K = {k_value:.4f}

规范判定：{overall}'''
        
        props = dict(boxstyle='round', facecolor='wheat', alpha=0.8)
        ax.text(0.02, 0.98, textstr, transform=ax.transAxes, fontsize=9,
               verticalalignment='top', bbox=props, family='monospace')
    
    def plot_multiple_conditions(self, conditions_data: Dict, 
                                output_path: Optional[str] = None) -> Tuple:
        """
        绘制多工况对比图
        
        参数：
            conditions_data: {工况名: gz_data}
            output_path: 输出路径（可选）
        
        返回：
            (fig, axes) 图形对象
        """
        n_conditions = len(conditions_data)
        n_cols = 2
        n_rows = (n_conditions + 1) // 2
        
        fig, axes = plt.subplots(n_rows, n_cols, figsize=(16, 4*n_rows), dpi=self.dpi)
        
        if n_conditions == 1:
            axes = [axes]
        else:
            axes = axes.flatten()
        
        for idx, (condition_name, gz_data) in enumerate(conditions_data.items()):
            ax = axes[idx]
            
            rows = gz_data.get('rows', [])
            if not rows:
                continue
            
            thetas = np.array([r['theta'] for r in rows])
            gzs = np.array([r['GZ'] for r in rows])
            
            # 绘制曲线
            ax.plot(thetas, gzs, 'b-', linewidth=2.5)
            ax.grid(True, alpha=0.3)
            ax.set_xlabel('横倾角 θ (°)', fontsize=11)
            ax.set_ylabel('复原力臂 GZ (m)', fontsize=11)
            ax.set_title(f'{condition_name}', fontsize=12, fontweight='bold')
            ax.set_xlim(0, 90)
            
            # 标注关键点
            gz_max = gz_data.get('GZ_max', 0)
            theta_max_gz = gz_data.get('theta_max_gz', 0)
            
            if gz_max > 0:
                ax.plot(theta_max_gz, gz_max, 'ro', markersize=8)
                ax.annotate(f'GZ_max={gz_max:.3f}m', 
                           xy=(theta_max_gz, gz_max),
                           xytext=(theta_max_gz + 5, gz_max + 0.03),
                           fontsize=9)
        
        # 隐藏多余的子图
        for idx in range(n_conditions, len(axes)):
            axes[idx].set_visible(False)
        
        fig.tight_layout()
        
        if output_path:
            fig.savefig(output_path, dpi=self.dpi, bbox_inches='tight')
        
        return fig, axes
    
    def export_to_png(self, fig, output_path: str, dpi: int = 300):
        """导出为 PNG"""
        fig.savefig(output_path, dpi=dpi, bbox_inches='tight', format='png')
    
    def export_to_svg(self, fig, output_path: str):
        """导出为 SVG"""
        fig.savefig(output_path, bbox_inches='tight', format='svg')
    
    def export_to_base64(self, fig) -> str:
        """导出为 Base64（用于网页显示）"""
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        image_base64 = base64.b64encode(buf.read()).decode('utf-8')
        buf.close()
        return image_base64


# ══════════════════════════════════════════════════════════
# 测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    print('=' * 70)
    print('增强的 GZ 曲线绘制模块测试')
    print('=' * 70)
    
    # 模拟 GZ 数据
    thetas = np.arange(0, 91, 5)
    gzs = 0.42 * np.sin(np.radians(thetas))
    
    mock_gz_data = {
        'GM': 0.8234,
        'GZ_max': 0.42,
        'theta_max_gz': 30.0,
        'theta_vanish': 85.0,
        'K': 1.2,
        'rows': [
            {'theta': float(t), 'GZ': float(gz)}
            for t, gz in zip(thetas, gzs)
        ],
        'criteria': {'overall': '全部通过'}
    }
    
    # 创建绘图器
    plotter = EnhancedGzPlotter()
    
    # 绘制单工况
    print('\n【绘制单工况 GZ 曲线】')
    fig, ax = plotter.plot_gz_curve_with_annotations(mock_gz_data, '满载出港')
    print('✓ 单工况 GZ 曲线绘制完成')
    
    # 导出为 Base64
    print('\n【导出为 Base64】')
    b64 = plotter.export_to_base64(fig)
    print(f'✓ Base64 长度: {len(b64)} 字符')
    
    print('\n✓ 增强的 GZ 曲线绘制模块测试通过!')
