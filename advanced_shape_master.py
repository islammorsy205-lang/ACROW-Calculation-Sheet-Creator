# ==============================================================================
# ADVANCED SHAPE MASTER - ENGINEERING CALCULATION MODULE (Ultimate V27)
# ==============================================================================
# This module acts as the core mathematical and physical engine for the chain 
# builder and advanced shapes structural analysis. It includes DXF parsing,
# manual parametric generation, exact FEA matrix solving (including Deflection), 
# parallel strut shifting AI, Bridge & Single-Sided L-Frame optimizations,
# and Acrow-Standard height-based minimum strut configuration.
# ==============================================================================

import streamlit as st
import numpy as np
import pandas as pd
import io
import os
import re
import math
import tempfile
import time
import itertools
from collections import deque
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from PIL import Image, ImageChops
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import streamlit.components.v1 as components

# =========================================================
# System Check & Imports
# =========================================================
try:
    import ezdxf
except ImportError:
    st.error("⚠️ مكتبة 'ezdxf' غير موجودة! برجاء كتابة الأمر 'pip install ezdxf' في التيرمينال.")
    ezdxf = None

try:
    from config import SECTIONS_DB, STRUTS_DB
    from report_builder import insert_blue_banner, add_eq, append_pdf_stream_to_word
except ImportError:
    st.error("⚠️ برجاء التأكد من وجود ملفات config.py و report_builder.py في نفس المسار.")

# =========================================================
# 0. Helper Functions & Styles
# =========================================================
def apply_plot_styles():
    """
    تطبيق تنسيقات الرسم البياني الهندسية الدقيقة والخطوط
    """
    mpl.rcParams['font.family'] = 'sans-serif'
    mpl.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans']
    mpl.rcParams['axes.linewidth'] = 0.3
    mpl.rcParams['font.size'] = 7
    mpl.rcParams['font.weight'] = 'normal'

def get_short_name(sec_name):
    """
    استخراج الاسم القصير للقطاع وحذف التفاصيل بين الأقواس
    """
    cleaned_name = re.sub(
        r'\s*\(.*?\)', 
        '', 
        sec_name
    )
    return cleaned_name.strip()

def crop_image_bbox(img_bytes):
    """
    قص حواف الصور الناتجة لجعلها متناسقة في تقرير الوورد
    """
    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    bg = Image.new(
        "RGBA", 
        img.size, 
        (255, 255, 255, 0)
    )
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    
    if bbox:
        img = img.crop(bbox)
        
    out = io.BytesIO()
    img.save(
        out, 
        format='PNG'
    )
    return out.getvalue()

def safe_render_fig(fig, ax=None, nodes=None):
    """
    العدسة الذكية (Auto-Scaling): 
    تعتمد على margins بدلاً من set_ylim لتفادي الانهيار عند القيم الصغيرة جداً
    """
    try:
        if ax is not None:
            ax.axis('equal')
            ax.margins(0.15)

        plt.subplots_adjust(
            left=0.01, 
            right=0.99, 
            top=0.99, 
            bottom=0.01
        )
        
        buf = io.BytesIO()
        fig.savefig(
            buf, 
            format='png', 
            dpi=300, 
            bbox_inches='tight', 
            pad_inches=0.05, 
            transparent=True
        )
        return crop_image_bbox(buf.getvalue())
    finally:
        plt.close(fig)

def draw_reaction_arrow(ax, node_x, node_y, force_mag, axis_nx, axis_ny, is_bridge_mode=False):
    """
    رسم أسهم ردود الأفعال بدقة مع تمييز قوة الشد.
    """
    if abs(force_mag) < 0.001:
        return
        
    if is_bridge_mode and abs(axis_nx) > 0.5:
        return
        
    arr_L = 0.5  
    sgn = np.sign(force_mag)
    
    dx = sgn * axis_nx
    dy = sgn * axis_ny
    
    start_x = node_x - arr_L * dx
    start_y = node_y - arr_L * dy
    
    if force_mag >= 0:
        arr_c = 'blue'
    else:
        arr_c = 'red'
    
    ax.arrow(
        start_x, 
        start_y, 
        arr_L * dx, 
        arr_L * dy, 
        length_includes_head=True, 
        head_width=0.08, 
        head_length=0.12, 
        fc=arr_c, 
        ec=arr_c, 
        lw=0.8, 
        zorder=5
    )
    
    ax.text(
        start_x - 0.15 * dx, 
        start_y - 0.15 * dy, 
        f"{force_mag:+.2f}", 
        color=arr_c, 
        fontsize=7, 
        fontname='Arial', 
        ha='center', 
        va='center'
    )

def eval_seg_point(seg, s_val, start_data=None):
    """
    حساب الإحداثيات المطلقة X,Y لأي نقطة على أي فريم بناء على المسافة (s)
    """
    if seg.get('is_divided'):
        actual_s = s_val + seg.get('parent_offset', 0.0)
        return eval_seg_point(
            seg['parent_seg'], 
            actual_s, 
            start_data
        )

    L = seg.get('L', 0.0)
    s_val = min(max(s_val, 0.0), L)
    
    if L > 1e-6:
        ratio = s_val / L 
    else:
        ratio = 0.0
        
    is_dxf = seg.get('is_dxf', False)
    shape_type = seg.get('Shape Type', 'Straight Line')
    
    if is_dxf:
        if shape_type == 'Straight Line' and 'abs_p1' in seg:
            p1 = seg['abs_p1']
            p2 = seg['abs_p2']
            
            px = p1[0] + ratio * (p2[0] - p1[0])
            py = p1[1] + ratio * (p2[1] - p1[1])
            th = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
            
            return px, py, th
            
        elif shape_type == 'Curve (Arc & Radius)' and 'abs_c' in seg:
            c = seg['abs_c']
            r = seg['abs_r']
            
            current_ang = seg['abs_sa'] + ratio * seg.get('sweep', 0.0)
            px = c[0] + r * math.cos(current_ang)
            py = c[1] + r * math.sin(current_ang)
            th = current_ang + math.pi / 2.0
            
            return px, py, th
            
    if start_data:
        x0 = start_data.get('x0', 0.0)
        y0 = start_data.get('y0', 0.0)
        th0 = start_data.get('th0', 0.0)
        kappa = start_data.get('kappa', 0.0)
        
        if abs(kappa) < 1e-6: 
            x = x0 + s_val * math.cos(th0)
            y = y0 + s_val * math.sin(th0)
            th = th0
        else: 
            x = x0 + (math.sin(th0 + kappa * s_val) - math.sin(th0)) / kappa
            y = y0 - (math.cos(th0 + kappa * s_val) - math.cos(th0)) / kappa
            th = th0 + kappa * s_val
            
        return x, y, th
    
    return 0.0, 0.0, 0.0

def get_closest_segment_exact(pt, segs):
    """
    إيجاد أقرب فريم لأي نقطة مرسومة للالتحام المغناطيسي
    """
    min_d = 9999.0
    best_idx = 0
    best_s = 0.0
    pt = np.array(pt)
    
    for idx, seg in enumerate(segs):
        if seg.get('is_divided'):
            temp_seg = seg['parent_seg']
            L_orig = temp_seg.get('L', 0.0)
        else:
            temp_seg = seg
            L_orig = temp_seg.get('L', 0.0)
            
        shape_type = temp_seg.get('Shape Type')
        
        if shape_type == 'Straight Line' and 'abs_p1' in temp_seg:
            p1 = np.array(temp_seg['abs_p1'])
            p2 = np.array(temp_seg['abs_p2'])
            
            v = p2 - p1
            w = pt - p1
            c2 = np.dot(v, v)
            
            if c2 > 1e-6:
                ratio = np.dot(w, v) / c2 
            else:
                ratio = 0.0
                
            ratio = max(0.0, min(1.0, ratio))
            proj = p1 + ratio * v
            d = np.linalg.norm(pt - proj)
            
            if d < min_d:
                min_d = d
                best_idx = idx
                best_s = ratio * L_orig
                
                if seg.get('is_divided'):
                    best_s -= seg.get('parent_offset', 0.0)
                
        elif shape_type == 'Curve (Arc & Radius)' and 'abs_c' in temp_seg:
            c = np.array(temp_seg['abs_c'])
            r = temp_seg['abs_r']
            
            v = pt - c
            ang = math.atan2(v[1], v[0])
            
            sa = temp_seg['abs_sa']
            sweep = temp_seg['sweep']
            
            ang_norm = (ang - sa) % (2.0 * math.pi)
            
            if ang_norm > abs(sweep):
                if ang_norm < math.pi:
                    ratio = 1.0 
                else:
                    ratio = 0.0
            else:
                if abs(sweep) > 1e-6:
                    ratio = ang_norm / sweep 
                else:
                    ratio = 0.0
                
            ratio = max(0.0, min(1.0, ratio))
            current_ang = sa + ratio * sweep
            proj = c + r * np.array([math.cos(current_ang), math.sin(current_ang)])
            d = np.linalg.norm(pt - proj)
            
            if d < min_d:
                min_d = d
                best_idx = idx
                best_s = ratio * L_orig
                
                if seg.get('is_divided'):
                    best_s -= seg.get('parent_offset', 0.0)
                
    return min_d, best_idx, best_s

def get_shifted_coords_along_segment(px, py, ds, segs):
    """
    الخوارزمية الجراحية للزحزحة الموازية (Parallel Micro-Shifting)
    """
    if abs(ds) < 1e-4: 
        return px, py
        
    d_min, best_idx, best_s = get_closest_segment_exact((px, py), segs)
    
    if d_min > 0.5: 
        return px + ds, py
        
    seg = segs[best_idx]
    new_s = best_s + ds
    new_s = max(0.0, min(new_s, seg.get('L', 0.0)))
    
    nx, ny, _ = eval_seg_point(seg, new_s)
    return nx, ny

def get_strut_priority(name):
    """
    تحديد أولوية القطاعات الخاصة بالنهايز. 
    💡 تم تفعيل الحظر لجميع المقاسات التي تنتهي بـ 1 أو 3 لتجاهلها تماماً.
    """
    name_u = name.upper()
    score = 100
    
    base_name = name.split('(')[0].strip()
    if bool(re.search(r'(1|3)$', base_name)):
        return 999  
    
    if "PPS" in name_u: 
        score = 10
    elif "PPH" in name_u: 
        score = 20
    elif "TILT" in name_u: 
        return 999
    elif "MMP" in name_u: 
        return 999 
    elif "MNB" in name_u: 
        score = 95
    elif "MIB" in name_u: 
        score = 99
        
    return score

def get_optimal_strut_section(req_length, req_axial_force, system_type="Bridge"):
    """
    الدالة الذكية للترقية، تلتزم بالفلتر أعلاه ولن تختار نحايز فردية أبداً
    """
    valid_struts = []
    
    for s_name, s_props in STRUTS_DB.items():
        if "TILT" in s_name.upper() or "MMP" in s_name.upper(): 
            continue
            
        priority = get_strut_priority(s_name)
        if priority == 999:  
            continue
            
        m = re.search(r'\((\d+\.\d+):(\d+\.\d+)m\)', s_name)
        if m:
            min_L = float(m.group(1))
            max_L = float(m.group(2))
            
            if min_L <= req_length <= max_L:
                allowable = s_props.get('allow', 0.0)
                if allowable >= abs(req_axial_force):
                    valid_struts.append({
                        'name': s_name, 
                        'allowable': allowable, 
                        'priority': priority
                    })
                    
    if not valid_struts: 
        return None
        
    valid_struts.sort(key=lambda x: (x['priority'], x['allowable'])) 
    return valid_struts[0]['name']

def extract_dxf_for_interactive(file_bytes):
    if ezdxf is None: 
        return None
        
    tmp_path = ""
    try:
        try: 
            dxf_str = file_bytes.decode('utf-8')
        except UnicodeDecodeError: 
            dxf_str = file_bytes.decode('cp1252', errors='ignore')
            
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf", mode='w', encoding='utf-8') as tmp:
            tmp.write(dxf_str)
            tmp_path = tmp.name
            
        doc = ezdxf.readfile(tmp_path)
        msp = doc.modelspace()
        
        frames = []
        struts = []
        supports = []
        
        for e in msp:
            lyr = e.dxf.layer.lower()
            dxftype = e.dxftype()
            
            if dxftype in ['POINT', 'CIRCLE', 'INSERT']:
                if dxftype == 'POINT': 
                    supports.append({
                        'x': e.dxf.location.x, 
                        'y': e.dxf.location.y
                    })
                elif dxftype == 'CIRCLE': 
                    supports.append({
                        'x': e.dxf.center.x, 
                        'y': e.dxf.center.y
                    })
                elif dxftype == 'INSERT': 
                    supports.append({
                        'x': e.dxf.insert.x, 
                        'y': e.dxf.insert.y
                    })
                continue 
                
            if dxftype in ['LWPOLYLINE', 'POLYLINE']:
                entities = list(e.virtual_entities())
            else:
                entities = [e]
                
            for sub_e in entities:
                sub_type = sub_e.dxftype()
                
                if "push" in lyr or "pull" in lyr:
                    if sub_type == 'LINE':
                        struts.append({
                            'p1': [sub_e.dxf.start.x, sub_e.dxf.start.y], 
                            'p2': [sub_e.dxf.end.x, sub_e.dxf.end.y]
                        })
                else:
                    if sub_type == 'LINE':
                        frames.append({
                            'type': 'line', 
                            'x1': sub_e.dxf.start.x, 
                            'y1': sub_e.dxf.start.y, 
                            'x2': sub_e.dxf.end.x, 
                            'y2': sub_e.dxf.end.y
                        })
                    elif sub_type == 'ARC':
                        frames.append({
                            'type': 'arc', 
                            'c': [sub_e.dxf.center.x, sub_e.dxf.center.y], 
                            'r': sub_e.dxf.radius, 
                            'sa': math.radians(sub_e.dxf.start_angle), 
                            'ea': math.radians(sub_e.dxf.end_angle)
                        })
                        
        return {
            'frames': frames, 
            'struts': struts, 
            'supports': supports
        }
        
    except Exception as e: 
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: 
                os.remove(tmp_path)
            except: 
                pass

def parse_dxf_to_data(file_bytes):
    raw_data = extract_dxf_for_interactive(file_bytes)
    
    if not raw_data or not raw_data['frames']: 
        return None
        
    raw_frames = raw_data.get('frames', [])
    raw_struts = raw_data.get('struts', [])
    raw_supports = raw_data.get('supports', [])

    def get_min_x(f):
        if f['type'] == 'line': 
            return min(f['x1'], f['x2'])
        return f['c'][0] - f['r']
        
    raw_frames.sort(key=get_min_x)

    base_segments = []
    for idx, f in enumerate(raw_frames):
        if f['type'] == 'line':
            p_start = (f['x1'], f['y1'])
            p_end = (f['x2'], f['y2'])
            
            if p_start[0] > p_end[0] + 1e-5 or (abs(p_start[0] - p_end[0]) < 1e-5 and p_start[1] > p_end[1]):
                p_start, p_end = p_end, p_start
                
            dx_line = p_end[0] - p_start[0]
            dy_line = p_end[1] - p_start[1]
            
            L = math.hypot(dx_line, dy_line)
            ang = math.degrees(math.atan2(dy_line, dx_line))
            
            base_segments.append({
                'name': f"S{idx+1}", 
                'master_idx': idx, 
                'type': 'Straight Line', 
                'Shape Type': 'Straight Line', 
                'L': L, 
                'start_angle': ang, 
                'smooth': False, 
                'is_dxf': True, 
                'abs_p1': p_start, 
                'abs_p2': p_end, 
                'kappa': 0.0, 
                'is_divided': False
            })
            
        elif f['type'] == 'arc':
            sa = f['sa']
            ea = f['ea']
            
            if ea < sa: 
                ea += 2.0 * math.pi
                
            sweep = ea - sa
            L = f['r'] * (ea - sa)
            
            base_segments.append({
                'name': f"S{idx+1}", 
                'master_idx': idx, 
                'type': 'Curve (Arc & Radius)', 
                'Shape Type': 'Curve (Arc & Radius)', 
                'L': L, 
                'Radius (R) (m)': f['r'], 
                'Curvature Direction': "Arching Up ⤴ (Concave)",
                'start_angle': math.degrees(sa + math.pi/2.0), 
                'smooth': False, 
                'is_dxf': True, 
                'abs_c': list(f['c']), 
                'abs_r': f['r'], 
                'abs_sa': sa, 
                'abs_ea': ea, 
                'sweep': sweep, 
                'kappa': 1.0 / f['r'], 
                'is_divided': False
            })

    struts_mapped = []
    for s in raw_struts:
        p1 = s['p1']
        p2 = s['p2']
        
        if p1[1] > p2[1]:
            top_p = p1
            bot_p = p2
        else:
            top_p = p2
            bot_p = p1
            
        struts_mapped.append({
            'tx': top_p[0], 
            'ty': top_p[1], 
            'bx': bot_p[0], 
            'by': bot_p[1], 
            'sec': list(STRUTS_DB.keys())[0] if STRUTS_DB else "PPH 353 (1.5:3.5m)"
        })

    supps_mapped = []
    for sp in raw_supports:
        supps_mapped.append({
            'x': sp['x'], 
            'y': sp['y'], 
            'type': 'Hinged', 
            'angle': 0.0
        })

    return {
        'base_segments': base_segments, 
        'struts': struts_mapped, 
        'supports': supps_mapped
    }
# =========================================================
# 2. Smart Division & Dynamic Meshing Engine
# =========================================================
def get_approx_xy(segs, s_idx, s_val):
    if s_idx < 0 or s_idx >= len(segs): 
        return 0.0, 0.0
        
    seg = segs[s_idx]
    
    if seg.get('is_dxf'):
        px, py, _ = eval_seg_point(seg, s_val)
        return px, py
        
    return 0.0, 0.0

def perform_smart_division(base_segments, supports, struts):
    """
    التقطيع البصري: يتم فقط لتجميل الرسم وتوقيع الأحمال
    """
    cut_points_dict = {}
    
    for i, seg in enumerate(base_segments):
        cut_points_dict[i] = {0.0, seg['L']}
    
    for sp in supports:
        d_min, w_seg, w_s = get_closest_segment_exact((sp['x'], sp['y']), base_segments)
        if d_min < 0.30:
            cut_points_dict[w_seg].add(min(max(w_s, 0.0), base_segments[w_seg]['L']))
            
    for st in struts:
        dt, wt_seg, wt_s = get_closest_segment_exact((st['tx'], st['ty']), base_segments)
        if dt < 0.30: 
            cut_points_dict[wt_seg].add(min(max(wt_s, 0.0), base_segments[wt_seg]['L']))
        
        db, wb_seg, wb_s = get_closest_segment_exact((st['bx'], st['by']), base_segments)
        if db < 0.30: 
            cut_points_dict[wb_seg].add(min(max(wb_s, 0.0), base_segments[wb_seg]['L']))

    divided_segments = []
    sub_letters = "abcdefghijklmnopqrstuvwxyz"
    
    for m_idx, s_vals_set in sorted(cut_points_dict.items()):
        master_seg = base_segments[m_idx]
        sorted_s = sorted(list(s_vals_set))
        num_sub = len(sorted_s) - 1
        
        for k in range(num_sub):
            s_start = sorted_s[k]
            s_end = sorted_s[k+1]
            
            if s_end - s_start < 1e-4: 
                continue
            
            if num_sub == 1:
                sub_name = master_seg['name']
            else:
                sub_name = f"{master_seg['name']}-{sub_letters[k % 26]}"
                
            new_seg = master_seg.copy()
            new_seg.update({
                'name': sub_name,
                'is_divided': True,
                'parent_seg': master_seg,
                'parent_offset': s_start,
                'L': s_end - s_start,
                'master_idx': m_idx
            })
            divided_segments.append(new_seg)
            
    return divided_segments

def build_chain_mesh(segments, seg_sections, loads, struts, supports, base_segments, mesh_size=0.25):
    """
    بناء مصفوفة الـ Finite Element مع اللحام المغناطيسي المطور
    """
    nodes = []
    elements = []
    nodal_loads = []
    
    node_tol = 0.05 
    
    def get_or_add_node(x, y):
        for i, n in enumerate(nodes):
            if abs(n[0] - x) < node_tol and abs(n[1] - y) < node_tol:
                return i
        nodes.append([x, y])
        return len(nodes) - 1

    key_nodes = set()
    
    support_injections = {}
    strut_injections = {}
    for i in range(len(base_segments)):
        support_injections[i] = []
        strut_injections[i] = []
        
    supports_list_out = []
    
    for sup in supports:
        sx = sup['x']
        sy = sup['y']
        
        d_min, w_seg, w_s = get_closest_segment_exact((sx, sy), base_segments)
        
        if d_min < 0.30:
            nx, ny, _ = eval_seg_point(base_segments[w_seg], w_s)
            support_injections[w_seg].append(w_s)
        else:
            nx = sx
            ny = sy
            
        nid = get_or_add_node(nx, ny)
        supports_list_out.append({
            'node': nid, 
            'type': sup.get('type', 'Hinged'), 
            'angle': sup.get('angle', 0.0)
        })

    for st_item in struts:
        dt, wt_seg, wt_s = get_closest_segment_exact((st_item['tx'], st_item['ty']), base_segments)
        if dt < 0.30:
            strut_injections[wt_seg].append(wt_s)
            st_item['tx'], st_item['ty'], _ = eval_seg_point(base_segments[wt_seg], wt_s)
            
        db, wb_seg, wb_s = get_closest_segment_exact((st_item['bx'], st_item['by']), base_segments)
        if db < 0.30:
            strut_injections[wb_seg].append(wb_s)
            st_item['bx'], st_item['by'], _ = eval_seg_point(base_segments[wb_seg], wb_s)

    for i, seg in enumerate(segments):
        L = seg.get('L', 0.0)
        key_s_vals = [0.0, L]
        seg_m_idx = seg.get('master_idx')
        parent_off = seg.get('parent_offset', 0.0)

        for s_val in strut_injections.get(seg_m_idx, []):
            s_local = s_val - parent_off
            if -1e-4 <= s_local <= L + 1e-4:
                key_s_vals.append(max(0.0, min(L, s_local)))
                
        for s_val in support_injections.get(seg_m_idx, []):
            s_local = s_val - parent_off
            if -1e-4 <= s_local <= L + 1e-4:
                key_s_vals.append(max(0.0, min(L, s_local)))

        for ld in loads:
            if ld.get('seg_idx') == i:
                key_s_vals.extend([ld['start'], ld['end']])
                
        num_sub = max(1, int(np.ceil(L / mesh_size)))
        for p in np.linspace(0, L, num_sub+1): 
            key_s_vals.append(p)
            
        keys = sorted(list(set([min(max(round(k, 4), 0.0), round(L, 4)) for k in key_s_vals])))
        node_indices = []
        
        for s_val in keys:
            px, py, _ = eval_seg_point(seg, s_val)
            nid = get_or_add_node(px, py)
            node_indices.append(nid)
            
            for kv in key_s_vals:
                if abs(s_val - round(kv, 4)) < 1e-3:
                    key_nodes.add(nid)
                    break
            
        if i < len(seg_sections):
            sec_props = seg_sections[i]
        else:
            sec_props = seg_sections[0]
        
        for j in range(len(keys)-1):
            n1 = node_indices[j]
            n2 = node_indices[j+1]
            if n1 == n2: 
                continue 
                
            s_mid = (keys[j] + keys[j+1]) / 2.0
            _, _, th_mid = eval_seg_point(seg, s_mid)
            
            c_t = np.cos(th_mid)
            s_t = np.sin(th_mid)
            
            p_x1 = 0.0
            p_y1 = 0.0
            p_x2 = 0.0
            p_y2 = 0.0
            
            for ld in loads:
                if ld.get('seg_idx') == i and ld.get('type') != 'Point Load':
                    if ld['start'] - 1e-4 <= s_mid <= ld['end'] + 1e-4:
                        L_ld = max(ld['end'] - ld['start'], 1e-5)
                        wa = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j] - ld['start']) / L_ld
                        wb = ld['w1'] + (ld['w2'] - ld['w1']) * (keys[j+1] - ld['start']) / L_ld
                        
                        dir_str = ld.get('dir', '')
                        if 'Global Z' in dir_str or 'Global Y' in dir_str:
                            p_x1 += wa * s_t
                            p_y1 += wa * c_t
                            p_x2 += wb * s_t
                            p_y2 += wb * c_t
                        elif 'Global X' in dir_str:
                            p_x1 += wa * c_t
                            p_y1 -= wa * s_t
                            p_x2 += wb * c_t
                            p_y2 -= wb * s_t
                        else:
                            p_y1 += wa
                            p_y2 += wb
                            
            elements.append({
                'type': 'frame', 
                'group': 'segment', 
                'sec': sec_props['name'],
                'n1': n1, 
                'n2': n2, 
                'px1': p_x1, 
                'py1': p_y1, 
                'px2': p_x2, 
                'py2': p_y2,
                'E': sec_props['E'] * 10000.0, 
                'A': sec_props['A'], 
                'I': sec_props['I'] / 100000000.0,
                'seg_idx': i, 
                's_start': keys[j], 
                's_end': keys[j+1], 
                'L': keys[j+1] - keys[j], 
                'th_mid': th_mid
            })
            
        for ld in loads:
            if ld.get('seg_idx') == i and ld.get('type') == 'Point Load':
                px, py, th_pt = eval_seg_point(seg, ld['start'])
                nid = get_or_add_node(px, py)
                dir_str = ld.get('dir', '')
                
                if 'Global Z' in dir_str or 'Global Y' in dir_str: 
                    nodal_loads.append({'node': nid, 'Fx': 0.0, 'Fy': ld['w1']})
                elif 'Global X' in dir_str: 
                    nodal_loads.append({'node': nid, 'Fx': ld['w1'], 'Fy': 0.0})
                else: 
                    c_pt = np.cos(th_pt)
                    s_pt = np.sin(th_pt)
                    nodal_loads.append({'node': nid, 'Fx': -ld['w1']*s_pt, 'Fy': ld['w1']*c_pt})

    for st_idx, st_item in enumerate(struts):
        top_node = get_or_add_node(st_item['tx'], st_item['ty'])
        bot_node = get_or_add_node(st_item['bx'], st_item['by'])
        
        elements.append({
            'type': 'truss', 
            'group': 'strut', 
            'sec': st_item.get('sec', 'Unknown'),
            'n1': bot_node, 
            'n2': top_node, 
            'strut_idx': st_idx, 
            'E': 21000000.0, 
            'A': 0.001
        })

    display_nodes = set()
    for s in supports_list_out:
        display_nodes.add(s['node'])
        
    for k in key_nodes:
        display_nodes.add(k)

    return nodes, elements, nodal_loads, display_nodes, supports_list_out
# =========================================================
# 3. Advanced FEA Solver (Exact Matrix Engine)
# =========================================================
def solve_fea_engine(nodes, elements, nodal_loads, supports_list):
    """
    محرك الـ FEA المتطور:
    يحل العزوم، الشير، الأكسيال، ويحسب الترخيم الموضعي (Local Deflection) لغرض العرض.
    """
    NDOF = len(nodes) * 3
    K = np.zeros((NDOF, NDOF))
    F = np.zeros(NDOF)
    
    for el in elements:
        n1 = el['n1']
        n2 = el['n2']
        
        x1 = nodes[n1][0]
        y1 = nodes[n1][1]
        x2 = nodes[n2][0]
        y2 = nodes[n2][1]
        
        L = np.hypot(x2 - x1, y2 - y1)
        
        if L < 1e-5: 
            el['c'] = 1.0
            el['s'] = 0.0
            el['L'] = 1e-5
            el['internal'] = {
                'N': [0.0, 0.0], 
                'V': [0.0, 0.0], 
                'M': [0.0, 0.0], 
                'D': [0.0, 0.0],
                'x': [0.0, 1e-5]
            }
            continue
            
        c = (x2 - x1) / L
        s = (y2 - y1) / L
        
        el['L'] = L
        el['c'] = c
        el['s'] = s
        
        E = el['E']
        A = el['A']
        I = el.get('I', 0.00005)
        
        T = np.array([
            [c, s, 0, 0, 0, 0], 
            [-s, c, 0, 0, 0, 0], 
            [0, 0, 1, 0, 0, 0],
            [0, 0, 0, c, s, 0], 
            [0, 0, 0, -s, c, 0], 
            [0, 0, 0, 0, 0, 1]
        ])
        
        if el['type'] == 'truss':
            k_loc = np.zeros((6, 6))
            k_loc[0, 0] = E * A / L
            k_loc[3, 3] = E * A / L
            k_loc[0, 3] = -E * A / L
            k_loc[3, 0] = -E * A / L
        else:
            k_loc = np.array([
                [E*A/L, 0, 0, -E*A/L, 0, 0], 
                [0, 12*E*I/L**3, 6*E*I/L**2, 0, -12*E*I/L**3, 6*E*I/L**2],
                [0, 6*E*I/L**2, 4*E*I/L, 0, -6*E*I/L**2, 2*E*I/L], 
                [-E*A/L, 0, 0, E*A/L, 0, 0],
                [0, -12*E*I/L**3, -6*E*I/L**2, 0, 12*E*I/L**3, -6*E*I/L**2], 
                [0, 6*E*I/L**2, 2*E*I/L, 0, -6*E*I/L**2, 4*E*I/L]
            ])
            
            px1 = el.get('px1', 0.0)
            py1 = el.get('py1', 0.0)
            px2 = el.get('px2', 0.0)
            py2 = el.get('py2', 0.0)
            
            f_loc = np.array([
                (2.0*px1 + px2)*L/6.0, 
                (7.0*py1 + 3.0*py2)*L/20.0, 
                (3.0*py1 + 2.0*py2)*L**2/60.0,
                (px1 + 2.0*px2)*L/6.0, 
                (3.0*py1 + 7.0*py2)*L/20.0, 
                -(2.0*py1 + 3.0*py2)*L**2/60.0
            ])
            f_glob = T.T @ f_loc
            
            dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
            for r in range(6): 
                F[dof[r]] += f_glob[r]
            
        k_glob = T.T @ k_loc @ T
        dof = [3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]
        
        for r in range(6):
            for col in range(6): 
                K[dof[r], dof[col]] += k_glob[r, col]
                
    for nl in nodal_loads:
        F[3*nl['node']] += nl['Fx']
        F[3*nl['node']+1] += nl['Fy']
            
    # حساب إجمالي الحمل الرأسي للتأكد من الاتزان
    net_load_y = abs(np.sum(F[1::3]))

    K_orig = K.copy()
    fixed_dofs = []
    K_pen = 1e12
    
    for sup in supports_list:
        n = sup['node']
        t = sup['type']
        a = sup.get('angle', 0.0)
        
        if t == 'Fixed': 
            fixed_dofs.extend([3*n, 3*n+1, 3*n+2])
        elif t == 'Hinged': 
            fixed_dofs.extend([3*n, 3*n+1])
        elif t == 'Roller':
            rad = np.radians(a)
            nx = -np.sin(rad)
            ny = np.cos(rad) 
            
            K[3*n, 3*n] += K_pen * nx**2
            K[3*n+1, 3*n+1] += K_pen * ny**2
            K[3*n, 3*n+1] += K_pen * nx * ny
            K[3*n+1, 3*n] += K_pen * nx * ny

    free_dof = []
    for i in range(NDOF):
        if i not in fixed_dofs:
            free_dof.append(i)
            
    K_ff = K[np.ix_(free_dof, free_dof)]
    F_f = F[free_dof]
    
    U = np.zeros(NDOF)
    try: 
        U[free_dof] = np.linalg.solve(K_ff, F_f)
    except np.linalg.LinAlgError: 
        U[free_dof] = np.linalg.lstsq(K_ff, F_f, rcond=None)[0]
    
    R_reactions = K_orig @ U - F 
    
    for el in elements:
        if el.get('L', 0) < 1e-5: 
            continue
            
        n1 = el['n1']
        n2 = el['n2']
        c = el['c']
        s = el['s']
        L = el['L']
        E = el['E']
        A = el['A']
        I = el.get('I', 0.00005)
        
        u_glob = U[[3*n1, 3*n1+1, 3*n1+2, 3*n2, 3*n2+1, 3*n2+2]]
        T = np.array([
            [c, s, 0, 0, 0, 0], 
            [-s, c, 0, 0, 0, 0], 
            [0, 0, 1, 0, 0, 0],
            [0, 0, 0, c, s, 0], 
            [0, 0, 0, -s, c, 0], 
            [0, 0, 0, 0, 0, 1]
        ])
        
        u_loc = T @ u_glob
        el['internal'] = {'u_loc': u_loc}
        
        if el['type'] == 'truss':
            N_val = (E * A / L) * (u_loc[3] - u_loc[0])
            xs = np.linspace(0, L, 51)
            el['internal'].update({
                'N': np.full_like(xs, N_val), 
                'V': np.zeros_like(xs), 
                'M': np.zeros_like(xs),
                'D': np.zeros_like(xs),
                'x': xs
            })
        else:
            k_loc = np.array([
                [E*A/L, 0, 0, -E*A/L, 0, 0], 
                [0, 12*E*I/L**3, 6*E*I/L**2, 0, -12*E*I/L**3, 6*E*I/L**2],
                [0, 6*E*I/L**2, 4*E*I/L, 0, -6*E*I/L**2, 2*E*I/L], 
                [-E*A/L, 0, 0, E*A/L, 0, 0],
                [0, -12*E*I/L**3, -6*E*I/L**2, 0, 12*E*I/L**3, -6*E*I/L**2], 
                [0, 6*E*I/L**2, 2*E*I/L, 0, -6*E*I/L**2, 4*E*I/L]
            ])
            
            px1 = el.get('px1', 0.0)
            py1 = el.get('py1', 0.0)
            px2 = el.get('px2', 0.0)
            py2 = el.get('py2', 0.0)
            
            f_loc = np.array([
                (2.0*px1 + px2)*L/6.0, 
                (7.0*py1 + 3.0*py2)*L/20.0, 
                (3.0*py1 + 2.0*py2)*L**2/60.0,
                (px1 + 2.0*px2)*L/6.0, 
                (3.0*py1 + 7.0*py2)*L/20.0, 
                -(2.0*py1 + 3.0*py2)*L**2/60.0
            ])
            
            f_end = k_loc @ u_loc - f_loc
            xs = np.linspace(0, L, 51) 
            
            N_arr = np.zeros_like(xs)
            V_arr = np.zeros_like(xs)
            M_arr = np.zeros_like(xs)
            D_arr = np.zeros_like(xs)
            
            v1 = u_loc[1]
            th1 = u_loc[2]
            v2 = u_loc[4]
            th2 = u_loc[5]
            
            for i, x in enumerate(xs):
                N_arr[i] = -f_end[0] - (px1*x + (px2-px1)*x**2/(2*L))
                V_arr[i] = f_end[1] + (py1*x + (py2-py1)*x**2/(2*L))
                M_arr[i] = -f_end[2] + f_end[1]*x + py1*x**2/2.0 + (py2-py1)*x**3/(6*L)
                
                # معادلة الترخيم (Deflection) باستخدام Shape Functions للعرض فقط
                xi = x / L if L > 0 else 0
                N1_shp = 1.0 - 3.0*xi**2 + 2.0*xi**3
                N2_shp = L * (xi - 2.0*xi**2 + xi**3)
                N3_shp = 3.0*xi**2 - 2.0*xi**3
                N4_shp = L * (-xi**2 + xi**3)
                
                v_x = v1*N1_shp + th1*N2_shp + v2*N3_shp + th2*N4_shp
                w_avg = (py1 + py2) / 2.0
                v_load = (w_avg * x**2 * (L - x)**2) / (24.0 * E * I) if (E * I) != 0 else 0
                
                D_arr[i] = (v_x + v_load) * 1000.0 # بالمليمتر
                
            el['internal'].update({
                'N': N_arr, 
                'V': V_arr, 
                'M': M_arr, 
                'D': D_arr,
                'x': xs
            })
            
    return U, R_reactions, net_load_y
# =========================================================
# 4. Plotting Engine & Word Report Generator
# =========================================================
def draw_base_geometry(ax, nodes, elements, supports_list, seg_sections=None, segments=None, show_seg_names=False, ss_spacings=None):
    for el in elements:
        if el['type'] not in ['frame', 'truss']: 
            continue
            
        n1 = nodes[el['n1']]
        n2 = nodes[el['n2']]
        
        if el['type'] == 'truss':
            ax.plot([n1[0], n2[0]], [n1[1], n2[1]], color='red', linestyle='-', linewidth=0.8, zorder=1)
        else:
            if el.get('group') == 'base' and el.get('sec') == "None (Direct to Ground)": 
                continue
            ax.plot([n1[0], n2[0]], [n1[1], n2[1]], color='royalblue', linestyle='-', linewidth=1.5, zorder=1)
            
    for i, sup in enumerate(supports_list):
        n = sup['node']
        t = sup['type']
        ang_deg = sup.get('angle', 0.0)
        x = nodes[n][0]
        y = nodes[n][1]
        
        ax.text(x, y - 0.4, f"J{i+1}", color='green', fontsize=7, ha='center', fontname='Arial')
        ang_rad = math.radians(ang_deg)
        c_a = math.cos(ang_rad)
        s_a = math.sin(ang_rad)
        
        def rot_pt(px, py): 
            rx = x + (px - x)*c_a - (py - y)*s_a
            ry = y + (px - x)*s_a + (py - y)*c_a
            return rx, ry

        if t == 'Fixed':
            p1, p2, p3, p4 = rot_pt(x-0.1, y-0.1), rot_pt(x+0.1, y-0.1), rot_pt(x+0.1, y+0.1), rot_pt(x-0.1, y+0.1)
            ax.add_patch(Polygon([p1, p2, p3, p4], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            l1, l2 = rot_pt(x-0.1, y), rot_pt(x+0.1, y)
            ax.plot([l1[0], l2[0]], [l1[1], l2[1]], color='limegreen', lw=1.0, zorder=4)
        elif t == 'Hinged':
            h = 0.15
            w = 0.12
            p1, p2, p3 = rot_pt(x, y), rot_pt(x+w, y-h), rot_pt(x-w, y-h)
            ax.add_patch(Polygon([p1, p2, p3], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            l1, l2 = rot_pt(x-w-0.05, y-h), rot_pt(x+w+0.05, y-h)
            ax.plot([l1[0], l2[0]], [l1[1], l2[1]], color='limegreen', lw=1.0, zorder=4)
        elif t == 'Roller':
            h = 0.15
            w = 0.12
            r = 0.04
            p1, p2, p3 = rot_pt(x, y), rot_pt(x+w, y-h), rot_pt(x-w, y-h)
            ax.add_patch(Polygon([p1, p2, p3], facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            c_pt = rot_pt(x, y-h-r)
            ax.add_patch(plt.Circle(c_pt, r, facecolor='none', edgecolor='limegreen', lw=1.0, zorder=5))
            l1, l2 = rot_pt(x-w-0.05, y-h-2*r), rot_pt(x+w+0.05, y-h-2*r)
            ax.plot([l1[0], l2[0]], [l1[1], l2[1]], color='limegreen', lw=1.0, zorder=4)

    if seg_sections and segments:
        for el in elements:
            if el['type'] == 'truss':
                n1 = nodes[el['n1']]
                n2 = nodes[el['n2']]
                mid_x = (n1[0]+n2[0])/2.0
                mid_y = (n1[1]+n2[1])/2.0
                dx = n2[0]-n1[0]
                dy = n2[1]-n1[1]
                
                rot = np.degrees(math.atan2(dy, dx))
                if rot > 90: 
                    rot -= 180
                elif rot < -90: 
                    rot += 180
                    
                L_hyp = np.hypot(dx, dy)
                if L_hyp > 1e-4:
                    nx_s = -dy/L_hyp
                    ny_s = dx/L_hyp
                    st_id = el.get('strut_idx', 0) + 1
                    sec_name_clean = get_short_name(el.get('sec', ''))
                    ax.text(mid_x + nx_s*0.1, mid_y + ny_s*0.1, f"P{st_id}: {sec_name_clean}", color='dimgray', fontsize=6, rotation=rot, ha='center', va='center', fontname='Arial')
        
        if show_seg_names:
            for i, seg in enumerate(segments):
                mx, my, mth = eval_seg_point(seg, seg.get('L', 0)/2.0)
                rot_deg = math.degrees(mth)
                if rot_deg > 90: 
                    rot_deg -= 180
                elif rot_deg < -90: 
                    rot_deg += 180
                seg_name = seg.get('name', f"S{i+1}")
                ax.text(mx - math.sin(mth)*0.3, my + math.cos(mth)*0.3, seg_name, color='dimgray', fontsize=6, ha='center', va='center', rotation=rot_deg, fontname='Arial')

    if len(supports_list) > 1 and not ss_spacings:
        sup_xs = []
        sup_ys = []
        for sup in supports_list:
            n_idx = sup['node']
            sup_xs.append(nodes[n_idx][0])
            sup_ys.append(nodes[n_idx][1])
            
        unique_xs = sorted(list(set([round(x, 3) for x in sup_xs])))
        if len(unique_xs) > 1:
            min_y = min(sup_ys)
            dim_y = min_y - 0.9 
            dyn_fontsize = max(4.0, min(7.0, 70.0 / len(unique_xs)))
            ax.plot([unique_xs[0], unique_xs[-1]], [dim_y, dim_y], color='gray', lw=0.5, zorder=1)
            
            for i in range(len(unique_xs)):
                ax.plot([unique_xs[i], unique_xs[i]], [dim_y - 0.15, dim_y + 0.15], color='gray', lw=0.5, zorder=1)
                if i < len(unique_xs) - 1:
                    dist = unique_xs[i+1] - unique_xs[i]
                    mid_x = (unique_xs[i] + unique_xs[i+1]) / 2.0
                    ax.text(mid_x, dim_y + 0.05, f"{dist:.2f}", color='dimgray', fontsize=dyn_fontsize, ha='center', va='bottom', fontname='Arial')

    if ss_spacings:
        cx = ss_spacings['cx']
        cy = ss_spacings['cy']
        strut_coords = ss_spacings['struts']
        supp_coords = ss_spacings['supps']
        
        dim_y = cy - 0.6
        hx_vals = [cx]
        for sc in strut_coords:
            hx_vals.append(sc['bx'])
        for sp in supp_coords:
            hx_vals.append(sp['x'])
            
        hx_vals = sorted(list(set(hx_vals)))
        
        ax.plot([hx_vals[0], hx_vals[-1]], [dim_y, dim_y], color='purple', lw=0.6, zorder=1)
        for i in range(len(hx_vals)):
            ax.plot([hx_vals[i], hx_vals[i]], [dim_y - 0.1, dim_y + 0.1], color='purple', lw=0.6, zorder=1)
            if i < len(hx_vals) - 1:
                dist = hx_vals[i+1] - hx_vals[i]
                mid_x = (hx_vals[i] + hx_vals[i+1]) / 2.0
                ax.text(mid_x, dim_y + 0.05, f"{dist:.2f}", color='purple', fontsize=6, ha='center', va='bottom', fontname='Arial')

        dim_x = cx - 0.6
        vy_vals = [cy]
        for sc in strut_coords:
            vy_vals.append(sc['ty'])
            
        vy_vals = sorted(list(set(vy_vals)))
        
        ax.plot([dim_x, dim_x], [vy_vals[0], vy_vals[-1]], color='purple', lw=0.6, zorder=1)
        for i in range(len(vy_vals)):
            ax.plot([dim_x - 0.1, dim_x + 0.1], [vy_vals[i], vy_vals[i]], color='purple', lw=0.6, zorder=1)
            if i < len(vy_vals) - 1:
                dist = vy_vals[i+1] - vy_vals[i]
                mid_y = (vy_vals[i] + vy_vals[i+1]) / 2.0
                ax.text(dim_x - 0.05, mid_y, f"{dist:.2f}", color='purple', fontsize=6, ha='right', va='center', fontname='Arial')

def draw_loads_and_geometry(ax, nodes, elements, supports_list, seg_sections, loads, segments=None, load_cat_filter=None, show_seg_names=False, ss_spacings=None):
    draw_base_geometry(ax, nodes, elements, supports_list, seg_sections, segments, show_seg_names, ss_spacings)
    
    scale_ld = 0.05
    color_map = {
        'Dead Load': ('blue', 0.15), 
        'Live Load': ('red', 0.15), 
        'Wind Load': ('green', 0.20)
    }
    
    for ld in loads:
        if load_cat_filter and ld.get('category') != load_cat_filter: 
            continue
            
        if abs(ld.get('w1', 0)) < 1e-4 and abs(ld.get('w2', 0)) < 1e-4: 
            continue
            
        ld_color, ld_alpha = color_map.get(ld.get('category', 'Dead Load'), ('blue', 0.15))
            
        if segments:
            i = ld.get('seg_idx', 0)
            if i >= len(segments): 
                continue
                
            w1 = ld.get('w1', 0.0)
            w2 = ld.get('w2', 0.0)
            
            num_pts = max(10, int((ld.get('end', 0) - ld.get('start', 0)) / 0.1))
            s_vals = np.linspace(ld.get('start', 0), ld.get('end', 0), num_pts)
            
            poly_pts = []
            top_pts = []
            
            for sv in s_vals:
                px, py, th = eval_seg_point(segments[i], sv)
                L_load = max(1e-5, (ld.get('end', 0) - ld.get('start', 0)))
                
                w_val = (w1 + (w2 - w1) * (sv - ld.get('start', 0)) / L_load) * scale_ld
                poly_pts.append((px, py))
                
                dir_str = ld.get('dir', '')
                if 'Global Z' in dir_str or 'Global Y' in dir_str: 
                    f_vx = 0.0
                    f_vy = w_val
                elif 'Global X' in dir_str: 
                    f_vx = w_val
                    f_vy = 0.0
                else: 
                    f_vx = -math.sin(th) * w_val
                    f_vy = math.cos(th) * w_val
                    
                top_pts.append((px - f_vx, py - f_vy))
                    
            poly_pts.extend(top_pts[::-1])
            
            if len(poly_pts) > 2:
                ax.add_patch(Polygon(poly_pts, facecolor=ld_color, edgecolor=ld_color, alpha=ld_alpha, lw=0.8, zorder=2))
                ax.add_patch(Polygon(poly_pts, facecolor='none', edgecolor=ld_color, lw=0.8, zorder=3))

def draw_live_preview(nodes, elements, supports_list, seg_sections, loads, segments=None):
    apply_plot_styles()
    fig, ax = plt.subplots(figsize=(8, 6))
    ax.axis('off')
    draw_loads_and_geometry(ax, nodes, elements, supports_list, seg_sections, loads, segments, show_seg_names=True)
    return safe_render_fig(fig, ax, nodes)

def plot_sap2000_diagrams(nodes, elements, R_reactions, scales, display_nodes, supports_list, seg_sections, loads, segments=None, ss_spacings=None, is_bridge_opt=False):
    apply_plot_styles()
    figs_dict = {}
    
    for cat in ['Dead Load', 'Live Load', 'Wind Load']:
        has_load = False
        for ld in loads:
            if ld.get('category') == cat and (abs(ld.get('w1', 0)) > 1e-4 or abs(ld.get('w2', 0)) > 1e-4):
                has_load = True
                break
                
        if has_load:
            fig_ld, ax_ld = plt.subplots(figsize=(7, 5))
            ax_ld.axis('off')
            draw_loads_and_geometry(ax_ld, nodes, elements, supports_list, seg_sections, loads, segments, load_cat_filter=cat, ss_spacings=ss_spacings)
            figs_dict[f'Load_{cat}'] = safe_render_fig(fig_ld, ax_ld, nodes)
    
    fig_r, ax_r = plt.subplots(figsize=(7, 5))
    ax_r.axis('off')
    draw_base_geometry(ax_r, nodes, elements, supports_list, seg_sections, segments, ss_spacings=ss_spacings)
    
    for sup in supports_list:
        n = sup['node']
        t = sup['type']
        ang = sup.get('angle', 0.0)
        
        Rx = R_reactions[3*n]
        Ry = R_reactions[3*n+1]
        x = nodes[n][0]
        y = nodes[n][1]
        
        c_a = math.cos(math.radians(ang))
        s_a = math.sin(math.radians(ang))
        
        R_loc_x = Rx * c_a + Ry * s_a
        R_loc_y = -Rx * s_a + Ry * c_a
        
        draw_reaction_arrow(ax_r, x, y, R_loc_x, c_a, s_a, is_bridge_mode=is_bridge_opt)
        draw_reaction_arrow(ax_r, x, y, R_loc_y, -s_a, c_a, is_bridge_mode=is_bridge_opt)
            
    figs_dict['React'] = safe_render_fig(fig_r, ax_r, nodes)
    
    def create_force_plot(val_key, scale, c_pos, c_neg):
        fig_f, ax_f = plt.subplots(figsize=(7, 5))
        ax_f.axis('off')
        draw_base_geometry(ax_f, nodes, elements, supports_list, seg_sections, segments, ss_spacings=ss_spacings)
        
        global_texts = []
        def is_far(tx, ty):
            for (px, py) in global_texts:
                if math.hypot(tx-px, ty-py) < 0.35: 
                    return False
            return True

        for el in elements:
            n1 = el['n1']
            n2 = el['n2']
            
            x1 = nodes[n1][0]
            y1 = nodes[n1][1]
            x2 = nodes[n2][0]
            y2 = nodes[n2][1]
            
            c = el.get('c', 1.0)
            s = el.get('s', 0.0)
            
            xs = el.get('internal', {}).get('x', [])
            vals = el.get('internal', {}).get(val_key, [])
            
            if len(vals) == 0 or np.all(np.abs(vals) < 1e-6): 
                continue
            
            plot_vals = -vals if val_key != 'N' else vals
            
            px = x1 + c * xs - s * plot_vals * scale
            py = y1 + s * xs + c * plot_vals * scale
            
            for k in range(len(px)-1):
                color = c_pos if vals[k] >= 0 else c_neg
                ax_f.plot([px[k], px[k+1]], [py[k], py[k+1]], color=color, lw=0.8)
                
            ax_f.plot([x1, px[0]], [y1, py[0]], color=c_pos if vals[0]>=0 else c_neg, lw=0.8)
            ax_f.plot([x2, px[-1]], [y2, py[-1]], color=c_pos if vals[-1]>=0 else c_neg, lw=0.8)

            def plot_val(idx):
                v_disp = vals[idx]
                if abs(v_disp) < 0.01: 
                    return
                tx = px[idx]
                ty = py[idx]
                
                sgn = 1 if plot_vals[idx] >= 0 else -1
                tx += -s * sgn * 0.15
                ty += c * sgn * 0.15
                
                if is_far(tx, ty):
                    v_color = c_pos if vals[idx] >= 0 else c_neg
                    ax_f.text(tx, ty, f"{v_disp:+.2f}", fontsize=6, color=v_color, ha='center', va='center', fontname='Arial')
                    global_texts.append((tx, ty))
                    
            if len(vals) > 0: 
                plot_val(len(vals)//2)
                plot_val(np.argmax(np.abs(vals)))
                
        return safe_render_fig(fig_f, ax_f, nodes)

    figs_dict['N'] = create_force_plot('N', scales.get('N', 0.015), 'blue', 'red')
    figs_dict['V'] = create_force_plot('V', scales.get('V', 0.015), 'blue', 'red')
    figs_dict['M'] = create_force_plot('M', scales.get('M', 0.015), 'blue', 'red')
    
    return figs_dict

def generate_chain_report(sys_data):
    if os.path.exists("Acrow_Template.docx"):
        doc = Document("Acrow_Template.docx") 
    else:
        doc = Document()
        
    def force_ltr_left(p):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._element.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '0')
        pPr.append(bidi)
        
    def add_line(text, bold=False):
        p = doc.add_paragraph()
        force_ltr_left(p)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        r.font.bold = bold
        r.font.rtl = False
        
    def add_large_diagram(doc, img_bytes, title):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_img = p_img.add_run()
        r_img.add_picture(io.BytesIO(img_bytes), width=Cm(15.0))
        
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run(title)
        r_title.font.name = 'Arial'
        r_title.font.size = Pt(12)
        r_title.bold = True
        
        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_line = p_line.add_run("_" * 60)
        r_line.font.size = Pt(8)
        r_line.font.color.rgb = RGBColor(150, 150, 150)
        doc.add_paragraph()

    p_title = doc.add_paragraph()
    force_ltr_left(p_title)
    run_title = p_title.add_run("CALCULATION SHEET FOR ADVANCED SHAPES")
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.rtl = False
    
    add_line("="*50, bold=True)
    add_line("1. Safety Checks (Moment, Shear & Struts):", bold=True)
    
    for df_row in sys_data['safety_df']:
        add_line(f"- {df_row['Component']} ({df_row['Force Type']}): {df_row['Actual']} vs {df_row['Allowable']} => {df_row['Status']}")
    
    doc.add_page_break()
    add_line("2. Analysis Diagrams:", bold=True)
    
    bufs = sys_data['img_bufs']
    for cat in ['Dead Load', 'Live Load', 'Wind Load']:
        key_name = f'Load_{cat}'
        if key_name in bufs: 
            add_large_diagram(doc, bufs[key_name], f"{cat} Distribution Diagram")
            
    add_large_diagram(doc, bufs['React'], "Reactions Diagram (kN)")
    add_large_diagram(doc, bufs['N'], "Axial Force Diagram (kN)")
    add_large_diagram(doc, bufs['V'], "Shear Force Diagram (kN)")
    add_large_diagram(doc, bufs['M'], "Bending Moment Diagram (kN.m)")
    
    out = io.BytesIO()
    doc.save(out)
    return out
# =====================================================================
# 🧠 THE HEURISTIC OPTIMIZER ENGINES (V26 - Dynamic Height Strut Logic)
# =====================================================================
def run_bridge_optimizer(base_segments, working_segments, active_seg_sections, ui_struts, ui_loads, target_rxn, spacings_str, view_plane, auto_mesh_size, dxf_v, is_symmetric, opt_mode, status_text, progress_bar):
    try: 
        spacings_raw = spacings_str.split(',')
        spacings = []
        for x in spacings_raw:
            spacings.append(float(x.strip()))
        spacings = sorted(spacings, reverse=True)
    except Exception as e: 
        return False, None, None, "❌ Format error in spacings."
    
    if not base_segments: 
        return False, None, None, "❌ No base segments found."
    
    min_y = 9999.0
    max_x = -9999.0
    min_x = 9999.0
    
    for seg in base_segments:
        if seg.get('Shape Type') == 'Straight Line':
            p1 = seg['abs_p1']
            p2 = seg['abs_p2']
            min_y = min(min_y, p1[1], p2[1])
            max_x = max(max_x, p1[0], p2[0])
            min_x = min(min_x, p1[0], p2[0])
            
    bottom_xs = []
    for seg in base_segments:
        if seg.get('Shape Type') == 'Straight Line':
            p1 = seg['abs_p1']
            p2 = seg['abs_p2']
            if abs(p1[1] - min_y) < 0.2 or abs(p2[1] - min_y) < 0.2:
                if abs(p1[1] - min_y) < 0.2: 
                    bottom_xs.append(p1[0])
                if abs(p2[1] - min_y) < 0.2: 
                    bottom_xs.append(p2[0])
                    
    if bottom_xs:
        soldier_min_x = min(bottom_xs)
        soldier_max_x = max(bottom_xs)
    else:
        soldier_min_x = min_x
        soldier_max_x = max_x
        
    def get_y_on_bottom_chord(test_x):
        for seg in base_segments:
            if seg.get('Shape Type') == 'Straight Line':
                p1 = seg['abs_p1']
                p2 = seg['abs_p2']
                min_px = min(p1[0], p2[0])
                max_px = max(p1[0], p2[0])
                
                if min_px - 1e-3 <= test_x <= max_px + 1e-3:
                    if abs(max_px - min_px) < 1e-5:
                        return min(p1[1], p2[1])
                    ratio = (test_x - p1[0]) / (p2[0] - p1[0])
                    return p1[1] + ratio * (p2[1] - p1[1])
        return min_y 
            
    center_x = (soldier_min_x + soldier_max_x) / 2.0
    bridge_width = soldier_max_x - soldier_min_x
    half_width = bridge_width / 2.0
    
    fac_d = float(st.session_state.get("cmb_d", 1.0))
    fac_l = float(st.session_state.get("cmb_l", 1.0))
    fac_w = float(st.session_state.get("cmb_w", 1.0))
    combo_factors = {'Dead Load': fac_d, 'Live Load': fac_l, 'Wind Load': fac_w} 
    
    base_seg_names = []
    for i, s in enumerate(base_segments):
        base_seg_names.append(s.get('name', f"S{i+1}"))
    
    test_combined_loads = []
    for i, ld in enumerate(ui_loads):
        t_mode = ld.get('target_mode', 'All Segments')
        target_base_indices = []
        
        if t_mode == "Single Segment":
            s_choice = st.session_state.get(f"ld_single_{i}", base_seg_names[0])
            if s_choice in base_seg_names: 
                target_base_indices.append(base_seg_names.index(s_choice))
        elif t_mode == "Multiple Segments":
            raw_def_segs = st.session_state.get(f"ld_multi_{i}", ld.get('target_segs', []))
            for s in raw_def_segs:
                if s in base_seg_names:
                    target_base_indices.append(base_seg_names.index(s))
        else: 
            for idx_r in range(len(base_segments)):
                target_base_indices.append(idx_r)
            
        w1 = float(st.session_state.get(f"ld_w1_{i}_{dxf_v}", ld.get('w1', 0.0)))
        if ld.get('type') == 'Trapezoidal':
            w2 = float(st.session_state.get(f"ld_w2_{i}_{dxf_v}", ld.get('w2', 0.0))) 
        else:
            w2 = w1
            
        loc_m = float(st.session_state.get(f"ld_loc_{i}_{dxf_v}", ld.get('loc', 0.0))) 
        
        target_working_indices = []
        for w_idx, w_seg in enumerate(working_segments):
            if w_seg.get('master_idx', 0) in target_base_indices:
                target_working_indices.append(w_idx)
        
        for s_idx_num in target_working_indices:
            w_len = float(working_segments[s_idx_num].get('L', 0.0))
            if ld['type'] == 'Point Load':
                start_val = min(loc_m, w_len)
                end_val = start_val
            else:
                start_val = 0.0
                end_val = w_len
                
            test_combined_loads.append({
                'seg_idx': s_idx_num, 
                'category': ld['category'], 
                'type': ld['type'], 
                'dir': ld['dir'], 
                'start': start_val, 
                'end': end_val, 
                'w1': w1 * combo_factors[ld['category']], 
                'w2': w2 * combo_factors[ld['category']]
            })

    def is_safe_soldier(elems, sections):
        # 💡 الدفلكشن ليس شرطاً للنجاح
        for i, sec in enumerate(sections):
            max_m = 0.0
            max_v = 0.0
            for el in elems:
                if el.get('group') == 'segment' and el.get('seg_idx') == i:
                    m_arr = el.get('internal', {}).get('M', [0])
                    v_arr = el.get('internal', {}).get('V', [0])
                    max_m = max(max_m, np.max(np.abs(m_arr)))
                    max_v = max(max_v, np.max(np.abs(v_arr)))
            if max_m > sec['Mall'] or max_v > sec['Qall']: 
                return False
        return True

    def run_trial(test_supps, dynamic_struts):
        nodes_t, elems_t, nloads_t, _, slist_t = build_chain_mesh(
            working_segments, active_seg_sections, test_combined_loads, 
            dynamic_struts, test_supps, base_segments, mesh_size=auto_mesh_size
        )
        U, R, net_load = solve_fea_engine(nodes_t, elems_t, nloads_t, slist_t)
        
        ry_list = []
        for sup in slist_t:
            ry_list.append(R[3*sup['node']+1])
            
        max_ry = max(ry_list)
        min_ry = min(ry_list)
        
        soldier_safe = is_safe_soldier(elems_t, active_seg_sections)
        
        struts_safe = True
        upgraded_struts = []
        for el in elems_t:
            if el['type'] == 'truss':
                N_max = np.max(np.abs(el.get('internal', {}).get('N', [0])))
                opt_sec = get_optimal_strut_section(el.get('L', 0.0), N_max, "Bridge")
                if not opt_sec: 
                    struts_safe = False
                    upgraded_struts.append(el.get('sec'))
                else: 
                    upgraded_struts.append(opt_sec)
        
        return max_ry, min_ry, net_load, soldier_safe, struts_safe, upgraded_struts

    dummy_supps = [
        {'x': soldier_min_x, 'y': min_y, 'type': 'Hinged', 'angle': 0.0}, 
        {'x': soldier_max_x, 'y': min_y, 'type': 'Hinged', 'angle': 0.0}
    ]
    _, _, total_system_load, _, _, _ = run_trial(dummy_supps, ui_struts)
    
    if target_rxn > 1e-4:
        min_required_props = max(2, int(math.ceil(total_system_load / target_rxn))) 
    else:
        min_required_props = 2

    valid_grids = []
    if is_symmetric:
        def build_sym_grids(current_grid):
            cantilever = half_width - current_grid[-1]
            if 0.15 <= cantilever <= 1.50:
                full_grid = set(current_grid)
                for x in current_grid:
                    if x > 1e-4: 
                        full_grid.add(-x)
                        
                sym_coords = []
                for x in sorted(list(full_grid)):
                    sym_coords.append(round(center_x + x, 3))
                    
                valid_grids.append(tuple(sym_coords))
                
            if cantilever < 0.15: 
                return
                
            for s in spacings: 
                build_sym_grids(current_grid + [current_grid[-1] + s])
                
        build_sym_grids([0.0])
        for s in spacings: 
            build_sym_grids([s / 2.0])
            
    else:
        def build_asym_grids(current_grid):
            cantilever = soldier_max_x - current_grid[-1]
            if 0.15 <= cantilever <= 1.50:
                asym_coords = []
                for x in current_grid:
                    if soldier_min_x - 0.05 <= x <= soldier_max_x + 0.05:
                        asym_coords.append(round(x, 3))
                valid_grids.append(tuple(asym_coords))
                
            if cantilever < 0.15: 
                return
                
            for s in spacings: 
                build_asym_grids(current_grid + [current_grid[-1] + s])
                
        cantilever_opts = np.arange(0.15, 1.51, 0.10)
        for lc in cantilever_opts: 
            build_asym_grids([soldier_min_x + lc])

    filtered_grids = []
    unique_v = list(set(valid_grids))
    for g in unique_v:
        if len(g) >= min_required_props:
            filtered_grids.append(list(g))
            
    if not filtered_grids: 
        return False, None, None, f"❌ Impossible! Requires at least {min_required_props} props with the given 0.15m minimum cantilever."

    grids_by_props = {}
    for g in filtered_grids:
        p_count = len(g)
        if p_count not in grids_by_props: 
            grids_by_props[p_count] = []
        grids_by_props[p_count].append(g)

    shift_options = [0.0, 0.10, -0.10, 0.20, -0.20]
    start_time = time.time()
    
    if "Quick" in opt_mode:
        max_time = 180.0
    else:
        max_time = 900.0
        
    best_fallback_grid = None
    best_fallback_struts = ui_struts
    best_fallback_score = 999999.0
    
    trials_count = 0
    total_estimated_trials = len(filtered_grids) * len(shift_options)
    
    sorted_p_keys = sorted(list(grids_by_props.keys()))
    
    for p_count in sorted_p_keys:
        for actual_coords in grids_by_props[p_count]:
            if time.time() - start_time > max_time: 
                break
            
            if is_symmetric:
                cantilever_L = half_width - (max(actual_coords) - center_x) 
            else:
                cantilever_L = soldier_max_x - max(actual_coords)
                
            excluded_zone_start = soldier_max_x - (cantilever_L / 3.0)
            excluded_zone_start_left = soldier_min_x + (cantilever_L / 3.0)
            
            test_supps = []
            for idx, gx in enumerate(actual_coords):
                gy = get_y_on_bottom_chord(gx)
                test_supps.append({'x': gx, 'y': round(gy, 3), 'type': 'Hinged', 'angle': 0.0})
                
            for shift_val in shift_options:
                if time.time() - start_time > max_time: 
                    break
                
                shifted_struts = []
                for strut in ui_struts:
                    new_strut = strut.copy()
                    
                    nx_b, ny_b = get_shifted_coords_along_segment(strut['bx'], strut['by'], shift_val, base_segments)
                    nx_t, ny_t = get_shifted_coords_along_segment(strut['tx'], strut['ty'], shift_val, base_segments)
                    
                    if nx_b > excluded_zone_start or nx_b < excluded_zone_start_left: 
                        nx_b, ny_b = strut['bx'], strut['by'] 
                        nx_t, ny_t = strut['tx'], strut['ty']
                        
                    new_strut['bx'], new_strut['by'] = nx_b, ny_b
                    new_strut['tx'], new_strut['ty'] = nx_t, ny_t
                    
                    if STRUTS_DB:
                        new_strut['sec'] = "PPS 252 (2.14:2.84m)"
                    else:
                        new_strut['sec'] = "Unknown" 
                        
                    shifted_struts.append(new_strut)
                
                max_ry, min_ry, _, soldier_safe, struts_safe, upg_secs = run_trial(test_supps, shifted_struts)
                
                if not struts_safe and len(upg_secs) == len(shifted_struts):
                    for idx_st in range(len(shifted_struts)): 
                        shifted_struts[idx_st]['sec'] = upg_secs[idx_st]
                    max_ry, min_ry, _, soldier_safe, struts_safe, _ = run_trial(test_supps, shifted_struts) 
                
                trials_count += 1
                if trials_count % 15 == 0:
                    ratio_prog = min(1.0, trials_count / float(total_estimated_trials))
                    progress_bar.progress(ratio_prog)
                    status_text.markdown(f"**⏳ Bridge Search:** Grid **{p_count} Props** | Rxn: **{best_fallback_score:.2f} kN**")
                
                if max_ry <= target_rxn and min_ry >= 0.5 and soldier_safe and struts_safe:
                    progress_bar.progress(1.0)
                    status_text.empty()
                    return True, test_supps, shifted_struts, f"✅ BOOM! Safe Grid Found: Max Rxn = {max_ry:.2f} kN. Props = {p_count}. Soldier & Struts are SAFE."
                    
                if max_ry < best_fallback_score:
                    best_fallback_score = max_ry
                    best_fallback_grid = test_supps
                    best_fallback_struts = shifted_struts
                        
        if time.time() - start_time > max_time: 
            break
                
    progress_bar.empty()
    status_text.empty()
    
    if best_fallback_grid:
        return False, best_fallback_grid, best_fallback_struts, f"⚠️ Notice: Best structure yields Max Rxn = {best_fallback_score:.2f} kN. Showing fallback for visual inspection."
            
    return False, None, None, f"❌ Failed! Cannot satisfy basic stability (Uplift)."
def run_single_sided_optimizer(base_segments, working_segments, active_seg_sections, ui_loads, target_rxn, view_plane, auto_mesh_size, dxf_v, status_text, progress_bar):
    """
    💡 الميزة الجديدة للـ Strongback بناءً على الكتالوج:
    لا تزيد أعداد النحايز إلا بزيادة الارتفاع لضمان عدم إهدار الموارد.
    """
    if not base_segments: 
        return False, None, None, None, "❌ No base segments."
    
    cx = 9999.0
    cy = 9999.0
    max_x = -9999.0
    max_y = -9999.0
    
    for seg in base_segments:
        if seg.get('Shape Type') == 'Straight Line':
            p1 = seg['abs_p1']
            p2 = seg['abs_p2']
            cx = min(cx, p1[0], p2[0])
            cy = min(cy, p1[1], p2[1])
            max_x = max(max_x, p1[0], p2[0])
            max_y = max(max_y, p1[1], p2[1])

    fac_d = float(st.session_state.get("cmb_d", 1.0))
    fac_l = float(st.session_state.get("cmb_l", 1.0))
    fac_w = float(st.session_state.get("cmb_w", 1.0))
    combo_factors = {'Dead Load': fac_d, 'Live Load': fac_l, 'Wind Load': fac_w}
    
    test_combined_loads = []
    for i, ld in enumerate(ui_loads):
        w1 = float(st.session_state.get(f"ld_w1_{i}_{dxf_v}", ld.get('w1', 0.0)))
        if ld.get('type') == 'Trapezoidal':
            w2 = float(st.session_state.get(f"ld_w2_{i}_{dxf_v}", ld.get('w2', 0.0))) 
        else:
            w2 = w1
            
        loc_m = float(st.session_state.get(f"ld_loc_{i}_{dxf_v}", ld.get('loc', 0.0))) 
        
        for s_idx_num in range(len(working_segments)):
            w_len = float(working_segments[s_idx_num].get('L', 0.0))
            if ld['type'] == 'Point Load':
                start_val = min(loc_m, w_len)
                end_val = start_val
            else:
                start_val = 0.0
                end_val = w_len
                
            test_combined_loads.append({
                'seg_idx': s_idx_num, 
                'category': ld['category'], 
                'type': ld['type'], 
                'dir': ld['dir'], 
                'start': start_val, 
                'end': end_val, 
                'w1': w1 * combo_factors[ld['category']], 
                'w2': w2 * combo_factors[ld['category']]
            })

    def run_trial(dynamic_struts, test_supps):
        nodes_t, elems_t, nloads_t, _, slist_t = build_chain_mesh(
            working_segments, active_seg_sections, test_combined_loads, 
            dynamic_struts, test_supps, base_segments, mesh_size=auto_mesh_size
        )
        U, R, _ = solve_fea_engine(nodes_t, elems_t, nloads_t, slist_t)
        
        corner_node = slist_t[0]['node']
        rx_c = R[3*corner_node]
        ry_c = R[3*corner_node+1]
        tie_axial = math.hypot(rx_c, ry_c) 
        
        soldier_safe = True
        for i, sec in enumerate(active_seg_sections):
            max_m = 0.0
            max_v = 0.0
            for el in elems_t:
                if el.get('group') == 'segment' and el.get('seg_idx') == i:
                    m_arr = el.get('internal', {}).get('M', [0])
                    v_arr = el.get('internal', {}).get('V', [0])
                    max_m = max(max_m, np.max(np.abs(m_arr)))
                    max_v = max(max_v, np.max(np.abs(v_arr)))
            if max_m > sec['Mall'] or max_v > sec['Qall']: 
                soldier_safe = False
        
        struts_safe = True
        upg_secs = []
        for el in elems_t:
            if el['type'] == 'truss':
                N_max = np.max(np.abs(el.get('internal', {}).get('N', [0])))
                opt_sec = get_optimal_strut_section(el.get('L', 0.0), N_max, "Single")
                if not opt_sec: 
                    struts_safe = False
                    upg_secs.append(el.get('sec'))
                else: 
                    upg_secs.append(opt_sec)
                    
        return tie_axial, soldier_safe, struts_safe, upg_secs

    # 💡 قاعدة الارتفاعات (Acrow Height Rule)
    wall_H_actual = max_y - cy
    if wall_H_actual <= 3.1:
        strut_counts = [2, 3, 4]
    elif wall_H_actual <= 4.3:
        strut_counts = [3, 4, 5]
    elif wall_H_actual <= 5.5:
        strut_counts = [4, 5, 6]
    else:
        base_k = 4 + int(math.floor((wall_H_actual - 4.3) / 1.5))
        strut_counts = [base_k, base_k + 1, base_k + 2]

    spacings = [2.4, 2.1, 1.8, 1.5, 1.2, 0.9, 0.6, 0.5, 0.4, 0.3]
    valid_h_grids = []
    def build_h_grids(current_grid):
        cantilever = max_x - current_grid[-1]
        if 0.0 <= cantilever <= 1.20:
            if len(current_grid) > 1: 
                valid_h_grids.append(tuple(current_grid[1:])) 
        if cantilever < 0.0: 
            return
        for s in spacings: 
            build_h_grids(current_grid + [current_grid[-1] + s])
            
    build_h_grids([cx])
    if not valid_h_grids: 
        mid_pt = cx + (max_x-cx)/2.0
        valid_h_grids = [(mid_pt, max_x)]
    
    y_opts = np.arange(0.6, max_y - cy + 0.1, 0.2) 
    x_opts = np.arange(0.6, max_x - cx + 0.1, 0.2) 
    shift_options = [0.0, 0.10, -0.10, 0.20, -0.20]
    
    best_fallback_tie = 999999.0
    best_struts = []
    best_supps = []
    
    start_time = time.time()
    
    for k in strut_counts:
        y_combos = list(itertools.combinations(y_opts, k))
        x_combos = list(itertools.combinations(x_opts, k))
        
        for h_grid in valid_h_grids:
            for y_tup in y_combos:
                for x_tup in x_combos:
                    if time.time() - start_time > 180.0: 
                        break 
                        
                    for shift_val in shift_options:
                        test_struts = []
                        test_supps = [{'x': cx, 'y': cy, 'type': 'Hinged', 'angle': -45.0}]
                        
                        for idx in range(k):
                            st_sec = "PPS 252 (2.14:2.84m)" if STRUTS_DB else "Unknown"
                            nx_b, ny_b = get_shifted_coords_along_segment(cx + x_tup[idx], cy, shift_val, base_segments)
                            nx_t, ny_t = get_shifted_coords_along_segment(cx, cy + y_tup[idx], shift_val, base_segments)
                            
                            test_struts.append({
                                'tx': nx_t, 
                                'ty': ny_t, 
                                'bx': nx_b, 
                                'by': ny_b, 
                                'sec': st_sec
                            })
                            test_supps.append({'x': round(nx_b, 3), 'y': cy, 'type': 'Roller', 'angle': 0.0})
                            
                        if abs(max_x - test_struts[-1]['bx']) > 0.5:
                            test_supps.append({'x': round(max_x, 3), 'y': cy, 'type': 'Roller', 'angle': 0.0})
                            
                        tie_rxn, sld_safe, st_safe, upg_secs = run_trial(test_struts, test_supps)
                        
                        if not st_safe and len(upg_secs) == k:
                            for idx_st in range(k): 
                                test_struts[idx_st]['sec'] = upg_secs[idx_st]
                            tie_rxn, sld_safe, st_safe, _ = run_trial(test_struts, test_supps)
                            
                        status_text.markdown(f"**⏳ Single-Sided Search:** Testing {k} Struts | Shift: {shift_val*100:+.0f}cm | Lowest Tie Rxn: **{best_fallback_tie:.2f} kN**")
                        
                        if sld_safe and st_safe and tie_rxn <= 180.0:
                            spacing_data = {'cx': cx, 'cy': cy, 'struts': test_struts, 'supps': test_supps[1:]}
                            status_text.empty()
                            return True, test_supps, test_struts, spacing_data, f"✅ BOOM! Perfect L-Frame. Tie Reaction = {tie_rxn:.2f} kN <= 180kN. Soldiers & Struts SAFE."
                        
                        if tie_rxn < best_fallback_tie:
                            best_fallback_tie = tie_rxn
                            best_struts = test_struts
                            best_supps = test_supps

    status_text.empty()
    if best_struts:
        spacing_data = {'cx': cx, 'cy': cy, 'struts': best_struts, 'supps': best_supps[1:]}
        return False, best_supps, best_struts, spacing_data, f"⚠️ Notice: Best safe structure yields Tie Reaction = {best_fallback_tie:.2f} kN (>180) or has un-safe members. Showing fallback."
    
    return False, None, None, None, "❌ Optimizer failed completely to find a stable geometry."
# =========================================================
# 6. Main Streamlit UI (Smart Topology & Dynamic UX)
# =========================================================
def render_advanced_shape_module():
    st.markdown("## 🎢 The Intelligent L-Frame & Chain Builder")
    
    st.markdown("""
        <style>
            div[data-testid="column"]:nth-of-type(2) { 
                position: sticky !important; 
                top: 4rem !important; 
                align-self: flex-start !important; 
                z-index: 999 !important; 
                height: max-content !important; 
            }
        </style>
    """, unsafe_allow_html=True)

    if 'dxf_version' not in st.session_state: 
        st.session_state.dxf_version = 0
    if 'opt_v' not in st.session_state: 
        st.session_state.opt_v = 0

    vp_opts = ["Section View (XZ Axes - Vertical)", "Plan View (XY Axes - Horizontal)"]
    idx_vp = vp_opts.index(st.session_state.get("view_plane_adv", vp_opts[0])) if st.session_state.get("view_plane_adv") in vp_opts else 0
    view_plane = st.radio("📐 Structural Analysis Plane / System Projection", vp_opts, index=idx_vp, key="view_plane_adv", horizontal=True)
    
    st.markdown("### 📥 Geometry Input Mode")
    geom_mode = st.radio(
        "Choose Input Method:", 
        ["Upload DXF File", "Parametric L-Frame (Single-Sided)"], 
        horizontal=True, 
        key="geom_mode_adv"
    )

    if geom_mode == "Upload DXF File":
        c_upload, c_mesh = st.columns([2, 1])
        with c_upload: 
            uploaded_dxf = st.file_uploader("📥 Upload DXF File (.dxf)", type=['dxf'], key="dxf_uploader")
        with c_mesh:
            st.write("")
            st.write("")
            auto_mesh_size = st.number_input(
                "Auto Frame Mesh Size (m)", 
                min_value=0.05, 
                max_value=5.0, 
                value=float(st.session_state.get("auto_mesh_size_adv", 0.25)), 
                step=0.05, 
                key="auto_mesh_size_adv"
            )

        if uploaded_dxf and st.button("Extract Data from DXF"):
            parsed = parse_dxf_to_data(uploaded_dxf.getvalue())
            if parsed:
                st.session_state.dxf_parsed = parsed
                st.session_state.ui_supports = parsed['supports']
                st.session_state.ui_struts = parsed['struts']
                st.session_state.ui_loads = [] 
                if 'divided_segments' in st.session_state: 
                    del st.session_state['divided_segments']
                st.session_state.dxf_version += 1
                st.session_state.opt_v += 1
                st.success("✅ DXF Parsed! Base frames extracted perfectly.")
                st.rerun()
            else: 
                st.error("❌ Failed to parse DXF. Please check layers.")
            
    else:
        st.info("📏 Generate a precise L-Frame without AutoCAD.")
        c_h, c_b, c_m, c_btn = st.columns([1, 1, 1, 1.5])
        wall_H = c_h.number_input("Vertical Height (m)", min_value=1.0, value=4.0, step=0.1)
        base_B = c_b.number_input("Horizontal Base Length (m)", min_value=1.0, value=3.0, step=0.1)
        auto_mesh_size = c_m.number_input("Mesh Size (m)", min_value=0.05, value=0.25, step=0.05)
        
        if c_btn.button("✨ Generate L-Frame Matrix", use_container_width=True, type="primary"):
            seg_v = {
                'name': 'S1', 
                'master_idx': 0, 
                'type': 'Straight Line', 
                'Shape Type': 'Straight Line', 
                'L': wall_H, 
                'start_angle': -90.0, 
                'smooth': False, 
                'is_dxf': True, 
                'abs_p1': (0.0, wall_H), 
                'abs_p2': (0.0, 0.0), 
                'kappa': 0.0, 
                'is_divided': False
            }
            seg_h = {
                'name': 'S2', 
                'master_idx': 1, 
                'type': 'Straight Line', 
                'Shape Type': 'Straight Line', 
                'L': base_B, 
                'start_angle': 0.0, 
                'smooth': False, 
                'is_dxf': True, 
                'abs_p1': (0.0, 0.0), 
                'abs_p2': (base_B, 0.0), 
                'kappa': 0.0, 
                'is_divided': False
            }
            parsed = {
                'base_segments': [seg_v, seg_h], 
                'struts': [], 
                'supports': [{'x':0.0, 'y':0.0, 'type':'Hinged', 'angle':-45.0}]
            }
            st.session_state.dxf_parsed = parsed
            st.session_state.ui_supports = parsed['supports']
            st.session_state.ui_struts = parsed['struts']
            st.session_state.ui_loads = []
            
            if 'divided_segments' in st.session_state: 
                del st.session_state['divided_segments']
                
            st.session_state.dxf_version += 1
            st.session_state.opt_v += 1
            st.success("✅ L-Frame Matrix Generated! You can now use the Single-Sided Optimizer.")
            st.rerun()

    dxf_data = st.session_state.get('dxf_parsed', None)
    dxf_v = st.session_state.dxf_version
    opt_v = st.session_state.opt_v
    
    if 'ui_supports' not in st.session_state: 
        if dxf_data:
            st.session_state.ui_supports = dxf_data['supports']
        else:
            st.session_state.ui_supports = [{'x': 0.0, 'y': 0.0, 'type': 'Hinged', 'angle': 0.0}, {'x': 3.0, 'y': 0.0, 'type': 'Hinged', 'angle': 0.0}]
            
    if 'ui_struts' not in st.session_state: 
        if dxf_data:
            st.session_state.ui_struts = dxf_data['struts']
        else:
            st.session_state.ui_struts = []
            
    if 'ui_loads' not in st.session_state: 
        st.session_state.ui_loads = []

    c_in, c_plot = st.columns([1.1, 1.9])
    
    with c_in:
        st.markdown("### 1. Supports")
        
        if dxf_data:
            base_segments = dxf_data['base_segments']
        else:
            base_segments = [{'name': 'S1', 'L': 3.0, 'type': 'Straight Line'}]
            
        base_seg_names = []
        for i, s in enumerate(base_segments):
            base_seg_names.append(s.get('name', f"S{i+1}"))

        for i, sup in enumerate(st.session_state.ui_supports):
            cc1, cc2, cc3, cc4, cc5 = st.columns([1.2, 1.2, 1.5, 1.2, 0.4])
            sup['x'] = cc1.number_input(f"J{i+1} X (m)", value=float(sup.get('x', 0.0)), format="%.4f", step=0.1, key=f"sx_{i}_{dxf_v}_{opt_v}")
            sup['y'] = cc2.number_input(f"J{i+1} Y (m)", value=float(sup.get('y', 0.0)), format="%.4f", step=0.1, key=f"sy_{i}_{dxf_v}_{opt_v}")
            
            type_opts = ["Hinged", "Roller", "Fixed"]
            idx_type = type_opts.index(sup.get('type', 'Hinged')) if sup.get('type') in type_opts else 0
            sup['type'] = cc3.selectbox(f"J{i+1} Type", type_opts, index=idx_type, key=f"sp_{i}_{dxf_v}_{opt_v}")
            sup['angle'] = cc4.number_input(f"J{i+1} Angle(°)", value=float(sup.get('angle', 0.0)), step=15.0, key=f"sa_{i}_{dxf_v}_{opt_v}")
            
            cc5.markdown("<br>", unsafe_allow_html=True)
            if cc5.button("❌", key=f"del_sup_{i}_{opt_v}"): 
                st.session_state.ui_supports.pop(i)
                st.rerun()

        if st.button("➕ Add Support", use_container_width=True): 
            st.session_state.ui_supports.append({'x': 0.0, 'y': 0.0, 'type': 'Hinged', 'angle': 0.0})
            st.rerun()

        st.markdown("### 2. Global & Override Sections")
        c_sec1, c_sec2 = st.columns([1, 1.5])
        g_sec_opts = ["Soldier U100", "Custom Section"]
        idx_gsec = g_sec_opts.index(st.session_state.get("g_sec", g_sec_opts[0])) if st.session_state.get("g_sec") in g_sec_opts else 0
        g_sec_type = c_sec1.selectbox("Global Frame Section", g_sec_opts, index=idx_gsec, key="g_sec")
        
        if g_sec_type == "Soldier U100":
            global_sec = {'name': "Soldier U100", 'E': 2100.0, 'A': 34.3 / 10000.0, 'I': 412.0 / 100000000.0, 'Mall': 13.1, 'Qall': 100.8}
            c_sec2.info("Default Acrow Soldier U100 Selected.")
        else:
            with c_sec2.expander("📝 Custom Properties", expanded=True):
                cs_A = st.number_input("Area (cm2)", value=float(st.session_state.get("cs_A_adv", 40.0)), key="cs_A_adv")
                cs_I = st.number_input("Inertia I (cm4)", value=float(st.session_state.get("cs_I_adv", 800.0)), key="cs_I_adv")
                cs_M = st.number_input("Mall (kN.m)", value=float(st.session_state.get("cs_M_adv", 20.0)), key="cs_M_adv")
                cs_Q = st.number_input("Qall (kN)", value=float(st.session_state.get("cs_Q_adv", 120.0)), key="cs_Q_adv")
                global_sec = {'name': "Custom Section", 'E': 2100.0, 'A': cs_A / 10000.0, 'I': cs_I / 100000000.0, 'Mall': cs_M, 'Qall': cs_Q}
        
        seg_sections = []
        for _ in range(len(base_segments)):
            seg_sections.append(global_sec.copy())
        
        with st.expander("🛠️ Override specific segments section", expanded=False):
            safe_override_segs = []
            for s in st.session_state.get("override_segs_adv", []):
                if s in base_seg_names:
                    safe_override_segs.append(s)
                    
            override_segs = st.multiselect("Select segments to override:", base_seg_names, default=safe_override_segs, key="override_segs_adv")
            
            if override_segs:
                o_rad_opts = ["Custom Section", "Soldier U100"]
                idx_orad = o_rad_opts.index(st.session_state.get("o_rad", o_rad_opts[0])) if st.session_state.get("o_rad") in o_rad_opts else 0
                o_sec_type = st.radio("Override Profile", o_rad_opts, index=idx_orad, key="o_rad", horizontal=True)
                
                if o_sec_type == "Soldier U100": 
                    o_sec = {'name': "Soldier U100", 'E': 2100.0, 'A': 34.3 / 10000.0, 'I': 412.0 / 100000000.0, 'Mall': 13.1, 'Qall': 100.8}
                else:
                    o1, o2, o3, o4 = st.columns(4)
                    oa_val = o1.number_input("A (cm2)", value=float(st.session_state.get("oa", 50.0)), key="oa")
                    oi_val = o2.number_input("I (cm4)", value=float(st.session_state.get("oi", 1200.0)), key="oi")
                    om_val = o3.number_input("Mall (kN.m)", value=float(st.session_state.get("om", 30.0)), key="om")
                    oq_val = o4.number_input("Qall (kN)", value=float(st.session_state.get("oq", 150.0)), key="oq")
                    o_sec = {'name': "Custom Override", 'E': 2100.0, 'A': oa_val / 10000.0, 'I': oi_val / 100000000.0, 'Mall': om_val, 'Qall': oq_val}
                    
                for s_name in override_segs: 
                    idx_seg = base_seg_names.index(s_name)
                    seg_sections[idx_seg] = o_sec.copy()

        st.markdown("### 3. Struts (Push-Pulls & Ties)")
        if STRUTS_DB:
            strut_opts_base = []
            for k in STRUTS_DB.keys():
                if "TILT" not in k.upper() and "MMP" not in k.upper():
                    strut_opts_base.append(k)
        else:
            strut_opts_base = ["PPH 353 (1.5:3.5m)"]
        
        for i, ds in enumerate(st.session_state.ui_struts):
            with st.expander(f"📏 Strut P{i+1}", expanded=True):
                c1, c2, c3, c4, c5 = st.columns([1.5, 1.5, 1.5, 1.5, 0.4])
                ds['tx'] = c1.number_input("Top X (m)", value=float(ds.get('tx', 0.0)), format="%.3f", step=0.1, key=f"st_tx_{i}_{dxf_v}_{opt_v}")
                ds['ty'] = c2.number_input("Top Y (m)", value=float(ds.get('ty', 0.0)), format="%.3f", step=0.1, key=f"st_ty_{i}_{dxf_v}_{opt_v}")
                ds['bx'] = c3.number_input("Bot X (m)", value=float(ds.get('bx', 0.0)), format="%.3f", step=0.1, key=f"st_bx_{i}_{dxf_v}_{opt_v}")
                ds['by'] = c4.number_input("Bot Y (m)", value=float(ds.get('by', 0.0)), format="%.3f", step=0.1, key=f"st_by_{i}_{dxf_v}_{opt_v}")
                
                actual_L = math.hypot(ds['bx'] - ds['tx'], ds['by'] - ds['ty'])
                
                valid_opts = []
                for opt in strut_opts_base:
                    m = re.search(r'\((\d+\.\d+):(\d+\.\d+)m\)', opt)
                    if m:
                        min_len = float(m.group(1))
                        max_len = float(m.group(2))
                        if min_len <= actual_L <= max_len:
                            valid_opts.append(opt)
                            
                if not valid_opts: 
                    valid_opts = strut_opts_base
                    
                valid_opts.sort(key=get_strut_priority)
                idx_sec = valid_opts.index(ds.get('sec')) if ds.get('sec') in valid_opts else 0
                ds['sec'] = st.selectbox(f"Type (L = {actual_L:.3f}m)", valid_opts, index=idx_sec, key=f"st_sec_{i}_{dxf_v}_{opt_v}")
                
                cc5.markdown("<br>", unsafe_allow_html=True)
                if c5.button("❌", key=f"del_st_{i}_{opt_v}"): 
                    st.session_state.ui_struts.pop(i)
                    st.rerun()

        if st.button("➕ Add Strut", use_container_width=True): 
            st.session_state.ui_struts.append({'tx': 0.0, 'ty': 3.0, 'bx': 1.0, 'by': 0.0, 'sec': strut_opts_base[0]})
            st.rerun()

        st.markdown("### ✂️ Smart Topology Division")
        if st.button("✂️ Divide Frames & Update Topology", use_container_width=True, type="primary"):
            st.session_state.divided_segments = perform_smart_division(base_segments, st.session_state.ui_supports, st.session_state.ui_struts)
            st.success("✅ Frames Divided Successfully! (e.g., S1 became S1-a, S1-b). You can now apply loads.")

        working_segments = st.session_state.get('divided_segments', base_segments)
        working_seg_names = []
        for i, s in enumerate(working_segments):
            working_seg_names.append(s.get('name', f"S{i+1}"))
            
        active_seg_sections = []
        for s in working_segments:
            m_idx = s.get('master_idx', 0)
            if m_idx < len(seg_sections):
                active_seg_sections.append(seg_sections[m_idx])
            else:
                active_seg_sections.append(global_sec)

        st.markdown("### 🤖 Generative AI Auto-Designers")
        opt_tab1, opt_tab2 = st.tabs(["🌉 Bridge Shoring", "📐 Single-Sided Wall"])
        
        with opt_tab1:
            st.info("The AI optimizes bridge supports. Checks Soldier safety, enforces 15cm min cantilever, and upgrades struts.")
            c_ai1, c_ai2 = st.columns(2)
            ai_target_rxn = c_ai1.number_input("Target Max Rxn/Leg (kN)", value=54.4, step=1.0, key='b_rxn')
            ai_spc_str = c_ai2.text_input("Spacings (m) [Comma Sep]", value="2.40, 2.10, 1.80, 1.50, 1.20, 0.90, 0.60", key='b_spc')
            is_sym = st.checkbox("Symmetric Bridge Layout", value=True)
            opt_mode = st.radio("Optimization Depth:", ["Quick Search", "Deep Search"], index=0, key='b_mode')
            
            if st.button("✨ Run Bridge AI Optimizer", type="primary", use_container_width=True):
                prog_b = st.progress(0)
                stat_b = st.empty()
                with st.spinner("🧠 AI is executing Bridge Architecture..."):
                    succ, res_supps, res_struts, msg = run_bridge_optimizer(
                        base_segments, working_segments, active_seg_sections, 
                        st.session_state.ui_struts, st.session_state.ui_loads, 
                        ai_target_rxn, ai_spc_str, view_plane, auto_mesh_size, 
                        dxf_v, is_sym, opt_mode, stat_b, prog_b
                    )
                    if res_supps:
                        st.session_state.ui_supports = res_supps
                        st.session_state.ui_struts = res_struts
                        st.session_state.opt_v += 1 
                        if succ:
                            st.success(msg)
                        else:
                            st.warning(msg)
                        time.sleep(1.5)
                        st.rerun()
                    else: 
                        st.error(msg)

        with opt_tab2:
            st.info("The AI builds Single-Sided Strongbacks dynamically. Enforces 180kN Max Tie-Rod Reaction and guarantees Soldier Safety.")
            if st.button("✨ Run Single-Sided L-Frame AI Optimizer", type="primary", use_container_width=True):
                prog_s = st.progress(0)
                stat_s = st.empty()
                with st.spinner("🧠 AI is building L-Frame Matrix & Deploying 180kN Gatekeeper..."):
                    succ, res_supps, res_struts, spc_data, msg = run_single_sided_optimizer(
                        base_segments, working_segments, active_seg_sections, 
                        st.session_state.ui_loads, 180.0, view_plane, 
                        auto_mesh_size, dxf_v, stat_s, prog_s
                    )
                    if res_supps:
                        st.session_state.ui_supports = res_supps
                        st.session_state.ui_struts = res_struts
                        st.session_state.ss_spacings = spc_data
                        st.session_state.opt_v += 1 
                        if succ:
                            st.success(msg)
                        else:
                            st.warning(msg)
                        time.sleep(2)
                        st.rerun()
                    else: 
                        st.error(msg)

    with c_plot:
        st.markdown("<h3 style='text-align: center; border-bottom: 2px solid #ddd; padding-bottom: 10px; color: #1e3d59;'>Live Geometry & Loads</h3>", unsafe_allow_html=True)
        preview_spot = st.empty()

        st.markdown("### 5. Applied Loads")
        
        with st.expander("📋 Smart Auto-Load Assigner (Excel Paste)", expanded=False):
            excel_text = st.text_area("📋 Paste Directly from Excel:", placeholder="S1 \t 25.5\nS2 \t 30.0")
            if st.button("⚡ Assign Loads from Text"):
                found_loads = 0
                for line in excel_text.split('\n'):
                    line_upper = line.strip().upper()
                    s_match = re.search(r'(S\d+)', line_upper)
                    nums = re.findall(r'-?\d+\.?\d*', line_upper)
                    
                    if s_match and nums:
                        vals = []
                        s_num_str = s_match.group(1)[1:]
                        for v in nums:
                            if v != s_num_str:
                                vals.append(v)
                                
                        if vals:
                            w_val = -abs(float(vals[-1]))
                            dir_vert = 'Global Z (Vertical)' if "XZ" in view_plane else 'Global Y (Vertical)'
                            st.session_state.ui_loads.append({
                                'category': 'Dead Load', 
                                'type': 'Uniform', 
                                'dir': dir_vert, 
                                'target_mode': 'Single Segment', 
                                'target_segs': [], 
                                'w1': w_val, 
                                'w2': w_val, 
                                'loc': 0.0,
                                '_auto_s_name': s_match.group(1)
                            })
                            found_loads += 1
                            
                if found_loads > 0: 
                    st.success(f"✅ Mapped {found_loads} Loads downwards (-Z).")
                    st.rerun()

        cc_d, cc_l, cc_w = st.columns(3)
        fac_d = cc_d.number_input("Dead Factor", value=float(st.session_state.get("cmb_d", 1.00)), step=0.1, format="%.2f", key="cmb_d")
        fac_l = cc_l.number_input("Live Factor", value=float(st.session_state.get("cmb_l", 1.00)), step=0.1, format="%.2f", key="cmb_l")
        fac_w = cc_w.number_input("Wind Factor", value=float(st.session_state.get("cmb_w", 1.00)), step=0.1, format="%.2f", key="cmb_w")
        combo_factors = {'Dead Load': fac_d, 'Live Load': fac_l, 'Wind Load': fac_w}
        
        combined_loads = []
        if "XY" in view_plane:
            dir_options = ["Global X (Horizontal)", "Global Y (Vertical)", "Local Y (Perpendicular)"] 
        else:
            dir_options = ["Global X (Horizontal)", "Global Z (Vertical)", "Local Z (Perpendicular)"]
        
        for i, ld in enumerate(st.session_state.ui_loads):
            # 💡 طي القوائم لعدم التشتت
            with st.expander(f"📥 Load Item {i+1}", expanded=False):
                col_l1, col_l2, col_l3, col_l4 = st.columns([1.5, 1.5, 1.5, 0.5])
                
                cat_opts = ["Dead Load", "Live Load", "Wind Load"]
                idx_cat = cat_opts.index(ld.get('category', 'Dead Load'))
                ld['category'] = col_l1.selectbox("Category", cat_opts, index=idx_cat, key=f"ld_cat_{i}")
                
                type_opts = ["Uniform", "Trapezoidal", "Point Load"]
                idx_typ = type_opts.index(ld.get('type', 'Uniform'))
                ld['type'] = col_l2.selectbox("Type", type_opts, index=idx_typ, key=f"ld_t_{i}")
                
                idx_dir = dir_options.index(ld.get('dir')) if ld.get('dir') in dir_options else 1
                ld['dir'] = col_l3.selectbox("Direction", dir_options, index=idx_dir, key=f"ld_d_{i}")
                
                col_l4.markdown("<br>", unsafe_allow_html=True)
                if col_l4.button("❌", key=f"del_ld_{i}"): 
                    st.session_state.ui_loads.pop(i)
                    st.rerun()
                
                t_mode_opts = ["Single Segment", "Multiple Segments", "All Segments"]
                idx_mode = t_mode_opts.index(ld.get('target_mode', 'All Segments'))
                ld['target_mode'] = st.radio("Apply Load To:", t_mode_opts, index=idx_mode, key=f"ld_mode_{i}", horizontal=True)
                
                target_base_indices = []
                
                if ld['target_mode'] == "Single Segment":
                    auto_name = ld.get('_auto_s_name', base_seg_names[0])
                    safe_auto = auto_name if auto_name in base_seg_names else base_seg_names[0]
                    idx_sing = base_seg_names.index(st.session_state.get(f"ld_single_{i}", safe_auto))
                    s_choice = st.selectbox("Select Base Segment", base_seg_names, index=idx_sing, key=f"ld_single_{i}")
                    target_base_indices.append(base_seg_names.index(s_choice))
                elif ld['target_mode'] == "Multiple Segments":
                    safe_multi = []
                    for s in st.session_state.get(f"ld_multi_{i}", ld.get('target_segs', [])):
                        if s in base_seg_names:
                            safe_multi.append(s)
                            
                    selected_segs = st.multiselect("Select Base Segments", base_seg_names, default=safe_multi, key=f"ld_multi_{i}")
                    ld['target_segs'] = selected_segs
                    for s in selected_segs: 
                        target_base_indices.append(base_seg_names.index(s))
                else: 
                    for idx_range in range(len(base_segments)):
                        target_base_indices.append(idx_range)
                
                sc1, sc2, sc3 = st.columns(3)
                ld['w1'] = sc1.number_input("Value W1 (kN/m)", value=float(ld.get('w1', 0.0)), format="%.3f", key=f"ld_w1_{i}")
                if ld['type'] == "Trapezoidal":
                    ld['w2'] = sc2.number_input("Value W2 (kN/m)", value=float(ld.get('w2', 0.0)), format="%.3f", key=f"ld_w2_{i}") 
                else:
                    ld['w2'] = ld['w1']
                    
                if ld['type'] == "Point Load":
                    ld['loc'] = sc3.number_input("Location (m)", value=float(ld.get('loc', 0.0)), format="%.3f", key=f"ld_loc_{i}")
                else:
                    ld['loc'] = 0.0
                
                for w_idx, w_seg in enumerate(working_segments):
                    if w_seg.get('master_idx', 0) in target_base_indices:
                        L_w = float(w_seg.get('L', 0.0))
                        if ld['type'] == 'Point Load':
                            start_val = min(ld['loc'], L_w)
                            end_val = start_val
                        else:
                            start_val = 0.0
                            end_val = L_w
                            
                        combined_loads.append({
                            'seg_idx': w_idx, 
                            'category': ld['category'], 
                            'type': ld['type'], 
                            'dir': ld['dir'], 
                            'start': start_val, 
                            'end': end_val, 
                            'w1': ld['w1'] * combo_factors[ld['category']], 
                            'w2': ld['w2'] * combo_factors[ld['category']]
                        })

        if st.button("➕ Add Manual Load Item", use_container_width=True): 
            st.session_state.ui_loads.append({
                'category': 'Dead Load', 
                'type': 'Uniform', 
                'dir': dir_options[1], 
                'target_mode': 'All Segments', 
                'target_segs': [], 
                'w1': 0.0, 
                'w2': 0.0,
                'loc': 0.0
            })
            st.rerun()

        nodes_base, elements_base, nodal_loads_base, display_nodes_base, supports_list_base = build_chain_mesh(
            working_segments, active_seg_sections, combined_loads, 
            st.session_state.ui_struts, st.session_state.ui_supports, 
            base_segments, mesh_size=auto_mesh_size
        )
        
        live_img = draw_live_preview(nodes_base, elements_base, supports_list_base, active_seg_sections, combined_loads, working_segments)
        preview_spot.image(live_img, use_container_width=True)
        
        st.write("")
        if st.button("🚀 Run Advanced FEA & Generate Diagrams", type="primary", use_container_width=True):
            if 'divided_segments' not in st.session_state: 
                st.warning("⚠️ Warning: You haven't clicked 'Divide Frames' yet. The analysis will run on base frames.")
                
            with st.spinner("Solving Finite Element Matrix..."):
                U_full, R_full, net_load_full = solve_fea_engine(nodes_base, elements_base, nodal_loads_base, supports_list_base) 
                
                tot_r = 0.0
                for sup in supports_list_base:
                    tot_r += R_full[3*sup['node']+1]
                    
                st.session_state.adv_fea_data = {
                    'U': U_full, 
                    'R': R_full, 
                    'nodes': nodes_base, 
                    'elements': elements_base, 
                    'display_nodes': display_nodes_base, 
                    'supports_list': supports_list_base, 
                    'seg_sections': active_seg_sections, 
                    'loads_data': combined_loads, 
                    'segments': working_segments, 
                    'total_load': net_load_full, 
                    'total_rxn': tot_r
                }
                st.session_state.adv_solved = True
                
            st.success("✅ Analysis Complete! Scroll down for accurate diagrams.")

    if getattr(st.session_state, 'adv_solved', False):
        st.markdown("---")
        fea_data = st.session_state.adv_fea_data
        st.markdown("### 📊 Equilibrium & Safety Summary")
        st.info(f"⚖️ **Physics Check:** Applied Vertical Load = **{fea_data['total_load']:.2f} kN** ➔ Vertical Reactions = **{fea_data['total_rxn']:.2f} kN**")
        
        with st.expander("⚙️ Diagram Scale Controls", expanded=False):
            c_s1, c_s2, c_s3, c_s4 = st.columns(4)
            sc_n = c_s1.slider("Axial Scale", 0.001, 0.100, float(st.session_state.get("adv_sc_n", 0.015)), step=0.001, key="adv_sc_n")
            sc_v = c_s2.slider("Shear Scale", 0.001, 0.100, float(st.session_state.get("adv_sc_v", 0.015)), step=0.001, key="adv_sc_v")
            sc_m = c_s3.slider("Moment Scale", 0.001, 0.100, float(st.session_state.get("adv_sc_m", 0.015)), step=0.001, key="adv_sc_m")
            
        is_b_opt = True if st.session_state.get('ss_spacings') is None else False
        img_bufs = plot_sap2000_diagrams(
            fea_data['nodes'], fea_data['elements'], fea_data['R'], 
            {'N': sc_n, 'V': sc_v, 'M': sc_m}, 
            fea_data['display_nodes'], fea_data['supports_list'], 
            fea_data['seg_sections'], loads=fea_data['loads_data'], 
            segments=fea_data.get('segments'), ss_spacings=st.session_state.get('ss_spacings'),
            is_bridge_opt=is_b_opt
        )

        titles = {
            'Load_Dead Load': "Dead Load Distribution", 
            'Load_Live Load': "Live Load Distribution", 
            'Load_Wind Load': "Wind Load Distribution", 
            'React': "Reactions (kN)", 
            'N': "Axial Force (kN)", 
            'V': "Shear Force (kN)", 
            'M': "Bending Moment (kN.m)"
        }
        
        load_keys = []
        for k in ['Load_Dead Load', 'Load_Live Load', 'Load_Wind Load']:
            if k in img_bufs:
                load_keys.append(k)
                
        if load_keys:
            st.markdown("#### 📥 Applied Loads Diagrams")
            cols_ld = st.columns(len(load_keys))
            for idx, lk in enumerate(load_keys):
                cols_ld[idx].image(img_bufs[lk], use_container_width=True)
                cols_ld[idx].markdown(f"<p align='center' style='font-weight:bold; border-bottom:1px solid #ccc;'>{titles[lk]}</p>", unsafe_allow_html=True)
                
        st.markdown("#### ⚙️ Internal Forces & Reactions")
        c_p1, c_p2 = st.columns(2)
        c_p1.image(img_bufs['React'], use_container_width=True)
        c_p1.markdown(f"<p align='center' style='font-weight:bold; border-bottom:1px solid #ccc;'>{titles['React']}</p>", unsafe_allow_html=True)
        
        c_p2.image(img_bufs['M'], use_container_width=True)
        c_p2.markdown(f"<p align='center' style='font-weight:bold; border-bottom:1px solid #ccc;'>{titles['M']}</p>", unsafe_allow_html=True)
        
        c_p3, c_p4 = st.columns(2)
        c_p3.image(img_bufs['V'], use_container_width=True)
        c_p3.markdown(f"<p align='center' style='font-weight:bold; border-bottom:1px solid #ccc;'>{titles['V']}</p>", unsafe_allow_html=True)
        
        c_p4.image(img_bufs['N'], use_container_width=True)
        c_p4.markdown(f"<p align='center' style='font-weight:bold; border-bottom:1px solid #ccc;'>{titles['N']}</p>", unsafe_allow_html=True)
        
        # 💡 الجداول الشاملة للأمان (بدون الدفلكشن)
        safety_data = []
        for i, sec in enumerate(fea_data['seg_sections']):
            max_m, max_v, max_d = 0.0, 0.0, 0.0
            for el in fea_data['elements']:
                if el.get('group') == 'segment' and el.get('seg_idx') == i:
                    max_m = max(max_m, np.max(np.abs(el.get('internal', {}).get('M', [0]))))
                    max_v = max(max_v, np.max(np.abs(el.get('internal', {}).get('V', [0]))))
                    max_d = max(max_d, np.max(np.abs(el.get('internal', {}).get('D', [0]))))
                    
            L_seg = float(fea_data['segments'][i].get('L', 3.0))
            
            s_stat = "SAFE" if (max_m <= sec['Mall'] and max_v <= sec['Qall']) else "UNSAFE"
            
            seg_name = fea_data['segments'][i].get('name', f'S{i+1}')
            sec_short = get_short_name(sec['name'])
            
            safety_data.append({
                "Component": f"{seg_name} ({sec_short})", 
                "Force Type": "M, V (Def: " + f"{max_d:.2f}" + "mm)", 
                "Actual": f"M={max_m:.2f}, V={max_v:.2f}", 
                "Allowable": f"M={sec['Mall']:.2f}, V={sec['Qall']:.2f}", 
                "Status": s_stat
            })
        
        if fea_data['supports_list']:
            corner_node = fea_data['supports_list'][0]['node']
            for sup in fea_data['supports_list']:
                if sup.get('angle') == -45.0: 
                    corner_node = sup['node']
                    break
                    
            tie_rxn = math.hypot(fea_data['R'][3*corner_node], fea_data['R'][3*corner_node+1])
            safety_data.append({
                "Component": "Corner Tie-Rod (2x15mm)", 
                "Force Type": "Axial Tension", 
                "Actual": f"{tie_rxn:.2f}", 
                "Allowable": "180.00", 
                "Status": "SAFE" if tie_rxn <= 180 else "UNSAFE"
            })
        
        for el in fea_data['elements']:
            if el['type'] == 'truss':
                N_max = np.max(np.abs(el.get('internal', {}).get('N', [0])))
                allow = STRUTS_DB.get(el['sec'], {}).get('allow', 999.0) if STRUTS_DB else 999.0
                st_id = el.get('strut_idx', 0) + 1
                safety_data.append({
                    "Component": f"Strut P{st_id} ({get_short_name(el['sec'])})", 
                    "Force Type": "Axial Load", 
                    "Actual": f"{N_max:.2f}", 
                    "Allowable": f"{allow:.2f}", 
                    "Status": "SAFE" if N_max <= allow else "UNSAFE"
                })

        st.table(pd.DataFrame(safety_data))
        fea_data['safety_df'] = safety_data
        fea_data['img_bufs'] = img_bufs
        
        doc_out = generate_chain_report(fea_data)
        st.download_button(
            "⬇️ Download Calculation Sheet (Word)", 
            data=doc_out.getvalue(), 
            file_name="Advanced_Shape_Calculation_Sheet.docx"
        )

if __name__ == "__main__":
    render_advanced_shape_module()