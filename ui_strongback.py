# ui_strongback.py

import streamlit as st
import pandas as pd
import numpy as np
from helpers import get_valid_struts, convert_transparent_to_pdf_stream
from config import STRUTS_DB, SECTIONS_DB
from math_solver import solve_fea, parse_loads_from_df
from plot_core import draw_sap_loads_single, draw_sap_axial_single, draw_sap_shear_single, draw_sap_moment_single, draw_sap_rxn_single, generate_s2k_file

def render_strongback_ui(i, w_tot, h_static, m_spc_val):
    sb_dict = {'active': False}
    st.divider()
    st.markdown("### 🏗️ Strong Back System")
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1: 
            val_spc = float(st.session_state.get(f"sb_spc_{i}", m_spc_val))
            sb_spc = st.number_input("Strongback Spacing (m)", value=val_spc, step=0.005, format="%.3f", key=f"sb_spc_{i}")
            w1_base = w_tot * sb_spc
            st.info(f"**Hydrostatic Base Load:** {w1_base:.2f} kN/m'")
        with c2: 
            sec_v_opts = ["Soldier U100", "Soldier ][8", "Soldier ][10", "Soldier ][12"]
            idx_v = sec_v_opts.index(st.session_state.get(f"sb_sec_v_{i}", sec_v_opts[2])) if st.session_state.get(f"sb_sec_v_{i}") in sec_v_opts else 2
            sb_sv = st.selectbox("Vert Soldier", sec_v_opts, index=idx_v, key=f"sb_sec_v_{i}")
            
            val_lv = float(st.session_state.get(f"sb_Lv_{i}", h_static+0.5))
            sb_Lv = st.number_input("Vert L (m)", value=val_lv, step=0.05, key=f"sb_Lv_{i}")
        with c3: 
            sec_h_opts = ["Soldier U100", "Soldier ][8", "Soldier ][10", "Soldier ][12"]
            idx_h = sec_h_opts.index(st.session_state.get(f"sb_sec_h_{i}", sec_h_opts[2])) if st.session_state.get(f"sb_sec_h_{i}") in sec_h_opts else 2
            sb_sh = st.selectbox("Horz Soldier", sec_h_opts, index=idx_h, key=f"sb_sec_h_{i}")
            
            val_lh = float(st.session_state.get(f"sb_Lh_{i}", 2.0))
            sb_Lh = st.number_input("Horz L (m)", value=val_lh, step=0.05, key=f"sb_Lh_{i}")
        with c4:
            sec_w_opts = ["Soldier U100", "Soldier ][8", "Soldier ][10", "Soldier ][12"]
            idx_w = sec_w_opts.index(st.session_state.get(f"sb_waler_sec_{i}", sec_w_opts[0])) if st.session_state.get(f"sb_waler_sec_{i}") in sec_w_opts else 0
            sb_waler_sec = st.selectbox("Waler Soldier", sec_w_opts, index=idx_w, key=f"sb_waler_sec_{i}")
            
            val_tie = float(st.session_state.get(f"sb_tie_h_{i}", 1.20))
            tie_h = st.number_input("Tie Rod Horiz. Spacing (m)", value=val_tie, step=0.05, key=f"sb_tie_h_{i}")
        
        st.divider()
        col_ld1, col_ld2 = st.columns([1.5, 2])
        
        with col_ld1:
            st.markdown("**Diagonals Setup**")
            cor_opts = ["Roller", "Hinged"]
            idx_cor = cor_opts.index(st.session_state.get(f"sb_corner_{i}", cor_opts[1])) if st.session_state.get(f"sb_corner_{i}") in cor_opts else 1
            corner_sup = st.selectbox("Corner Support", cor_opts, index=idx_cor, key=f"sb_corner_{i}")
            
            val_npp = int(st.session_state.get(f"num_pp_sb_{i}", 2))
            n_pp = st.number_input("Number of Diagonals:", min_value=1, max_value=6, value=val_npp, key=f"num_pp_sb_{i}")
            
            diags = []
            for d in range(int(n_pp)):
                d1, d2, d3, d4 = st.columns([1, 1, 1.5, 1])
                def_y = (d+1)*1.5
                def_x = sb_Lh-0.2
                
                val_dy = float(st.session_state.get(f"sb_dy_{i}_{d}", def_y if def_y>0 else 1.5))
                dy = d1.number_input(f"Y{d+1} (m)", value=val_dy, step=0.05, key=f"sb_dy_{i}_{d}")
                
                val_dx = float(st.session_state.get(f"sb_dx_{i}_{d}", def_x if def_x>0 else 1.8))
                dx = d2.number_input(f"X{d+1} (m)", value=val_dx, step=0.05, key=f"sb_dx_{i}_{d}")
                
                L_req = np.hypot(dx, dy)
                valid_struts = get_valid_struts(L_req, mode="strongback")
                idx_dt = valid_struts.index(st.session_state.get(f"sb_dt_{i}_{d}", valid_struts[0])) if st.session_state.get(f"sb_dt_{i}_{d}") in valid_struts else 0
                dt = d3.selectbox(f"Type ({L_req:.2f}m)", valid_struts, index=idx_dt, key=f"sb_dt_{i}_{d}")
                
                sup_opts = ["Roller", "Hinged"]
                idx_sup = sup_opts.index(st.session_state.get(f"sb_sup_{i}_{d}", sup_opts[0])) if st.session_state.get(f"sb_sup_{i}_{d}") in sup_opts else 0
                sup_type = d4.selectbox("Support", sup_opts, index=idx_sup, key=f"sb_sup_{i}_{d}")
                
                allow_val = STRUTS_DB.get(dt, {}).get('allow', 999.0) if "No strut fits" not in dt else 0.0
                diags.append({'y': dy, 'x': dx, 'type': dt, 'allow': allow_val, 'support': sup_type})
        
        with col_ld2:
            st.markdown("**Load Assignment on Vert Soldier**")
            
            # 💡 الإصلاح الجذري لحماية DataFrame
            default_sbdf = [{"Load Type": "Linear", "WA (kN/m) or P (kN)": round(w1_base, 2), "WB (kN/m)": round(w1_base, 2), "LA (m) or X (m)": 0.0, "LB (m)": round(h_static, 2)}]
            saved_sbdf = st.session_state.get(f"sb_loads_{i}")
            if not isinstance(saved_sbdf, list): 
                saved_sbdf = default_sbdf
                
            sb_df = pd.DataFrame(saved_sbdf)
            
            sb_loads_df = st.data_editor(
                sb_df, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"sb_loads_{i}",
                column_config={"Load Type": st.column_config.SelectboxColumn("Load Type", options=["Linear", "Trapezoidal", "Point"], required=True)}
            )
            sb_loads_parsed = parse_loads_from_df(sb_loads_df)
            
            sb_custom_loads = []
            sb_dist_loads = []
            for ld in sb_loads_parsed:
                if ld['type'] == 'point': sb_custom_loads.append({'p': ld['p'], 'y': ld['x']})
                elif ld['type'] in ['linear', 'Trapezoidal']: sb_dist_loads.append({'w1': ld['w1'], 'w2': ld['w2'], 'y1': ld['x1'], 'y2': ld['x2']})
        
        if all("No strut fits" not in d['type'] for d in diags):
            y_ns = [0.0, sb_Lv] + [d['y'] for d in diags]
            if sb_custom_loads: y_ns.extend([ld['y'] for ld in sb_custom_loads])
            if sb_dist_loads:
                for dld in sb_dist_loads:
                    y_ns.extend([dld['y1'], dld['y2']])
                    if dld['y2'] > dld['y1']: y_ns.extend(list(np.arange(dld['y1'], dld['y2'], 0.1))) 
                        
            y_ns = sorted(list(set([round(y, 4) for y in y_ns if 0 <= y <= sb_Lv])))
            x_ns = sorted(list(set([0.0, sb_Lh] + [d['x'] for d in diags])))
            
            nodes = []
            for y in y_ns: 
                if y >= 0: nodes.append([0.0, y, False, False, False])
                    
            base_supports = {x: {'fix_x': False, 'fix_y': True, 'fix_r': False} for x in x_ns if x > 0}
            for d in diags:
                if d['support'] == 'Hinged': base_supports[d['x']]['fix_x'] = True
                    
            for x in x_ns:
                if x > 0: nodes.append([x, 0.0, base_supports[x]['fix_x'], base_supports[x]['fix_y'], base_supports[x]['fix_r']])
                    
            nodes.append([-1.0, -1.0, True, True, True]) 
            corner_is_hinged = (corner_sup == "Hinged")
            
            if corner_is_hinged:
                nodes[0][2], nodes[0][3], nodes[0][4] = True, True, False 
            else:
                nodes[0][2], nodes[0][3], nodes[0][4] = False, False, False 
            
            def get_n(x, y): return next((j for j, n in enumerate(nodes) if abs(n[0]-x)<1e-3 and abs(n[1]-y)<1e-3), -1)
                
            els = []
            vy = [n[1] for n in nodes if n[0]==0.0]
            for j in range(len(vy)-1): els.append({'n1': get_n(0, vy[j]), 'n2': get_n(0, vy[j+1]), 'sec': sb_sv, 'mem': 'V', 'type': 'frame'})
                
            hx = [n[0] for n in nodes if n[1]==0.0]
            for j in range(len(hx)-1): els.append({'n1': get_n(hx[j], 0), 'n2': get_n(hx[j+1], 0), 'sec': sb_sh, 'mem': 'H', 'type': 'frame'})
                
            for d in diags: els.append({'n1': get_n(0, d['y']), 'n2': get_n(d['x'], 0), 'sec': d['type'].split()[0], 'mem': 'D', 'type': 'truss'})
                
            if not corner_is_hinged:
                els.append({'n1': get_n(0, 0), 'n2': get_n(-1, -1), 'sec': 'Tie', 'mem': 'Tie', 'type': 'truss'})
            
            R_tot, U = solve_fea(nodes, els, sb_custom_loads, sb_dist_loads)
            
            max_M_v = max([max(abs(el['M'])) for el in els if el['mem']=='V'] + [0])
            max_V_v = max([max(abs(el['V'])) for el in els if el['mem']=='V'] + [0])
            max_M_h = max([max(abs(el['M'])) for el in els if el['mem']=='H'] + [0])
            max_V_h = max([max(abs(el['V'])) for el in els if el['mem']=='H'] + [0])
            
            v_nodes_idx = [j for j, n in enumerate(nodes) if n[0]==0]
            max_def_v = max([abs(U[3*idx])*1000 for idx in v_nodes_idx])
            y_vals = sorted([nodes[idx][1] for idx in v_nodes_idx])
            allw_def_v = ((max(np.diff(y_vals)) if len(y_vals)>1 else sb_Lv) * 1000) / 400.0
            
            corner_idx = get_n(0, 0)
            if corner_is_hinged:
                if corner_idx != -1:
                    rx_corner = R_tot[3 * corner_idx]
                    ry_corner = R_tot[3 * corner_idx + 1]
                    tie_force_total = abs(rx_corner * 0.70710678 + ry_corner * 0.70710678)
                else:
                    tie_force_total = 0.0
            else:
                tie_el = next((e for e in els if e['mem'] == 'Tie'), None)
                tie_force_total = abs(tie_el['N_ax']) if tie_el else 0.0
                
            tie_force_single = tie_force_total / 2.0
            waler_M = (tie_force_single * tie_h) / 4.0
            
            max_diag_force = max([abs(e['N_ax']) for e in els if e['type'] == 'truss' and e['mem'] == 'D'] + [0])
            
            sb_dict.update({
                'active': True, 'spc': sb_spc, 'Lv': sb_Lv, 'Lh': sb_Lh, 'w': w1_base, 'sv': sb_sv, 'sh': sb_sh, 'diags': diags,
                'M_v': max_M_v, 'V_v': max_V_v, 'D_v': max_def_v, 'allw_D_v': allw_def_v, 'M_h': max_M_h, 'V_h': max_V_h, 
                'tie_T_single': tie_force_single, 'tie_force_total': tie_force_total, 'waler_sec': sb_waler_sec, 'waler_M': waler_M, 'tie_h': tie_h,
                'max_diag_force': max_diag_force, 'elements': els
            })

            twsb_val = st.session_state.get(f"twsb_{i}", False)
            if st.toggle("📊 Show Analysis Diagrams", value=twsb_val, key=f"twsb_{i}"):
                
                st.markdown("**Diagrams View Settings (Independent Scales)**")
                c_sc1, c_sc2, c_sc3, c_sc4 = st.columns(4)
                sc_ld = c_sc1.slider("Loads Scale", min_value=0.1, max_value=5.0, value=float(st.session_state.get(f"sc_ld_{i}", 1.0)), step=0.1, key=f"sc_ld_{i}")
                sc_ax = c_sc2.slider("Axial Scale", min_value=0.1, max_value=5.0, value=float(st.session_state.get(f"sc_ax_{i}", 1.0)), step=0.1, key=f"sc_ax_{i}")
                sc_sh = c_sc3.slider("Shear Scale", min_value=0.1, max_value=5.0, value=float(st.session_state.get(f"sc_sh_{i}", 1.0)), step=0.1, key=f"sc_sh_{i}")
                sc_mo = c_sc4.slider("Moment Scale", min_value=0.1, max_value=5.0, value=float(st.session_state.get(f"sc_mo_{i}", 1.0)), step=0.1, key=f"sc_mo_{i}")
                
                img_ld_bytes = draw_sap_loads_single(nodes, els, sb_custom_loads, sb_dist_loads, sc_ld)
                img_ax_bytes = draw_sap_axial_single(nodes, els, sc_ax)
                img_sh_bytes = draw_sap_shear_single(nodes, els, sc_sh)
                img_mo_bytes = draw_sap_moment_single(nodes, els, sc_mo)
                img_rx_bytes = draw_sap_rxn_single(nodes, els, R_tot, corner_sup)
                
                sb_dict.update({
                    'img_ld_single': img_ld_bytes, 'img_ax_single': img_ax_bytes,
                    'img_sh_single': img_sh_bytes, 'img_mo_single': img_mo_bytes, 'img_rx_single': img_rx_bytes
                })
                
                st.markdown("#### SAP2000 Analysis Output")
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.image(img_ld_bytes, caption="Loads", use_container_width=True)
                    st.download_button("📥 PDF", convert_transparent_to_pdf_stream(img_ld_bytes), "Loads.pdf", key=f"dl_ld_{i}")
                with c2:
                    st.image(img_ax_bytes, caption="Axial", use_container_width=True)
                    st.download_button("📥 PDF", convert_transparent_to_pdf_stream(img_ax_bytes), "Axial.pdf", key=f"dl_ax_{i}")
                with c3:
                    st.image(img_sh_bytes, caption="Shear", use_container_width=True)
                    st.download_button("📥 PDF", convert_transparent_to_pdf_stream(img_sh_bytes), "Shear.pdf", key=f"dl_sh_{i}")
                with c4:
                    st.image(img_mo_bytes, caption="Moment", use_container_width=True)
                    st.download_button("📥 PDF", convert_transparent_to_pdf_stream(img_mo_bytes), "Moment.pdf", key=f"dl_mo_{i}")
                with c5:
                    st.image(img_rx_bytes, caption="Reactions", use_container_width=True)
                    st.download_button("📥 PDF", convert_transparent_to_pdf_stream(img_rx_bytes), "Reactions.pdf", key=f"dl_rx_{i}")
                
                s2k_data = generate_s2k_file(nodes, els, sb_custom_loads, sb_dist_loads, corner_sup)
                st.download_button("📥 Export S2K", data=s2k_data, file_name=f"Strongback_System_{i+1}.s2k", mime="text/plain", key=f"s2k_sb_{i}")
                
        else: 
            st.error("⚠️ Invalid Diagonals!")
            sb_dict['active'] = False
            
    return sb_dict