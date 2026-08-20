# ui_standard.py

import streamlit as st
import pandas as pd
import fitz
import re
from helpers import convert_transparent_to_pdf_stream
from config import SECTIONS_DB, STD_LENGTHS, SHORING_OPTIONS_SLAB, ECO_FORM_ALLOW, TECH_FORM_ALLOW, CIRCULAR_ALLOW
from math_solver import parse_loads_from_df, generate_hydrostatic_loads, get_scaffold_allowable, get_prop_allowable, solve_beam_advanced
from plot_core import draw_system_sketch, generate_acrow_diagrams

from ui_strongback import render_strongback_ui
from ui_wind import render_wind_tilting_ui

def render_slab_element(i, gamma_c, live_load, fw_load, def_sec, def_main):
    # 💡 Helper function to inherit defaults from the immediately previous Element (i-1)
    def get_def(key_base, default_val):
        if i > 0:
            return st.session_state.get(f"{key_base}_{i-1}", default_val)
        return default_val

    etype_opts = ["Slab", "Drop Panel", "Beam"]
    
    def_etype = get_def("etype", "Slab")
    if st.session_state.get(f"etype_{i}") in etype_opts:
        idx_etype = etype_opts.index(st.session_state.get(f"etype_{i}"))
    else:
        idx_etype = etype_opts.index(def_etype) if def_etype in etype_opts else 0
        
    element_type = st.selectbox("Element Type:", etype_opts, key=f"etype_{i}", index=idx_etype)
    
    st.markdown("### 🧱 Load Config")
    if element_type == "Beam":
        cb1, cb2 = st.columns(2)
        with cb1: 
            b_width = st.number_input("Beam Width (m)", value=float(st.session_state.get(f"bw_{i}", get_def("bw", 0.30))), step=0.05, key=f"bw_{i}")
        with cb2: 
            ts = st.number_input("Beam Depth (m)", value=float(st.session_state.get(f"ts_beam_{i}", get_def("ts_beam", 0.60))), step=0.05, key=f"ts_beam_{i}")
        
        bpos_opts = ["Edge", "Middle"]
        def_bpos = get_def("bpos", "Edge")
        if st.session_state.get(f"bpos_{i}") in bpos_opts:
            idx_bpos = bpos_opts.index(st.session_state.get(f"bpos_{i}"))
        else:
            idx_bpos = bpos_opts.index(def_bpos) if def_bpos in bpos_opts else 0
            
        beam_pos = st.radio("Beam Position:", bpos_opts, key=f"bpos_{i}", horizontal=True, index=idx_bpos)
    else: 
        ts = st.number_input(f"{element_type} Thickness (m)", value=float(st.session_state.get(f"ts_thick_{i}", get_def("ts_thick", 0.30))), step=0.05, key=f"ts_thick_{i}")
        b_width = 0.0
        beam_pos = "None"
        
    w_tot = (gamma_c * ts) + live_load + fw_load
    st.info(f"**Total Basic Load** = {w_tot:.2f} kN/m²")
    p_th = "18.00 mm"
    p_mal = 54.0

    st.divider()
    with st.container(border=True):
        st.markdown("### 🪵 Secondary Beam")
        col_s1, col_s2 = st.columns([1, 1.2])
        with col_s1:
            def_ssec = get_def("ss", list(SECTIONS_DB.keys())[def_sec])
            if st.session_state.get(f"ss_{i}") in SECTIONS_DB:
                idx_ssec = list(SECTIONS_DB.keys()).index(st.session_state.get(f"ss_{i}"))
            else:
                idx_ssec = list(SECTIONS_DB.keys()).index(def_ssec) if def_ssec in SECTIONS_DB else 0
                
            s_sec = st.selectbox("Sec Section", list(SECTIONS_DB.keys()), index=idx_ssec, key=f"ss_{i}")
            
            s_spc = st.number_input("Spacing (Loaded Width) (m)", value=float(st.session_state.get(f"ssp_{i}", get_def("ssp", 0.350))), step=0.005, format="%.3f", key=f"ssp_{i}")
            s_w_calc = w_tot * s_spc
            
            s_l1_opts = STD_LENGTHS.get(s_sec, [3.0])
            c_sl1, c_sl2 = st.columns([3, 1])
            with c_sl2:
                st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
                cust_sL = st.checkbox("✏️ Custom", value=st.session_state.get(f"cust_sL_{i}", get_def("cust_sL", False)), key=f"cust_sL_{i}")
            
            with c_sl1:
                if cust_sL:
                    s_L = st.number_input("Total Length (m)", min_value=0.1, value=float(st.session_state.get(f"num_sL_{i}", get_def("num_sL", s_l1_opts[-1]))), step=0.1, key=f"num_sL_{i}")
                else:
                    def_sl1 = get_def("sl1", s_l1_opts[-1])
                    if st.session_state.get(f"sl1_{i}") in s_l1_opts:
                        idx_sl1 = s_l1_opts.index(st.session_state.get(f"sl1_{i}"))
                    else:
                        idx_sl1 = s_l1_opts.index(def_sl1) if def_sl1 in s_l1_opts else (len(s_l1_opts)-1 if s_sec in STD_LENGTHS else 0)
                        
                    s_L = st.selectbox("Total Length (m)", s_l1_opts, index=idx_sl1, key=f"sl1_{i}")
            
            s_cl = st.number_input("L. Cant (m)", value=float(st.session_state.get(f"scl_{i}", get_def("scl", 0.50))), step=0.05, key=f"scl_{i}")
            s_spns = st.text_input("Spans (m) [Comma sep]", value=str(st.session_state.get(f"sspn_{i}", get_def("sspn", "1.30, 1.30"))), key=f"sspn_{i}")
            
            s_spans_list = [float(x.strip()) for x in s_spns.split(',') if x.strip()]
            s_cr = s_L - s_cl - sum(s_spans_list)
            
            if s_cr < -0.01: 
                st.error(f"❌ Error: Spans exceed total length!")
                
            auto_m_spc = 1.10
            if element_type == "Beam":
                # 💡 التعديل الخاص بمسافة تحميل المين بيم لكمرة السقف
                auto_m_spc = b_width / 2.0 if beam_pos == "Middle" else b_width
            else:
                if len(s_spans_list) > 0:
                    trib_widths = [s_cl + s_spans_list[0] / 2.0]
                    for j in range(1, len(s_spans_list)):
                        trib_widths.append((s_spans_list[j-1] + s_spans_list[j]) / 2.0)
                    right_cant = max(0.0, s_cr)
                    trib_widths.append(s_spans_list[-1] / 2.0 + right_cant)
                    auto_m_spc = max(trib_widths)
        
        with col_s2:
            st.markdown("**Load Assignment & Interactive Sketch**")
            if element_type == "Beam":
                def_s_la = max(0.0, (s_L / 2.0) - (b_width / 2.0) if beam_pos == "Middle" else s_cl - 0.05)
                def_s_lb = min(s_L, max(def_s_la, def_s_la + b_width))
            else: 
                def_s_la, def_s_lb = 0.0, s_L
                
            default_sdf = [{
                "Load Type": "Linear", 
                "WA (kN/m) or P (kN)": round(s_w_calc, 2), 
                "WB (kN/m)": round(s_w_calc, 2), 
                "LA (m) or X (m)": round(def_s_la, 2), 
                "LB (m)": round(def_s_lb, 2)
            }]
            
            saved_sdf = st.session_state.get(f"sdf_{i}")
            if not isinstance(saved_sdf, list): 
                if i > 0 and isinstance(st.session_state.get(f"sdf_{i-1}"), list):
                    saved_sdf = st.session_state.get(f"sdf_{i-1}")
                else:
                    saved_sdf = default_sdf
                
            s_df = pd.DataFrame(saved_sdf)
            
            s_loads_df = st.data_editor(
                s_df, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"sdf_{i}",
                column_config={"Load Type": st.column_config.SelectboxColumn("Load Type", options=["Linear", "Trapezoidal", "Point"], required=True)}
            )
            s_loads_parsed = parse_loads_from_df(s_loads_df)
            s_supports = [s_cl] + [s_cl + sum(s_spans_list[:j+1]) for j in range(len(s_spans_list))]
                
            if s_L > 0 and s_cr >= -0.01: 
                c_pad1, c_img, c_pad2 = st.columns([1, 4, 1])
                with c_img: 
                    s_sketch_bytes = draw_system_sketch(s_L, s_supports, s_loads_parsed, transparent_bg=True)
                    st.image(s_sketch_bytes, use_container_width=True)
                    
                if st.toggle("📊 Show Analysis Diagrams", value=st.session_state.get(f"tgl_diag_s_slab_{i}", False), key=f"tgl_diag_s_slab_{i}"):
                    with st.spinner("Calculating..."):
                        s_img_bytes, _, _, _, _, _, _ = generate_acrow_diagrams(
                            s_sec, s_L, s_supports, s_loads_parsed, SECTIONS_DB[s_sec]['E'], SECTIONS_DB[s_sec]['I'], 
                            SECTIONS_DB[s_sec]['Mall'], SECTIONS_DB[s_sec]['Qall'], Rall=None, transparent_bg=True
                        )
                        col_dwn, img_col, _ = st.columns([1, 3, 1])
                        with col_dwn: 
                            st.download_button("📥 PDF", convert_transparent_to_pdf_stream(s_img_bytes), f"{s_sec}_Diagram.pdf", "application/pdf", key=f"dwn_s_slab_{i}")
                        with img_col: 
                            st.image(s_img_bytes, use_container_width=True)

        try:
            _, _, _, _, s_R = solve_beam_advanced(s_L, s_supports, s_loads_parsed, SECTIONS_DB[s_sec]['E'], SECTIONS_DB[s_sec]['I'])
            max_s_reaction = max(s_R) if len(s_R) > 0 else 0.0
        except:
            max_s_reaction = 0.0

    st.divider()
    with st.container(border=True):
        st.markdown("### 🏗️ Main Beam & Shoring")
        col_m1, col_m2 = st.columns([1, 1.2])
        with col_m1:
            def_msec = get_def("ms", list(SECTIONS_DB.keys())[def_main])
            if st.session_state.get(f"ms_{i}") in SECTIONS_DB:
                idx_msec = list(SECTIONS_DB.keys()).index(st.session_state.get(f"ms_{i}"))
            else:
                idx_msec = list(SECTIONS_DB.keys()).index(def_msec) if def_msec in SECTIONS_DB else 0
                
            m_sec = st.selectbox("Main Section", list(SECTIONS_DB.keys()), index=idx_msec, key=f"ms_{i}")
            
            m_spc_def = round(auto_m_spc, 3)
            saved_msp = st.session_state.get(f"msp_{i}")
            if saved_msp is None or st.session_state.get(f"last_auto_msp_{i}") != m_spc_def:
                if saved_msp is None and i > 0:
                    val_msp = get_def("msp", m_spc_def)
                else:
                    val_msp = m_spc_def
                st.session_state[f"last_auto_msp_{i}"] = m_spc_def
            else:
                val_msp = float(saved_msp)
                
            m_spc = st.number_input("Spacing (Loaded Width) (m)", value=val_msp, step=0.005, format="%.3f", key=f"msp_{i}")
            
            m_l1_opts = STD_LENGTHS.get(m_sec, [3.0])
            c_ml1, c_ml2 = st.columns([3, 1])
            with c_ml2:
                st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
                cust_mL = st.checkbox("✏️ Custom", value=st.session_state.get(f"cust_mL_{i}", get_def("cust_mL", False)), key=f"cust_mL_{i}")
            
            with c_ml1:
                if cust_mL:
                    m_L = st.number_input("Total Length (m)", min_value=0.1, value=float(st.session_state.get(f"num_mL_{i}", get_def("num_mL", m_l1_opts[-1]))), step=0.1, key=f"num_mL_{i}")
                else:
                    def_wml1 = get_def("wml1", m_l1_opts[-1])
                    if st.session_state.get(f"wml1_{i}") in m_l1_opts:
                        idx_ml1 = m_l1_opts.index(st.session_state.get(f"wml1_{i}"))
                    else:
                        idx_ml1 = m_l1_opts.index(def_wml1) if def_wml1 in m_l1_opts else (len(m_l1_opts)-1 if m_sec in STD_LENGTHS else 0)
                        
                    m_L = st.selectbox("Total Length (m)", m_l1_opts, index=idx_ml1, key=f"wml1_{i}")
            
            m_cl = st.number_input("L. Cant (m)", value=float(st.session_state.get(f"mcl_{i}", get_def("mcl", 0.50))), step=0.05, key=f"mcl_{i}")
            m_spns = st.text_input("Spans (m) [Comma sep]", value=str(st.session_state.get(f"mspn_{i}", get_def("mspn", "1.20"))), key=f"mspn_{i}")
            
            m_spans_list = [float(x.strip()) for x in m_spns.split(',') if x.strip()]
            m_cr = m_L - m_cl - sum(m_spans_list)
            
            if m_cr < -0.01: 
                st.error(f"❌ Error: Spans exceed total length!")
                
            def_tn = get_def("tn", SHORING_OPTIONS_SLAB[0])
            if st.session_state.get(f"tn_{i}") in SHORING_OPTIONS_SLAB:
                idx_tn = SHORING_OPTIONS_SLAB.index(st.session_state.get(f"tn_{i}"))
            else:
                idx_tn = SHORING_OPTIONS_SLAB.index(def_tn) if def_tn in SHORING_OPTIONS_SLAB else 0
                
            t_nm = st.selectbox("Shoring Type", SHORING_OPTIONS_SLAB, index=idx_tn, key=f"tn_{i}")
            
            t_sub = ""
            t_unb = 0.0

            if t_nm == "Shorebrace Frame":
                st.info("✅ **Auto-Assigned:** Allowable Load = **54.40 kN / Leg**")
                t_al = 54.40
                
            elif t_nm == "Acrow Frame":
                st.info("✅ **Auto-Assigned:** Allowable Load = **22.25 kN / Leg**")
                t_al = 22.25
                
            elif t_nm == "Cup-lock":
                c_sup1, c_sup2 = st.columns(2)
                sub_opts = ["S355 (st.52)", "S235"]
                def_cup_sub = get_def("cup_sub", "S355 (st.52)")
                if st.session_state.get(f"cup_sub_{i}") in sub_opts:
                    idx_cup = sub_opts.index(st.session_state.get(f"cup_sub_{i}"))
                else:
                    idx_cup = sub_opts.index(def_cup_sub) if def_cup_sub in sub_opts else 0
                    
                subtype = c_sup1.selectbox("Steel Grade", sub_opts, index=idx_cup, key=f"cup_sub_{i}")
                unbraced = c_sup2.number_input("Unbraced Length (Lcr) (m)", min_value=0.5, max_value=3.0, value=float(st.session_state.get(f"cup_unb_{i}", get_def("cup_unb", 1.5))), step=0.5, key=f"cup_unb_{i}")
                t_al = get_scaffold_allowable("Cup-lock", subtype, unbraced)
                st.info(f"✅ **Auto-Calculated Load:** {t_al:.2f} kN / Leg")
                t_sub = subtype
                t_unb = unbraced
                
            elif t_nm == "Ring-lock":
                c_sup1, c_sup2 = st.columns(2)
                r_opts = ["Ringlock 1.5\"", "Ringlock 2.0\""]
                def_ring_sub = get_def("ring_sub", "Ringlock 1.5\"")
                if st.session_state.get(f"ring_sub_{i}") in r_opts:
                    idx_ring = r_opts.index(st.session_state.get(f"ring_sub_{i}"))
                else:
                    idx_ring = r_opts.index(def_ring_sub) if def_ring_sub in r_opts else 0
                    
                subtype = c_sup1.selectbox("Diameter", r_opts, index=idx_ring, key=f"ring_sub_{i}")
                unbraced = c_sup2.number_input("Unbraced Length (Lcr) (m)", min_value=1.0, max_value=3.0, value=float(st.session_state.get(f"ring_unb_{i}", get_def("ring_unb", 1.5))), step=0.5, key=f"ring_unb_{i}")
                t_al = get_scaffold_allowable("Ring-lock", subtype, unbraced)
                st.info(f"✅ **Auto-Calculated Load:** {t_al:.2f} kN / Leg")
                t_sub = subtype
                t_unb = unbraced
                
            elif t_nm == "Acrow Prop":
                c_prop1, c_prop2, c_prop3 = st.columns([1.2, 1.5, 1])
                req_ext = c_prop1.number_input("Prop Extension (m)", min_value=0.95, max_value=5.50, value=float(st.session_state.get(f"prop_ext_{i}", get_def("prop_ext", 3.20))), step=0.1, key=f"prop_ext_{i}")
                
                from config import PROP_DB
                valid_props = [k for k, v in PROP_DB.items() if v['min'] <= req_ext <= v['max']]
                if not valid_props:
                    st.error("❌ No Acrow Prop fits this extension!")
                    t_al = 0.0
                else:
                    def_prop_sel = get_def("prop_sel", valid_props[0])
                    if st.session_state.get(f"prop_sel_{i}") in valid_props:
                        idx_psel = valid_props.index(st.session_state.get(f"prop_sel_{i}"))
                    else:
                        idx_psel = valid_props.index(def_prop_sel) if def_prop_sel in valid_props else 0
                        
                    sel_prop = c_prop2.selectbox("Select Valid Prop", valid_props, index=idx_psel, key=f"prop_sel_{i}")
                    inner_up = c_prop3.toggle("Inner Tube UP?", value=bool(st.session_state.get(f"prop_dir_{i}", get_def("prop_dir", True))), key=f"prop_dir_{i}")
                    
                    t_al = get_prop_allowable(sel_prop, req_ext, inner_up)
                    st.success(f"✅ **Capacity:** {t_al:.2f} kN")
                    t_sub = sel_prop
                    
            else:
                t_al = st.number_input("Allowable (kN)", value=float(st.session_state.get(f"ta_man_{i}", get_def("ta_man", 20.0))), step=0.5, key=f"ta_man_{i}")

        with col_m2:
            st.markdown("**Load Calculation Method:**")
            meth_opts = ["Surface Pressure (W_tot × Main Spacing)", "Secondary Reaction (Max Reaction / Sec Spacing)"]
            def_meth = get_def("m_l_meth", meth_opts[0])
            if st.session_state.get(f"m_l_meth_{i}") in meth_opts:
                idx_meth = meth_opts.index(st.session_state.get(f"m_l_meth_{i}"))
            else:
                idx_meth = meth_opts.index(def_meth) if def_meth in meth_opts else 0
                
            m_load_method = st.radio("Main Beam Load Source:", meth_opts, horizontal=True, key=f"m_l_meth_{i}", label_visibility="collapsed", index=idx_meth)
            
            if "Surface Pressure" in m_load_method:
                m_w_calc = w_tot * m_spc
            else:
                m_w_calc = max_s_reaction / s_spc if s_spc > 0 else 0.0
                
            st.markdown("**Load Assignment & Interactive Sketch**")
            
            default_mdf = [{
                "Load Type": "Linear", "WA (kN/m) or P (kN)": round(m_w_calc, 2), 
                "WB (kN/m)": round(m_w_calc, 2), "LA (m) or X (m)": 0.0, "LB (m)": m_L
            }]
            saved_mdf = st.session_state.get(f"mdf_{i}")
            if not isinstance(saved_mdf, list): 
                if i > 0 and isinstance(st.session_state.get(f"mdf_{i-1}"), list):
                    saved_mdf = st.session_state.get(f"mdf_{i-1}")
                else:
                    saved_mdf = default_mdf
                
            m_df = pd.DataFrame(saved_mdf)
            
            m_loads_df = st.data_editor(
                m_df, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"mdf_{i}",
                column_config={"Load Type": st.column_config.SelectboxColumn("Load Type", options=["Linear", "Trapezoidal", "Point"], required=True)}
            )
            m_loads_parsed = parse_loads_from_df(m_loads_df)
            m_supports = [m_cl] + [m_cl + sum(m_spans_list[:j+1]) for j in range(len(m_spans_list))]
                
            if m_L > 0 and m_cr >= -0.01: 
                c_pad1, c_img, c_pad2 = st.columns([1, 4, 1])
                with c_img: 
                    m_sketch_bytes = draw_system_sketch(m_L, m_supports, m_loads_parsed, transparent_bg=True)
                    st.image(m_sketch_bytes, use_container_width=True)
                    
                if st.toggle("📊 Show Analysis Diagrams", value=st.session_state.get(f"tgl_diag_m_slab_{i}", False), key=f"tgl_diag_m_slab_{i}"):
                    with st.spinner("Calculating..."):
                        m_img_bytes, _, _, _, _, _, _ = generate_acrow_diagrams(
                            m_sec, m_L, m_supports, m_loads_parsed, SECTIONS_DB[m_sec]['E'], SECTIONS_DB[m_sec]['I'], 
                            SECTIONS_DB[m_sec]['Mall'], SECTIONS_DB[m_sec]['Qall'], Rall=t_al, transparent_bg=True
                        )
                        col_dwn, img_col, _ = st.columns([1, 3, 1])
                        with col_dwn: 
                            st.download_button("📥 PDF", convert_transparent_to_pdf_stream(m_img_bytes), f"{m_sec}_Diagram.pdf", "application/pdf", key=f"dwn_m_slab_{i}")
                        with img_col: 
                            st.image(m_img_bytes, use_container_width=True)
    
    return {
        "cat": "horizontal", "sub_cat": element_type, "ts": ts, "beam_b": b_width, "w": w_tot, 
        "ply_thick": p_th, "ply_mall": p_mal, "s_sec": s_sec, "s_spc": s_spc, "s_L": s_L, 
        "s_cl": s_cl, "s_sp": s_spns, "s_cr": s_cr, "s_ld": s_loads_parsed, "s_sup": s_supports, 
        "s_ld_img": s_sketch_bytes, "m_sec": m_sec, "m_spc": m_spc, "m_L": m_L, "m_cl": m_cl, 
        "m_sp": m_spns, "m_cr": m_cr, "m_ld": m_loads_parsed, "m_sup": m_supports, 
        "m_ld_img": m_sketch_bytes, "t_name": t_nm, "t_allow": t_al, "t_sub": t_sub, "t_unb": t_unb,
        "m_load_method": m_load_method, "max_s_rxn": max_s_reaction
    }

def render_vertical_element(i, element_subtype, def_sec, def_main):
    def get_def(key_base, default_val):
        if i > 0:
            return st.session_state.get(f"{key_base}_{i-1}", default_val)
        return default_val

    if f"hp_{i}" not in st.session_state: 
        st.session_state[f"hp_{i}"] = 3.50
        
    wall_pdf_curr = None
    t_nm = "Tie rod 15mm"
    t_al = 90.0
    section_data = {}
    
    if element_subtype == "Wall":
        wt_opts = ["Double Sided Wall", "Single Sided Wall (Strongback)"]
        def_wt = get_def("wt", wt_opts[0])
        if st.session_state.get(f"wt_{i}") in wt_opts:
            idx_wt = wt_opts.index(st.session_state.get(f"wt_{i}"))
        else:
            idx_wt = wt_opts.index(def_wt) if def_wt in wt_opts else 0
            
        wall_type = st.radio("Wall Type:", wt_opts, horizontal=True, key=f"wt_{i}", index=idx_wt) 
    else:
        wall_type = "Double Sided Wall"
        
    is_single_sided = (wall_type == "Single Sided Wall (Strongback)")
    
    if is_single_sided:
        vsys_opts = ["Timber H20 & Soldier System", "Acrow Beam S12 & Soldier System", "Eco-form Panel System", "Tech-form Panel System"]
    elif element_subtype == "Wall":
        vsys_opts = ["H20 & Soldier System", "VMC Panel System", "Eco-form Panel System", "Tech-form Panel System", "Curved Steel Panel System"]
    else:
        vsys_opts = ["H20 & Soldier System", "VMC Panel System", "Eco-form Panel System", "Tech-form Panel System", "Circular Steel Panel System"]
        
    def_vsys = get_def("vsys", vsys_opts[0])
    if st.session_state.get(f"vsys_{i}") in vsys_opts:
        idx_vsys = vsys_opts.index(st.session_state.get(f"vsys_{i}"))
    else:
        idx_vsys = vsys_opts.index(def_vsys) if def_vsys in vsys_opts else 0
        
    vert_system = st.selectbox("Formwork System", vsys_opts, index=idx_vsys, key=f"vsys_{i}")
    is_panel_sys = "Panel" in vert_system
    
    st.markdown("### 🧱 Concrete Pressure Configuration")
    
    col_w1, col_w2, col_w3, col_w4 = st.columns(4)
    with col_w1:
        st.link_button("🔗 Open Acrow Concrete Pressure Calculator", "https://acrow-sdt.github.io/pressure-calculator/")
        wall_pdf_curr = st.file_uploader("Upload Pressure PDF (Auto-extracts Value):", type=['pdf'], key=f"wall_pdf_{i}")
    
    if wall_pdf_curr:
        text = "".join([page.get_text() for page in fitz.open(stream=wall_pdf_curr.read(), filetype="pdf")])
        match_p = re.search(r'(?i)(pressure|pmax|p\s*=)[\s:=]*([\d.]+)', text)
        w_tot = float(match_p.group(2)) if match_p else 47.16
        h_static = w_tot / 25.0
        st.success(f"✅ Auto-extracted: Pmax = {w_tot} kN/m², H static = {h_static:.2f} m")
        wall_pdf_curr.seek(0)
        with col_w2: 
            wall_h = st.number_input("Pouring Height (m)", value=float(st.session_state.get(f"wall_ph_{i}", get_def("wall_ph", st.session_state.get(f"hp_{i}", 3.50)))), step=0.05, key=f"wall_ph_{i}")
    else:
        with col_w2: 
            w_tot = st.number_input("Concrete Pressure Pmax (kN/m²)", value=float(st.session_state.get(f"wall_p_{i}", get_def("wall_p", 47.16))), step=0.05, key=f"wall_p_{i}")
        with col_w3: 
            h_static = st.number_input("H static (m)", value=float(st.session_state.get(f"wall_hs_{i}", get_def("wall_hs", w_tot/25.0))), step=0.05, key=f"wall_hs_{i}")
        with col_w4: 
            wall_h = st.number_input("Pouring Height (m)", value=float(st.session_state.get(f"wall_ph_{i}", get_def("wall_ph", st.session_state.get(f"hp_{i}", 3.50)))), step=0.05, key=f"wall_ph_{i}")
            
    st.session_state[f"hp_{i}"] = wall_h
    
    m_spc_val = 1.0
    p_al_sys = 999.0
    panel_width = 0.0
    
    if is_panel_sys:
        if "Eco-form" in vert_system: 
            panel_w_opts = [0.30, 0.45, 0.60, 0.75, 0.90, 1.05]
            def_pw = get_def("panel_w", 0.60)
            if st.session_state.get(f"panel_w_{i}") in panel_w_opts:
                idx_pw = panel_w_opts.index(st.session_state.get(f"panel_w_{i}"))
            else:
                idx_pw = panel_w_opts.index(def_pw) if def_pw in panel_w_opts else 2
                
            panel_width = st.selectbox("Select Panel Width (m)", panel_w_opts, index=idx_pw, key=f"panel_w_{i}")
            p_al_sys = ECO_FORM_ALLOW.get(element_subtype, {}).get(panel_width, 90.0)
        elif "Tech-form" in vert_system: 
            p_al_sys = 80.0 if element_subtype == "Wall" else 100.0
        elif "Circular" in vert_system: 
            p_al_sys = 150.0
        elif "VMC" in vert_system: 
            p_al_sys = 70.0
        else: 
            p_al_sys = 80.0
        st.info(f"**Allowable Pressure for this Panel:** {p_al_sys} kN/m²")
    
    if not is_panel_sys:
        p_th = "18.00 mm"
        p_mal = 54.0
        st.divider()
        with st.container(border=True):
            st.markdown("### 🪵 Secondary Beam")
            col_s1, col_s2 = st.columns([1, 1.2])
            with col_s1:
                s_sec_opts = list(SECTIONS_DB.keys())
                def_s_idx = s_sec_opts.index("Timber H20") if "H20" in vert_system else s_sec_opts.index("Acrow Beam S12")
                def_wss = get_def("wss", s_sec_opts[def_s_idx])
                if st.session_state.get(f"wss_{i}") in s_sec_opts:
                    idx_wss = s_sec_opts.index(st.session_state.get(f"wss_{i}"))
                else:
                    idx_wss = s_sec_opts.index(def_wss) if def_wss in s_sec_opts else def_s_idx
                    
                s_sec = st.selectbox("Sec Section", s_sec_opts, index=idx_wss, key=f"wss_{i}")
                
                s_spc = st.number_input("Spacing (Loaded Width) (m)", value=float(st.session_state.get(f"wssp_{i}", get_def("wssp", 0.310))), step=0.005, format="%.3f", key=f"wssp_{i}")
                
                wsl1_opts = STD_LENGTHS.get(s_sec, [3.0])
                c_wsl1, c_wsl2 = st.columns([3, 1])
                with c_wsl2:
                    st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
                    cust_wsl = st.checkbox("✏️ Custom", value=st.session_state.get(f"cust_wsl_{i}", get_def("cust_wsl", False)), key=f"cust_wsl_{i}")
                
                with c_wsl1:
                    if cust_wsl:
                        s_L = st.number_input("Total Length (m)", min_value=0.1, value=float(st.session_state.get(f"num_wsl1_{i}", get_def("num_wsl1", wsl1_opts[-1]))), step=0.1, key=f"num_wsl1_{i}")
                    else:
                        def_wsl1 = get_def("wsl1", wsl1_opts[-1])
                        if st.session_state.get(f"wsl1_{i}") in wsl1_opts:
                            idx_wsl1 = wsl1_opts.index(st.session_state.get(f"wsl1_{i}"))
                        else:
                            idx_wsl1 = wsl1_opts.index(def_wsl1) if def_wsl1 in wsl1_opts else (len(wsl1_opts)-1 if s_sec in STD_LENGTHS else 0)
                            
                        s_L = st.selectbox("Total Length (m)", wsl1_opts, index=idx_wsl1, key=f"wsl1_{i}")
                
                s_cl = st.number_input("L. Cant (m)", value=float(st.session_state.get(f"wscl_{i}", get_def("wscl", 0.50))), step=0.05, key=f"wscl_{i}")
                s_spns = st.text_input("Spans (m) [Comma sep]", value=str(st.session_state.get(f"wsspn_{i}", get_def("wsspn", "1.30, 1.30"))), key=f"wsspn_{i}")
                
                s_spans_list = [float(x.strip()) for x in s_spns.split(',') if x.strip()]
                s_cr = s_L - s_cl - sum(s_spans_list)
                
                if s_cr < -0.01: 
                    st.error(f"❌ Error: Spans exceed total length!")
                    
                auto_m_spc = 1.30
                if len(s_spans_list) > 0:
                    trib_widths = [s_cl + s_spans_list[0] / 2.0]
                    for j in range(1, len(s_spans_list)):
                        trib_widths.append((s_spans_list[j-1] + s_spans_list[j]) / 2.0)
                    right_cant = max(0.0, s_cr)
                    trib_widths.append(s_spans_list[-1] / 2.0 + right_cant)
                    auto_m_spc = max(trib_widths)
                
            with col_s2:
                st.markdown("**Load Assignment & Interactive Sketch**")
                
                col_tog_s, col_inp_s = st.columns([1.2, 1])
                with col_tog_s:
                    is_hydro_s = st.toggle("Apply Hydrostatic Load (Trapezoidal) ✅", value=bool(st.session_state.get(f"wshydro_{i}", get_def("wshydro", False))), key=f"wshydro_{i}")
                
                s_top_empty = 0.0
                if is_hydro_s:
                    with col_inp_s:
                        s_top_empty = st.number_input("Top Empty Dist. (m)", min_value=0.0, max_value=float(s_L), value=float(st.session_state.get(f"s_top_empty_{i}", get_def("s_top_empty", 0.0))), step=0.05, key=f"s_top_empty_{i}")

                if is_hydro_s: 
                    s_w_max = w_tot * s_spc
                    eff_h = s_L - s_top_empty
                    h_const = max(0.0, eff_h - h_static)
                    s_loads_data = []
                    
                    if h_const > 0:
                        s_loads_data.append({"Load Type": "Linear", "WA (kN/m) or P (kN)": round(s_w_max, 2), "WB (kN/m)": round(s_w_max, 2), "LA (m) or X (m)": 0.0, "LB (m)": round(h_const, 2)})
                    
                    if h_static > 0 and eff_h > h_const:
                        s_loads_data.append({"Load Type": "Trapezoidal", "WA (kN/m) or P (kN)": round(s_w_max, 2), "WB (kN/m)": 0.0, "LA (m) or X (m)": round(h_const, 2), "LB (m)": round(eff_h, 2)})
                        
                    if eff_h < s_L:
                        s_loads_data.append({"Load Type": "Linear", "WA (kN/m) or P (kN)": 0.0, "WB (kN/m)": 0.0, "LA (m) or X (m)": round(eff_h, 2), "LB (m)": round(s_L, 2)})
                        
                    if not s_loads_data:
                        s_loads_data = [{"Load Type": "Linear", "WA (kN/m) or P (kN)": 0.0, "WB (kN/m)": 0.0, "LA (m) or X (m)": 0.0, "LB (m)": round(s_L, 2)}]
                else: 
                    s_w_calc = w_tot * s_spc
                    s_loads_data = [{"Load Type": "Linear", "WA (kN/m) or P (kN)": round(s_w_calc, 2), "WB (kN/m)": round(s_w_calc, 2), "LA (m) or X (m)": 0.0, "LB (m)": s_L}]
                    
                saved_wsdf = st.session_state.get(f"wsdf_{i}")
                if not isinstance(saved_wsdf, list): 
                    if i > 0 and isinstance(st.session_state.get(f"wsdf_{i-1}"), list):
                        saved_wsdf = st.session_state.get(f"wsdf_{i-1}")
                    else:
                        saved_wsdf = s_loads_data
                    
                s_df = pd.DataFrame(saved_wsdf)
                s_loads_df = st.data_editor(
                    s_df, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"wsdf_{i}",
                    column_config={"Load Type": st.column_config.SelectboxColumn("Load Type", options=["Linear", "Trapezoidal", "Point"], required=True)}
                )
                s_loads_parsed = parse_loads_from_df(s_loads_df)
                s_supports = [s_cl] + [s_cl + sum(s_spans_list[:j+1]) for j in range(len(s_spans_list))]
                
                if s_L > 0 and s_cr >= -0.01: 
                    c_pad1, c_img, c_pad2 = st.columns([1, 4, 1])
                    with c_img: 
                        s_sketch_bytes = draw_system_sketch(s_L, s_supports, s_loads_parsed, transparent_bg=True)
                        st.image(s_sketch_bytes, use_container_width=True)
                        
                    if st.toggle("📊 Show Analysis Diagrams", value=bool(st.session_state.get(f"tgl_diag_s_wall_{i}", False)), key=f"tgl_diag_s_wall_{i}"):
                        with st.spinner("Calculating..."):
                            ws_img_bytes, _, _, _, _, _, _ = generate_acrow_diagrams(
                                s_sec, s_L, s_supports, s_loads_parsed, SECTIONS_DB[s_sec]['E'], SECTIONS_DB[s_sec]['I'], 
                                SECTIONS_DB[s_sec]['Mall'], SECTIONS_DB[s_sec]['Qall'], Rall=None, transparent_bg=True
                            )
                            col_dwn, img_col, _ = st.columns([1, 3, 1])
                            with col_dwn: st.download_button("📥 PDF", convert_transparent_to_pdf_stream(ws_img_bytes), f"{s_sec}_Diagram.pdf", "application/pdf", key=f"dwn_s_wall_{i}")
                            with img_col: st.image(ws_img_bytes, use_container_width=True)

        try:
            _, _, _, _, s_R = solve_beam_advanced(s_L, s_supports, s_loads_parsed, SECTIONS_DB[s_sec]['E'], SECTIONS_DB[s_sec]['I'])
            max_s_reaction = max(s_R) if len(s_R) > 0 else 0.0
        except:
            max_s_reaction = 0.0

        st.divider()
        with st.container(border=True):
            st.markdown("### 🏗️ Main Beam" + ("" if is_single_sided else " & Tie Rod"))
            col_m1, col_m2 = st.columns([1, 1.2])
            with col_m1:
                def_wms = get_def("wms", list(SECTIONS_DB.keys())[def_main])
                if st.session_state.get(f"wms_{i}") in SECTIONS_DB:
                    idx_wms = list(SECTIONS_DB.keys()).index(st.session_state.get(f"wms_{i}"))
                else:
                    idx_wms = list(SECTIONS_DB.keys()).index(def_wms) if def_wms in SECTIONS_DB else def_main
                    
                m_sec = st.selectbox("Main Section", list(SECTIONS_DB.keys()), index=idx_wms, key=f"wms_{i}")
                
                m_spc_def = round(auto_m_spc, 3)
                saved_wmsp = st.session_state.get(f"wmsp_{i}")
                if saved_wmsp is None or st.session_state.get(f"last_auto_wmsp_{i}") != m_spc_def:
                    if saved_wmsp is None and i > 0:
                        val_wmsp = get_def("wmsp", m_spc_def)
                    else:
                        val_wmsp = m_spc_def
                    st.session_state[f"last_auto_wmsp_{i}"] = m_spc_def
                else:
                    val_wmsp = float(saved_wmsp)
                    
                m_spc = st.number_input("Spacing (Loaded Width) (m)", value=val_wmsp, step=0.005, format="%.3f", key=f"wmsp_{i}")
                m_spc_val = m_spc
                
                wml1_opts = STD_LENGTHS.get(m_sec, [3.0])
                c_wml1, c_wml2 = st.columns([3, 1])
                with c_wml2:
                    st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
                    cust_wml = st.checkbox("✏️ Custom", value=st.session_state.get(f"cust_wml_{i}", get_def("cust_wml", False)), key=f"cust_wml_{i}")
                
                with c_wml1:
                    if cust_wml:
                        m_L = st.number_input("Total Length (m)", min_value=0.1, value=float(st.session_state.get(f"num_wml1_{i}", get_def("num_wml1", wml1_opts[-1]))), step=0.1, key=f"num_wml1_{i}")
                    else:
                        def_wml1_w = get_def("wml1_w", wml1_opts[-1])
                        if st.session_state.get(f"wml1_w_{i}") in wml1_opts:
                            idx_wml1_w = wml1_opts.index(st.session_state.get(f"wml1_w_{i}"))
                        else:
                            idx_wml1_w = wml1_opts.index(def_wml1_w) if def_wml1_w in wml1_opts else (len(wml1_opts)-1 if m_sec in STD_LENGTHS else 0)
                            
                        m_L = st.selectbox("Total Length (m)", wml1_opts, index=idx_wml1_w, key=f"wml1_w_{i}")
                
                m_cl = st.number_input("L. Cant (m)", value=float(st.session_state.get(f"wmcl_w_{i}", get_def("wmcl_w", 0.50))), step=0.05, key=f"wmcl_w_{i}")
                m_spns = st.text_input("Spans (m) [Comma sep]", value=str(st.session_state.get(f"wmspn_w_{i}", get_def("wmspn_w", "1.32, 1.32"))), key=f"wmspn_w_{i}")
                
                m_spans_list = [float(x.strip()) for x in m_spns.split(',') if x.strip()]
                m_cr = m_L - m_cl - sum(m_spans_list)
                
                if m_cr < -0.01: 
                    st.error(f"❌ Error: Spans exceed total length!")
                    
                if not is_single_sided: 
                    t_nm = "Tie rod 15mm"
                    st.text_input("Support Type", value=t_nm, disabled=True, key=f"wtn_{i}")
                    t_al = st.number_input("Allowable (kN)", value=90.0, disabled=True, key=f"wta_{i}")
                else: 
                    t_nm = "Strongback Truss"
                    t_al = 999.0
                
            with col_m2:
                st.markdown("**Load Calculation Method:**")
                meth_opts = ["Surface Pressure (W_tot × Main Spacing)", "Secondary Reaction (Max Reaction / Sec Spacing)"]
                def_meth = get_def("wm_l_meth", meth_opts[0])
                if st.session_state.get(f"wm_l_meth_{i}") in meth_opts:
                    idx_meth = meth_opts.index(st.session_state.get(f"wm_l_meth_{i}"))
                else:
                    idx_meth = meth_opts.index(def_meth) if def_meth in meth_opts else 0
                    
                m_load_method = st.radio("Main Beam Load Source:", meth_opts, horizontal=True, key=f"wm_l_meth_{i}", label_visibility="collapsed", index=idx_meth)
                
                if "Surface Pressure" in m_load_method:
                    m_w_max = w_tot * m_spc
                else:
                    m_w_max = max_s_reaction / s_spc if s_spc > 0 else 0.0
                
                st.markdown("**Load Assignment & Interactive Sketch**")
                
                col_tog_m, col_inp_m = st.columns([1.2, 1])
                with col_tog_m:
                    is_hydro_m = st.toggle("Apply Hydrostatic Load (Trapezoidal) ✅", value=bool(st.session_state.get(f"wmhydro_{i}", get_def("wmhydro", False))), key=f"wmhydro_tog_{i}")
                
                m_top_empty = 0.0
                if is_hydro_m:
                    with col_inp_m:
                        m_top_empty = st.number_input("Top Empty Dist. (m)", min_value=0.0, max_value=float(m_L), value=float(st.session_state.get(f"m_top_empty_{i}", get_def("m_top_empty", 0.0))), step=0.05, key=f"m_top_empty_inp_{i}")

                if is_hydro_m: 
                    eff_h = m_L - m_top_empty
                    h_const = max(0.0, eff_h - h_static)
                    m_loads_data = []
                    
                    if h_const > 0:
                        m_loads_data.append({"Load Type": "Linear", "WA (kN/m) or P (kN)": round(m_w_max, 2), "WB (kN/m)": round(m_w_max, 2), "LA (m) or X (m)": 0.0, "LB (m)": round(h_const, 2)})
                    
                    if h_static > 0 and eff_h > h_const:
                        m_loads_data.append({"Load Type": "Trapezoidal", "WA (kN/m) or P (kN)": round(m_w_max, 2), "WB (kN/m)": 0.0, "LA (m) or X (m)": round(h_const, 2), "LB (m)": round(eff_h, 2)})
                        
                    if eff_h < m_L:
                        m_loads_data.append({"Load Type": "Linear", "WA (kN/m) or P (kN)": 0.0, "WB (kN/m)": 0.0, "LA (m) or X (m)": round(eff_h, 2), "LB (m)": round(m_L, 2)})
                        
                    if not m_loads_data:
                        m_loads_data = [{"Load Type": "Linear", "WA (kN/m) or P (kN)": 0.0, "WB (kN/m)": 0.0, "LA (m) or X (m)": 0.0, "LB (m)": round(m_L, 2)}]
                else: 
                    m_loads_data = [{"Load Type": "Linear", "WA (kN/m) or P (kN)": round(m_w_max, 2), "WB (kN/m)": round(m_w_max, 2), "LA (m) or X (m)": 0.0, "LB (m)": m_L}]
                    
                saved_wmdf = st.session_state.get(f"wmdf_{i}")
                if not isinstance(saved_wmdf, list): 
                    if i > 0 and isinstance(st.session_state.get(f"wmdf_{i-1}"), list):
                        saved_wmdf = st.session_state.get(f"wmdf_{i-1}")
                    else:
                        saved_wmdf = m_loads_data
                        
                m_df = pd.DataFrame(saved_wmdf)
                
                m_loads_df = st.data_editor(
                    m_df, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"wmdf_{i}",
                    column_config={"Load Type": st.column_config.SelectboxColumn("Load Type", options=["Linear", "Trapezoidal", "Point"], required=True)}
                )
                m_loads_parsed = parse_loads_from_df(m_loads_df)
                m_supports = [m_cl] + [m_cl + sum(m_spans_list[:j+1]) for j in range(len(m_spans_list))]
                
                if m_L > 0 and m_cr >= -0.01: 
                    c_pad1, c_img, c_pad2 = st.columns([1, 4, 1])
                    with c_img: 
                        m_sketch_bytes = draw_system_sketch(m_L, m_supports, m_loads_parsed, transparent_bg=True)
                        st.image(m_sketch_bytes, use_container_width=True)
                        
                    if st.toggle("📊 Show Analysis Diagrams", value=bool(st.session_state.get(f"tgl_diag_m_wall_{i}", False)), key=f"tgl_diag_m_wall_{i}"):
                        with st.spinner("Calculating..."):
                            wm_img_bytes, _, _, _, _, _, _ = generate_acrow_diagrams(
                                m_sec, m_L, m_supports, m_loads_parsed, SECTIONS_DB[m_sec]['E'], SECTIONS_DB[m_sec]['I'], 
                                SECTIONS_DB[m_sec]['Mall'], SECTIONS_DB[m_sec]['Qall'], Rall=t_al if not is_single_sided else None, transparent_bg=True
                            )
                            col_dwn, img_col, _ = st.columns([1, 3, 1])
                            with col_dwn: st.download_button("📥 PDF", convert_transparent_to_pdf_stream(wm_img_bytes), f"{m_sec}_Diagram.pdf", "application/pdf", key=f"dwn_m_wall_{i}")
                            with img_col: st.image(wm_img_bytes, use_container_width=True)

    if is_single_sided: 
        section_data['strongback'] = render_strongback_ui(i, w_tot, h_static, m_spc_val)
        
    st.divider()
    if st.toggle("🌬️ Include Wind Load Analysis & Tilting Check", value=bool(st.session_state.get(f"wind_tog_{i}", False)), key=f"wind_tog_{i}"):
        section_data['tilting'] = render_wind_tilting_ui(i, st.session_state.get(f"hp_{i}", h_static+0.5), w_tot)
    
    if is_panel_sys and not is_single_sided:
        st.divider()
        col_t, col_b = st.columns(2)
        tie_h, tie_v, bolt_h, bolt_v = 0.0, 0.0, 0.0, 0.0
        if vert_system != "Circular Steel Panel System":
            with col_t:
                st.markdown("#### 🔗 Tie Rod 15mm Configuration")
                tie_h = st.number_input("Tie Rod Horiz. Spacing (m)", value=float(st.session_state.get(f"tie_h_{i}", get_def("tie_h", 1.20))), step=0.05, key=f"tie_h_{i}")
                tie_v = st.number_input("Tie Rod Vert. Spacing (m)", value=float(st.session_state.get(f"tie_v_{i}", get_def("tie_v", 1.20))), step=0.05, key=f"tie_v_{i}")
                st.info("**Tie Rod Allowable Load:** 90.00 kN")
        with col_b:
            st.markdown("#### 🔩 Acrow Bolts Configuration")
            bolt_h = st.number_input("Bolt Horiz. Spacing (m)", value=float(st.session_state.get(f"bolt_h_{i}", get_def("bolt_h", 0.30))), step=0.05, key=f"bolt_h_{i}")
            bolt_v = st.number_input("Bolt Vert. Spacing (m)", value=float(st.session_state.get(f"bolt_v_{i}", get_def("bolt_v", 1.20))), step=0.05, key=f"bolt_v_{i}")
            st.info("**Bolt Allowable Tension/Shear:** 50.00 kN")
        section_data.update({'tie_h': tie_h, 'tie_v': tie_v, 'bolt_h': bolt_h, 'bolt_v': bolt_v})

    if not is_panel_sys:
        section_data.update({
            "cat": "vertical", "sub_cat": element_subtype, "sys_name": vert_system, "is_panel_system": False, 
            "height": st.session_state.get(f"hp_{i}", h_static+0.5) if not is_single_sided else h_static, 
            "w": w_tot, "ply_thick": p_th, "ply_mall": p_mal, "s_sec": s_sec, "s_spc": s_spc, "s_L": s_L, 
            "s_cl": s_cl, "s_sp": s_spns, "s_cr": s_cr, "s_ld": s_loads_parsed, "s_sup": s_supports, 
            "s_ld_img": s_sketch_bytes, "m_sec": m_sec, "m_spc": m_spc, "m_L": m_L, "m_cl": m_cl, 
            "m_sp": m_spns, "m_cr": m_cr, "m_ld": m_loads_parsed, "m_sup": m_supports, 
            "m_ld_img": m_sketch_bytes, "t_name": t_nm, "t_allow": t_al, "wall_pdf_curr": wall_pdf_curr,
            "m_load_method": m_load_method, "max_s_rxn": max_s_reaction
        })
    else:
        section_data.update({
            "cat": "vertical", "sub_cat": element_subtype, "sys_name": vert_system, "is_panel_system": True, 
            "w": w_tot, "panel_w": panel_width, "panel_allowable": p_al_sys, "wall_pdf_curr": wall_pdf_curr, 
            "height": st.session_state.get(f"hp_{i}", h_static+0.5) if not is_single_sided else h_static
        })
        
    return section_data