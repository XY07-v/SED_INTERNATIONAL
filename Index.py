import os
import pandas as pd
import unicodedata

def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn').upper()

file_path = r"C:\Users\edwar.vanegas_PANASA\OneDrive - ManpowerGroup Colombia\Escritorio\SED INTERNATIONAL\DATA.xlsx"
base_dir = r"C:\Users\edwar.vanegas_PANASA\OneDrive - ManpowerGroup Colombia\Escritorio\SED INTERNATIONAL"
output_path = os.path.join(base_dir, "index.html")

excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names

df_base = pd.read_excel(file_path, sheet_name="BASE") if "BASE" in sheet_names else pd.DataFrame()
df_asistencia = pd.read_excel(file_path, sheet_name="MALLA") if "MALLA" in sheet_names else pd.DataFrame()
df_inventarios = pd.read_excel(file_path, sheet_name="INVENTARIOS") if "INVENTARIOS" in sheet_names else pd.DataFrame()
df_ventas = pd.read_excel(file_path, sheet_name="VENTAS") if "VENTAS" in sheet_names else pd.DataFrame()
df_exhibiciones = pd.read_excel(file_path, sheet_name="EXHIBICION") if "EXHIBICION" in sheet_names else pd.DataFrame()

visita_sheet = next((s for s in sheet_names if "VISITA" in s.upper() or "POC" in s.upper()), None)
df_visita_pocs = pd.read_excel(file_path, sheet_name=visita_sheet) if visita_sheet else pd.DataFrame()

for df in [df_base, df_asistencia, df_inventarios, df_ventas, df_exhibiciones, df_visita_pocs]:
    if not df.empty:
        df.columns = df.columns.str.strip()

def clean_df_dates_and_months(df):
    if df.empty: return df
    for col in df.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df[col] = df[col].fillna("")
            if not any("MES" in strip_accents(c) for c in df.columns):
                df["Mes"] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d").str.slice(0, 7)
                df["Mes"] = df["Mes"].fillna("")
    return df

df_base = clean_df_dates_and_months(df_base)
df_inventarios = clean_df_dates_and_months(df_inventarios)
df_ventas = clean_df_dates_and_months(df_ventas)

if not df_asistencia.empty:
    for col in df_asistencia.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean:
            df_asistencia[col] = pd.to_datetime(df_asistencia[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df_asistencia[col] = df_asistencia[col].fillna("")
            if not any("MES" in strip_accents(c) for c in df_asistencia.columns):
                df_asistencia["Mes"] = pd.to_datetime(df_asistencia[col], errors="coerce").dt.strftime("%Y-%m-%d").str.slice(0, 7)
                df_asistencia["Mes"] = df_asistencia["Mes"].fillna("")

def fix_coordinates(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            try:
                lat_raw = parts[0].strip()
                lon_raw = parts[1].strip()
                
                if "." not in lat_raw and len(lat_raw) > 6:
                    lat_val = float(lat_raw)
                    if abs(lat_val) > 1000:
                        lat_val = lat_val / (10 ** (len(lat_raw) - 2))
                else:
                    lat_val = float(lat_raw)
                    
                if "." not in lon_raw and len(lon_raw) > 6:
                    lon_val = float(lon_raw)
                    if abs(lon_val) > 1000:
                        lon_val = lon_val / (10 ** (len(lon_raw) - 3))
                else:
                    lon_val = float(lon_raw)
                    
                return f"{lat_val:.6f}, {lon_val:.6f}"
            except:
                return s
    return s

if not df_visita_pocs.empty:
    for col in df_visita_pocs.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean:
            df_visita_pocs[col] = pd.to_datetime(df_visita_pocs[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df_visita_pocs[col] = df_visita_pocs[col].fillna("")
            if not any("MES" in strip_accents(c) for c in df_visita_pocs.columns):
                df_visita_pocs["Mes"] = pd.to_datetime(df_visita_pocs[col], errors="coerce").dt.strftime("%Y-%m-%d").str.slice(0, 7)
                df_visita_pocs["Mes"] = df_visita_pocs["Mes"].fillna("")
        if "GEO" in c_clean or "UBICACION" in c_clean or "COORDENADA" in c_clean:
            df_visita_pocs[col] = df_visita_pocs[col].apply(fix_coordinates)

if not df_exhibiciones.empty:
    for col in df_exhibiciones.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean:
            df_exhibiciones[col] = pd.to_datetime(df_exhibiciones[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df_exhibiciones[col] = df_exhibiciones[col].replace("NaT", "").fillna("")
            if not any("MES" in strip_accents(c) for c in df_exhibiciones.columns):
                df_exhibiciones["Mes"] = pd.to_datetime(df_exhibiciones[col], errors="coerce").dt.strftime("%Y-%m-%d").str.slice(0, 7)
                df_exhibiciones["Mes"] = df_exhibiciones["Mes"].fillna("")

fechas = sorted(df_base["Fecha"].dropna().astype(str).unique()) if "Fecha" in df_base.columns else []
meses = sorted(df_base["Mes"].dropna().astype(str).unique()) if "Mes" in df_base.columns else []
nombres = sorted(df_base["Nombre completo"].dropna().astype(str).unique()) if "Nombre completo" in df_base.columns else []

asistencia_val = round(df_base["Asistencia"].mean() * 100) if "Asistencia" in df_base.columns and pd.api.types.is_numeric_dtype(df_base["Asistencia"]) else 0
inventarios_val = round(df_base["Inventarios"].mean() * 100) if "Inventarios" in df_base.columns and pd.api.types.is_numeric_dtype(df_base["Inventarios"]) else 0
ventas_val = round(df_base["Ventas"].mean() * 100) if "Ventas" in df_base.columns and pd.api.types.is_numeric_dtype(df_base["Ventas"]) else 0
exhibiciones_val = round(df_base["Exhibiciones"].mean() * 100) if "Exhibiciones" in df_base.columns and pd.api.types.is_numeric_dtype(df_base["Exhibiciones"]) else 0
visita_pocs_val = round(df_base["Visita Pocs"].mean() * 100) if "Visita Pocs" in df_base.columns and pd.api.types.is_numeric_dtype(df_base["Visita Pocs"]) else 0

def get_filter_options(df, col_keywords):
    for col in df.columns:
        c_clean = strip_accents(col)
        if any(kw in c_clean for kw in col_keywords):
            return sorted(df[col].dropna().astype(str).unique())
    return []

inv_fechas = get_filter_options(df_inventarios, ["FECHA"])
inv_meses = get_filter_options(df_inventarios, ["MES"])
inv_nombres = get_filter_options(df_inventarios, ["NOMBRE"])
inv_pvs = get_filter_options(df_inventarios, ["PUNTO", "VENTA", "PDV", "CLIENTE"])
inv_refs = get_filter_options(df_inventarios, ["REFERENCIA", "REF", "SKU", "PRODUCTO"])

ven_fechas = get_filter_options(df_ventas, ["FECHA"])
ven_meses = get_filter_options(df_ventas, ["MES"])
ven_nombres = get_filter_options(df_ventas, ["NOMBRE"])
ven_pvs = get_filter_options(df_ventas, ["PUNTO", "VENTA", "PDV", "CLIENTE"])
ven_refs = get_filter_options(df_ventas, ["REFERENCIA", "REF", "SKU", "PRODUCTO"])

ast_fechas = get_filter_options(df_asistencia, ["FECHA"])
ast_meses = get_filter_options(df_asistencia, ["MES"])
ast_nombres = get_filter_options(df_asistencia, ["NOMBRE"])
ast_estados = get_filter_options(df_asistencia, ["ESTADO"])

exh_fechas = get_filter_options(df_exhibiciones, ["FECHA"])
exh_meses = get_filter_options(df_exhibiciones, ["MES"])
exh_nombres = get_filter_options(df_exhibiciones, ["NOMBRE"])
exh_pvs = get_filter_options(df_exhibiciones, ["PUNTO", "VENTA", "PDV", "CLIENTE"])

vis_fechas = get_filter_options(df_visita_pocs, ["FECHA"])
vis_meses = get_filter_options(df_visita_pocs, ["MES"])
vis_nombres = get_filter_options(df_visita_pocs, ["NOMBRE"])
vis_pvs = get_filter_options(df_visita_pocs, ["PUNTO", "VENTA", "PDV", "CLIENTE"])

icons = {
    "general": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z"/></svg>',
    "asistencia": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"/></svg>',
    "inventarios": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/></svg>',
    "ventas": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>',
    "exhibiciones": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 6a2 2 0 012-2h12a2 2 0 012 2v12a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM9 4v16M15 4v16M4 9h16M4 15h16"/></svg>',
    "visita_pocs": '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/><path d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/></svg>'
}

def format_pct_with_color(val):
    try:
        if pd.isna(val) or val == "":
            return "", ""
        num = float(val) * 100
        pct_int = round(num)
        text_display = f"{pct_int}%"
        
        if pct_int == 100:
            bg_color, text_color = "#d1e7dd", "#0f5132"
        elif pct_int >= 90:
            bg_color, text_color = "#fff3cd", "#664d03"
        else:
            bg_color, text_color = "#f8d7da", "#842029"
            
        style = f"background-color: {bg_color}; color: {text_color}; font-weight: 700; padding: 3px 6px; border-radius: 4px; display: inline-block;"
        return text_display, style
    except:
        return val, ""

def generate_table_with_metadata(df, table_prefix):
    if df.empty:
        return '<tr><td colspan="100%" style="text-align:center; padding: 12px; color: #64748b;">No hay datos disponibles en esta hoja.</td></tr>', 0
    
    col_map = {}
    total_cantidad = 0
    for col in df.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean and "fecha" not in col_map: col_map["fecha"] = col
        elif "MES" in c_clean and "mes" not in col_map: col_map["mes"] = col
        elif ("NOMBRE" in c_clean) and "nombre" not in col_map: col_map["nombre"] = col
        elif ("PUNTO" in c_clean or "VENTA" in c_clean or "PDV" in c_clean or "CLIENTE" in c_clean) and "pv" not in col_map: col_map["pv"] = col
        elif ("REFERENCIA" in c_clean or "REF" in c_clean or "SKU" in c_clean) and "ref" not in col_map: col_map["ref"] = col
        elif "CANTIDAD" in c_clean and "cantidad" not in col_map: col_map["cantidad"] = col
        elif ("GEO" in c_clean or "UBICACION" in c_clean or "COORDENADA" in c_clean) and "geo" not in col_map: col_map["geo"] = col

    headers = "".join([f"<th>{col}</th>" for col in df.columns])
    rows = ""
    
    for idx, row in df.iterrows():
        f_val = str(row.get(col_map.get("fecha", ""), ""))
        m_val = str(row.get(col_map.get("mes", ""), ""))
        if not m_val and f_val and len(f_val) >= 7:
            m_val = f_val[:7]
            
        n_val = str(row.get(col_map.get("nombre", ""), ""))
        pv_val = str(row.get(col_map.get("pv", ""), ""))
        ref_val = str(row.get(col_map.get("ref", ""), ""))
        
        cant_val = 0
        if "cantidad" in col_map:
            try:
                cant_val = float(row.get(col_map["cantidad"], 0))
                total_cantidad += cant_val
            except:
                pass

        rows += f'''<tr data-fecha="{f_val}" data-mes="{m_val}" data-nombre="{n_val}" data-pv="{pv_val}" data-ref="{ref_val}" data-cantidad="{cant_val}">'''
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                val_str = ""
            else:
                val_str = str(val)
                if val_str.endswith(".0") and val_str.replace(".", "", 1).isdigit():
                    val_str = val_str[:-2]
            
            c_clean_col = strip_accents(col)
            if ("GEO" in c_clean_col or "UBICACION" in c_clean_col or "COORDENADA" in c_clean_col) and "," in val_str:
                coords = val_str.strip()
                map_id_str = f"map_generic_{table_prefix}_{idx}"
                cell_content = f'''<div style="display: flex; align-items: center; gap: 6px;">
                    <span style="font-size:11px; color:#0f2b48; font-weight:600;">{val_str}</span>
                    <div class="mini-map-card" onclick="openMapModal('{coords}')"><div id="{map_id_str}" class="mini-leaflet-map"></div><div class="mini-map-overlay-text">📍</div></div>
                </div><script>setTimeout(() => {{ if(typeof initMiniMap === "function") initMiniMap("{map_id_str}", "{coords}", 0.9); }}, 300);</script>'''
            elif ("SOPORTE" in c_clean_col and "FOTO" in c_clean_col) or (val_str.startswith("http") and "," in val_str):
                urls = [u.strip() for u in val_str.split(",") if u.strip().startswith("http")]
                if urls:
                    thumbs_html = '<div style="display: flex; gap: 4px; flex-wrap: wrap;">'
                    for u in urls:
                        thumbs_html += f'<div class="img-thumb-container"><img src="{u}" class="img-thumb" onclick="openImageModal(\'{u}\')" alt="Img"></div>'
                    thumbs_html += '</div>'
                    cell_content = thumbs_html
                else:
                    cell_content = val_str
            elif val_str.startswith("http"):
                cell_content = f'<div class="img-thumb-container"><img src="{val_str}" class="img-thumb" onclick="openImageModal(\'{val_str}\')" alt="Img"></div>'
            elif "OBSERVACION" in c_clean_col or "COMENTARIO" in c_clean_col:
                cell_content = f'<div style="white-space: normal; word-break: break-word; max-width: 300px; font-size: 11px;">{val_str}</div>'
            else:
                cell_content = val_str
                
            rows += f"<td>{cell_content}</td>"
        rows += "</tr>"
    
    table_html = f"<thead><tr>{headers}</tr></thead><tbody id='{table_prefix}TableBody'>{rows}</tbody>"
    return table_html, total_cantidad

def generate_asistencia_rows(df):
    if df.empty:
        return '<tr><td colspan="100%" style="text-align:center;">No hay datos disponibles en esta hoja.</td></tr>'
    
    col_map = {}
    for col in df.columns:
        c_clean = strip_accents(col)
        if "FECHA" in c_clean and "fecha" not in col_map: col_map["fecha"] = col
        elif "MES" in c_clean and "mes" not in col_map: col_map["mes"] = col
        elif "NOMBRE" in c_clean and "nombre" not in col_map: col_map["nombre"] = col
        elif c_clean == "CHECK IN": col_map["check_in"] = col
        elif "UBICACION" in c_clean and "CHECK IN" in c_clean: col_map["ubi_ci"] = col
        elif "IMAGEN" in c_clean and "CHECK IN" in c_clean: col_map["img_ci"] = col
        elif c_clean == "CHECK OUT": col_map["check_out"] = col
        elif "UBICACION" in c_clean and "CHECK OUT" in c_clean: col_map["ubi_co"] = col
        elif "IMAGEN" in c_clean and "CHECK OUT" in c_clean: col_map["img_co"] = col
        elif "ALMUERZO" in c_clean and "IN" in c_clean: col_map["alm_in"] = col
        elif "UBICACION" in c_clean and "BREAK" in c_clean and "IN" in c_clean: col_map["ubi_ai"] = col
        elif "ALMUERZO" in c_clean and "OUT" in c_clean: col_map["alm_out"] = col
        elif "UBICACION" in c_clean and "BREAK" in c_clean and "OUT" in c_clean: col_map["ubi_ao"] = col
        elif "ESTADO" in c_clean: col_map["estado"] = col

    rows_html = ""
    for row_idx, row in df.iterrows():
        is_inasistencia = False
        val_est = str(row.get(col_map.get("estado", ""), "")).strip()
        if "estado" in col_map and val_est.lower() == "inasistencia":
            is_inasistencia = True
        
        f_val = str(row.get(col_map.get("fecha", ""), ""))
        m_val = str(row.get(col_map.get("mes", ""), ""))
        if not m_val and f_val and len(f_val) >= 7:
            m_val = f_val[:7]
            
        n_val = str(row.get(col_map.get("nombre", ""), ""))
        
        row_style = 'style="background-color: #f8d7da; color: #842029;"' if is_inasistencia else ''
        rows_html += f'<tr data-fecha="{f_val}" data-mes="{m_val}" data-nombre="{n_val}" data-estado="{val_est}" {row_style}>'
        
        if " " in f_val: f_val = f_val.split()[0]
        rows_html += f'<td style="white-space: nowrap; font-size: 11px;">{f_val}</td>'
        rows_html += f'<td style="font-weight: 600; font-size: 11px; word-break: break-word;">{n_val}</td>'
        
        def render_img(val_img):
            if pd.notna(val_img) and str(val_img).startswith("http"):
                return f'<div class="img-thumb-container"><img src="{val_img}" class="img-thumb" onclick="openImageModal(\'{val_img}\')" alt="Img"></div>'
            return f'<span style="font-size:9px; color:#888;">{val_img}</span>'
        
        def render_map(val_coords, map_id_str):
            if pd.notna(val_coords) and "," in str(val_coords):
                coords = str(val_coords).strip()
                return f'<div class="mini-map-card" onclick="openMapModal(\'{coords}\')"><div id="{map_id_str}" class="mini-leaflet-map"></div><div class="mini-map-overlay-text">📍</div></div><script>setTimeout(() => {{ if(typeof initMiniMap === "function") initMiniMap("{map_id_str}", "{coords}", 0.9); }}, 300);</script>'
            return f'<span style="font-size:9px; color:#888;">{val_coords}</span>'

        v_ci = str(row.get(col_map.get("check_in", ""), "nan"))
        u_ci = row.get(col_map.get("ubi_ci", ""), "nan")
        i_ci = row.get(col_map.get("img_ci", ""), "nan")
        v_co = str(row.get(col_map.get("check_out", ""), "nan"))
        u_co = row.get(col_map.get("ubi_co", ""), "nan")
        i_co = row.get(col_map.get("img_co", ""), "nan")
        
        m_id_ci = f"m_ci_{row_idx}"
        m_id_co = f"m_co_{row_idx}"

        malla_html = f"""
        <td style="padding: 4px; vertical-align: middle;">
            <div style="display: flex; align-items: center; gap: 8px; justify-content: flex-start; white-space: nowrap;">
                <div style="display: flex; align-items: center; gap: 3px; background: rgba(0,0,0,0.02); padding: 2px 4px; border-radius: 3px;">
                    <span style="font-size: 10px; font-weight: 700; color: #0f2b48;">IN: {v_ci}</span>
                    {render_map(u_ci, m_id_ci)}
                    {render_img(i_ci)}
                </div>
                <div style="display: flex; align-items: center; gap: 3px; background: rgba(0,0,0,0.02); padding: 2px 4px; border-radius: 3px;">
                    <span style="font-size: 10px; font-weight: 700; color: #b94025;">OUT: {v_co}</span>
                    {render_map(u_co, m_id_co)}
                    {render_img(i_co)}
                </div>
            </div>
        </td>
        """
        rows_html += malla_html

        v_ai = str(row.get(col_map.get("alm_in", ""), "nan"))
        u_ai = row.get(col_map.get("ubi_ai", ""), "nan")
        v_ao = str(row.get(col_map.get("alm_out", ""), "nan"))
        u_ao = row.get(col_map.get("ubi_ao", ""), "nan")
        
        m_id_ai = f"m_ai_{row_idx}"
        m_id_ao = f"m_ao_{row_idx}"

        almuerzo_html = f"""
        <td style="padding: 4px; vertical-align: middle;">
            <div style="display: flex; align-items: center; gap: 8px; justify-content: flex-start; white-space: nowrap;">
                <div style="display: flex; align-items: center; gap: 3px; background: rgba(0,0,0,0.02); padding: 2px 4px; border-radius: 3px;">
                    <span style="font-size: 10px; font-weight: 700; color: #0f2b48;">IN: {v_ai}</span>
                    {render_map(u_ai, m_id_ai)}
                </div>
                <div style="display: flex; align-items: center; gap: 3px; background: rgba(0,0,0,0.02); padding: 2px 4px; border-radius: 3px;">
                    <span style="font-size: 10px; font-weight: 700; color: #b94025;">OUT: {v_ao}</span>
                    {render_map(u_ao, m_id_ao)}
                </div>
            </div>
        </td>
        """
        rows_html += almuerzo_html
        rows_html += f'<td style="font-size: 11px; font-weight: 700; text-align: center; white-space: nowrap;">{val_est}</td>'
        rows_html += "</tr>"
    return rows_html

main_table_rows = ""
for _, row in df_base.iterrows():
    f_val = str(row.get("Fecha", ""))
    m_val = str(row.get("Mes", ""))
    if not m_val and f_val and len(f_val) >= 7:
        m_val = f_val[:7]
    n_val = str(row.get("Nombre completo", ""))
    
    ast_text, ast_style = format_pct_with_color(row.get("Asistencia", ""))
    inv_text, inv_style = format_pct_with_color(row.get("Inventarios", ""))
    ven_text, ven_style = format_pct_with_color(row.get("Ventas", ""))
    exh_text, exh_style = format_pct_with_color(row.get("Exhibiciones", ""))
    vis_text, vis_style = format_pct_with_color(row.get("Visita Pocs", ""))
    
    main_table_rows += f"""
                    <tr data-fecha="{f_val}" data-mes="{m_val}" data-nombre="{n_val}">
                        <td>{f_val}</td>
                        <td>{n_val}</td>
                        <td><span style="{ast_style}">{ast_text}</span></td>
                        <td><span style="{inv_style}">{inv_text}</span></td>
                        <td><span style="{ven_style}">{ven_text}</span></td>
                        <td><span style="{exh_style}">{exh_text}</span></td>
                        <td><span style="{vis_style}">{vis_text}</span></td>
                    </tr>"""

asistencia_rows_html = generate_asistencia_rows(df_asistencia)
inventarios_table_html, inv_total_qty = generate_table_with_metadata(df_inventarios, "inventarios")
ventas_table_html, ven_total_qty = generate_table_with_metadata(df_ventas, "ventas")
exhibiciones_table_html, _ = generate_table_with_metadata(df_exhibiciones, "exhibiciones")
visita_pocs_table_html, _ = generate_table_with_metadata(df_visita_pocs, "visita_pocs")

html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SED International - Dashboard Único</title>
    <link href="https://fonts.cdnfonts.com/css/bahnschrift" rel="stylesheet">
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); }}
        body {{ background-color: #f4f7fb; color: #1c2d42; padding: 16px; font-size: 13px; width: 100%; min-height: 100vh; overflow-x: hidden; }}
        
        .container {{ width: 100%; max-width: 100%; margin: 0; padding: 0 10px; }}
        
        .top-header {{
            background: linear-gradient(135deg, #103254 0%, #1e5a8a 35%, #8b2332 75%, #b94025 100%);
            border-radius: 10px; padding: 16px 24px; color: white; margin-bottom: 16px;
            display: flex; align-items: center; justify-content: space-between;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15);
            animation: fadeInDown 0.6s cubic-bezier(0.16, 1, 0.3, 1);
            width: 100%;
        }}
        .header-title {{ font-size: 18px; font-weight: 700; letter-spacing: 0.5px; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        
        .tabs-nav {{
            display: flex; gap: 8px; margin-bottom: 16px; background: #ffffff; padding: 8px;
            border-radius: 10px; border: 1px solid #e2e8f0; box-shadow: 0 6px 15px rgba(0,0,0,0.05);
            width: 100%; flex-wrap: wrap;
            animation: fadeInUp 0.6s cubic-bezier(0.16, 1, 0.3, 1);
        }}
        .tab-link {{
            flex: 1; min-width: 130px; display: flex; align-items: center; justify-content: center; gap: 6px;
            padding: 10px 10px; border-radius: 6px; text-decoration: none; font-size: 11px; font-weight: 700;
            text-transform: uppercase; color: #0f2b48; background: #f8fafc; border: 1px solid #cbd5e1;
            white-space: nowrap; text-align: center; cursor: pointer;
            position: relative; overflow: hidden; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif;
        }}
        .tab-link::after {{
            content: ''; position: absolute; bottom: 0; left: 50%; width: 0; height: 3px; background: #0070ba; transition: width 0.3s ease, left 0.3s ease;
        }}
        .tab-link:hover {{ background: #e2e8f0; transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.08); }}
        .tab-link:hover::after {{ width: 80%; left: 10%; }}
        .tab-link.active {{ background: #0f2b48; color: white; border-color: #0f2b48; box-shadow: 0 6px 15px rgba(15,43,72,0.3); transform: translateY(-2px); }}
        .tab-link.active::after {{ background: #b94025; width: 100%; left: 0; }}
        
        .tab-content {{ display: none; opacity: 0; transform: translateY(12px); transition: opacity 0.4s ease, transform 0.4s ease; width: 100%; }}
        .tab-content.active {{ display: block; opacity: 1; transform: translateY(0); }}

        .filters-bar {{
            background: #ffffff; border-radius: 10px; padding: 14px;
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 16px;
            box-shadow: 0 6px 15px rgba(0,0,0,0.04); border: 1px solid #e2e8f0; width: 100%;
        }}
        .filters-bar-5 {{
            display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 16px;
            background: #ffffff; border-radius: 10px; padding: 14px; border: 1px solid #e2e8f0; box-shadow: 0 6px 15px rgba(0,0,0,0.04); width: 100%;
        }}
        .filters-bar-4 {{
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 16px;
            background: #ffffff; border-radius: 10px; padding: 14px; border: 1px solid #e2e8f0; box-shadow: 0 6px 15px rgba(0,0,0,0.04); width: 100%;
        }}
        .filter-item {{ display: flex; flex-direction: column; gap: 4px; width: 100%; }}
        .filter-label {{ font-size: 11px; font-weight: 700; color: #0f2b48; text-transform: uppercase; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        .filter-control {{
            width: 100%; padding: 9px 12px; border-radius: 6px; border: 1px solid #cbd5e1;
            background: #f8fafc; font-size: 12px; color: #1c2d42; outline: none; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif;
        }}
        .filter-control:focus {{ border-color: #0070ba; box-shadow: 0 0 0 3px rgba(0,112,186,0.15); background: #ffffff; }}

        .kpi-container {{
            display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 16px; width: 100%;
        }}
        .kpi-card {{
            background: #ffffff; padding: 16px; border-radius: 10px; border: 1px solid #e2e8f0;
            box-shadow: 0 6px 15px rgba(0,0,0,0.04); text-align: center; text-decoration: none; display: block; cursor: pointer;
            position: relative; overflow: hidden;
        }}
        .kpi-card::before {{
            content: ''; position: absolute; top: 0; left: 0; width: 4px; height: 100%; background: #0070ba; transition: width 0.3s;
        }}
        .kpi-card:hover {{ border-color: #0070ba; transform: translateY(-4px); box-shadow: 0 10px 22px rgba(0,112,186,0.15); }}
        .kpi-card:hover::before {{ width: 100%; opacity: 0.05; }}
        .kpi-title {{ font-size: 11px; font-weight: 700; color: #0f2b48; text-transform: uppercase; margin-bottom: 6px; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        .kpi-value {{ font-size: 18px; font-weight: 700; color: #0070ba; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        
        .section-box {{ background: #ffffff; border-radius: 10px; padding: 16px; margin-bottom: 16px; box-shadow: 0 6px 15px rgba(0,0,0,0.04); border: 1px solid #e2e8f0; width: 100%; }}
        .section-title {{ font-size: 15px; font-weight: 700; color: #0f2b48; margin-bottom: 12px; border-bottom: 2px solid #edf2f7; padding-bottom: 8px; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        
        .table-container {{ background: #ffffff; border-radius: 10px; box-shadow: 0 6px 15px rgba(0,0,0,0.04); overflow-x: auto; border: 1px solid #e2e8f0; width: 100%; max-height: 70vh; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; white-space: nowrap; }}
        th {{ background-color: #0f2b48; color: #ffffff; padding: 11px 14px; font-weight: 700; text-transform: uppercase; font-size: 11px; position: sticky; top: 0; z-index: 10; letter-spacing: 0.5px; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        td {{ padding: 11px 14px; border-bottom: 1px solid #edf2f7; vertical-align: middle; font-family: 'Bahnschrift', 'Bahnschrift Light', sans-serif; }}
        tbody tr {{ transition: background-color 0.2s ease, transform 0.2s ease; }}
        tbody tr:hover {{ background-color: #f1f5f9; }}
        
        .img-thumb-container {{ width: 42px; height: 42px; border-radius: 6px; overflow: hidden; border: 1px solid #cbd5e1; cursor: pointer; display: inline-block; flex-shrink: 0; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }}
        .img-thumb {{ width: 100%; height: 100%; object-fit: cover; transition: transform 0.4s cubic-bezier(0.16, 1, 0.3, 1); }}
        .img-thumb:hover {{ transform: scale(1.2); }}
        
        .mini-map-card {{
            width: 42px; height: 42px; border-radius: 6px; border: 1px solid #cbd5e1;
            cursor: pointer; position: relative; overflow: hidden; background: #e2e8f0; flex-shrink: 0;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1); display: inline-block; box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }}
        .mini-map-card:hover {{ border-color: #0070ba; transform: scale(1.12); box-shadow: 0 6px 16px rgba(0,0,0,0.25); }}
        .mini-leaflet-map {{ width: 100%; height: 100%; pointer-events: none; }}
        .mini-map-overlay-text {{
            position: absolute; bottom: 2px; right: 2px; background: rgba(15, 43, 72, 0.9); color: white;
            font-size: 6px; padding: 1px 3px; border-radius: 3px; z-index: 400; font-weight: 700;
        }}

        .modal {{
            display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%;
            background-color: rgba(0,0,0,0.8); align-items: center; justify-content: center;
            opacity: 0; transition: opacity 0.4s cubic-bezier(0.16, 1, 0.3, 1); backdrop-filter: blur(5px);
        }}
        .modal.show {{ display: flex; opacity: 1; }}
        .modal-content {{
            background: #ffffff; padding: 22px; border-radius: 12px; width: 90%; max-width: 650px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.4); transform: scale(0.85); transition: transform 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275); position: relative;
        }}
        .modal.show .modal-content {{ transform: scale(1); }}
        .close-modal {{
            position: absolute; right: 16px; top: 14px; font-size: 22px; font-weight: bold; color: #64748b; cursor: pointer; z-index: 2010;
            transition: color 0.2s, transform 0.2s;
        }}
        .close-modal:hover {{ color: #b94025; transform: rotate(90deg); }}
        #fullMap {{ width: 100%; height: 420px; border-radius: 8px; margin-top: 10px; border: 1px solid #cbd5e1; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05); }}
        #fullImage {{ width: 100%; max-height: 70vh; object-fit: contain; border-radius: 8px; margin-top: 10px; cursor: zoom-in; transition: transform 0.3s ease; }}
        #fullImage.zoomed {{ transform: scale(1.6); cursor: zoom-out; }}

        @keyframes fadeInDown {{
            from {{ opacity: 0; transform: translateY(-20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        @keyframes fadeInUp {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header class="top-header">
            <h1 class="header-title">SED International - Dashboard Integrado</h1>
        </header>

        <div class="tabs-nav">
            <a class="tab-link active" onclick="switchTab('indicadores', this)">{icons["general"]} Indicadores</a>
            <a class="tab-link" onclick="switchTab('asistencia', this)">{icons["asistencia"]} Asistencia</a>
            <a class="tab-link" onclick="switchTab('inventarios', this)">{icons["inventarios"]} Inventarios</a>
            <a class="tab-link" onclick="switchTab('ventas', this)">{icons["ventas"]} Ventas</a>
            <a class="tab-link" onclick="switchTab('exhibiciones', this)">{icons["exhibiciones"]} Exhibiciones</a>
            <a class="tab-link" onclick="switchTab('visita_pocs', this)">{icons["visita_pocs"]} Visita Pocs</a>
        </div>

        <!-- SECCIÓN: INDICADORES -->
        <div id="tab-indicadores" class="tab-content active">
            <div class="filters-bar">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="filtroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="filtroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre:</label>
                    <select id="filtroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in nombres])}
                    </select>
                </div>
            </div>

            <div class="kpi-container">
                <div class="kpi-card" onclick="switchTab('asistencia', document.querySelectorAll('.tab-link')[1])">
                    <div class="kpi-title">Asistencia</div>
                    <div class="kpi-value">{asistencia_val}%</div>
                </div>
                <div class="kpi-card" onclick="switchTab('inventarios', document.querySelectorAll('.tab-link')[2])">
                    <div class="kpi-title">Inventarios</div>
                    <div class="kpi-value">{inventarios_val}%</div>
                </div>
                <div class="kpi-card" onclick="switchTab('ventas', document.querySelectorAll('.tab-link')[3])">
                    <div class="kpi-title">Ventas</div>
                    <div class="kpi-value">{ventas_val}%</div>
                </div>
                <div class="kpi-card" onclick="switchTab('exhibiciones', document.querySelectorAll('.tab-link')[4])">
                    <div class="kpi-title">Exhibiciones</div>
                    <div class="kpi-value">{exhibiciones_val}%</div>
                </div>
                <div class="kpi-card" onclick="switchTab('visita_pocs', document.querySelectorAll('.tab-link')[5])">
                    <div class="kpi-title">Visita Pocs</div>
                    <div class="kpi-value">{visita_pocs_val}%</div>
                </div>
            </div>

            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Nombre Completo</th>
                            <th>Asistencia</th>
                            <th>Inventarios</th>
                            <th>Ventas</th>
                            <th>Exhibiciones</th>
                            <th>Visita Pocs</th>
                        </tr>
                    </thead>
                    <tbody id="mainTableBody">
                        {main_table_rows}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- SECCIÓN: ASISTENCIA -->
        <div id="tab-asistencia" class="tab-content">
            <div class="filters-bar-4">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="astFiltroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in ast_fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="astFiltroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in ast_meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre:</label>
                    <select id="astFiltroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in ast_nombres])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Estado:</label>
                    <select id="astFiltroEstado" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{e}">{e}</option>' for e in ast_estados])}
                    </select>
                </div>
            </div>
            <div class="section-box">
                <div class="section-title">Asistencia</div>
                <div class="table-container">
                    <table>
                        <thead>
                            <tr>
                                <th>Fecha</th>
                                <th>Nombre</th>
                                <th>Malla</th>
                                <th>Almuerzo</th>
                                <th style="text-align: center;">Estado</th>
                            </tr>
                        </thead>
                        <tbody id="asistenciaTableBody">
                            {asistencia_rows_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- SECCIÓN: INVENTARIOS -->
        <div id="tab-inventarios" class="tab-content">
            <div class="filters-bar-5">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="inventariosFiltroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in inv_fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="inventariosFiltroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in inv_meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre completo:</label>
                    <select id="inventariosFiltroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in inv_nombres])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Punto de Venta:</label>
                    <select id="inventariosFiltroPV" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{pv}">{pv}</option>' for pv in inv_pvs])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Referencias:</label>
                    <select id="inventariosFiltroRef" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{ref}">{ref}</option>' for ref in inv_refs])}
                    </select>
                </div>
            </div>

            <div class="kpi-container" style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); max-width: 320px;">
                <div class="kpi-card" style="cursor: default;">
                    <div class="kpi-title">Suma Total Cantidad</div>
                    <div class="kpi-value" id="inventariosTotalCantidad">{inv_total_qty:,.0f}</div>
                </div>
            </div>

            <div class="section-box">
                <div class="section-title">Inventarios</div>
                <div class="table-container">
                    <table>
                        {inventarios_table_html}
                    </table>
                </div>
            </div>
        </div>

        <!-- SECCIÓN: VENTAS -->
        <div id="tab-ventas" class="tab-content">
            <div class="filters-bar-5">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="ventasFiltroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in ven_fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="ventasFiltroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in ven_meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre completo:</label>
                    <select id="ventasFiltroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in ven_nombres])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Punto de Venta:</label>
                    <select id="ventasFiltroPV" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{pv}">{pv}</option>' for pv in ven_pvs])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Referencias:</label>
                    <select id="ventasFiltroRef" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{ref}">{ref}</option>' for ref in ven_refs])}
                    </select>
                </div>
            </div>

            <div class="kpi-container" style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); max-width: 320px;">
                <div class="kpi-card" style="cursor: default;">
                    <div class="kpi-title">Suma Total Cantidad</div>
                    <div class="kpi-value" id="ventasTotalCantidad">{ven_total_qty:,.0f}</div>
                </div>
            </div>

            <div class="section-box">
                <div class="section-title">Ventas</div>
                <div class="table-container">
                    <table>
                        {ventas_table_html}
                    </table>
                </div>
            </div>
        </div>

        <!-- SECCIÓN: EXHIBICIONES -->
        <div id="tab-exhibiciones" class="tab-content">
            <div class="filters-bar-4">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="exhibicionesFiltroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in exh_fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="exhibicionesFiltroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in exh_meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre completo:</label>
                    <select id="exhibicionesFiltroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in exh_nombres])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Punto de Venta:</label>
                    <select id="exhibicionesFiltroPV" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{pv}">{pv}</option>' for pv in exh_pvs])}
                    </select>
                </div>
            </div>

            <div class="section-box">
                <div class="section-title">Exhibiciones</div>
                <div class="table-container">
                    <table>
                        {exhibiciones_table_html}
                    </table>
                </div>
            </div>
        </div>

        <!-- SECCIÓN: VISITA POCS -->
        <div id="tab-visita_pocs" class="tab-content">
            <div class="filters-bar-4">
                <div class="filter-item">
                    <label class="filter-label">Fecha:</label>
                    <select id="visita_pocsFiltroFecha" class="filter-control">
                        <option value="">Todas</option>
                        {"".join([f'<option value="{f}">{f}</option>' for f in vis_fechas])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Mes:</label>
                    <select id="visita_pocsFiltroMes" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{m}">{m}</option>' for m in vis_meses])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Nombre completo:</label>
                    <select id="visita_pocsFiltroNombre" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{n}">{n}</option>' for n in vis_nombres])}
                    </select>
                </div>
                <div class="filter-item">
                    <label class="filter-label">Punto de Venta:</label>
                    <select id="visita_pocsFiltroPV" class="filter-control">
                        <option value="">Todos</option>
                        {"".join([f'<option value="{pv}">{pv}</option>' for pv in vis_pvs])}
                    </select>
                </div>
            </div>

            <div class="section-box">
                <div class="section-title">Visita Pocs</div>
                <div class="table-container">
                    <table>
                        {visita_pocs_table_html}
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal para Mapa -->
    <div id="mapModal" class="modal">
        <div class="modal-content">
            <span class="close-modal" onclick="closeMapModal()">&times;</span>
            <h3 style="font-size: 15px; color: #0f2b48; margin-bottom: 6px; font-weight: 700;">Ubicación Geográfica</h3>
            <div id="fullMap"></div>
        </div>
    </div>

    <!-- Modal para Imagen -->
    <div id="imageModal" class="modal">
        <div class="modal-content" style="text-align: center;">
            <span class="close-modal" onclick="closeImageModal()">&times;</span>
            <h3 style="font-size: 15px; color: #0f2b48; margin-bottom: 6px; font-weight: 700;">Visualización de Imagen</h3>
            <div style="overflow: hidden; border-radius: 8px; display: inline-block; width: 100%;">
                <img id="fullImage" src="" alt="Imagen ampliada" onclick="toggleImageZoom(this)">
            </div>
            <p style="font-size: 11px; color: #64748b; margin-top: 6px;">Haz clic sobre la imagen para hacer zoom</p>
        </div>
    </div>

    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
        function switchTab(tabId, element) {{
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-link').forEach(el => el.classList.remove('active'));
            
            document.getElementById('tab-' + tabId).classList.add('active');
            if(element) element.classList.add('active');

            setTimeout(() => {{
                window.dispatchEvent(new Event('resize'));
            }}, 200);
        }}

        document.addEventListener('DOMContentLoaded', function() {{
            // Filtros Indicadores y actualización dinámica de KPIs
            const filtroFecha = document.getElementById('filtroFecha');
            const filtroMes = document.getElementById('filtroMes');
            const filtroNombre = document.getElementById('filtroNombre');
            const mainRows = document.querySelectorAll('#mainTableBody tr');
            const kpiValues = document.querySelectorAll('#tab-indicadores .kpi-card .kpi-value');

            function filterMainTable() {{
                const fFecha = (filtroFecha.value || '').toLowerCase().trim();
                const fMes = (filtroMes.value || '').toLowerCase().trim();
                const fNombre = (filtroNombre.value || '').toLowerCase().trim();

                let count = 0;
                let sumAst = 0, sumInv = 0, sumVen = 0, sumExh = 0, sumVis = 0;

                mainRows.forEach(row => {{
                    const fecha = (row.getAttribute('data-fecha') || '').toLowerCase().trim();
                    const mes = (row.getAttribute('data-mes') || '').toLowerCase().trim();
                    const nombre = (row.getAttribute('data-nombre') || '').toLowerCase().trim();

                    const matchFecha = !fFecha || fecha === fFecha;
                    const matchMes = !fMes || mes === fMes;
                    const matchNombre = !fNombre || nombre === fNombre;

                    if (matchFecha && matchMes && matchNombre) {{
                        row.style.display = '';
                        count++;

                        const cells = row.querySelectorAll('td');
                        if (cells.length >= 7) {{
                            const parseVal = (cell) => {{
                                const txt = cell.textContent.replace('%', '').trim();
                                const val = parseFloat(txt);
                                return isNaN(val) ? 0 : val;
                            }};
                            sumAst += parseVal(cells[2]);
                            sumInv += parseVal(cells[3]);
                            sumVen += parseVal(cells[4]);
                            sumExh += parseVal(cells[5]);
                            sumVis += parseVal(cells[6]);
                        }}
                    }} else {{
                        row.style.display = 'none';
                    }}
                }});

                if (count > 0) {{
                    if(kpiValues[0]) kpiValues[0].textContent = Math.round(sumAst / count) + '%';
                    if(kpiValues[1]) kpiValues[1].textContent = Math.round(sumInv / count) + '%';
                    if(kpiValues[2]) kpiValues[2].textContent = Math.round(sumVen / count) + '%';
                    if(kpiValues[3]) kpiValues[3].textContent = Math.round(sumExh / count) + '%';
                    if(kpiValues[4]) kpiValues[4].textContent = Math.round(sumVis / count) + '%';
                }} else {{
                    kpiValues.forEach(el => el.textContent = '0%');
                }}
            }}

            if(filtroFecha) filtroFecha.addEventListener('change', filterMainTable);
            if(filtroMes) filtroMes.addEventListener('change', filterMainTable);
            if(filtroNombre) filtroNombre.addEventListener('change', filterMainTable);

            // Filtros Asistencia en Cascada cruzada
            function setupAsistenciaCrossFilter() {{
                const astFecha = document.getElementById('astFiltroFecha');
                const astMes = document.getElementById('astFiltroMes');
                const astNombre = document.getElementById('astFiltroNombre');
                const astEstado = document.getElementById('astFiltroEstado');
                const astRows = document.querySelectorAll('#asistenciaTableBody tr');
                const selects = [astFecha, astMes, astNombre, astEstado];

                function updateAstDropdownOptions(activeSelect) {{
                    selects.forEach(sel => {{
                        if (!sel || sel === activeSelect) return;
                        const selKey = sel.id.replace('astFiltro', '').toLowerCase();
                        const validOptions = new Set();

                        astRows.forEach(row => {{
                            const rFecha = row.getAttribute('data-fecha') || '';
                            const rMes = row.getAttribute('data-mes') || '';
                            const rNombre = row.getAttribute('data-nombre') || '';
                            const rEstado = row.getAttribute('data-estado') || '';

                            const matchF = !astFecha.value || rFecha === astFecha.value;
                            const matchM = !astMes.value || rMes === astMes.value;
                            const matchN = !astNombre.value || rNombre === astNombre.value;
                            const matchE = !astEstado.value || rEstado === astEstado.value;

                            let passOthers = true;
                            if (selKey === 'fecha') passOthers = (matchM && matchN && matchE);
                            if (selKey === 'mes') passOthers = (matchF && matchN && matchE);
                            if (selKey === 'nombre') passOthers = (matchF && matchM && matchE);
                            if (selKey === 'estado') passOthers = (matchF && matchM && matchN);

                            if (passOthers) {{
                                if (selKey === 'fecha' && rFecha) validOptions.add(rFecha);
                                if (selKey === 'mes' && rMes) validOptions.add(rMes);
                                if (selKey === 'nombre' && rNombre) validOptions.add(rNombre);
                                if (selKey === 'estado' && rEstado) validOptions.add(rEstado);
                            }}
                        }});

                        Array.from(sel.options).forEach((opt, idx) => {{
                            if (idx === 0) return;
                            if (validOptions.has(opt.value)) {{
                                opt.hidden = false;
                            }} else {{
                                opt.hidden = true;
                                if (sel.value === opt.value) sel.value = "";
                            }}
                        }});
                    }});
                }}

                function applyAstFilters(changedSelect) {{
                    updateAstDropdownOptions(changedSelect);

                    const fFecha = (astFecha.value || '').toLowerCase().trim();
                    const fMes = (astMes.value || '').toLowerCase().trim();
                    const fNombre = (astNombre.value || '').toLowerCase().trim();
                    const fEstado = (astEstado.value || '').toLowerCase().trim();

                    astRows.forEach(row => {{
                        const fecha = (row.getAttribute('data-fecha') || '').toLowerCase().trim();
                        const mes = (row.getAttribute('data-mes') || '').toLowerCase().trim();
                        const nombre = (row.getAttribute('data-nombre') || '').toLowerCase().trim();
                        const estado = (row.getAttribute('data-estado') || '').toLowerCase().trim();

                        const matchFecha = !fFecha || fecha.includes(fFecha);
                        const matchMes = !fMes || mes === fMes;
                        const matchNombre = !fNombre || nombre === fNombre;
                        const matchEstado = !fEstado || estado === fEstado;

                        row.style.display = (matchFecha && matchMes && matchNombre && matchEstado) ? '' : 'none';
                    }});
                }}

                selects.forEach(sel => {{
                    if(sel) sel.addEventListener('change', () => applyAstFilters(sel));
                }});
            }}
            setupAsistenciaCrossFilter();

            // Motor general de Filtros en Cascada para Inventarios, Ventas, Exhibiciones y Visita Pocs
            function setupCrossFilter(prefix, hasRef = false) {{
                const fFecha = document.getElementById(prefix + 'FiltroFecha');
                const fMes = document.getElementById(prefix + 'FiltroMes');
                const fNombre = document.getElementById(prefix + 'FiltroNombre');
                const fPV = document.getElementById(prefix + 'FiltroPV');
                const fRef = hasRef ? document.getElementById(prefix + 'FiltroRef') : null;
                const rows = document.querySelectorAll('#' + prefix + 'TableBody tr');
                const totalQtyElem = document.getElementById(prefix + 'TotalCantidad');

                const selects = hasRef ? [fFecha, fMes, fNombre, fPV, fRef] : [fFecha, fMes, fNombre, fPV];

                function updateDropdownOptions(activeChangedSelect) {{
                    if (activeChangedSelect === fFecha && fFecha && fFecha.value !== "") {{
                        const matchedRow = Array.from(rows).find(r => r.getAttribute('data-fecha') === fFecha.value);
                        if (matchedRow && fMes) {{
                            fMes.value = matchedRow.getAttribute('data-mes');
                        }}
                    }}

                    selects.forEach(sel => {{
                        if (!sel || sel === activeChangedSelect) return;
                        const selKey = sel.id.replace(prefix + 'Filtro', '').toLowerCase();
                        const validOptions = new Set();

                        rows.forEach(row => {{
                            const rFecha = row.getAttribute('data-fecha') || '';
                            const rMes = row.getAttribute('data-mes') || '';
                            const rNombre = row.getAttribute('data-nombre') || '';
                            const rPV = row.getAttribute('data-pv') || '';
                            const rRef = hasRef ? (row.getAttribute('data-ref') || '') : '';

                            const matchF = !fFecha || !fFecha.value || rFecha === fFecha.value;
                            const matchM = !fMes || !fMes.value || rMes === fMes.value;
                            const matchN = !fNombre || !fNombre.value || rNombre === fNombre.value;
                            const matchP = !fPV || !fPV.value || rPV === fPV.value;
                            const matchR = !hasRef || !fRef || !fRef.value || rRef === fRef.value;

                            let passOthers = true;
                            if (selKey === 'fecha') passOthers = (matchM && matchN && matchP && matchR);
                            if (selKey === 'mes') passOthers = (matchF && matchN && matchP && matchR);
                            if (selKey === 'nombre') passOthers = (matchF && matchM && matchP && matchR);
                            if (selKey === 'pv') passOthers = (matchF && matchM && matchN && matchR);
                            if (hasRef && selKey === 'ref') passOthers = (matchF && matchM && matchN && matchP);

                            if (passOthers) {{
                                if (selKey === 'fecha' && rFecha) validOptions.add(rFecha);
                                if (selKey === 'mes' && rMes) validOptions.add(rMes);
                                if (selKey === 'nombre' && rNombre) validOptions.add(rNombre);
                                if (selKey === 'pv' && rPV) validOptions.add(rPV);
                                if (hasRef && selKey === 'ref' && rRef) validOptions.add(rRef);
                            }}
                        }});

                        Array.from(sel.options).forEach((opt, idx) => {{
                            if (idx === 0) return;
                            if (validOptions.has(opt.value)) {{
                                opt.hidden = false;
                            }} else {{
                                opt.hidden = true;
                                if (sel.value === opt.value) sel.value = "";
                            }}
                        }});
                    }});
                }}

                function applyFilters(changedSelect) {{
                    updateDropdownOptions(changedSelect);

                    const valFecha = fFecha ? (fFecha.value || '').toLowerCase().trim() : '';
                    const valMes = fMes ? (fMes.value || '').toLowerCase().trim() : '';
                    const valNombre = fNombre ? (fNombre.value || '').toLowerCase().trim() : '';
                    const valPV = fPV ? (fPV.value || '').toLowerCase().trim() : '';
                    const valRef = hasRef && fRef ? (fRef.value || '').toLowerCase().trim() : '';

                    let runningTotal = 0;

                    rows.forEach(row => {{
                        const fecha = (row.getAttribute('data-fecha') || '').toLowerCase().trim();
                        const mes = (row.getAttribute('data-mes') || '').toLowerCase().trim();
                        const nombre = (row.getAttribute('data-nombre') || '').toLowerCase().trim();
                        const pv = (row.getAttribute('data-pv') || '').toLowerCase().trim();
                        const ref = hasRef ? (row.getAttribute('data-ref') || '').toLowerCase().trim() : '';
                        const cantidad = parseFloat(row.getAttribute('data-cantidad') || '0');

                        const matchFecha = !valFecha || fecha === valFecha;
                        const matchMes = !valMes || mes === valMes;
                        const matchNombre = !valNombre || nombre === valNombre;
                        const matchPV = !valPV || pv === valPV;
                        const matchRef = !hasRef || !valRef || ref === valRef;

                        if (matchFecha && matchMes && matchNombre && matchPV && matchRef) {{
                            row.style.display = '';
                            runningTotal += cantidad;
                        }} else {{
                            row.style.display = 'none';
                        }}
                    }});

                    if(totalQtyElem) {{
                        totalQtyElem.textContent = runningTotal.toLocaleString('en-US', {{maximumFractionDigits: 2}});
                    }}
                }}

                selects.forEach(sel => {{
                    if(sel) {{
                        sel.addEventListener('change', () => applyFilters(sel));
                    }}
                }});
            }}

            setupCrossFilter('inventarios', true);
            setupCrossFilter('ventas', true);
            setupCrossFilter('exhibiciones', false);
            setupCrossFilter('visita_pocs', false);
        }});

        let modalMapInstance = null;
        const activeMiniMaps = {{}};

        function initMiniMap(mapId, coordsStr, zoomFactor = 0.9) {{
            const parts = coordsStr.split(',');
            if(parts.length !== 2) return;
            const lat = parseFloat(parts[0].trim());
            const lon = parseFloat(parts[1].trim());

            if(document.getElementById(mapId) && !activeMiniMaps[mapId]) {{
                const map = L.map(mapId, {{
                    zoomControl: false,
                    attributionControl: false,
                    dragging: false,
                    scrollWheelZoom: false,
                    doubleClickZoom: false
                }}).setView([lat, lon], 14);

                L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                    maxZoom: 19,
                    subdomains: 'abcd',
                    attribution: '&copy; OpenStreetMap & CARTO'
                }}).addTo(map);

                L.circleMarker([lat, lon], {{
                    radius: 4,
                    color: '#0f2b48',
                    fillColor: '#b94025',
                    fillOpacity: 0.9,
                    weight: 1
                }}).addTo(map);

                activeMiniMaps[mapId] = map;
            }}
        }}

        function openMapModal(coordsStr) {{
            const modal = document.getElementById('mapModal');
            modal.classList.add('show');

            const parts = coordsStr.split(',');
            if(parts.length === 2) {{
                const lat = parseFloat(parts[0].trim());
                const lon = parseFloat(parts[1].trim());

                setTimeout(() => {{
                    if(!modalMapInstance) {{
                        modalMapInstance = L.map('fullMap').setView([lat, lon], 15);
                        L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                            maxZoom: 19,
                            subdomains: 'abcd',
                            attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
                        }}).addTo(modalMapInstance);
                    }} else {{
                        modalMapInstance.setView([lat, lon], 15);
                        modalMapInstance.invalidateSize();
                    }}
                    
                    modalMapInstance.eachLayer((layer) => {{
                        if (layer instanceof L.Marker || layer instanceof L.CircleMarker) {{
                            modalMapInstance.removeLayer(layer);
                        }}
                    }});

                    L.marker([lat, lon]).addTo(modalMapInstance)
                        .bindPopup("Coordenadas: " + coordsStr)
                        .openPopup();
                }}, 250);
            }}
        }}

        function closeMapModal() {{
            document.getElementById('mapModal').classList.remove('show');
        }}

        function openImageModal(imgSrc) {{
            const modal = document.getElementById('imageModal');
            const img = document.getElementById('fullImage');
            img.classList.remove('zoomed');
            img.src = imgSrc;
            img.style.transform = "scale(0.9)";
            modal.classList.add('show');
        }}

        function closeImageModal() {{
            const modal = document.getElementById('imageModal');
            const img = document.getElementById('fullImage');
            img.classList.remove('zoomed');
            img.style.transform = "none";
            modal.classList.remove('show');
        }}

        function toggleImageZoom(img) {{
            if (img.classList.contains('zoomed')) {{
                img.classList.remove('zoomed');
                img.style.transform = "scale(0.9)";
            }} else {{
                img.classList.add('zoomed');
                img.style.transform = "scale(1.4)";
            }}
        }}

        window.onclick = function(event) {{
            const mapModal = document.getElementById('mapModal');
            const imgModal = document.getElementById('imageModal');
            if (event.target === mapModal) closeMapModal();
            if (event.target === imgModal) closeImageModal();
        }}
    </script>
</body>
</html>
"""

os.makedirs(base_dir, exist_ok=True)
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_content)