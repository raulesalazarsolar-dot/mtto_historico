import io
import base64
import json
import os
import unicodedata
from urllib.parse import urlparse, unquote
from datetime import datetime
from zoneinfo import ZoneInfo
from PIL import Image

# Librerías para cruce de datos
import pandas as pd

# Librerías de SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN (SHAREPOINT + GITHUB)
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/EquipoPlanificacin"
LIST_NAME = "Seguimiento Infraestructura"

USERNAME = os.environ.get("SP_USERNAME", "r0r0noi@cl.wal-mart.com")
PASSWORD = os.environ.get("SP_PASSWORD", "dEbit.spLiT+9")

OUTPUT_HTML = "index.html"

# ==========================================
# 2. PROCESAMIENTO EXCEL (PLAN MATRIZ SAP)
# ==========================================
def procesar_bases_sap():
    print("⏳ Leyendo Excel de Equipos y OTs SAP para el Plan Matriz...")
    try:
        df_eq = pd.read_excel('Copia de EVR-06-01 Equipos Críticos Planta Masas (5).xlsx', 
                              sheet_name='Clasificación ABC', 
                              usecols=[1, 2, 3, 6, 14])
        df_eq.columns = ['Eq', 'Nom', 'UbiSAP', 'UbiSP', 'Clase']

        def limpiar_codigo(codigo):
            c = str(codigo).strip().upper()
            if c in ['NAN', 'NONE', 'NAT', '']: return 'SIN_EQUIPO'
            if c.endswith('.0'): c = c[:-2]
            return c.lstrip('0')

        df_eq['Eq'] = df_eq['Eq'].apply(limpiar_codigo)
        df_eq['Nom'] = df_eq['Nom'].astype(str).str.strip()
        df_eq['UbiSAP'] = df_eq['UbiSAP'].astype(str).str.strip()
        df_eq['UbiSP'] = df_eq['UbiSP'].astype(str).str.strip()
        df_eq['Clase'] = df_eq['Clase'].astype(str).str.strip().str.upper()

        mapa_sp_sap = {}
        for _, r in df_eq.iterrows():
            usp = r['UbiSP'].upper()
            usap = r['UbiSAP']
            if usp not in ['NAN', '', 'NONE'] and usap not in ['NAN', '', 'NONE']:
                mapa_sp_sap[usp] = usap

        mapa_eq = df_eq[df_eq['Clase'].isin(['A','B','C'])].drop_duplicates(subset=['Eq'])
        mapa_eq = mapa_eq[mapa_eq['Eq'] != 'SIN_EQUIPO']

        df_ot = pd.read_excel('OT 2025 A 22.07.26.xlsx', dtype={'Equipo': str})
        df_ot['Eq'] = df_ot['Equipo'].apply(limpiar_codigo)
        df_ot_cruzado = pd.merge(df_ot, mapa_eq, on='Eq', how='inner')

        df_ot_cruzado['Fecha'] = pd.to_datetime(df_ot_cruzado['Inic.prog.'], errors='coerce', dayfirst=True)
        df_ot_cruzado = df_ot_cruzado.dropna(subset=['Fecha'])
        df_ot_cruzado['Semana'] = df_ot_cruzado['Fecha'].dt.isocalendar().week.astype(int)
        df_ot_cruzado['Año'] = df_ot_cruzado['Fecha'].dt.year.astype(int)
        df_ot_cruzado = df_ot_cruzado[df_ot_cruzado['Año'].isin([2025, 2026])]

        def det_st(st):
            s = str(st).upper()
            if 'CERR' in s or 'CTEC' in s: return 'realizada'
            if 'LIB.' in s or 'TRAB' in s: return 'en proceso'
            return 'pendiente'

        df_ot_cruzado['status_limpio'] = df_ot_cruzado['Stat.sist.'].apply(det_st)

        matriz_sap = { '2025': {'A':{}, 'B':{}, 'C':{}}, '2026': {'A':{}, 'B':{}, 'C':{}} }

        for _, r in mapa_eq.iterrows():
            c = r['Clase']
            u = str(r['UbiSAP']).strip()
            if u.lower() in ['nan', 'none', '']: u = 'Sin Ubicación Asignada'
            
            nom = str(r['Nom'])
            eq_disp = f"{r['Eq']} - {nom}" if nom.lower() != 'nan' else str(r['Eq'])

            for a in ['2025', '2026']:
                if u not in matriz_sap[a][c]: matriz_sap[a][c][u] = {}
                if eq_disp not in matriz_sap[a][c][u]: matriz_sap[a][c][u][eq_disp] = {}

        for _, row in df_ot_cruzado.iterrows():
            a = str(row['Año'])
            c = row['Clase']
            
            u = str(row['UbiSAP']).strip()
            if u.lower() in ['nan', 'none', '']:
                u_ot = str(row.get('Ubicación técnica', '')).strip()
                denom = str(row.get('Denom.ubic.técnica', '')).strip()
                if u_ot and u_ot.lower() != 'nan': u = u_ot
                elif denom and denom.lower() != 'nan': u = denom
                else: u = 'Sin Ubicación Asignada'
                
            nom = str(row['Nom'])
            eq_disp = f"{row['Eq']} - {nom}" if nom.lower() != 'nan' else str(row['Eq'])
            sem = str(row['Semana'])

            if u not in matriz_sap[a][c]: matriz_sap[a][c][u] = {}
            if eq_disp not in matriz_sap[a][c][u]: matriz_sap[a][c][u][eq_disp] = {}
            if sem not in matriz_sap[a][c][u][eq_disp]: matriz_sap[a][c][u][eq_disp][sem] = []

            matriz_sap[a][c][u][eq_disp][sem].append({
                "id": str(row['Orden']),
                "ot": str(row['Orden']),
                "titulo": str(row['Texto breve']).replace('"', '').replace("'", ""),
                "fecha_prog": row['Fecha'].strftime('%Y-%m-%d'),
                "status": row['status_limpio']
            })

        print("✅ Base SAP procesada. Mapeo SP->SAP creado exitosamente.")
        return matriz_sap, mapa_sp_sap

    except Exception as e:
        print(f"❌ Error procesando Excel SAP: {e}")
        import traceback
        traceback.print_exc()
        return {}, {}

# ==========================================
# 3. UTILIDADES Y "SABUESO DE FOTOS"
# ==========================================
def limpiar(val):
    if val is None: return ""
    s = str(val).strip()
    if s == "0" or s == "0.0": return "0"
    if s.lower() == "nan": return "" 
    return s.replace(".0", "")

def normalizar_texto(texto):
    if not texto: return ""
    s = str(texto).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def formatear_fecha(texto_fecha):
    if not texto_fecha: return "--"
    try:
        s_fecha = str(texto_fecha)
        if "T" in s_fecha: return datetime.strptime(s_fecha.split("T")[0], "%Y-%m-%d").strftime("%d-%m-%Y")
        if isinstance(texto_fecha, datetime): return texto_fecha.strftime("%d-%m-%Y")
        if " " in s_fecha: return s_fecha.split(" ")[0]
        return s_fecha
    except: return str(texto_fecha)

def descargar_foto_por_url(ctx, url):
    try:
        url = unquote(url)
        if url.startswith("http"): url = urlparse(url).path
        
        file_content = io.BytesIO()
        ctx.web.get_file_by_server_relative_url(url).download(file_content).execute_query()
        file_content.seek(0)
        
        if len(file_content.getvalue()) > 0:
            with Image.open(file_content) as img:
                if img.mode != "RGB": img = img.convert("RGB")
                img.thumbnail((400, 400)) # Compresión
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=60)
                return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
    except Exception:
        pass
    return None

def extraer_foto_columna(ctx, p, col_name, item_id):
    img_b64 = None
    json_raw = p.get(col_name)
    if json_raw:
        try:
            data = json.loads(json_raw) if isinstance(json_raw, str) else json_raw
            if isinstance(data, dict):
                url = data.get("serverRelativeUrl") or data.get("serverUrl") or data.get("Url")
                filename = data.get("fileName")
                if url: 
                    img_b64 = descargar_foto_por_url(ctx, url)
                if not img_b64 and filename:
                    rel_site = SITE_URL.replace("https://teams.wal-mart.com", "")
                    url_adj = f"{rel_site}/Lists/{LIST_NAME}/Attachments/{item_id}/{filename}"
                    img_b64 = descargar_foto_por_url(ctx, url_adj)
        except: pass
    return img_b64

# ==========================================
# 4. EXTRACCIÓN PRINCIPAL (SHAREPOINT API)
# ==========================================
def main():
    try:
        matriz_sap_json, mapa_sp_sap = procesar_bases_sap()

        print("🚀 INICIANDO EXTRACCIÓN DIRECTA DESDE SHAREPOINT...")
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        sp_list = ctx.web.lists.get_by_title(LIST_NAME)
        
        print("   ⏳ Solicitando registros y adjuntos...")
        columnas_req = [
            "Id", "Title", "LinkTitle", "field_2", "field_3", "field_4", 
            "field_5", "field_6", "field_7", "Responsable", "field_10", 
            "field_11", "field_14", "field_15", "Antes", "Despues", 
            "field_1", "ClaseM", "Zona", "Planta", "Attachments", "AttachmentFiles", "HH", 
            "Duraci_x00f3_n_x0028_HR_x0029_", "CantidadPersonas",
            "Dia", "Tecnico", "Tecnico2", "Tecnico3", "CRITICIDAD", "Colaborador"
        ]
        
        try:
            items = sp_list.items.select(columnas_req).expand(["AttachmentFiles"]).top(5000).get().execute_query()
        except Exception:
            columnas_req.remove("AttachmentFiles")
            items = sp_list.items.select(columnas_req).top(5000).get().execute_query()
            
        total_main = len(items)
        print(f"   ✅ Se descargaron {total_main} registros brutos.")
        
        db_json = {}
        for idx, item in enumerate(items):
            print(f"      ... Procesando OT {idx+1} de {total_main}", end='\r')
            p = item.properties
            
            semana_val = limpiar(p.get("field_1"))

            item_id = int(p.get("Id", 0))
            planta_raw = limpiar(p.get("Planta")).lower()
            planta_final = "carne" if "carne" in planta_raw else "masas"

            act_str = limpiar(p.get("field_4")) 
            tag_id = limpiar(p.get("LinkTitle"))
            titulo_final = act_str if act_str else (tag_id or f"OT #{item_id}")

            status_txt = normalizar_texto(limpiar(p.get("field_11"))) 
            if any(k in status_txt for k in ['ok', 'listo', 'cerrad', 'realiza', 'complet']): status = "realizada"
            elif any(k in status_txt for k in ['prog', 'planif']): status = "programado"
            elif any(k in status_txt for k in ['proceso', 'tratando', 'curso']): status = "en proceso"
            else: status = "pendiente"

            crit_raw = normalizar_texto(limpiar(p.get("CRITICIDAD")))
            if "crit" in crit_raw: crit_final = "Critica"
            elif "mayor" in crit_raw: crit_final = "Mayor"
            elif "menor" in crit_raw: crit_final = "Menor"
            else: crit_final = "Sin Asignar"

            colab_raw = limpiar(p.get("Colaborador")).strip().title()
            clase_str = limpiar(p.get("ClaseM")).title()
            clase_final = clase_str if clase_str and clase_str.lower() != "none" else "General"

            img_antes = extraer_foto_columna(ctx, p, "Antes", item_id)
            img_despues = extraer_foto_columna(ctx, p, "Despues", item_id)

            duracion_val = limpiar(p.get("Duraci_x00f3_n_x0028_HR_x0029_"))
            dotacion_val = limpiar(p.get("CantidadPersonas"))
            
            hh_raw = p.get("HH")
            try: hh_val = float(str(hh_raw).replace(',', '.')) if hh_raw else 0.0
            except: hh_val = 0.0

            key_id = f"MTTO_{item_id}"
            db_json[key_id] = {
                "key_id": key_id,
                "id_real": item_id,
                "titulo": titulo_final,
                "tag": tag_id,
                "semana": semana_val,
                "planta": planta_final,
                "ejecutor": limpiar(p.get("Responsable")) or "Sin Asignar",
                "criticidad": crit_final,
                "colaborador": colab_raw or "Interno",
                "ubicacion": limpiar(p.get("field_5")),
                "sub_ubi": limpiar(p.get("field_6")),
                "ot": limpiar(p.get("field_7")),
                "zona": limpiar(p.get("Zona")),
                "f_lev": formatear_fecha(p.get("field_2")),
                "f_cie": formatear_fecha(p.get("field_3")),
                "actividad": act_str or "Sin descripción",
                "observacion": limpiar(p.get("field_14")),
                "obs2": limpiar(p.get("field_15")),
                "status": status,
                "clase": clase_final,
                "origen": "act",
                "img_antes": img_antes,
                "img_despues": img_despues,
                "hh": hh_val,
                "duracion": duracion_val,
                "dotacion": dotacion_val,
                "dia": limpiar(p.get("Dia")).title(),
                "tecnico": limpiar(p.get("Tecnico")).upper(),
                "tecnico1": limpiar(p.get("Tecnico2")).upper(),
                "tecnico2": limpiar(p.get("Tecnico3")).upper()
            }
            
        print("\n ✅ Procesamiento finalizado. Construyendo HTML...")
        generar_html_moderno(db_json, matriz_sap_json, mapa_sp_sap)

    except Exception as e: 
        print(f"\n❌ Error Fatal: {e}")
        import traceback
        traceback.print_exc()

# ==========================================
# 5. GENERADOR HTML FRONTEND
# ==========================================
def generar_html_moderno(db_json, matriz_sap_json, mapa_sp_sap):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    
    html_template = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard Mantenimiento</title>
    <link rel="icon" type="image/x-icon" href="https://www.walmart.com/favicon.ico">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root { --primary: #0f172a; --secondary: #334155; --accent: #2563eb; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; --success: #10b981; --warn: #f59e0b; --danger: #ef4444; --info: #3b82f6; }
        * { box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }
        body { background: transparent; color: var(--text); margin: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        
        .top-bar { background: var(--primary); color: white; padding: 0 20px; height: 60px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; z-index: 10; box-shadow: 0 2px 5px rgba(0,0,0,0.1); transition: background-color 0.4s; }
        .brand { flex: 1; }
        .brand h2 { margin: 0; font-size: 1.2rem; display:flex; align-items:center; gap: 8px; } 
        .brand span { opacity: 0.7; font-weight: 300; font-size: 0.95rem; }

        .planta-switch { display: flex; align-items: center; justify-content: center; gap: 12px; background: rgba(255,255,255,0.15); padding: 5px 20px; border-radius: 30px; font-weight: bold; flex: 1; font-size: 0.9rem;}
        .switch { position: relative; display: inline-block; width: 44px; height: 24px; }
        .switch input { opacity: 0; width: 0; height: 0; }
        .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #3b82f6; transition: .4s; border-radius: 34px; box-shadow: inset 0 1px 3px rgba(0,0,0,0.4); }
        .slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background-color: white; transition: .4s; border-radius: 50%; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
        input:checked + .slider { background-color: #7f1d1d; } 
        input:checked + .slider:before { transform: translateX(20px); }
        
        .tabs-container { background: white; border-bottom: 1px solid var(--border); padding: 0 20px; flex-shrink: 0; display:flex; justify-content: space-between; align-items: center; z-index: 5; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .tabs-nav { display: flex; gap: 15px; }
        .tab-btn { background: none; border: none; padding: 15px 5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 0.95rem; }
        .tab-btn:hover { color: var(--accent); } .tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
        
        .app-layout { display: flex; height: calc(100vh - 110px); width: 100%; overflow: hidden; position:relative; }
        
        /* ESTILOS PANEL DE FILTROS FLOTANTE */
        .col-filters { 
            position: absolute; left: -320px; top: 0; bottom: 0; 
            width: 320px; background: #fff; border-right: 1px solid var(--border); 
            display: flex; flex-direction: column; z-index: 100; 
            box-shadow: 5px 0 15px rgba(0,0,0,0.1); 
            transition: left 0.4s cubic-bezier(0.4, 0, 0.2, 1); 
        }
        .col-filters.active { left: 0; }
        
        .btn-toggle-filters {
            position: absolute; left: 20px; bottom: 20px; z-index: 90;
            background: var(--accent); color: white; border: none; 
            padding: 12px 20px; border-radius: 30px; font-weight: bold; 
            cursor: pointer; box-shadow: 0 4px 10px rgba(37, 99, 235, 0.3);
            display: flex; align-items: center; gap: 8px; transition: 0.3s;
        }
        .btn-toggle-filters:hover { transform: translateY(-2px); box-shadow: 0 6px 15px rgba(37, 99, 235, 0.4); }
        
        .filters-header { padding: 15px 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--primary); font-size: 0.9rem; text-transform: uppercase; background: #f8fafc; display: flex; justify-content: space-between; align-items: center; }
        .filters-body { flex: 1; overflow-y: auto; padding: 20px; min-height: 0; } 
        .filters-footer { padding: 20px; border-top: 1px solid var(--border); background: #f8fafc; flex-shrink: 0; }
        
        .f-group { margin-bottom: 15px; }
        .f-group label { font-size: 0.75rem; font-weight: 700; color: var(--muted); display: block; margin-bottom: 6px; text-transform: uppercase; }
        select, input[type="text"] { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; color: var(--text); }
        select:focus, input[type="text"]:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1); }
        
        .btn-clean { background: white; border: 1px solid var(--danger); color: var(--danger); padding: 10px; border-radius: 6px; cursor: pointer; font-weight: 700; transition: 0.2s; width: 100%; text-transform: uppercase; font-size: 0.8rem; letter-spacing: 0.5px; }
        .btn-clean:hover { background: var(--danger); color: white; }
        
        .kpi-row-mini { display: flex; justify-content: space-between; margin-bottom: 15px; }
        .kpi-box { text-align: center; } .k-label { display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; }
        .k-num { display: block; font-size: 1.3rem; font-weight: 800; color: var(--primary); } .k-ok { color: var(--success); } .k-pend { color: var(--danger); }
        .prog-title { display: flex; justify-content: space-between; font-size: 0.75rem; font-weight: 700; color: var(--muted); margin-bottom: 6px; }
        .progress-bar-container { width: 100%; height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }
        .progress-bar-fill { height: 100%; background: var(--success); width: 0%; transition: width 1s cubic-bezier(0.4, 0, 0.2, 1); }
        
        .col-list { width: 380px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; }
        .list-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 600; background: #f8fafc; color: var(--secondary); font-size: 0.9rem; flex-shrink: 0; display:flex; flex-direction:column; gap:12px; }
        .list-scroll-area { flex: 1; overflow-y: auto; min-height: 0; }
        
        .list-item { padding: 15px 20px; border-bottom: 1px solid var(--border); cursor: pointer; transition: 0.2s; border-left: 4px solid transparent; }
        .list-item:hover { background: #f8fafc; } .list-item.selected { background: #eff6ff; border-left-color: var(--accent); }
        .li-top { display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 0.75rem; color: var(--muted); font-weight: 600; }
        .li-title { font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 10px; line-height: 1.4; }
        .li-btm { display: flex; justify-content: space-between; font-size: 0.75rem; align-items: center; }
        
        .tag { padding: 4px 8px; border-radius: 4px; font-weight: 700; font-size: 0.7rem; letter-spacing: 0.3px; }
        .st-ok { background: #dcfce7; color: #166534; } .st-pend { background: #fee2e2; color: #991b1b; } .st-prog { background: #e0f2fe; color: #075985; } .st-proc { background: #fef3c7; color: #92400e; }
        
        .col-detail { flex: 1; overflow-y: auto; padding: 40px; }
        .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--muted); opacity: 0.7; }
        .detail-content { background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); overflow: hidden; max-width: 1000px; margin: 0 auto; border: 1px solid var(--border); }
        .detail-header { padding: 30px; border-bottom: 1px solid var(--border); background: #fff; }
        .dh-top { display: flex; justify-content: space-between; margin-bottom: 15px; align-items:center; }
        .detail-header h2 { margin: 0 0 5px 0; font-size: 1.6rem; color: var(--primary); }
        
        .data-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 25px; padding: 30px; background: #fff; border-bottom: 1px solid var(--border); }
        .dg-item small { display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; margin-bottom: 6px; text-transform: uppercase; }
        .dg-item strong { font-size: 1rem; color: var(--text); }
        
        .obs-box { padding: 30px; border-bottom: 1px solid var(--border); }
        .obs-box h4 { margin: 0 0 12px; color: var(--secondary); font-size: 0.9rem; text-transform: uppercase; }
        .obs-box p { background: #f8fafc; padding: 20px; border-radius: 8px; border: 1px solid var(--border); margin: 0; line-height: 1.6; color: #334155; }
        
        .gallery-section { padding: 30px; background: #f8fafc; display:flex; flex-direction: column; align-items: center; gap: 15px; }
        .gallery-section h4 { margin:0; color:var(--secondary); font-size:0.9rem; text-transform:uppercase; align-self: flex-start; }
        .gallery-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; width: 100%; }
        .gal-box { background: white; border: 1px solid var(--border); border-radius: 8px; padding: 15px; display: flex; flex-direction: column; align-items: center; gap: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .gal-box span { font-weight: 700; font-size: 0.85rem; color: var(--secondary); text-transform: uppercase; padding-bottom: 5px; border-bottom: 2px solid var(--accent); margin-bottom: 5px; }
        .gal-img { max-width: 100%; max-height: 350px; border-radius: 6px; cursor: zoom-in; box-shadow: 0 2px 5px rgba(0,0,0,0.1); transition: transform 0.2s; object-fit: contain; }
        .gal-img:hover { transform: scale(1.02); }
        
        .graficos-layout { flex: 1; padding: 30px; display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); grid-auto-rows: min-content; gap: 25px; overflow-y: auto; align-content: start; }
        .chart-card { background: white; padding: 25px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; height: 400px; width: 100%; }
        .chart-card.wide { grid-column: 1 / -1; height: 450px; }
        .chart-title { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; text-transform: uppercase; text-align: center; }
        .canvas-container { position: relative; flex: 1 1 auto; width: 100%; min-height: 0; }
        
        .prio-flag { padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; text-transform: uppercase; }
        .p-crit { background: #fee2e2; color: #dc2626; border: 1px solid #f87171; }
        .p-alta { background: #ffedd5; color: #ea580c; border: 1px solid #fdba74; }
        .p-baja { background: #dcfce7; color: #166534; border: 1px solid #86efac; }
        
        .modal { display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.85); align-items: center; justify-content: center; backdrop-filter: blur(4px); }
        .modal img { max-width: 90%; max-height: 90vh; border-radius: 8px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        
        #data_modal_content { background: white; width: 90%; max-width: 1200px; max-height: 85vh; border-radius: 12px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        .dm-header { padding: 20px 25px; background: var(--primary); color: white; display: flex; justify-content: space-between; align-items: center; }
        .dm-header h3 { margin: 0; font-size: 1.2rem; font-weight: 600; }
        .dm-close { background: none; border: none; color: white; font-size: 1.8rem; cursor: pointer; opacity: 0.8; transition: 0.2s; line-height: 1; }
        .dm-close:hover { opacity: 1; transform: scale(1.1); }
        .dm-body { padding: 0; overflow-y: auto; flex: 1; background: var(--bg); }
        .dm-table { width: 100%; border-collapse: collapse; background: white; font-size: 0.9rem; text-align: left; }
        .dm-table th { background: #f8fafc; padding: 15px 20px; font-weight: 700; color: var(--secondary); border-bottom: 2px solid var(--border); position: sticky; top: 0; z-index: 10; text-transform: uppercase; font-size: 0.8rem; cursor: pointer; user-select: none; transition: background 0.2s; }
        .dm-table td { padding: 15px 20px; border-bottom: 1px solid var(--border); color: var(--text); }
        .dm-table tr { transition: background 0.2s; }
        .dm-table tr:hover td { background: #eff6ff; cursor: pointer; }
        
        .summary-block { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:15px; margin-bottom:12px; }
        .summary-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; }
        .summary-title { font-weight:700; font-size: 0.95rem; }
        .summary-perc { font-weight:800; font-size: 1.1rem; }
        .summary-sub { font-size:0.8rem; color:#64748b; }
        .summary-bar-bg { width:100%; height:6px; background:#e2e8f0; border-radius:3px; margin-top:8px; overflow:hidden; }
        .summary-bar-fill { height:100%; transition:width 1s cubic-bezier(0.4, 0, 0.2, 1); }
        
        /* GANTT TURNOS */
        .gantt-day-col { flex:1; min-width:320px; background:white; border:1px solid var(--border); border-radius:8px; display:flex; flex-direction:column; overflow:hidden; box-shadow:0 2px 4px rgba(0,0,0,0.02); }
        .gantt-day-header { background:var(--secondary); color:white; padding:15px; text-align:center; }
        .gantt-day-title { margin:0; font-size:1.1rem; text-transform:uppercase; font-weight:700; letter-spacing:0.5px;}
        .gantt-day-hh { font-size:0.85rem; opacity:0.9; font-weight:600; display:block; margin-top:4px;}
        .gantt-shift-box { border-radius:6px; padding:12px; margin-bottom:12px; }
        .gantt-shift-header { display:flex; justify-content:space-between; font-weight:700; font-size:0.85rem; margin-bottom:10px; padding-bottom:6px; text-transform:uppercase; }
        .gantt-card { background:white; border-left:4px solid transparent; padding:10px; margin-bottom:8px; border-radius:6px; font-size:0.8rem; cursor:pointer; box-shadow:0 2px 4px rgba(0,0,0,0.06); transition:transform 0.15s; }
        .gantt-card:hover { transform: translateY(-2px); box-shadow:0 4px 6px rgba(0,0,0,0.1); }
        
        /* ESTILOS GANTT PLAN MATRIZ (RENOVADO) */
        .tab-pm { overflow: hidden; border: 1px solid #ccc; background-color: #f1f1f1; border-radius: 8px 8px 0 0; display: flex; }
        .tab-pm button { background-color: inherit; border: none; outline: none; cursor: pointer; padding: 12px 20px; transition: 0.3s; font-size: 14px; font-weight: bold; color: #555; flex-grow: 1;}
        .tab-pm button.active { background-color: #2c3e50; color: white; }
        .year-select-pm { padding: 6px 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 1em; font-weight: bold; color: #1a237e; background-color: #fff; cursor: pointer; margin-left:15px; }
        
        .pm-table { border-collapse: collapse; width: 100%; font-size: 0.8em; text-align: center; }
        .pm-table th { background: var(--secondary); color: white; padding: 8px 4px; position: sticky; top: 0; z-index: 10; font-weight: 600; min-width: 30px; border: 1px solid #475569; }
        .pm-table th.pm-tag-col { position: sticky; left: 0; background: var(--primary); z-index: 20; min-width: 380px; max-width: 380px; text-align: left; padding-left: 15px; }
        .pm-table td { border: 1px solid var(--border); padding: 0; height: 35px; position: relative; }
        
        .pm-table .location-row .pm-tag-col { position: sticky; left: 0; background-color: #e8eaf6; color: #1a237e; font-weight: 800; border-top: 3px solid #b0bec5; font-size: 0.95em; text-align: left; padding-left: 15px; z-index: 15; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}
        .pm-table .location-row td { background-color: #e8eaf6; border-top: 3px solid #b0bec5; }
        
        .pm-table .eq-row .pm-tag-col { position: sticky; left: 0; background-color: #fdfdfd; font-weight: 600; color: #333; padding-left: 35px; border-left: 4px solid #1a237e; text-align: left; z-index: 15; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}
        
        .pm-table tr:hover .pm-tag-col.loc-bg { background-color: #d1d5db; }
        .pm-table tr:hover { background: #f1f5f9; }
        
        .pm-cell-inner { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: 0.2s; }
        .pm-cell-inner:hover { opacity: 0.8; transform: scale(1.15); }
        .pm-ok { background: #22c55e; border-radius: 4px; width: 20px; height: 20px; margin: auto; color: white; display:flex; align-items:center; justify-content:center; font-weight:bold; position:relative; box-shadow: 0 1px 3px rgba(0,0,0,0.2);}
        .pm-pend { background: #ef4444; border-radius: 4px; width: 20px; height: 20px; margin: auto; color: white; display:flex; align-items:center; justify-content:center; font-weight:bold; position:relative; box-shadow: 0 1px 3px rgba(0,0,0,0.2);}
        .pm-proc { background: #f59e0b; border-radius: 4px; width: 20px; height: 20px; margin: auto; color: white; display:flex; align-items:center; justify-content:center; font-weight:bold; position:relative; box-shadow: 0 1px 3px rgba(0,0,0,0.2);}
        .pm-multi { position:absolute; top:-6px; right:-6px; background:#e74c3c; color:white; border-radius:50%; width:12px; height:12px; font-size:9px; display:flex; align-items:center; justify-content:center; font-weight:bold; box-shadow: 0 1px 2px rgba(0,0,0,0.3);}
    </style>
</head>
<body>
    <div id="modal" class="modal" onclick="if(event.target===this) this.style.display='none'"><img id="modalImg"></div>
    
    <div id="data_modal" class="modal" onclick="if(event.target===this) this.style.display='none'">
        <div id="data_modal_content"></div>
    </div>

    <div class="top-bar">
        <div class="brand"><h2>⚙️ Panel Gestión de Actividades <span>SubGerencia de Mantenimiento</span></h2></div>
        
        <div class="planta-switch">
            <span style="opacity:0.9;">Masas</span>
            <label class="switch">
                <input type="checkbox" id="planta_toggle" onchange="togglePlanta()">
                <span class="slider"></span>
            </label>
            <span style="opacity:0.9;">Carne</span>
        </div>

        <div style="flex: 1; display: flex; justify-content: flex-end; align-items: center; padding-right: 10px;">
            <img src="https://upload.wikimedia.org/wikipedia/commons/b/b1/Walmart_logo_%282008%29.svg" alt="Walmart Logo" style="height: 35px; object-fit: contain; opacity: 0.95; filter: brightness(0) invert(1); cursor: pointer;" ondblclick="descargarHTML()" title="Doble clic para descargar el archivo HTML del Dashboard">
        </div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs-nav">
            <button class="tab-btn active" onclick="setView('list', this)" id="btn_tab_list">📋 Visor de OTs</button>
            <button class="tab-btn" onclick="setView('charts', this)">📊 Análisis y Tendencias</button>
            <button class="tab-btn" onclick="setView('row', this)">📈 ROW</button>
            <button class="tab-btn" onclick="setView('gantt', this)">📅 Gantt / Turnos</button>
            <button class="tab-btn" onclick="setView('gantt_pm', this)">⚙️ Plan Matriz (Integrado)</button>
        </div>
        <div style="display:flex; gap:10px;">
            <button onclick="descargarExcel()" class="btn-clean" style="margin: 0; padding: 8px 15px; width: auto; border-color: #10b981; color: #10b981; display: flex; align-items: center; gap: 8px;" title="Descargar datos filtrados">
                <span style="font-size:1.2rem;">📊</span> Exportar Excel
            </button>
        </div>
    </div>
    
    <div class="app-layout">
        
        <!-- BOTÓN PARA MOSTRAR FILTROS -->
        <button class="btn-toggle-filters" onclick="toggleFiltersPanel()">⚙️ Filtros Principales</button>
        
        <div class="col-filters" id="main_filters">
            <div class="filters-header">
                <span>🔍 Filtros</span>
                <div style="display:flex; gap:8px;">
                    <button onclick="resetFilters()" class="btn-clean" style="margin: 0; padding: 4px 8px; width: auto; font-size: 0.7rem; border-color: #ef4444; color: #ef4444; text-transform: none;">🧹 Borrar</button>
                    <button onclick="toggleFiltersPanel()" class="btn-clean" style="margin: 0; padding: 4px 8px; width: auto; font-size: 0.7rem; border-color: var(--secondary); color: var(--secondary); text-transform: none;">❌ Cerrar</button>
                </div>
            </div>
            
            <div class="filters-body" id="filters_dynamic"></div>
            <div class="filters-footer">
                <div class="kpi-row-mini">
                    <div class="kpi-box"><span class="k-label">TOTAL OT</span><span class="k-num" id="k_total">0</span></div>
                    <div class="kpi-box"><span class="k-label">CERRADAS</span><span class="k-num k-ok" id="k_ok">0</span></div>
                    <div class="kpi-box"><span class="k-label">BACKLOG</span><span class="k-num k-pend" id="k_pend">0</span></div>
                </div>
                <div class="prog-title"><span>Cumplimiento Global</span><span id="k_perc">0%</span></div>
                <div class="progress-bar-container"><div id="bar_fill" class="progress-bar-fill"></div></div>
                
                <div style="margin-top: 15px; padding-top: 12px; border-top: 1px dashed var(--border); text-align: center; font-size: 0.75rem; color: var(--muted); font-weight: 700; text-transform: uppercase;">
                    <span id="top_week_indicator" style="color: var(--primary);">Actualizando...</span><br>
                    <span style="opacity: 0.7; font-size: 0.65rem; display: inline-block; margin-top: 4px; font-weight: 600;">Actualizado: __FECHA_ACTUAL__</span>
                </div>
            </div>
        </div>

        <div id="view_list" style="display:flex; flex:1; overflow:hidden;">
            <div class="col-list">
                <div class="list-header">
                    <div>📋 Listado de Actividades</div>
                    <input type="text" id="search_input" placeholder="🔍 Buscar TAG, Título o OT..." onkeyup="applyFilters()">
                </div>
                <div id="list_container" class="list-scroll-area"></div>
            </div>
            <div class="col-detail">
                <div id="empty_state" class="empty-state"><div style="font-size:4rem; margin-bottom:15px;">📋</div><h3 style="margin:0">Selecciona una OT</h3><p>Usa la lista izquierda para ver detalles.</p></div>
                <div id="detail_view" class="detail-content" style="display:none">
                    <div class="detail-header">
                        <div class="dh-top">
                            <div style="display:flex; align-items:center; gap:15px; flex-wrap:wrap;">
                                <span id="d_status" class="tag st-ok">STATUS</span>
                                <div id="d_extra_info" style="display:none; gap:12px; font-size:0.8rem; color:var(--secondary); font-weight:700; align-items:center; background:#f8fafc; padding:4px 12px; border-radius:6px; border:1px solid var(--border);"></div>
                            </div>
                            <div id="d_prio_lbl">PRIO</div>
                        </div>
                        <h2 id="d_title">Título de la Actividad</h2>
                        <p style="color:var(--accent); font-weight: 600; font-size: 1.05rem; margin:0;" id="d_tag">TAG</p>
                    </div>
                    <div class="data-grid" id="d_grid"></div>
                    <div class="obs-box" id="box_obs1"><h4 id="lbl_obs_title">📝 Observación Técnica</h4><p id="d_obs">--</p></div>
                    <div class="obs-box" id="box_obs2" style="display:none;"><h4 id="lbl_obs_title2">📝 Observación Adicional</h4><p id="d_obs2">--</p></div>
                    
                    <div class="gallery-section" id="d_gallery_sec">
                        <h4>📸 Registro Fotográfico</h4>
                        <div id="d_img_container" style="width: 100%;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="view_charts" class="graficos-layout" style="display:none;">
            <div class="chart-card"><div class="chart-title">Status del Backlog</div><div class="canvas-container"><canvas id="chart1"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Clase de Mantenimiento</div><div class="canvas-container"><canvas id="chart2"></canvas></div></div>
            
            <div class="chart-card">
                <div class="chart-title">Resumen de Actividades</div>
                <div id="summary_content" style="display:flex; flex-direction:column; justify-content:space-around; flex:1;"></div>
            </div>
            
            <div class="chart-card wide">
                <div class="chart-title">Desglose de Tiempos (HH) por Área</div>
                <div class="canvas-container"><canvas id="chart_hh_area"></canvas></div>
            </div>            
            
            <div class="chart-card wide"><div class="chart-title">Avance por Área</div><div class="canvas-container"><canvas id="chart5"></canvas></div></div>
            
            <div class="chart-card wide"><div class="chart-title">Top Ubicaciones Críticas</div><div class="canvas-container"><canvas id="chart4"></canvas></div></div>

            <div class="chart-card wide"><div class="chart-title">Carga Laboral por Responsable</div><div class="canvas-container"><canvas id="chart3"></canvas></div></div>
        </div>
        
        <div id="view_row" style="display:none; flex:1; flex-direction:column; overflow-y:auto; padding:30px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; flex-wrap:wrap; gap:15px;">
                <h2 style="color:var(--primary); margin:0; font-size:1.8rem;">Planificación Mantenimiento <span id="row_week_title" style="color:var(--accent);">--</span></h2>
                <button id="btn_descargar_row" class="btn-clean" style="width:auto; margin:0; padding: 8px 15px; border-color:var(--accent); color:var(--accent); display:flex; align-items:center; gap:8px;" onclick="descargarROW()">
                    <span style="font-size:1.2rem;">📸</span> Descargar Dashboard ROW
                </button>
            </div>
            
            <div style="display:flex; gap:25px; margin-bottom:30px; flex-wrap:wrap;">
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Distribución MTTO vs Aseo / Sanitización y Seguridad</div><div class="canvas-container"><canvas id="row_chart1"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Cumplimiento Mantenimiento General</div><div class="canvas-container"><canvas id="row_chart2"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:350px; min-width:300px;"><div class="chart-title">Cumplimiento Aseo / Sanitización y Seguridad</div><div class="canvas-container"><canvas id="row_chart3"></canvas></div></div>
            </div>
            
            <div style="display:flex; gap:25px; flex-wrap:wrap;">
                <div class="chart-card" style="flex:1; height:450px; min-width:400px;"><div class="chart-title">Panadería: Cumplimiento por Línea</div><div class="canvas-container"><canvas id="row_chart4"></canvas></div></div>
                <div class="chart-card" style="flex:1; height:450px; min-width:400px;"><div class="chart-title">Dely: Cumplimiento por Área</div><div class="canvas-container"><canvas id="row_chart5"></canvas></div></div>
            </div>
        </div>

        <div id="view_gantt" style="display:none; flex:1; flex-direction:column; overflow-y:auto; padding:30px; background:#f8fafc;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; flex-wrap:wrap; gap:15px;">
                <h2 style="color:var(--primary); margin:0; font-size:1.8rem;">Planificación de Turnos y Carga Técnica</h2>
            </div>
            <div id="gantt_container" style="display:flex; gap:20px; overflow-x:auto; padding-bottom:15px; flex:1;">
                </div>
        </div>

        <div id="view_gantt_pm" style="display:none; flex:1; flex-direction:column; overflow:hidden; padding:20px; background:#f8fafc;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; flex-shrink:0;">
                <h2 style="color:var(--primary); margin:0; font-size:1.5rem;">Plan Matriz de Mantenimiento</h2>
                
                <div style="display:flex; align-items:center; gap: 15px;">
                    <input type="text" id="search_pm" class="search-input" placeholder="🔍 Buscar OT, Ubicación o Equipo..." onkeyup="filtrarGanttPM()" style="width:300px;">
                    <div style="display:flex; gap:10px; font-size:0.8rem; font-weight:700;">
                        <span style="background:#dcfce7; color:#166534; padding:4px 8px; border-radius:4px;">🟩 Realizada</span>
                        <span style="background:#fef3c7; color:#92400e; padding:4px 8px; border-radius:4px;">🟨 En Proceso</span>
                        <span style="background:#fee2e2; color:#991b1b; padding:4px 8px; border-radius:4px;">🟥 Pendiente</span>
                    </div>
                </div>
            </div>
            
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0;">
                <div class="tab-pm">
                    <button class="tablinks-pm active" onclick="setPmClass('A', this)">Clase A (Críticos)</button>
                    <button class="tablinks-pm" onclick="setPmClass('B', this)">Clase B</button>
                    <button class="tablinks-pm" onclick="setPmClass('C', this)">Clase C</button>
                </div>
                <select id="pm_year_select" class="year-select-pm" onchange="setPmYear(this.value)">
                    <option value="2026">2026</option>
                    <option value="2025">2025</option>
                </select>
            </div>

            <div id="gantt_pm_container" style="flex:1; overflow:auto; background:white; border:1px solid var(--border); border-top:none; border-radius:0 0 8px 8px; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                </div>
        </div>

    </div>

    <script>
    Chart.register(ChartDataLabels);
    Chart.defaults.plugins.datalabels.display = false; 

    // DATOS DE SHAREPOINT Y SAP
    const db = __DB_JSON_DATA__;
    const records = Object.values(db).sort((a,b) => b.id_real - a.id_real);
    const weeks = [...new Set(records.map(x=>x.semana).filter(x=>x!=="S/N"))].sort((a,b)=>{ let na=parseInt(a), nb=parseInt(b); return (isNaN(na)||isNaN(nb)) ? a.localeCompare(b) : na-nb; });
    const matrizSAP = __MATRIZ_SAP_JSON__;
    const mapaSpSap = __MAPA_SP_SAP__;

    let appState = { statusFilter: 'all', view: 'list' };
    let currentChartData = [];
    let chartInstances = {};
    let currentPlanta = 'masas';
    let currentPmYear = '2026';
    let currentPmClass = 'A';
    let filterTimer;
    
    Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
    Chart.defaults.color = '#64748b';

    // Función Helper para fechas invertidas (YYYY-MM-DD -> DD-MM-YY)
    function formatDateDDMMYY(dateStr) {
        if(!dateStr || !dateStr.includes('-')) return dateStr;
        let parts = dateStr.split('-');
        if(parts.length === 3) {
            let yy = parts[0].length === 4 ? parts[0].substring(2) : parts[0];
            return `${parts[2]}-${parts[1]}-${yy}`;
        }
        return dateStr;
    }

    // --- MANEJO DEL PANEL DE FILTROS ---
    function toggleFiltersPanel() {
        const panel = document.getElementById('main_filters');
        panel.classList.toggle('active');
        if (panel.classList.contains('active')) {
            startFilterTimer();
        } else {
            clearTimeout(filterTimer);
        }
    }

    function startFilterTimer() {
        clearTimeout(filterTimer);
        filterTimer = setTimeout(() => {
            const panel = document.getElementById('main_filters');
            if(panel.classList.contains('active')) panel.classList.remove('active');
        }, 8000);
    }

    document.getElementById('main_filters').addEventListener('mousemove', startFilterTimer);
    document.getElementById('main_filters').addEventListener('click', startFilterTimer);

    const getAreaResp = (ejecutor) => {
        let ejL = (ejecutor || '').toLowerCase();
        let mecanicos = ['luis lagos', 'luis guajardo', 'rubén carrasco', 'ruben carrasco', 
                         'marcelo rivera', 'vladimir berrios', 'rubén briceño', 'ruben briceño', 
                         'mantenimiento', 'javier cordova', 'nicolás chandia', 'nicolas chandia'];
        if (mecanicos.some(nombre => ejL.includes(nombre))) return 'Mecánico';
        if (ejL.includes('autómata') || ejL.includes('automata')) return 'Autómata';
        if (ejL.includes('edward corona') || ejL.includes('frio') || ejL.includes('frío')) return 'Frio';
        if (ejL.includes('infraestructura')) return 'Infraestructura';
        return 'Otros';
    };

    function togglePlanta() {
        const cb = document.getElementById('planta_toggle');
        const topBar = document.querySelector('.top-bar');
        if(cb.checked) { currentPlanta = 'carne'; topBar.style.backgroundColor = '#ef4444'; } 
        else { currentPlanta = 'masas'; topBar.style.backgroundColor = '#0f172a'; }
        applyFilters();
    }

    function buildFilters() {
        const fDiv = document.getElementById('filters_dynamic');
        const createSelect = (id, label, options, defValue = 'ALL') => {
            let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()">`;
            sel += `<option value="ALL">Todos</option>`;
            options.forEach(o => { if(o) sel += `<option value="${o}" ${o === defValue ? 'selected' : ''}>${o}</option>`; });
            sel += `</select></div>`;
            return sel;
        };

        let html = '';
        html += createSelect('f_semana', '📆 Semana', weeks, 'ALL');
        html += createSelect('f_zona', '📍 Zona', [...new Set(records.map(x=>x.zona))].filter(Boolean).sort());
        html += createSelect('f_clase', '🛠️ Clase MTTO', [...new Set(records.map(x=>x.clase))].sort());
        html += createSelect('f_especialidad', '🧑‍🔧 Especialidad', ['Mecánico', 'Autómata', 'Frio', 'Infraestructura', 'Otros'].sort());
        html += `<div class="f-group"><label>👥 Colaborador</label><select id="f_colaborador" onchange="applyFilters()">
            <option value="ALL">Todos</option><option value="Interno">Interno</option><option value="Externo">Externo</option></select></div>`;
        html += createSelect('f_exec', '👷 Responsable', [...new Set(records.map(x=>x.ejecutor))].sort());
        html += createSelect('f_ubi', '🏭 Línea / Área', [...new Set(records.map(x=>x.ubicacion))].sort());
        html += `<div class="f-group"><label>⚠️ Criticidad</label><select id="f_criticidad" onchange="applyFilters()">
            <option value="ALL">Todas</option><option value="Critica">🚨 Crítica</option><option value="Mayor">🔴 Mayor</option><option value="Menor">🟢 Menor</option></select></div>`;
        html += `<div class="f-group"><label>🚦 Estado</label><select id="f_status" onchange="applyFilters()">
            <option value="ALL">Todas las OTs</option><option value="abiertas">Backlog (No Cerradas)</option><option value="pendiente">Solo Pendientes</option><option value="en proceso">Solo En Proceso</option><option value="programado">Solo Programadas</option><option value="realizada">Solo Cerradas</option></select></div>`;
        fDiv.innerHTML = html;
    }

    function resetFilters() {
        if(document.getElementById('search_input')) document.getElementById('search_input').value = '';
        document.querySelectorAll('.f-group select').forEach(sel => sel.value = "ALL");
        applyFilters();
    }

    function setView(view, btn) {
        appState.view = view;
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active');
        else document.getElementById('btn_tab_list').classList.add('active');
        
        document.getElementById('view_list').style.display = 'none';
        document.getElementById('view_charts').style.display = 'none';
        document.getElementById('view_row').style.display = 'none';
        document.getElementById('view_gantt').style.display = 'none';
        document.getElementById('view_gantt_pm').style.display = 'none';

        if (view === 'list') document.getElementById('view_list').style.display = 'flex';
        else if (view === 'charts') document.getElementById('view_charts').style.display = 'grid';
        else if (view === 'row') document.getElementById('view_row').style.display = 'flex';
        else if (view === 'gantt') document.getElementById('view_gantt').style.display = 'flex';
        else if (view === 'gantt_pm') {
            document.getElementById('view_gantt_pm').style.display = 'flex';
            drawGanttPM();
        }
        applyFilters();
    }

    function getFilteredData() {
        const eVal = document.getElementById('f_exec') ? document.getElementById('f_exec').value : 'ALL';
        const uVal = document.getElementById('f_ubi') ? document.getElementById('f_ubi').value : 'ALL';
        const cVal = document.getElementById('f_clase') ? document.getElementById('f_clase').value : 'ALL';
        const stVal = document.getElementById('f_status') ? document.getElementById('f_status').value : 'ALL';
        const semVal = document.getElementById('f_semana') ? document.getElementById('f_semana').value : 'ALL';
        const zVal = document.getElementById('f_zona') ? document.getElementById('f_zona').value : 'ALL';
        const espVal = document.getElementById('f_especialidad') ? document.getElementById('f_especialidad').value : 'ALL';
        const critVal = document.getElementById('f_criticidad') ? document.getElementById('f_criticidad').value : 'ALL';
        const colabVal = document.getElementById('f_colaborador') ? document.getElementById('f_colaborador').value : 'ALL';
        const searchVal = document.getElementById('search_input') ? document.getElementById('search_input').value.toLowerCase().trim() : '';

        let topWeekTitle = "Semanas Cargadas: " + (weeks.length > 0 ? weeks.join(', ') : "Ninguna");
        if (semVal !== "ALL") topWeekTitle = "Semana " + semVal;
        document.getElementById('top_week_indicator').innerText = topWeekTitle;

        return records.filter(d => {
            if (d.planta !== currentPlanta) return false;
            if (stVal !== 'ALL') {
                if (stVal === 'abiertas' && d.status === 'realizada') return false;
                else if (stVal !== 'abiertas' && d.status !== stVal) return false;
            }
            if (searchVal !== '') {
                const text = `${d.titulo} ${d.ot} ${d.tag}`.toLowerCase();
                if (!text.includes(searchVal)) return false;
            }
            if (cVal !== 'ALL' && d.clase !== cVal) return false;
            if (eVal !== 'ALL' && d.ejecutor !== eVal) return false;
            if (uVal !== 'ALL' && d.ubicacion !== uVal) return false;
            if (semVal !== 'ALL' && d.semana !== semVal) return false;
            if (zVal !== 'ALL' && d.zona !== zVal) return false;
            if (espVal !== 'ALL' && getAreaResp(d.ejecutor) !== espVal) return false;
            if (critVal !== 'ALL' && d.criticidad !== critVal) return false;
            if (colabVal !== 'ALL' && d.colaborador !== colabVal) return false;
            return true;
        });
    }

    function applyFilters() {
        currentChartData = getFilteredData();
        let ok = 0; currentChartData.forEach(d => { if(d.status === 'realizada') ok++; });
        const total = currentChartData.length;
        
        document.getElementById('k_total').innerText = total;
        document.getElementById('k_ok').innerText = ok;
        document.getElementById('k_pend').innerText = total - ok;
        let perc = total > 0 ? Math.round((ok/total)*100) : 0;
        document.getElementById('k_perc').innerText = perc + '%';
        const bar = document.getElementById('bar_fill');
        bar.style.width = perc + '%';
        bar.style.backgroundColor = perc > 80 ? '#10b981' : (perc > 40 ? '#f59e0b' : '#ef4444');

        if(appState.view === 'list') renderList(currentChartData);
        else if (appState.view === 'charts') drawCharts(currentChartData);
        else if (appState.view === 'row') drawRowCharts(currentChartData);
        else if (appState.view === 'gantt') drawGantt(currentChartData);
        else if (appState.view === 'gantt_pm') drawGanttPM();
    }

    function renderList(data) {
        const container = document.getElementById('list_container');
        container.innerHTML = '';
        data.forEach(d => {
            const item = document.createElement('div');
            item.className = 'list-item';
            item.onclick = function() { 
                renderDetail(d.key_id); 
                document.querySelectorAll('.list-item').forEach(i=>i.classList.remove('selected'));
                item.classList.add('selected');
            };
            let stText = '⚠️ PEND'; let stClass = 'st-pend';
            if (d.status === 'realizada') { stText='✅ CERRADA'; stClass='st-ok'; }
            else if (d.status === 'programado') { stText='📅 PROG'; stClass='st-prog'; }
            else if (d.status === 'en proceso') { stText='🔨 PROCESO'; stClass='st-proc'; }
            let idDisplay = d.ot ? `OT: ${d.ot}` : (d.tag ? d.tag : '#' + d.id_real);
            item.innerHTML = `<div class="li-top"><span>${idDisplay}</span><span>Sem: ${d.semana}</span></div>
                <div class="li-title">${d.titulo}</div>
                <div class="li-btm"><span class="tag ${stClass}">${stText}</span><span style="color:var(--muted); font-weight:700;">👷 ${d.ejecutor.split(' ')[0]}</span></div>`;
            container.appendChild(item);
        });
    }

    function renderDetail(key) {
        document.getElementById('empty_state').style.display='none';
        document.getElementById('detail_view').style.display='block';
        const d = db[key];
        
        document.getElementById('d_title').innerText = d.titulo;
        document.getElementById('d_tag').innerText = d.tag ? `TAG / Equipo: ${d.tag}` : (d.ot ? `OT: ${d.ot}` : 'Sin TAG');
        
        const stBadge = document.getElementById('d_status');
        if (d.status === 'realizada') { stBadge.innerText = '✅ CERRADA'; stBadge.className = 'tag st-ok'; }
        else if (d.status === 'programado') { stBadge.innerText = '📅 PROGRAMADA'; stBadge.className = 'tag st-prog'; }
        else if (d.status === 'en proceso') { stBadge.innerText = '🔨 EN PROCESO'; stBadge.className = 'tag st-proc'; }
        else { stBadge.innerText = '⚠️ PENDIENTE'; stBadge.className = 'tag st-pend'; }
        
        let durText = d.duracion ? `<span title="Tiempo de Ejecución (h)">⏱️ Ejecución: ${d.duracion}h</span>` : '';
        let dotText = d.dotacion ? `<span title="Dotación (Cantidad de Personas)">👥 Dotación: ${d.dotacion}</span>` : '';
        let hhText = d.hh > 0 ? `<span title="Duración Total HH">⌛ Total: ${d.hh % 1 === 0 ? d.hh : d.hh.toFixed(1)} HH</span>` : '';
        let extraHtml = [durText, dotText, hhText].filter(Boolean).join('<span style="color:#cbd5e1; margin:0 4px;">|</span>');
        
        const extraDiv = document.getElementById('d_extra_info');
        if(extraHtml) { extraDiv.style.display = 'flex'; extraDiv.innerHTML = extraHtml; } 
        else extraDiv.style.display = 'none';

        let crit = d.criticidad; let pl = '';
        if(crit === 'Critica') pl='<span class="prio-flag p-crit">🚨 CRÍTICA</span>';
        else if(crit === 'Mayor') pl='<span class="prio-flag p-alta">🔴 MAYOR</span>';
        else if(crit === 'Menor') pl='<span class="prio-flag p-baja">🟢 MENOR</span>';
        else pl='<span class="prio-flag" style="background:#f1f5f9; color:#64748b; border:1px solid #cbd5e1;">⚪ S/A</span>';
        document.getElementById('d_prio_lbl').innerHTML = pl;

        const grid = document.getElementById('d_grid');
        grid.innerHTML = '';
        const createItem = (label, val) => `<div class="dg-item"><small>${label}</small><strong>${val||'--'}</strong></div>`;
        grid.innerHTML += createItem('🛠️ Clase MTTO', d.clase);
        grid.innerHTML += createItem('📍 Zona', d.zona);
        grid.innerHTML += createItem('👷 Responsable', d.ejecutor);
        grid.innerHTML += createItem('🏭 Línea / Área', d.ubicacion);
        grid.innerHTML += createItem('📌 Sub Ubicación', d.sub_ubi);
        grid.innerHTML += createItem('🟢 Levantamiento', d.f_lev);
        grid.innerHTML += createItem('🏁 Cierre', d.f_cie);
        grid.innerHTML += createItem('📆 Semana', d.semana);
        grid.innerHTML += createItem('🧾 OT SAP', d.ot);
        grid.innerHTML += createItem('⚠️ Criticidad', d.criticidad);
        grid.innerHTML += createItem('👥 Colaborador', d.colaborador);
        
        if (d.dia) grid.innerHTML += createItem('📅 Día Asignado', d.dia);
        let techsList = [d.tecnico, d.tecnico1, d.tecnico2].filter(t => t && t !== 'ALL').join(', ');
        if (techsList) grid.innerHTML += createItem('🦺 Técnicos (Códigos)', techsList);

        document.getElementById('box_obs1').style.display = 'block';
        document.getElementById('d_obs').innerText = d.observacion || 'Sin observaciones registradas.';
        if(d.obs2) { document.getElementById('box_obs2').style.display = 'block'; document.getElementById('d_obs2').innerText = d.obs2; }
        else { document.getElementById('box_obs2').style.display = 'none'; }
        
        const imgContainer = document.getElementById('d_img_container');
        let htmlImgs = '<div class="gallery-grid">';
        if (d.img_antes) htmlImgs += `<div class="gal-box"><span>📸 Antes</span><img src="${d.img_antes}" class="gal-img" onclick="openModal(this.src)"></div>`;
        else htmlImgs += `<div class="gal-box"><span>📸 Antes</span><div style="height:150px; display:flex; align-items:center; justify-content:center; color:#cbd5e1; font-style:italic; font-weight:600; font-size:0.9rem;">Sin foto</div></div>`;
        if (d.img_despues) htmlImgs += `<div class="gal-box"><span>📸 Después</span><img src="${d.img_despues}" class="gal-img" onclick="openModal(this.src)"></div>`;
        else htmlImgs += `<div class="gal-box"><span>📸 Después</span><div style="height:150px; display:flex; align-items:center; justify-content:center; color:#cbd5e1; font-style:italic; font-weight:600; font-size:0.9rem;">Sin foto</div></div>`;
        htmlImgs += '</div>';
        imgContainer.innerHTML = htmlImgs;
        document.getElementById('d_gallery_sec').style.display = 'flex';
    }

    function openModal(src) {
        document.getElementById('modalImg').src = src;
        document.getElementById('modal').style.display = 'flex';
    }

    let currentSortCol = -1; let currentSortDir = 'asc';
    function sortModalTable(colIndex, thElement) {
        const table = document.querySelector('.dm-table'); const tbody = table.querySelector('tbody'); const rows = Array.from(tbody.querySelectorAll('tr'));
        table.querySelectorAll('th').forEach(th => th.innerText = th.innerText.replace(/ [▼▲]/g, ' ↕'));
        if (currentSortCol === colIndex) currentSortDir = currentSortDir === 'asc' ? 'desc' : 'asc'; else { currentSortDir = 'asc'; currentSortCol = colIndex; }
        thElement.innerText = thElement.innerText.replace(' ↕', currentSortDir === 'asc' ? ' ▲' : ' ▼');
        rows.sort((a, b) => {
            const aCol = a.querySelectorAll('td')[colIndex]; const bCol = b.querySelectorAll('td')[colIndex];
            if(!aCol || !bCol || a.cells.length === 1) return 0;
            let aText = aCol.innerText.trim().toLowerCase(); let bText = bCol.innerText.trim().toLowerCase();
            let aNum = parseFloat(aText); let bNum = parseFloat(bText);
            if (!isNaN(aNum) && !isNaN(bNum)) return currentSortDir === 'asc' ? aNum - bNum : bNum - aNum;
            if (aText < bText) return currentSortDir === 'asc' ? -1 : 1;
            if (aText > bText) return currentSortDir === 'asc' ? 1 : -1;
            return 0;
        });
        rows.forEach(row => tbody.appendChild(row));
    }

    function showDataModal(title, filterFn, colProp = 'ubicacion') {
        let colHeader = colProp === 'clase' ? 'Clase de Actividad' : 'Ubicación';
        let html = `<div class="dm-header"><h3>📊 Desglose: ${title}</h3><button class="dm-close" onclick="document.getElementById('data_modal').style.display='none'">&times;</button></div>
        <div class="dm-body"><table class="dm-table"><thead><tr>
            <th onclick="sortModalTable(0, this)">OT / TAG ↕</th><th onclick="sortModalTable(1, this)">${colHeader} ↕</th><th onclick="sortModalTable(2, this)">Título / Actividad ↕</th><th onclick="sortModalTable(3, this)">Responsable ↕</th><th onclick="sortModalTable(4, this)">Estado ↕</th><th onclick="sortModalTable(5, this)">Observación ↕</th>
        </tr></thead><tbody>`;

        let datosFiltrados = currentChartData.filter(filterFn);
        datosFiltrados.sort((a, b) => { let valA = a[colProp] ? String(a[colProp]).toLowerCase() : ""; let valB = b[colProp] ? String(b[colProp]).toLowerCase() : ""; return valA.localeCompare(valB); });
        
        let found = datosFiltrados.length > 0;
        datosFiltrados.forEach(d => {
            let stColor = d.status === 'realizada' ? '#166534' : (d.status === 'pendiente' ? '#991b1b' : (d.status === 'en proceso' ? '#92400e' : '#075985'));
            let idDisplay = d.ot ? d.ot : (d.tag ? d.tag : '#' + d.id_real);
            let obsText = d.observacion ? (d.observacion.length > 45 ? d.observacion.substring(0, 42) + '...' : d.observacion) : '-';
            html += `<tr onclick="document.getElementById('data_modal').style.display='none'; document.getElementById('btn_tab_list').click(); setTimeout(() => renderDetail('${d.key_id}'), 100);">
                <td style="font-weight:700;">${idDisplay}</td><td>${d[colProp]||'-'}</td><td>${d.titulo}</td><td>${d.ejecutor.split(' ')[0]}</td><td style="color:${stColor}; font-weight:700; text-transform:uppercase;">${d.status}</td>
                <td style="max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${d.observacion}">${obsText}</td></tr>`;
        });
        if (!found) html += `<tr><td colspan="6" style="text-align:center; padding: 30px; color:var(--muted);">No hay OTs</td></tr>`;
        html += `</tbody></table></div>`;
        document.getElementById('data_modal_content').innerHTML = html;
        document.getElementById('data_modal').style.display = 'flex';
        currentSortCol = -1; currentSortDir = 'asc';
    }

    function getFreshCanvas(id) {
        const old = document.getElementById(id);
        if(!old) return null;
        const container = old.parentElement;
        container.innerHTML = `<canvas id="${id}"></canvas>`;
        return document.getElementById(id);
    }

    function descargarROW() {
        const btn = document.getElementById('btn_descargar_row');
        const originalText = btn.innerHTML;
        btn.innerHTML = "⏳ Generando Imagen...";
        html2canvas(document.getElementById('view_row'), { scale: 2, backgroundColor: "#f1f5f9" }).then(canvas => {
            let link = document.createElement('a'); link.download = 'Dashboard_ROW.png'; link.href = canvas.toDataURL('image/png'); link.click(); btn.innerHTML = originalText;
        }).catch(err => { alert("Error al capturar la pantalla."); btn.innerHTML = originalText; });
    }

    function descargarExcel() {
        if (!currentChartData || currentChartData.length === 0) return alert("No hay datos para exportar.");
        const datosExcel = currentChartData.map(d => ({
            "Levantamiento": d.f_lev, "Cierre": d.f_cie, "Actividad": d.actividad, "Planta": d.planta, "Clase": d.clase, "Zona": d.zona,
            "Ubicación": d.ubicacion, "Sub Ubicación": d.sub_ubi, "OT": d.ot, "Criticidad": d.criticidad, "Colaborador": d.colaborador,
            "Ejecutor": d.ejecutor, "Status": d.status.toUpperCase(), "Observación": d.observacion, "Semana": d.semana, "Tiempo Ejecución (h)": d.duracion,
            "Dotación": d.dotacion, "Duración Total (HH)": d.hh, "Día Planificado": d.dia, "Código Turno 1": d.tecnico, "Código Turno 2": d.tecnico1, "Código Turno 3": d.tecnico2
        }));
        const worksheet = XLSX.utils.json_to_sheet(datosExcel);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Base de Datos");
        XLSX.writeFile(workbook, `Reporte_MTTO_${new Date().toISOString().split('T')[0]}.xlsx`);
    }

    function descargarHTML() {
        var a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob(['<!DOCTYPE html>\\n' + document.documentElement.outerHTML], { type: 'text/html;charset=utf-8' }));
        a.download = 'Dashboard_Mantenimiento_' + new Date().toISOString().split('T')[0] + '.html';
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
    }

    const isAseoAct = (d) => { let l = (d.clase || '').toLowerCase(); return l.includes('aseo') || l.includes('limpieza') || l.includes('sanitizacion'); };
    const getPLoc = (d) => { let t = (d.ubicacion + " " + (d.sub_ubi || "") + " " + (d.titulo || "")).toLowerCase(); if(t.includes('l1') || t.includes('panadería 1') || t.includes('panaderia 1')) return 'L1'; if (t.includes('l2') || t.includes('panaderia 2')) return 'L2'; if (t.includes('l3') || t.includes('panaderia 3')) return 'L3'; if (t.includes('l4') || t.includes('panaderia 4')) return 'L4'; if (t.includes('l5') || t.includes('panaderia 5')) return 'L5'; return null; };
    const getDLoc = (d) => { let t = (d.ubicacion + " " + (d.sub_ubi || "") + " " + (d.titulo || "")).toLowerCase(); if(t.includes('pizza')) return 'Pizza'; if(t.includes('bolleri')) return 'Bolleria'; if(t.includes('empanada')) return 'Empanadas'; return null; };

    function drawCharts(data) {
        if(!data) return;
        let stats = { ok:0, pend:0, proc:0, prog:0, ex:{}, loc:{}, wCounts:{}, cCounts:{} };
        let totAseo = 0, okAseo = 0, hhAseo = 0; let totMtto = 0, okMtto = 0, hhMtto = 0; let totGen = data.length, okGen = 0, hhGen = 0;
        let statsArea = { 'Mecánico': { total: 0, ok: 0, proc: 0, pend: 0 }, 'Autómata': { total: 0, ok: 0, proc: 0, pend: 0 }, 'Frio': { total: 0, ok: 0, proc: 0, pend: 0 }, 'Infraestructura': { total: 0, ok: 0, proc: 0, pend: 0 } };

        weeks.forEach(w => stats.wCounts[w] = {total:0, ok:0});
        data.forEach(d => {
            let isOk = (d.status === 'realizada'); let isProc = (d.status === 'en proceso'); let isProg = (d.status === 'programado'); let isPend = (d.status === 'pendiente');
            let hhActual = parseFloat(d.hh) || 0; hhGen += hhActual;

            if(isOk) { stats.ok++; okGen++; } else if(isProg) { stats.prog++; } else if(isProc) { stats.proc++; } else { stats.pend++; }
            stats.cCounts[d.clase] = (stats.cCounts[d.clase]||0)+1;
            
            if(isAseoAct(d)) { totAseo++; if(isOk) okAseo++; hhAseo += hhActual; } else { totMtto++; if(isOk) okMtto++; hhMtto += hhActual; }

            const e = d.ejecutor || 'Sin Asignar'; if(!stats.ex[e]) stats.ex[e]={ok:0, proc:0, pend:0};
            if(isOk) stats.ex[e].ok++; else if(isProc) stats.ex[e].proc++; else stats.ex[e].pend++; 

            const l = d.ubicacion || 'Sin Ubicación'; if(!stats.loc[l]) stats.loc[l]=0; stats.loc[l]++;
            if(d.semana!=="S/N" && stats.wCounts[d.semana]) { stats.wCounts[d.semana].total++; if(isOk) stats.wCounts[d.semana].ok++; }
            
            let area = getAreaResp(d.ejecutor);
            if (statsArea[area]) { statsArea[area].total++; if (isOk) statsArea[area].ok++; else if (isProc) statsArea[area].proc++; else statsArea[area].pend++; }
        });

        let percAseo = totAseo > 0 ? Math.round((okAseo/totAseo)*100) : 0;
        let percMtto = totMtto > 0 ? Math.round((okMtto/totMtto)*100) : 0;
        let percGen = totGen > 0 ? Math.round((okGen/totGen)*100) : 0;

        let colAseo = percAseo >= 80 ? '#10b981' : (percAseo >= 40 ? '#f59e0b' : '#ef4444');
        let colMtto = percMtto >= 80 ? '#10b981' : (percMtto >= 40 ? '#f59e0b' : '#ef4444');
        let colGen = percGen >= 80 ? '#1d4ed8' : (percGen >= 40 ? '#f59e0b' : '#ef4444');

        document.getElementById('summary_content').innerHTML = `
            <div class="summary-block"><div class="summary-header"><span class="summary-title" style="color:#eab308;">🧹 Aseo / Sanitización y Seg.</span><span class="summary-perc" style="color:${colAseo};">${percAseo}%</span></div><div class="summary-sub">De un total de <b>${totAseo}</b>, <b>${okAseo}</b> realizadas</div><div class="summary-bar-bg"><div class="summary-bar-fill" style="width:${percAseo}%; background:${colAseo};"></div></div></div>
            <div class="summary-block"><div class="summary-header"><span class="summary-title" style="color:#8b5cf6;">🔧 Mantenimiento</span><span class="summary-perc" style="color:${colMtto};">${percMtto}%</span></div><div class="summary-sub">De un total de <b>${totMtto}</b>, <b>${okMtto}</b> realizadas</div><div class="summary-bar-bg"><div class="summary-bar-fill" style="width:${percMtto}%; background:${colMtto};"></div></div></div>
            <div class="summary-block" style="background:#eff6ff; border-color:#bfdbfe; text-align:center; padding: 20px 15px; margin-top: auto; margin-bottom: 0;"><div style="font-size:0.8rem; color:#1e40af; font-weight:700; text-transform:uppercase; margin-bottom:5px;">Cumplimiento Plan FDS Total</div><div style="font-size:2rem; font-weight:800; color:${colGen};">${percGen}%</div></div>
        `;

        const chartOpts = { maintainAspectRatio:false, responsive:true, animation: { duration: 1200, easing: 'easeOutQuart' }, layout: { padding: 10 } };
        new Chart(getFreshCanvas('chart1'), { type: 'doughnut', data: { labels:['Cerradas', 'En Proceso', 'Pendientes', 'Programadas'], datasets:[{ data:[stats.ok, stats.proc, stats.pend, stats.prog], backgroundColor:['#10b981', '#f59e0b', '#ef4444', '#3b82f6'], borderWidth: 2, borderColor: '#fff' }] }, options: { ...chartOpts, cutout: '65%', plugins: { legend: { position: 'bottom', labels: { padding: 20, usePointStyle: true } } } } });
        let cLabels = Object.keys(stats.cCounts);
        let baseColors = ['#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316', '#6366f1', '#10b981'];
        let cBgColors = cLabels.map((lbl, idx) => { let l = lbl.toLowerCase(); if(l.includes('aseo') || l.includes('sanit')) return '#eab308'; return baseColors[idx % baseColors.length]; });
        new Chart(getFreshCanvas('chart2'), { type: 'pie', data: { labels: cLabels, datasets:[{ data:Object.values(stats.cCounts), backgroundColor: cBgColors, borderWidth: 2, borderColor: '#fff' }] }, options: { ...chartOpts, plugins: { legend: { position: 'right', labels: { padding: 15, usePointStyle: true } } } } });

        let statsAreaHH = { 'Mecánico': { ok: 0, proc: 0, pend: 0, sin_hh: 0 }, 'Autómata': { ok: 0, proc: 0, pend: 0, sin_hh: 0 }, 'Frio': { ok: 0, proc: 0, pend: 0, sin_hh: 0 }, 'Infraestructura': { ok: 0, proc: 0, pend: 0, sin_hh: 0 } };
        data.forEach(d => { let a = getAreaResp(d.ejecutor); if (statsAreaHH[a]) { let hh = parseFloat(d.hh) || 0; if (hh > 0) { if (d.status === 'realizada') statsAreaHH[a].ok += hh; else if (d.status === 'en proceso') statsAreaHH[a].proc += hh; else statsAreaHH[a].pend += hh; } else { statsAreaHH[a].sin_hh += 1; } } });
        let labelsAreaHH = Object.keys(statsAreaHH).sort((a, b) => (statsAreaHH[b].ok + statsAreaHH[b].proc + statsAreaHH[b].pend) - (statsAreaHH[a].ok + statsAreaHH[a].proc + statsAreaHH[a].pend));
        new Chart(getFreshCanvas('chart_hh_area'), { type: 'bar', data: { labels: labelsAreaHH, datasets: [ { label: 'Pendientes (HH)', data: labelsAreaHH.map(l => statsAreaHH[l].pend), backgroundColor: '#ef4444' }, { label: 'En Proceso (HH)', data: labelsAreaHH.map(l => statsAreaHH[l].proc), backgroundColor: '#f59e0b' }, { label: 'Cerradas (HH)', data: labelsAreaHH.map(l => statsAreaHH[l].ok), backgroundColor: '#10b981' } ] }, options: { ...chartOpts, indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true } } } });

        const areaLabels = ['Mecánico', 'Autómata', 'Frio', 'Infraestructura'];
        new Chart(getFreshCanvas('chart5'), { type: 'bar', data: { labels: areaLabels, datasets: [ { label: 'Pendientes', data: areaLabels.map(l => statsArea[l].pend), backgroundColor: '#ef4444' }, { label: 'En Proceso', data: areaLabels.map(l => statsArea[l].proc), backgroundColor: '#f59e0b' }, { label: 'Cerradas', data: areaLabels.map(l => statsArea[l].ok), backgroundColor: '#10b981' } ] }, options: { ...chartOpts, indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true } } } });
        
        const sortedEx = Object.entries(stats.ex).sort((a,b)=>(b[1].ok+b[1].pend+b[1].proc)-(a[1].ok+a[1].pend+a[1].proc)).slice(0,12);
        new Chart(getFreshCanvas('chart3'), { type: 'bar', data: { labels: sortedEx.map(x=>x[0]), datasets: [ { label:'Pendientes', data:sortedEx.map(x=>x[1].pend), backgroundColor:'#ef4444'}, { label:'En Proceso', data:sortedEx.map(x=>x[1].proc), backgroundColor:'#f59e0b'}, { label:'Cerradas', data:sortedEx.map(x=>x[1].ok), backgroundColor:'#10b981'} ] }, options: { ...chartOpts, indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true } } } });

        const sortedLocs = Object.entries(stats.loc).sort((a,b)=>b[1]-a[1]).slice(0,12);
        new Chart(getFreshCanvas('chart4'), { type: 'bar', data: { labels: sortedLocs.map(x=>x[0]), datasets: [ { label: 'Total Hallazgos', data: sortedLocs.map(x=>x[1]), backgroundColor:'#8b5cf6' } ]}, options: { ...chartOpts, indexAxis: 'y' } });
    }

    function drawRowCharts(data) {
        if(!data) return;
        let stats = { mtto: { total: 0, ok: 0 }, aseo: { total: 0, ok: 0 }, panaderia: { 'L1': { mtto: {tot:0, ok:0} }, 'L2': { mtto: {tot:0, ok:0} }, 'L3': { mtto: {tot:0, ok:0} }, 'L4': { mtto: {tot:0, ok:0} }, 'L5': { mtto: {tot:0, ok:0} } }, dely: { 'Pizza': { mtto: {tot:0, ok:0} }, 'Bolleria': { mtto: {tot:0, ok:0} }, 'Empanadas': { mtto: {tot:0, ok:0} } } };
        data.forEach(d => {
            let isOk = (d.status === 'realizada'); let isAseo = isAseoAct(d); let isMtto = !isAseo; 
            if (isAseo) { stats.aseo.total++; if(isOk) stats.aseo.ok++; } if (isMtto) { stats.mtto.total++; if(isOk) stats.mtto.ok++; }
            let pLoc = getPLoc(d); if (pLoc && isMtto) { stats.panaderia[pLoc].mtto.tot++; if(isOk) stats.panaderia[pLoc].mtto.ok++; }
            let dLoc = getDLoc(d); if (dLoc && isMtto) { stats.dely[dLoc].mtto.tot++; if(isOk) stats.dely[dLoc].mtto.ok++; }
        });
        const getPerc = (ok, tot) => tot > 0 ? Math.round((ok/tot)*100) : 0;
        const cOpts = { maintainAspectRatio: false, responsive: true, plugins: { legend: { position: 'top' } } };
        let totalAct = stats.mtto.total + stats.aseo.total;
        new Chart(getFreshCanvas('row_chart1'), { type: 'pie', data: { labels: ['Mantenimiento', 'Aseo'], datasets: [{ data: [getPerc(stats.mtto.total, totalAct), getPerc(stats.aseo.total, totalAct)], backgroundColor: ['#3b82f6', '#eab308'] }] }, options: cOpts });
        new Chart(getFreshCanvas('row_chart2'), { type: 'bar', data: { labels: ['Cumplimiento MTTO'], datasets: [{ label: 'Cerradas', data: [getPerc(stats.mtto.ok, stats.mtto.total)], backgroundColor: '#3b82f6' }] }, options: { ...cOpts, indexAxis: 'y', scales: { x: { max: 100 } } } });
        new Chart(getFreshCanvas('row_chart3'), { type: 'bar', data: { labels: ['Cumpl. Aseo'], datasets: [{ label: 'Cerradas', data: [getPerc(stats.aseo.ok, stats.aseo.total)], backgroundColor: '#eab308' }] }, options: { ...cOpts, indexAxis: 'y', scales: { x: { max: 100 } } } });
        const pLabels = ['L1', 'L2', 'L3', 'L4', 'L5']; new Chart(getFreshCanvas('row_chart4'), { type: 'bar', data: { labels: pLabels, datasets: [ { label: '% Cumpl. Mtto', data: pLabels.map(l => getPerc(stats.panaderia[l].mtto.ok, stats.panaderia[l].mtto.tot)), backgroundColor: '#3b82f6' } ] }, options: { ...cOpts, indexAxis: 'y', scales: { x: { max: 100 } } } });
        const dLabels = ['Pizza', 'Bolleria', 'Empanadas']; new Chart(getFreshCanvas('row_chart5'), { type: 'bar', data: { labels: dLabels, datasets: [ { label: '% Cumpl. Mtto', data: dLabels.map(l => getPerc(stats.dely[l].mtto.ok, stats.dely[l].mtto.tot)), backgroundColor: '#3b82f6' } ] }, options: { ...cOpts, indexAxis: 'y', scales: { x: { max: 100 } } } });
    }

    function drawGantt(data) {
        const container = document.getElementById('gantt_container');
        container.innerHTML = '';
        if (!data || data.length === 0) return;
        const diasOrden = ['Jueves', 'Viernes', 'Sabado', 'Domingo', 'Lunes', 'Martes', 'Miercoles'];
        let schedule = {};
        diasOrden.forEach(d => { schedule[d] = { totalHH: 0, turnos: { 'Mañana': { hh: 0, act: [] }, 'Tarde': { hh: 0, act: [] }, 'Noche': { hh: 0, act: [] }, 'Cuarto Turno': { hh: 0, act: [] }, 'Sin Turno Asignado': { hh: 0, act: [] } } }; });
        const getShift = (t1, t2, t3) => { let c = (String(t1) + " " + String(t2) + " " + String(t3)).toUpperCase(); if (/(AM|MM|LM|EM)/.test(c)) return 'Mañana'; if (/(AT|MT)/.test(c)) return 'Tarde'; if (/(AN|MN)/.test(c)) return 'Noche'; if (/(M4)/.test(c)) return 'Cuarto Turno'; return 'Sin Turno Asignado'; };
        const normDia = (d) => { if(!d || d === 'ALL') return 'Sin Día Asignado'; let l = String(d).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); if(l.includes('lun')) return 'Lunes'; if(l.includes('mar')) return 'Martes'; if(l.includes('mier')) return 'Miercoles'; if(l.includes('jue')) return 'Jueves'; if(l.includes('vie')) return 'Viernes'; if(l.includes('sab')) return 'Sabado'; if(l.includes('dom')) return 'Domingo'; return 'Sin Día Asignado'; };
        
        data.forEach(item => {
            let d = normDia(item.dia);
            if (!schedule[d]) { schedule[d] = { totalHH: 0, turnos: { 'Mañana': { hh: 0, act: [] }, 'Tarde': { hh: 0, act: [] }, 'Noche': { hh: 0, act: [] }, 'Cuarto Turno': { hh: 0, act: [] }, 'Sin Turno Asignado': { hh: 0, act: [] } } }; if (!diasOrden.includes(d)) diasOrden.push(d); }
            let shift = getShift(item.tecnico, item.tecnico1, item.tecnico2);
            let hh = parseFloat(item.hh) || 0;
            schedule[d].totalHH += hh; schedule[d].turnos[shift].hh += hh; schedule[d].turnos[shift].act.push(item);
        });

        let htmlFinal = "";
        diasOrden.forEach(dia => {
            let dayData = schedule[dia]; if(!dayData) return;
            if (Object.values(dayData.turnos).every(t => t.act.length === 0) && (dia === 'Sabado' || dia === 'Domingo' || dia === 'Sin Día Asignado')) return; 
            htmlFinal += `<div class="gantt-day-col"><div class="gantt-day-header"><h3 class="gantt-day-title">${dia}</h3><span class="gantt-day-hh">Total Carga: ${dayData.totalHH.toFixed(1)} HH</span></div><div style="padding:15px; display:flex; flex-direction:column; overflow-y:auto; flex:1; background:#f8fafc;">`;
            ['Mañana', 'Tarde', 'Noche', 'Cuarto Turno', 'Sin Turno Asignado'].forEach(turno => {
                let tData = dayData.turnos[turno];
                if (tData.act.length > 0) {
                    let tColor = turno === 'Mañana' ? '#3b82f6' : (turno === 'Tarde' ? '#f59e0b' : (turno === 'Noche' ? '#8b5cf6' : (turno === 'Cuarto Turno' ? '#ec4899' : '#94a3b8')));
                    htmlFinal += `<div class="gantt-shift-box" style="border:1px solid ${tColor}40; background:${tColor}10;"><div class="gantt-shift-header" style="color:${tColor}; border-bottom:1px solid ${tColor}40;"><span>${turno}</span><span>${tData.hh.toFixed(1)} HH</span></div>`;
                    tData.act.forEach(a => {
                        let statusColor = a.status === 'realizada' ? '#10b981' : (a.status === 'en proceso' ? '#f59e0b' : '#ef4444');
                        htmlFinal += `<div class="gantt-card" style="border-left-color:${statusColor};" onclick="document.getElementById('btn_tab_list').click(); setTimeout(() => renderDetail('${a.key_id}'), 100);"><div style="font-weight:700; color:var(--primary); margin-bottom:4px; display:flex; justify-content:space-between;"><span>${a.ot || a.tag || '#'+a.id_real}</span><span style="font-size:0.7rem; color:${statusColor};">${a.status.toUpperCase()}</span></div><div style="color:var(--secondary); margin-bottom:6px; line-height:1.3;">${a.titulo}</div></div>`;
                    });
                    htmlFinal += `</div>`;
                }
            });
            htmlFinal += `</div></div>`;
        });
        container.innerHTML = htmlFinal;
    }

    // ==========================================
    // PLAN MATRIZ (Integración SP + SAP)
    // ==========================================
    window.setPmClass = function(c, btn) { 
        currentPmClass = c; 
        document.querySelectorAll('.tablinks-pm').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active');
        drawGanttPM(); 
    };
    window.setPmYear = function(y) { currentPmYear = y; drawGanttPM(); };

    window.showModalPM = function(encodedJson) {
        let ots = JSON.parse(decodeURIComponent(encodedJson));
        let html = `<div class="dm-header">
            <h3>📊 Desglose de Actividades</h3>
            <button class="dm-close" onclick="document.getElementById('data_modal').style.display='none'">&times;</button>
        </div>
        <div class="dm-body" style="padding:15px;"><table class="dm-table"><thead><tr>
            <th>OT / ID</th><th>Equipo / Ubicación</th><th>Actividad</th><th>Fechas / Sem</th><th>Estado</th>
        </tr></thead><tbody>`;

        ots.forEach(o => {
            let stColor = o.status === 'realizada' ? '#166534' : (o.status === 'pendiente' ? '#991b1b' : (o.status === 'en proceso' ? '#92400e' : '#075985'));
            let clicAction = o.key_id ? `onclick="document.getElementById('data_modal').style.display='none'; document.getElementById('btn_tab_list').click(); setTimeout(() => renderDetail('${o.key_id}'), 100);"` : "";
            
            html += `<tr ${clicAction}>
                <td style="font-weight:700;">${o.ot}</td>
                <td>${o.equipo}<br><small style="color:var(--muted);">${o.ubicacion}</small></td>
                <td>${o.actividad}</td>
                <td>${formatDateDDMMYY(o.fecha_prog)}</td>
                <td style="color:${stColor}; font-weight:700; text-transform:uppercase;">${o.status}</td>
            </tr>`;
        });
        html += `</tbody></table></div>`;
        document.getElementById('data_modal_content').innerHTML = html;
        document.getElementById('data_modal').style.display = 'flex';
    };

    window.filtrarGanttPM = function() {
        let input = document.getElementById("search_pm").value.toUpperCase();
        let tr = document.querySelectorAll('#pm_table_render tbody tr');
        
        for (let i = 0; i < tr.length; i++) {
            let td = tr[i].getElementsByTagName("td")[0]; // Ubicacion/Equipo
            let hiddenData = tr[i].getAttribute('data-search') || ""; // Datos Ocultos (OT, TAG)

            if (td) {
                let textToSearch = ((td.textContent || td.innerText) + " " + hiddenData).toUpperCase();
                let isMatch = textToSearch.indexOf(input) > -1;
                tr[i].dataset.isMatch = isMatch;
                tr[i].style.display = isMatch ? "" : "none";
            }
        }
        
        if(input !== "") {
            let currentLocIndex = -1;
            let showChildrenOfLoc = false;
            
            for (let i = 0; i < tr.length; i++) {
                if (tr[i].classList.contains('location-row')) {
                    currentLocIndex = i;
                    showChildrenOfLoc = (tr[i].dataset.isMatch === 'true');
                } else if (tr[i].classList.contains('eq-row')) {
                    if (tr[i].dataset.isMatch === 'true' && currentLocIndex !== -1) {
                        tr[currentLocIndex].style.display = "";
                    } else if (showChildrenOfLoc) {
                        tr[i].style.display = "";
                    }
                }
            }
        }
    };

    function drawGanttPM() {
        const container = document.getElementById('gantt_pm_container');
        
        // Obtener el máximo de semanas dinámicamente
        let maxSemanas = 52;
        if(weeks && weeks.length > 0) {
            let semInts = weeks.map(w => parseInt(w)).filter(n => !isNaN(n));
            if(semInts.length > 0) {
                let maxEncontrada = Math.max(...semInts);
                if(maxEncontrada > maxSemanas) maxSemanas = maxEncontrada;
            }
        }

        let html = `
            <table class="pm-table" id="pm_table_render">
                <thead>
                    <tr>
                        <th class="pm-tag-col">📍 UBICACIÓN / ↳ EQUIPO</th>
                        ${Array.from({length:maxSemanas}, (_,i)=>`<th>S${i+1}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
        `;

        const dataAnio = matrizSAP[currentPmYear] || {};
        const dataClase = dataAnio[currentPmClass] || {};
        const ubicacionesSAP = Object.keys(dataClase).sort();

        let spOtsBySapLoc = {};
        currentChartData.forEach(d => {
            let ubiSP = (d.ubicacion || "").toUpperCase();
            let sapLoc = mapaSpSap[ubiSP];
            if (sapLoc && d.semana && d.semana !== "S/N") {
                if(!spOtsBySapLoc[sapLoc]) spOtsBySapLoc[sapLoc] = {};
                if(!spOtsBySapLoc[sapLoc][d.semana]) spOtsBySapLoc[sapLoc][d.semana] = [];
                spOtsBySapLoc[sapLoc][d.semana].push(d);
            }
        });

        ubicacionesSAP.forEach(ubi => {
            let spOts = spOtsBySapLoc[ubi] || {};
            
            // Recolectar datos ocultos para búsqueda en SP
            let hiddenSearchSP = [];
            Object.values(spOts).forEach(semanaArr => {
                semanaArr.forEach(ot => {
                    hiddenSearchSP.push(ot.ot || ot.key_id);
                    hiddenSearchSP.push(ot.tag || "");
                });
            });
            let dataSearchAttrSP = hiddenSearchSP.join(" ");

            html += `<tr class="location-row" data-search="${dataSearchAttrSP}"><td class="pm-tag-col loc-bg" title="${ubi}">📍 ${ubi}</td>`;
            
            for(let sem=1; sem<=maxSemanas; sem++) {
                let otsSemana = spOts[sem.toString()];
                if (otsSemana && otsSemana.length > 0) {
                    let st = otsSemana[0].status; 
                    let css = st==='realizada'?'pm-ok':(st==='en proceso'?'pm-proc':'pm-pend');
                    let icon = st==='realizada'?'✓':'!';
                    let multi = otsSemana.length > 1 ? '<div class="pm-multi">+</div>' : '';
                    
                    let dataJson = encodeURIComponent(JSON.stringify(otsSemana.map(o=>({
                        equipo: ubi + " (Actividad de SharePoint)",
                        ubicacion: o.ubicacion,
                        actividad: o.titulo,
                        ot: o.ot || "SP_ID: " + o.id_real,
                        fecha_prog: "Semana " + o.semana,
                        status: o.status,
                        key_id: o.key_id
                    }))));
                    html += `<td class="loc-bg"><div class="pm-cell-inner"><div class="${css}" onclick="showModalPM('${dataJson}')">${icon}${multi}</div></div></td>`;
                } else {
                    html += `<td class="loc-bg"></td>`;
                }
            }
            html += `</tr>`;

            let equipos = Object.keys(dataClase[ubi]).sort();
            equipos.forEach(eq => {
                let eqSemanas = dataClase[ubi][eq];
                
                // Recolectar datos ocultos para búsqueda en SAP
                let hiddenSearchSAP = [];
                Object.values(eqSemanas).forEach(semanaArr => {
                    semanaArr.forEach(ot => hiddenSearchSAP.push(ot.ot || ""));
                });
                let dataSearchAttrSAP = hiddenSearchSAP.join(" ");

                html += `<tr class="eq-row" data-search="${dataSearchAttrSAP}"><td class="pm-tag-col" title="${eq}">↳ ${eq}</td>`;
                
                for(let sem=1; sem<=maxSemanas; sem++) {
                    let otsSemana = eqSemanas[sem.toString()];
                    if (otsSemana && otsSemana.length > 0) {
                        let st = otsSemana[0].status;
                        let css = st==='realizada'?'pm-ok':(st==='en proceso'?'pm-proc':'pm-pend');
                        let icon = st==='realizada'?'✓':'!';
                        let multi = otsSemana.length > 1 ? '<div class="pm-multi">+</div>' : '';

                        let dataJson = encodeURIComponent(JSON.stringify(otsSemana.map(o=>({
                            equipo: eq,
                            ubicacion: ubi,
                            actividad: o.titulo,
                            ot: o.ot,
                            fecha_prog: o.fecha_prog,
                            status: o.status
                        }))));
                        html += `<td><div class="pm-cell-inner"><div class="${css}" onclick="showModalPM('${dataJson}')">${icon}${multi}</div></div></td>`;
                    } else {
                        html += `<td></td>`;
                    }
                }
                html += `</tr>`;
            });
        });

        html += `</tbody></table>`;
        container.innerHTML = html;
        filtrarGanttPM(); 
    }

    const initAntigravity = () => {
        const canvas = document.createElement('canvas');
        canvas.id = 'antigravity-bg'; document.body.prepend(canvas); const ctx = canvas.getContext('2d');
        canvas.style.position = 'fixed'; canvas.style.top = '0'; canvas.style.left = '0'; canvas.style.width = '100vw'; canvas.style.height = '100vh'; canvas.style.zIndex = '-1'; canvas.style.pointerEvents = 'none'; canvas.style.backgroundColor = '#f8fafc';
        let particles = []; const colors = ['#4285F4', '#EA4335', '#FBBC05', '#34A853', '#A0C3FF', '#FCA297']; let mouse = { x: null, y: null, radius: 120 };
        window.addEventListener('mousemove', (e) => { mouse.x = e.x; mouse.y = e.y; });
        window.addEventListener('mouseout', () => { mouse.x = undefined; mouse.y = undefined; });
        window.addEventListener('resize', () => { canvas.width = window.innerWidth; canvas.height = window.innerHeight; initParticles(); });

        class Particle {
            constructor(x, y) { this.x = x; this.y = y; this.baseX = x; this.baseY = y; this.size = Math.random() * 2 + 1.5; this.color = colors[Math.floor(Math.random() * colors.length)]; this.density = (Math.random() * 20) + 2; }
            draw() { ctx.fillStyle = this.color; ctx.beginPath(); ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2); ctx.closePath(); ctx.fill(); }
            update() { let dx = mouse.x - this.x; let dy = mouse.y - this.y; let distance = Math.sqrt(dx * dx + dy * dy); if (distance < mouse.radius) { let forceDirectionX = dx / distance; let forceDirectionY = dy / distance; let force = (mouse.radius - distance) / mouse.radius; let directionX = forceDirectionX * force * this.density; let directionY = forceDirectionY * force * this.density; this.x -= directionX; this.y -= directionY; } else { if (this.x !== this.baseX) { let dx = this.x - this.baseX; this.x -= dx / 15; } if (this.y !== this.baseY) { let dy = this.y - this.baseY; this.y -= dy / 15; } } this.draw(); }
        }
        function initParticles() { particles = []; canvas.width = window.innerWidth; canvas.height = window.innerHeight; let numberOfParticles = (canvas.width * canvas.height) / 7000; for (let i = 0; i < numberOfParticles; i++) { let x = Math.random() * canvas.width; let y = Math.random() * canvas.height; particles.push(new Particle(x, y)); } }
        function animateParticles() { ctx.clearRect(0, 0, canvas.width, canvas.height); for (let i = 0; i < particles.length; i++) { particles[i].update(); } requestAnimationFrame(animateParticles); }
        initParticles(); animateParticles();
    };

    window.onload = () => { buildFilters(); applyFilters(); initAntigravity(); };
    </script>
</body></html>"""

    full_html = html_template.replace("__FECHA_ACTUAL__", fecha_actual)
    full_html = full_html.replace("__DB_JSON_DATA__", json.dumps(db_json))
    full_html = full_html.replace("__MATRIZ_SAP_JSON__", json.dumps(matriz_sap_json))
    full_html = full_html.replace("__MAPA_SP_SAP__", json.dumps(mapa_sp_sap))
    full_html = full_html.replace('\xa0', ' ')
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: 
        f.write(full_html)
        
    print(f"\n✅ REPORTE GENERADO CON ÉXITO: {OUTPUT_HTML}")

if __name__ == "__main__":
    main()
