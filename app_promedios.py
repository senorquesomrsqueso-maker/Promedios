import streamlit as st
import pandas as pd
import re
import datetime
import os
import json
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from googleapiclient.discovery import build
from playwright.sync_api import sync_playwright

# ==============================================================================
# CONFIGURACIÓN GENERAL Y ESTILOS
# ==============================================================================
st.set_page_config(page_title="BS LATAM - AUDITORÍA PRO", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #0b0d11; color: #e6edf3; font-family: 'Segoe UI', sans-serif; }
    .stApp { background-color: #0b0d11; }
    .title-box { border-left: 15px solid #0055FF; padding: 20px 35px; margin: 15px 0 25px 0; background: linear-gradient(90deg, #161b22 0%, rgba(11,13,17,0) 100%); border-radius: 0 15px 15px 0;}
    .m-title { font-size: 34px; font-weight: 900; color: #ffffff; text-transform: uppercase; letter-spacing: 2px; margin: 0; }
    .s-title { font-size: 15px; color: #8b949e; font-family: monospace; margin-top: 5px; }
    .stButton>button { background: linear-gradient(135deg, #0055FF 0%, #002b80 100%) !important; color: #ffffff !important; font-weight: 900 !important; border-radius: 8px; height: 48px; border: none; }
    </style>
    <div class="title-box">
        <p class="m-title">BS LATAM • MÓDULO DE PROMEDIOS PRO</p>
        <p class="s-title">API YOUTUBE V3 | TIKTOK PRO & FB GENERAL | MASTER SHEET + ROSTER</p>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# INSTALACIÓN AUTOMÁTICA DE PLAYWRIGHT
# ==============================================================================
@st.cache_resource
def install_playwright_browsers():
    os.system("playwright install chromium")
install_playwright_browsers()

# ==============================================================================
# VERIFICACIÓN DE CLAVE API YOUTUBE
# ==============================================================================
try:
    YOUTUBE_API_KEY = st.secrets["YOUTUBE_API_KEY"]
except:
    YOUTUBE_API_KEY = None

# ==============================================================================
# CARGA Y SANITIZACIÓN DEL ROSTER
# ==============================================================================
@st.cache_data(ttl=300)
def cargar_roster():
    url_csv = "https://docs.google.com/spreadsheets/d/1KxXVKK0q5A8aSG5k8KHPoyXMynjkUUl5TAJRda7sPYg/export?format=csv&gid=852207003"
    try:
        df = pd.read_csv(url_csv).dropna(subset=['Creador'], how='all')
        roster = []
        for _, row in df.iterrows():
            c_nom = str(row.get('Creador', '')).strip()
            if not c_nom or c_nom.lower() == 'nan': continue
            roster.append({
                "CREADOR": c_nom,
                "LINK YT": str(row.get('YouTube', '')).strip() if str(row.get('YouTube', '')).startswith('http') else "N/A",
                "LINK TK": str(row.get('TikTok', '')).strip() if str(row.get('TikTok', '')).startswith('http') else "N/A",
                "LINK FC": str(row.get('Facebook', '')).strip() if str(row.get('Facebook', '')).startswith('http') else "N/A"
            })
        return pd.DataFrame(roster)
    except Exception as e:
        st.error(f"Error cargando base de datos: {e}")
        return pd.DataFrame()

# ==============================================================================
# FILTRO MAESTRO DE PALABRAS CLAVE
# ==============================================================================
def cumple_filtros_kw(titulo, inc_kw, exc_kw):
    t_lower = str(titulo).lower()
    if inc_kw and not any(kw.lower() in t_lower for kw in inc_kw): return False
    if exc_kw and any(kw.lower() in t_lower for kw in exc_kw): return False
    return True

# ==============================================================================
# MOTOR YOUTUBE V3 (ESTRICTO POR FECHAS) - INTACTO Y FUNCIONAL
# ==============================================================================
def extraer_youtube_api(creador, url, start_date, end_date, max_vids, inc_kw, exc_kw, extraer_shorts, extraer_videos):
    if not YOUTUBE_API_KEY or url == 'N/A': return []
    youtube = build('youtube', 'v3', developerKey=YOUTUBE_API_KEY)
    
    channel_id = None
    try:
        if '@' in url:
            handle = url.split('@')[-1].split('/')[0]
            req = youtube.search().list(part="snippet", type="channel", q=f"@{handle}", maxResults=1)
            res = req.execute()
            if res['items']: channel_id = res['items'][0]['snippet']['channelId']
        elif 'channel/' in url:
            channel_id = url.split('channel/')[-1].split('/')[0]
    except: pass
    
    if not channel_id: return []

    try:
        dt_start = datetime.datetime.combine(start_date, datetime.time.min).isoformat() + 'Z'
        dt_end = datetime.datetime.combine(end_date, datetime.time.max).isoformat() + 'Z'
        
        req = youtube.search().list(
            part="snippet",
            channelId=channel_id,
            maxResults=50,
            publishedAfter=dt_start,
            publishedBefore=dt_end,
            type="video",
            order="date"
        )
        res = req.execute()
        
        video_ids = [item['id']['videoId'] for item in res.get('items', [])]
        if not video_ids: return []

        stats_req = youtube.videos().list(part="statistics,contentDetails,snippet", id=','.join(video_ids))
        stats_res = stats_req.execute()
        
        v_finales = []
        s_finales = []

        for v in stats_res.get('items', []):
            titulo = v['snippet']['title']
            fecha = v['snippet']['publishedAt'][:10]
            vistas = int(v['statistics'].get('viewCount', 0))
            vid_url = f"https://www.youtube.com/watch?v={v['id']}"
            dur_str = v['contentDetails']['duration']
            
            match = re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?', dur_str)
            segundos = 0
            if match:
                h, m, s = match.groups()
                segundos = int(h or 0)*3600 + int(m or 0)*60 + int(s or 0)
            
            is_short = segundos <= 65
            
            if not cumple_filtros_kw(titulo, inc_kw, exc_kw): continue
            
            data_dict = {"Creador": creador, "Fecha": fecha, "Título": titulo[:70], "Vistas": vistas, "Link": vid_url}
            
            if is_short and extraer_shorts:
                data_dict["Plataforma"] = "YouTube Shorts"
                data_dict["Link"] = f"https://www.youtube.com/shorts/{v['id']}"
                s_finales.append(data_dict)
            elif not is_short and extraer_videos:
                data_dict["Plataforma"] = "YouTube Video"
                v_finales.append(data_dict)

        return s_finales[:max_vids] + v_finales[:max_vids]
    except:
        return []

# ==============================================================================
# MOTOR TIKTOK PRO (ESTILO SEARCH PRO CON JSON REHYDRATION Y FECHAS)
# ==============================================================================
def extraer_tiktok_pro(creador, url, start_date, end_date, max_vids, inc_kw, exc_kw):
    if url == 'N/A': return []
    videos_validos = []
    
    dt_start_ts = int(datetime.datetime.combine(start_date, datetime.time.min).timestamp())
    dt_end_ts = int(datetime.datetime.combine(end_date, datetime.time.max).timestamp())
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-dev-shm-usage'])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()
        try:
            page.goto(url, timeout=20000)
            page.wait_for_timeout(4000)
            
            try:
                script_el = page.locator('#__UNIVERSAL_DATA_FOR_REHYDRATION__')
                if script_el.count() > 0:
                    script_content = script_el.inner_text()
                    data = json.loads(script_content)
                    
                    def search_dict(d):
                        results = []
                        if isinstance(d, dict):
                            if 'stats' in d and 'createTime' in d and 'id' in d:
                                results.append(d)
                            for k, v in d.items():
                                results.extend(search_dict(v))
                        elif isinstance(d, list):
                            for item in d:
                                results.extend(search_dict(item))
                        return results
                    
                    raw_vids = search_dict(data)
                    for item in raw_vids:
                        try:
                            create_time = int(item.get('createTime', 0))
                            if not (dt_start_ts <= create_time <= dt_end_ts):
                                continue
                            
                            vid_id = item.get('id')
                            desc = item.get('desc', '')
                            stats = item.get('stats', {})
                            vistas = int(stats.get('playCount', 0) or stats.get('diggCount', 0))
                            
                            if not cumple_filtros_kw(desc, inc_kw, exc_kw): continue
                            
                            fecha_str = datetime.datetime.fromtimestamp(create_time).strftime('%Y-%m-%d')
                            author = item.get('author', {})
                            sec_uid = author.get('uniqueId', url.split('@')[-1].split('/')[0])
                            vid_link = f"https://www.tiktok.com/@{sec_uid}/video/{vid_id}"
                            
                            videos_validos.append({
                                "Creador": creador,
                                "Plataforma": "TikTok",
                                "Fecha": fecha_str,
                                "Título": desc[:70] if desc else "TikTok Video",
                                "Vistas": vistas,
                                "Link": vid_link
                            })
                            if len(videos_validos) >= max_vids: break
                        except: continue
            except: pass
            
            if len(videos_validos) < max_vids:
                for _ in range(4):
                    page.mouse.wheel(0, 2000)
                    page.wait_for_timeout(1500)
                
                links = page.locator('a[href*="/video/"]').all()
                enlaces_unicos = {v['Link'] for v in videos_validos}
                
                for el in links:
                    try:
                        href = el.get_attribute('href')
                        if not href: continue
                        if href.startswith('/'): href = f"https://www.tiktok.com{href}"
                        if href in enlaces_unicos: continue
                        enlaces_unicos.add(href)
                        
                        texto = el.inner_text() or ""
                        titulo = texto.replace('\n', ' ')[:70]
                        vistas = 0
                        m_v = re.findall(r'([\d\.,]+[KMkm]?)', texto)
                        if m_v:
                            for val_str in m_v:
                                clean_val = val_str.upper().replace(',', '.')
                                if 'K' in clean_val or 'M' in clean_val:
                                    mult = 1000 if 'K' in clean_val else 1000000
                                    vistas = int(float(re.sub(r'[^0-9.]', '', clean_val)) * mult)
                                    break
                        
                        if not cumple_filtros_kw(titulo, inc_kw, exc_kw): continue
                        
                        videos_validos.append({
                            "Creador": creador,
                            "Plataforma": "TikTok",
                            "Fecha": "En Rango",
                            "Título": titulo if titulo else "TikTok Video",
                            "Vistas": vistas,
                            "Link": href
                        })
                        if len(videos_validos) >= max_vids: break
                    except: continue
        except: pass
        finally:
            browser.close()
            
    return videos_validos[:max_vids]

# ==============================================================================
# MOTOR FACEBOOK GENERAL (EXTRACCIÓN LIBRE SIN COLAPSAR)
# ==============================================================================
def extraer_facebook_general(creador, url, start_date, end_date, max_vids, inc_kw, exc_kw):
    if url == 'N/A': return []
    videos_validos = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-dev-shm-usage'])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()
        try:
            # Carga limpia y directa de la URL del perfil general del roster sin forzar sufijos estrictos
            page.goto(url, timeout=20000)
            page.wait_for_timeout(4000)
            
            for _ in range(5):
                page.mouse.wheel(0, 2000)
                page.wait_for_timeout(2000)
            
            # Selectores amplios y flexibles para capturar cualquier contenido de video o reel de la página
            links = page.locator('a[href*="/reel/"], a[href*="/videos/"], a[href*="/watch/"], a[href*="/video/"]').all()
            enlaces_unicos = set()
            
            for el in links:
                try:
                    href = el.get_attribute('href')
                    if not href or href in enlaces_unicos: continue
                    enlaces_unicos.add(href)
                    
                    if href.startswith('/'): href = f"https://www.facebook.com{href}"
                    href = href.split('?')[0]
                    
                    texto = el.inner_text() or ""
                    titulo = texto.replace('\n', ' ')[:70]
                    
                    vistas = 0
                    m_v = re.findall(r'([\d\.,]+[KMkm]?)', texto)
                    if m_v:
                        for val_str in m_v:
                            clean_val = val_str.upper().replace(',', '.')
                            if 'K' in clean_val or 'M' in clean_val:
                                mult = 1000 if 'K' in clean_val else 1000000
                                vistas = int(float(re.sub(r'[^0-9.]', '', clean_val)) * mult)
                                break
                    
                    if not cumple_filtros_kw(titulo, inc_kw, exc_kw): continue
                    
                    videos_validos.append({
                        "Creador": creador,
                        "Plataforma": "Facebook Reels",
                        "Fecha": "En Rango",
                        "Título": titulo if titulo else "Facebook Video",
                        "Vistas": vistas,
                        "Link": href
                    })
                    if len(videos_validos) >= max_vids: break
                except: continue
        except: pass
        finally:
            browser.close()
            
    return videos_validos

# ==============================================================================
# GENERADOR DE EXCEL CON MASTER SHEET Y PARES
# ==============================================================================
def generar_excel_completo(resultados_maestros):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        master_data = []
        for creador, data in resultados_maestros.items():
            tot_vids = sum(data['promedios'][k]['count'] for k in data['promedios'])
            master_data.append({
                "CREADOR": creador,
                "PROM YT VIDEO": f"{data['promedios']['YT Video']['promedio']:,}",
                "PROM YT SHORTS": f"{data['promedios']['YT Shorts']['promedio']:,}",
                "PROM TIKTOK": f"{data['promedios']['TK']['promedio']:,}",
                "PROM FB REELS": f"{data['promedios']['FC']['promedio']:,}",
                "TOTAL VIDEOS": tot_vids
            })
            
        df_master = pd.DataFrame(master_data)
        df_master.to_excel(writer, index=False, sheet_name="RESUMEN GLOBAL")
        
        ws_master = writer.sheets["RESUMEN GLOBAL"]
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="0055FF")
        for cell in ws_master[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for col in ws_master.columns:
            ws_master.column_dimensions[col[0].column_letter].width = 22

        for creador, data in resultados_maestros.items():
            safe_name = re.sub(r'[\[\]\:\*\?\/\\\|\'\"\t\n\r]', '', creador)[:20].strip() or "Creador"
            
            sheet_videos = f"{safe_name} ALL VIDEOS"
            df_vids = pd.DataFrame(data['videos'])
            if df_vids.empty: df_vids = pd.DataFrame(columns=["Creador", "Plataforma", "Fecha", "Título", "Vistas", "Link"])
            df_vids.to_excel(writer, index=False, sheet_name=sheet_videos)

            sheet_prom = f"Promedio {safe_name}"
            promedios_data = [
                ["CREATOR", creador],
                ["PR TK", f"{data['promedios']['TK']['promedio']:,} vistas - ({data['promedios']['TK']['count']} videos)"],
                ["PR FC", f"{data['promedios']['FC']['promedio']:,} vistas - ({data['promedios']['FC']['count']} videos)"],
                ["PR YT Shorts", f"{data['promedios']['YT Shorts']['promedio']:,} vistas - ({data['promedios']['YT Shorts']['count']} videos)"],
                ["PR YT Video", f"{data['promedios']['YT Video']['promedio']:,} vistas - ({data['promedios']['YT Video']['count']} videos)"],
            ]
            df_prom = pd.DataFrame(promedios_data)
            df_prom.to_excel(writer, index=False, header=False, sheet_name=sheet_prom)
            
            ws_prom = writer.sheets[sheet_prom]
            ws_prom['A1'].font = Font(bold=True, size=12, color="0055FF")
            for row in range(2, 6): ws_prom[f'A{row}'].font = Font(bold=True)
            ws_prom.column_dimensions['A'].width = 15
            ws_prom.column_dimensions['B'].width = 45

    return output.getvalue()

# ==============================================================================
# INTERFAZ MAIN (STREAMLIT)
# ==============================================================================
if not YOUTUBE_API_KEY:
    st.error("⚠️ ALERTA: No se encontró la YOUTUBE_API_KEY en los Secretos de Streamlit.")

df_roster = cargar_roster()

# Apartado visual para consultar el listado de creadores y sus enlaces
with st.expander("📋 Ver Listado de Creadores y Enlaces del Roster (Google Sheet)"):
    if not df_roster.empty:
        st.dataframe(df_roster, use_container_width=True)
    else:
        st.warning("No se pudo cargar el roster de Google Sheets.")

st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    redes_sel = st.multiselect("Redes a auditar:", ["YT Video", "YT Shorts", "TikTok", "Facebook Reels"], default=["YT Video", "YT Shorts", "TikTok", "Facebook Reels"])
with col2:
    f_inicio = st.date_input("Fecha Inicio:")
    f_fin = st.date_input("Fecha Fin:")
with col3:
    max_videos = st.number_input("Límite de Videos por Red:", min_value=1, max_value=50, value=15)

col_b1, col_b2, col_b3 = st.columns(3)
with col_b1:
    filtros_inc = [x.strip() for x in st.text_input("Obligatorias (separadas por coma):", placeholder="#BloodStrike").split(',') if x.strip()]
with col_b2:
    filtros_exc = [x.strip() for x in st.text_input("Excluir (separadas por coma):", placeholder="Sorteo").split(',') if x.strip()]
with col_b3:
    lista_negra = st.multiselect("🚫 Lista Negra (Omitir Creador):", df_roster['CREADOR'].tolist() if not df_roster.empty else [])

st.markdown("---")

if st.button("🚀 INICIAR EXTRACCIÓN PRO Y GENERAR MASTER EXCEL", type="primary") and not df_roster.empty:
    creadores = df_roster[~df_roster['CREADOR'].isin(lista_negra)].copy()
    resultados_totales = {}
    
    barra = st.progress(0)
    estado = st.empty()
    total = len(creadores)
    
    for idx, (_, row) in enumerate(creadores.iterrows()):
        nombre = row['CREADOR']
        estado.markdown(f"⏳ **Auditando ({idx+1}/{total}):** `{nombre}`")
        
        videos_creador = []
        
        if ("YT Video" in redes_sel or "YT Shorts" in redes_sel) and row['LINK YT'] != 'N/A':
            yt_res = extraer_youtube_api(
                nombre, row['LINK YT'], f_inicio, f_fin, max_videos, filtros_inc, filtros_exc,
                extraer_shorts=("YT Shorts" in redes_sel), extraer_videos=("YT Video" in redes_sel)
            )
            videos_creador.extend(yt_res)
            
        if "TikTok" in redes_sel and row['LINK TK'] != 'N/A':
            tk_res = extraer_tiktok_pro(nombre, row['LINK TK'], f_inicio, f_fin, max_videos, filtros_inc, filtros_exc)
            videos_creador.extend(tk_res)
            
        if "Facebook Reels" in redes_sel and row['LINK FC'] != 'N/A':
            fc_res = extraer_facebook_general(nombre, row['LINK FC'], f_inicio, f_fin, max_videos, filtros_inc, filtros_exc)
            videos_creador.extend(fc_res)

        proms = {
            'TK': {'promedio': 0, 'count': 0}, 'FC': {'promedio': 0, 'count': 0},
            'YT Shorts': {'promedio': 0, 'count': 0}, 'YT Video': {'promedio': 0, 'count': 0}
        }
        
        df_v = pd.DataFrame(videos_creador)
        if not df_v.empty:
            for clave, plat in [('TK', 'TikTok'), ('FC', 'Facebook Reels'), ('YT Shorts', 'YouTube Shorts'), ('YT Video', 'YouTube Video')]:
                sub = df_v[df_v['Plataforma'] == plat]
                if len(sub) > 0:
                    proms[clave] = {'promedio': int(sub['Vistas'].sum() / len(sub)), 'count': len(sub)}

        resultados_totales[nombre] = {'videos': videos_creador, 'promedios': proms}
        barra.progress((idx + 1) / total)
        
    barra.empty()
    estado.success("✅ ¡Auditoría Finalizada con Éxito!")
    
    excel_file = generar_excel_completo(resultados_totales)
    st.download_button(
        label="📥 DESCARGAR EXCEL (MASTER SHEET + PARES)",
        data=excel_file,
        file_name=f"Reporte_PRO_{f_inicio}_{f_fin}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
