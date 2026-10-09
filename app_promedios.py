import streamlit as st
import pandas as pd
import re
import datetime
import os
import yt_dlp
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from googleapiclient.discovery import build
from concurrent.futures import ThreadPoolExecutor, as_completed

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
        <p class="s-title">YOUTUBE API V3 | YT-DLP TIKTOK & FB | MULTITHREADING | ANALYTICS DASHBOARD</p>
    </div>
""", unsafe_allow_html=True)

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
# MOTOR YOUTUBE V3 (ESTRICTO POR FECHAS) - 100% INTACTO Y FUNCIONAL
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
# MOTOR YT-DLP PROFESIONAL (PARA TIKTOK Y FACEBOOK)
# ==============================================================================
def extraer_ytdlp_red(creador, url, plataforma, start_date, end_date, max_vids, inc_kw, exc_kw):
    if url == 'N/A': return []
    videos_validos = []
    
    dt_start = datetime.datetime.combine(start_date, datetime.time.min)
    dt_end = datetime.datetime.combine(end_date, datetime.time.max)
    
    ydl_opts = {
        'extract_flat': False,
        'skip_download': True,
        'quiet': True,
        'no_warnings': True,
        'playlistend': max_vids * 3
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            if not info: return []
            
            entries = info.get('entries', [info])
            for entry in entries:
                if not entry: continue
                
                upload_date_str = entry.get('upload_date')
                fecha_vid = dt_start
                fecha_str = "Reciente"
                
                if upload_date_str and len(upload_date_str) == 8:
                    try:
                        fecha_vid = datetime.datetime.strptime(upload_date_str, '%Y%m%d')
                        fecha_str = fecha_vid.strftime('%Y-%m-%d')
                        if not (dt_start <= fecha_vid <= dt_end):
                            continue
                    except: pass
                
                titulo = entry.get('title', '') or entry.get('description', '')
                if not cumple_filtros_kw(titulo, inc_kw, exc_kw): continue
                
                vistas = int(entry.get('view_count', 0) or entry.get('play_count', 0) or 0)
                vid_url = entry.get('webpage_url', '') or entry.get('url', '')
                
                videos_validos.append({
                    "Creador": creador,
                    "Plataforma": plataforma,
                    "Fecha": fecha_str,
                    "Título": titulo[:70] if titulo else f"{plataforma} Video",
                    "Vistas": vistas,
                    "Link": vid_url
                })
                
                if len(videos_validos) >= max_vids: break
    except:
        pass
        
    return videos_validos

# ==============================================================================
# FUNCIÓN DE PROCESAMIENTO POR CREADOR (PARA CONCURRENCIA / MULTITHREADING)
# ==============================================================================
def procesar_un_creador(row, redes_sel, f_inicio, f_fin, max_videos, filtros_inc, filtros_exc):
    nombre = row['CREADOR']
    videos_creador = []
    
    # 1. YouTube
    if ("YT Video" in redes_sel or "YT Shorts" in redes_sel) and row['LINK YT'] != 'N/A':
        yt_res = extraer_youtube_api(
            nombre, row['LINK YT'], f_inicio, f_fin, max_videos, filtros_inc, filtros_exc,
            extraer_shorts=("YT Shorts" in redes_sel), extraer_videos=("YT Video" in redes_sel)
        )
        videos_creador.extend(yt_res)
        
    # 2. TikTok (yt-dlp)
    if "TikTok" in redes_sel and row['LINK TK'] != 'N/A':
        tk_res = extraer_ytdlp_red(nombre, row['LINK TK'], "TikTok", f_inicio, f_fin, max_videos, filtros_inc, filtros_exc)
        videos_creador.extend(tk_res)
        
    # 3. Facebook Reels (yt-dlp)
    if "Facebook Reels" in redes_sel and row['LINK FC'] != 'N/A':
        fc_res = extraer_ytdlp_red(nombre, row['LINK FC'], "Facebook Reels", f_inicio, f_fin, max_videos, filtros_inc, filtros_exc)
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

    return nombre, {'videos': videos_creador, 'promedios': proms}

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

if st.button("🚀 INICIAR EXTRACCIÓN PRO EN PARALELO Y GENERAR REPORTE", type="primary") and not df_roster.empty:
    creadores_df = df_roster[~df_roster['CREADOR'].isin(lista_negra)].copy()
    resultados_totales = {}
    
    barra = st.progress(0)
    estado = st.empty()
    total = len(creadores_df)
    completados = 0
    
    estado.markdown(f"⚡ **Iniciando auditoría paralela para {total} creadores...**")
    
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {
            executor.submit(procesar_un_creador, row, redes_sel, f_inicio, f_fin, max_videos, filtros_inc, filtros_exc): row['CREADOR']
            for _, row in creadores_df.iterrows()
        }
        
        for future in as_completed(futures):
            nombre, res_data = future.result()
            resultados_totales[nombre] = res_data
            completados += 1
            barra.progress(completados / total)
            estado.markdown(f"⏳ **Completados ({completados}/{total}):** Último procesado `{nombre}`")
            
    barra.empty()
    estado.success("✅ ¡Auditoría Paralela Finalizada con Éxito!")
    
    # ==============================================================================
    # PANEL DE ANALÍTICA VISUAL (DASHBOARD)
    # ==============================================================================
    st.markdown("### 📊 Panel de Analítica Global")
    
    total_vids_global = sum(sum(d['promedios'][k]['count'] for k in d['promedios']) for d in resultados_totales.values())
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Creadores Auditados", len(resultados_totales))
    m2.metric("Videos Recopilados", total_vids_global)
    m3.metric("Rango analizado", f"{f_inicio} al {f_fin}")
    
    # Gráfico resumen por plataforma
    totales_plat = {'YouTube Video': 0, 'YouTube Shorts': 0, 'TikTok': 0, 'Facebook Reels': 0}
    for d in resultados_totales.values():
        for v in d['videos']:
            p = v['Plataforma']
            if p in totales_plat: totales_plat[p] += 1
            
    st.bar_chart(pd.Series(totales_plat))
    
    # Descarga del archivo Excel
    excel_file = generar_excel_completo(resultados_totales)
    st.download_button(
        label="📥 DESCARGAR EXCEL (MASTER SHEET + PARES)",
        data=excel_file,
        file_name=f"Reporte_PRO_{f_inicio}_{f_fin}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
