import streamlit as st
import pandas as pd
import re
import random
import requests
import datetime
from io import BytesIO
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import yt_dlp
from concurrent.futures import ThreadPoolExecutor, as_completed

# ==============================================================================
# CONFIGURACIÓN GENERAL Y ESTILOS
# ==============================================================================
st.set_page_config(
    page_title="BS LATAM - MÓDULO DE PROMEDIOS DE CREADORES",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #0b0d11; color: #e6edf3; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stApp { background-color: #0b0d11; }
    
    .title-box { 
        border-left: 15px solid #0055FF; 
        padding: 20px 35px; 
        margin: 15px 0 25px 0; 
        background: linear-gradient(90deg, #161b22 0%, rgba(11,13,17,0) 100%);
        border-radius: 0 15px 15px 0;
    }
    .m-title { font-size: 36px; font-weight: 900; color: #ffffff; text-transform: uppercase; letter-spacing: 4px; margin: 0; }
    .s-title { font-size: 15px; color: #8b949e; font-family: 'Courier New', monospace; margin-top: 5px; }
    
    .stButton>button { background: linear-gradient(135deg, #0055FF 0%, #002b80 100%) !important; color: #ffffff !important; font-weight: 900 !important; text-transform: uppercase; border-radius: 12px; height: 48px; border: none; }
    .stTextArea textarea, .stTextInput input, .stNumberInput input { background-color: #161b22 !important; color: #e6edf3 !important; border: 1px solid #30363d !important; border-radius: 8px; }
    </style>

    <div class="title-box">
        <p class="m-title">BS LATAM • MÓDULO PROMEDIOS</p>
        <p class="s-title">EXTRACCIÓN SECUENCIAL POR CREADOR & REPORTES EXCEL PAREADOS</p>
    </div>
""", unsafe_allow_html=True)

USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0'
]

# ==============================================================================
# CARGA Y SANITIZACIÓN DEL LISTADO DE EMBAJADORES (124+ CREADORES)
# ==============================================================================
@st.cache_data(ttl=300)
def cargar_roster_embajadores_oficial():
    url_csv = "https://docs.google.com/spreadsheets/d/1KxXVKK0q5A8aSG5k8KHPoyXMynjkUUl5TAJRda7sPYg/export?format=csv&gid=852207003"
    try:
        df = pd.read_csv(url_csv)
        df = df.dropna(subset=['Creador'], how='all')
        
        # Mapeo unificado
        roster = []
        for _, row in df.iterrows():
            c_nombre = str(row.get('Creador', '')).strip()
            if not c_nombre or c_nombre.lower() == 'nan': continue
            
            link_yt = str(row.get('YouTube', '')).strip()
            link_tk = str(row.get('TikTok', '')).strip()
            link_fc = str(row.get('Facebook', '')).strip()
            
            roster.append({
                "CREADOR": c_nombre,
                "LINK YT": link_yt if link_yt.startswith('http') else "N/A",
                "LINK TK": link_tk if link_tk.startswith('http') else "N/A",
                "LINK FC": link_fc if link_fc.startswith('http') else "N/A"
            })
        return pd.DataFrame(roster)
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return pd.DataFrame()

# ==============================================================================
# MOTOR DE EXTRACCIÓN ESPECIALIZADO
# ==============================================================================
def extraer_facebook_reels_views(url):
    """Extrae las vistas de Facebook Reels de forma rápida."""
    try:
        headers = {'User-Agent': random.choice(USER_AGENTS), 'Accept-Language': 'es-ES,es;q=0.9'}
        res = requests.get(url, headers=headers, timeout=8)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            meta_desc = soup.find('meta', property='og:description')
            if meta_desc and 'content' in meta_desc.attrs:
                desc = meta_desc['content']
                m = re.search(r"([\d\.,]+[KMkm]?)\s*(views|reproducciones|vistas)", desc, re.I)
                if m:
                    v_str = m.group(1).upper().replace(',', '.')
                    mult = 1000 if 'K' in v_str else 1000000 if 'M' in v_str else 1
                    clean = re.sub(r'[^0-9.]', '', v_str)
                    return int(float(clean) * mult)
    except: pass
    return 0

def escanear_perfil_red(creador, url, plataforma, f_start_int, f_end_int, max_limit, filtro_inc, filtro_exc):
    """Extrae secuencialmente los contenidos de la red elegida aplicando todos los filtros."""
    if not url or url == 'N/A' or not url.startswith('http'):
        return []

    # Construir endpoint exacto para YouTube
    target_url = url
    if plataforma == "YouTube Video":
        target_url = url.rstrip('/') + '/videos'
    elif plataforma == "YouTube Shorts":
        target_url = url.rstrip('/') + '/shorts'

    ydl_opts = {
        'quiet': True,
        'skip_download': True,
        'ignoreerrors': True,
        'no_warnings': True,
        'playlistend': max_limit * 3, # Traemos un margen mayor para aplicar filtros
        'extract_flat': 'in_playlist',
        'socket_timeout': 10,
        'http_headers': {'User-Agent': random.choice(USER_AGENTS)}
    }

    videos_validos = []
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(target_url, download=False)
            if info and 'entries' in info:
                entries = [e for e in info['entries'] if e]
                
                for entry in entries:
                    if len(videos_validos) >= max_limit:
                        break # Límite alcanzado
                        
                    v_url = entry.get('url') or entry.get('webpage_url')
                    if not v_url and entry.get('id'):
                        v_url = f"https://www.youtube.com/watch?v={entry.get('id')}" if "YouTube" in plataforma else f"https://www.tiktok.com/@video/{entry.get('id')}"
                    
                    if not v_url: continue

                    titulo = str(entry.get('title', '')).strip()
                    v_date_str = entry.get('upload_date')
                    vistas = entry.get('view_count')

                    # Evaluación de palabras clave inclusivas y exclusivas
                    titulo_lower = titulo.lower()
                    if filtro_inc and not any(fi.lower() in titulo_lower for fi in filtro_inc):
                        continue
                    if filtro_exc and any(fe.lower() in titulo_lower for fe in filtro_exc):
                        continue

                    # Evaluación de rango de fechas
                    if v_date_str:
                        try:
                            d_int = int(str(v_date_str))
                            if not (f_start_int <= d_int <= f_end_int):
                                continue
                        except: pass

                    # Si faltan vistas, realizamos un fetch puntual
                    if vistas is None:
                        if plataforma == "Facebook Reels":
                            vistas = extraer_facebook_reels_views(v_url)
                        else:
                            try:
                                with yt_dlp.YoutubeDL({'quiet': True, 'skip_download': True, 'socket_timeout': 6}) as ydl_v:
                                    v_i = ydl_v.extract_info(v_url, download=False)
                                    vistas = v_i.get('view_count', 0) if v_i else 0
                            except: vistas = 0

                    vistas = int(vistas or 0)
                    fecha_fmt = f"{str(v_date_str)[:4]}-{str(v_date_str)[4:6]}-{str(v_date_str)[6:]}" if v_date_str else "N/A"

                    videos_validos.append({
                        "Creador": creador,
                        "Plataforma": plataforma,
                        "Fecha": fecha_fmt,
                        "Título": titulo[:70],
                        "Vistas": vistas,
                        "Link": v_url
                    })
    except Exception:
        pass

    return videos_validos

# ==============================================================================
# GENERADOR DE EXCEL PAREADO POR CREADOR
# ==============================================================================
def generar_excel_pareado_promedios(resultados_por_creador):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for creador, data in resultados_por_creador.items():
            safe_name = re.sub(r'[\[\]\:\*\?\/\\\|\'\"\t\n\r]', '', creador)[:20].strip() or "Creador"
            
            # 1. Hoja "ALL VIDEOS"
            sheet_videos = f"{safe_name} ALL VIDEOS"
            df_vids = pd.DataFrame(data['videos'])
            if df_vids.empty:
                df_vids = pd.DataFrame(columns=["Creador", "Plataforma", "Fecha", "Título", "Vistas", "Link"])
            df_vids.to_excel(writer, index=False, sheet_name=sheet_videos)

            # 2. Hoja "PROMEDIO"
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
            
            # Estilos visuals en la hoja de promedio
            ws_prom = writer.sheets[sheet_prom]
            ws_prom['A1'].font = Font(bold=True, size=12, color="0055FF")
            ws_prom['B1'].font = Font(bold=True, size=12, color="FFFFFF")
            
            for row in range(2, 6):
                ws_prom[f'A{row}'].font = Font(bold=True)

    return output.getvalue()

# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
df_embajadores = cargar_roster_embajadores_oficial()

st.sidebar.markdown("### 📋 DIRECTORIO OFICIAL")
if not df_embajadores.empty:
    st.sidebar.success(f"Cargados {len(df_embajadores)} creadores en vivo desde Google Sheets.")
    st.sidebar.dataframe(df_embajadores[['CREADOR', 'LINK YT', 'LINK TK', 'LINK FC']], use_container_width=True, hide_index=True)

st.markdown("### 🛠️ PARÁMETROS DE AUDITORÍA Y BARRIDO SECUENCIAL")

col_a1, col_a2, col_a3 = st.columns(3)

with col_a1:
    redes_seleccionadas = st.multiselect(
        "Redes a extraer:",
        ["YT Video", "YT Shorts", "TK", "FC Reels"],
        default=["YT Video", "YT Shorts", "TK", "FC Reels"]
    )
    
with col_a2:
    f_inicio = st.date_input("Fecha Inicio:", value=datetime.date(2026, 2, 1))
    f_fin = st.date_input("Fecha Fin:", value=datetime.date(2026, 2, 28))

with col_a3:
    max_videos_limite = st.number_input("Últimos N videos a evaluar por red:", min_value=1, max_value=50, value=15)

st.markdown("---")
col_b1, col_b2, col_b3 = st.columns(3)

with col_b1:
    inc_str = st.text_input("Filtro INCLUSIVO (Títulos/Hashtags separadas por coma):", placeholder="#BloodStrike, Blood Strike")
    filtros_inc = [x.strip() for x in inc_str.split(',') if x.strip()]

with col_b2:
    exc_str = st.text_input("Filtro EXCLUSIVO (Ignorar palabras separadas por coma):", placeholder="Dlss5, Sorteo")
    filtros_exc = [x.strip() for x in exc_str.split(',') if x.strip()]

with col_b3:
    lista_creadores_nombres = df_embajadores['CREADOR'].tolist() if not df_embajadores.empty else []
    lista_negra = st.multiselect("🚫 LISTA NEGRA (Ignorar en esta ronda):", lista_creadores_nombres)

if st.button("🚀 EJECUTAR BARRIDO Y CALCULAR PROMEDIOS PAREADOS", type="primary"):
    if df_embajadores.empty:
        st.error("Error: No se pudo cargar la lista de creadores.")
    else:
        # Filtrar creadores no excluidos
        creadores_a_procesar = df_embajadores[~df_embajadores['CREADOR'].isin(lista_negra)].copy()
        
        f_start_int = int(f_inicio.strftime('%Y%m%d'))
        f_end_int = int(f_fin.strftime('%Y%m%d'))
        
        resultados_maestros = {}
        p_bar = st.progress(0)
        status_txt = st.empty()
        
        total_creadores = len(creadores_a_procesar)
        
        for idx, (_, row) in enumerate(creadores_a_procesar.iterrows()):
            c_name = row['CREADOR']
            status_txt.markdown(f"🛰️ **PROCESANDO CREADOR ({idx+1}/{total_creadores}):** `{c_name}`")
            
            c_videos = []
            
            # Extracción Secuencial por Red
            if "YT Video" in redes_seleccionadas and row['LINK YT'] != 'N/A':
                vids = escanear_perfil_red(c_name, row['LINK YT'], "YouTube Video", f_start_int, f_end_int, max_videos_limite, filtros_inc, filtros_exc)
                c_videos.extend(vids)
                
            if "YT Shorts" in redes_seleccionadas and row['LINK YT'] != 'N/A':
                vids = escanear_perfil_red(c_name, row['LINK YT'], "YouTube Shorts", f_start_int, f_end_int, max_videos_limite, filtros_inc, filtros_exc)
                c_videos.extend(vids)

            if "TK" in redes_seleccionadas and row['LINK TK'] != 'N/A':
                vids = escanear_perfil_red(c_name, row['LINK TK'], "TikTok", f_start_int, f_end_int, max_videos_limite, filtros_inc, filtros_exc)
                c_videos.extend(vids)

            if "FC Reels" in redes_seleccionadas and row['LINK FC'] != 'N/A':
                vids = escanear_perfil_red(c_name, row['LINK FC'], "Facebook Reels", f_start_int, f_end_int, max_videos_limite, filtros_inc, filtros_exc)
                c_videos.extend(vids)

            # Cálculo Matemático de Promedios
            proms = {
                'TK': {'promedio': 0, 'count': 0},
                'FC': {'promedio': 0, 'count': 0},
                'YT Shorts': {'promedio': 0, 'count': 0},
                'YT Video': {'promedio': 0, 'count': 0}
            }
            
            df_c_vids = pd.DataFrame(c_videos)
            if not df_c_vids.empty:
                for plat_key, target_name in [('TK', 'TikTok'), ('FC', 'Facebook Reels'), ('YT Shorts', 'YouTube Shorts'), ('YT Video', 'YouTube Video')]:
                    sub = df_c_vids[df_c_vids['Plataforma'] == target_name]
                    cnt = len(sub)
                    if cnt > 0:
                        tot_v = sub['Vistas'].sum()
                        proms[plat_key] = {'promedio': int(tot_v / cnt), 'count': cnt}

            resultados_maestros[c_name] = {
                'videos': c_videos,
                'promedios': proms
            }
            
            p_bar.progress((idx + 1) / total_creadores)
            
        p_bar.empty()
        status_txt.empty()
        
        st.success("✅ Extracción y cálculo de promedios finalizado exitosamente.")
        
        # Generar Excel pareado
        excel_data = generar_excel_pareado_promedios(resultados_maestros)
        
        st.download_button(
            label="📊 DESCARGAR REPORTES PAREADOS EXCEL (.XLSX)",
            data=excel_data,
            file_name=f"Promedios_BSLT_{f_inicio}_al_{f_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
)
      
