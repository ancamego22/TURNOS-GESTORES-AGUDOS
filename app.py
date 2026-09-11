import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import io
import os
import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from PIL import Image

st.set_page_config(page_title="Sistema Integral de Turnos - Salud en Casa SURA", layout="wide", page_icon="🏢")

# -------------------------------------------------------------
# APLICAR ESTILOS CORPORATIVOS SURA (AZUL INSTITUCIONAL)
# -------------------------------------------------------------
st.markdown("""
    <style>
    /* Color principal de botones y acentos */
    .stButton>button {
        background-color: #003366;
        color: white;
        border-radius: 5px;
        font-weight: bold;
        border: none;
    }
    .stButton>button:hover {
        background-color: #002244;
        color: white;
    }
    /* Estilos de títulos */
    h1, h2, h3 {
        color: #003366 !important;
    }
    /* Contenedores con borde fino */
    div.st-emotion-cache-1r6slb0 {
        border-color: #003366 !important;
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 0. CONFIGURACIÓN DE ALMACENAMIENTO PERSISTENTE LOCAL
# -------------------------------------------------------------
DATA_DIR = "datos_sistema"
os.makedirs(DATA_DIR, exist_ok=True)

PATH_PERSONAL = os.path.join(DATA_DIR, "catalogo_personal.xlsx")
PATH_HISTORIAL = os.path.join(DATA_DIR, "ultimo_historial.xlsx")
PATH_SALARIOS = os.path.join(DATA_DIR, "salarios_base.json")
PATH_AUDITORIA = os.path.join(DATA_DIR, "registro_auditoria.json")
PATH_CUADRO_ACTUAL = os.path.join(DATA_DIR, "cuadro_activo.json")
PATH_VACACIONES = os.path.join(DATA_DIR, "control_vacaciones.json")
PATH_LOGO = "logo_sura.png"

def cargar_json(path, default={}):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default

def guardar_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def registrar_auditoria(tipo, detalle, responsable):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    hist = cargar_json(PATH_AUDITORIA, [])
    registro = {
        "ID": str(datetime.datetime.now().timestamp()),
        "Fecha_Hora": now_str,
        "Tipo": tipo,
        "Detalle": detalle,
        "Aprobado_Por": responsable if responsable else "Coordinación de Turnos"
    }
    hist.insert(0, registro)
    guardar_json(PATH_AUDITORIA, hist)

def obtener_auditoria():
    if os.path.exists(PATH_AUDITORIA):
        try:
            with open(PATH_AUDITORIA, "r", encoding="utf-8") as f:
                return pd.DataFrame(json.load(f))
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

# Constantes operativas y legales colombianas
SHIFT_HOURS = {"D": 0, "T1": 8, "T2": 8, "T3": 8, "T4": 8, "T1E": 10, "T2E": 10, "INC": 0, "V": 0, "CAL": 0}
WORKING_SHIFTS = ["T1", "T2", "T3", "T4", "T1E", "T2E"]
dias_es = {0: "Lun", 1: "Mar", 2: "Mié", 3: "Jue", 4: "Vie", 5: "Sáb", 6: "Dom"}

def is_festivo_or_domingo(dt):
    return dt.weekday() == 6 or dt.strftime('%Y-%m-%d') == '2026-10-12'

def contar_dias_habiles_vacaciones(f_inicio, f_fin):
    dias_h = 0
    curr = f_inicio
    while curr <= f_fin:
        if not is_festivo_or_domingo(curr):
            dias_h += 1
        curr += datetime.timedelta(days=1)
    return dias_h

# -------------------------------------------------------------
# 1. GESTIÓN DE SESIÓN Y MEMORIA EN VIVO
# -------------------------------------------------------------
if "df_c20" not in st.session_state: st.session_state.df_c20 = None
if "df_c21" not in st.session_state: st.session_state.df_c21 = None
if "dias_28" not in st.session_state: st.session_state.dias_28 = None
if "novedades_list" not in st.session_state: st.session_state.novedades_list = []
if "salarios" not in st.session_state: st.session_state.salarios = cargar_json(PATH_SALARIOS, {})
if "vacaciones_db" not in st.session_state: st.session_state.vacaciones_db = cargar_json(PATH_VACACIONES, {})
if "admin_autenticado" not in st.session_state: st.session_state.admin_autenticado = False

if st.session_state.df_c20 is None and os.path.exists(PATH_CUADRO_ACTUAL):
    try:
        with open(PATH_CUADRO_ACTUAL, "r", encoding="utf-8") as f:
            saved_data = json.load(f)
            st.session_state.df_c20 = pd.DataFrame(saved_data["c20"])
            st.session_state.df_c21 = pd.DataFrame(saved_data["c21"])
            st.session_state.dias_28 = pd.date_range(saved_data["fecha_inicio"], periods=28)
    except Exception:
        pass

nombres_emp = []
df_p_global = None
if os.path.exists(PATH_PERSONAL):
    try:
        try:
            df_p_global = pd.read_excel(PATH_PERSONAL, sheet_name='PERSONAL ')
        except Exception:
            df_p_global = pd.read_excel(PATH_PERSONAL, sheet_name=0)
            
        df_p_global.columns = [str(c).strip() for c in df_p_global.iloc[0].values]
        df_p_global = df_p_global.iloc[1:].dropna(subset=['EMPLEADO'])
        nombres_emp = sorted(df_p_global['EMPLEADO'].astype(str).str.strip().unique().tolist())
    except Exception:
        pass
# -------------------------------------------------------------
# 2. BARRA LATERAL: SELECCIÓN DE PERFIL (EMPLEADO VS ADMIN SEGURO)
# -------------------------------------------------------------
with st.sidebar:
    if os.path.exists(PATH_LOGO):
        st.image(PATH_LOGO, use_container_width=True)
    else:
        st.header("🏥 Salud en Casa — SURA")

    perfil_ingreso = st.radio("Selecciona tu Perfil:", ["📅 Portal de Empleados (Público)", "🔒 Panel Administrativo (Requiere Clave)"])

    if perfil_ingreso == "🔒 Panel Administrativo (Requiere Clave)":
        st.markdown("---")
        if not st.session_state.admin_autenticado:
            clave_ingreso = st.text_input("Contraseña de Administrador:", type="password")
            if st.button("🔓 Ingresar al Sistema"):
                if clave_ingreso == "sura2026":
                    st.session_state.admin_autenticado = True
                    st.success("¡Acceso concedido!")
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
            st.stop()
        else:
            st.success("🔒 Sesión Admin Activa")
            if st.button("🔒 Cerrar Sesión"):
                st.session_state.admin_autenticado = False
                st.rerun()
            
            st.markdown("---")
            st.header("👤 Responsable Activo")
            resp_guardados = ["Andrés Medina (Coordinador)", "Gestión Agudos", "Jefatura de Operaciones", "Talento Humano", "Otro (Escribir nombre manual)"]
            resp_sel = st.selectbox("Firma en Auditoría:", resp_guardados)
            if resp_sel == "Otro (Escribir nombre manual)":
                responsable_activo = st.text_input("Escribe tu Nombre y Rol:", value="Supervisor Operativo")
            else:
                responsable_activo = resp_sel

            st.header("💾 Base de Datos Fija")
            if os.path.exists(PATH_PERSONAL):
                st.success("✅ Catálogo Guardado")
                if st.button("🗑️ Cambiar Catálogo"):
                    os.remove(PATH_PERSONAL); st.rerun()
            else:
                f_p_init = st.file_uploader("Catálogo de Personal (.xls/.xlsx)", type=["xls", "xlsx"])
                if f_p_init:
                    with open(PATH_PERSONAL, "wb") as f: f.write(f_p_init.getbuffer())
                    st.success("¡Guardado!"); st.rerun()

            if os.path.exists(PATH_HISTORIAL):
                st.success("✅ Historial Guardado")
                if st.button("🗑️ Cambiar Historial"):
                    os.remove(PATH_HISTORIAL); st.rerun()
            else:
                f_h_init = st.file_uploader("Historial Ciclo Anterior (.xlsx)", type=["xlsx"])
                if f_h_init:
                    with open(PATH_HISTORIAL, "wb") as f: f.write(f_h_init.getbuffer())
                    st.success("¡Guardado!"); st.rerun()

            st.header("⚙️ Parámetros")
            fecha_inicio = st.date_input("Lunes Inicio de Ciclo", datetime.date(2026, 9, 21))

# -------------------------------------------------------------
# VISTA 1: PORTAL DE AUTOGESTIÓN PARA EMPLEADOS (PÚBLICO)
# -------------------------------------------------------------
if perfil_ingreso == "📅 Portal de Empleados (Público)":
    c_logo_t, c_title_t = st.columns([1, 4])
    with c_logo_t:
        if os.path.exists(PATH_LOGO):
            st.image(PATH_LOGO, use_container_width=True)
    with c_title_t:
        st.title("📅 Portal de Autogestión — Mis Turnos SURA")
    
    st.markdown("Consulta tu programación quincenal y descarga tu tarjeta corporativa oficial de **Salud en Casa SURA**.")
    
    if st.session_state.df_c20 is not None:
        c20_cols = [c for c in st.session_state.df_c20.columns if c not in ['EMPLEADO', 'ROL']]
        c21_cols = [c for c in st.session_state.df_c21.columns if c not in ['EMPLEADO', 'ROL']]
        todos_emps = st.session_state.df_c20['EMPLEADO'].tolist()
        
        emp_portal = st.selectbox("Selecciona tu Nombre para ver tu Calendario:", todos_emps)
        if emp_portal:
            r20 = st.session_state.df_c20.loc[st.session_state.df_c20['EMPLEADO'] == emp_portal].iloc[0]
            r21 = st.session_state.df_c21.loc[st.session_state.df_c21['EMPLEADO'] == emp_portal].iloc[0]
            
            st.markdown(f"#### Cronograma de Turnos para: {emp_portal}")
            df_portal_view = pd.DataFrame({
                "Fecha / Día": c20_cols + c21_cols,
                "Turno Asignado": [r20[c] for c in c20_cols] + [r21[c] for c in c21_cols]
            })
            st.dataframe(df_portal_view, use_container_width=True)
            
            fig, ax = plt.subplots(figsize=(14, 6))
            ax.axis('off')
            
            bg = patches.Rectangle((0, 0), 1, 1, transform=ax.transAxes, facecolor='#F8F9FA', edgecolor='#003366', linewidth=2, zorder=0)
            ax.add_patch(bg)
            
            banner = patches.Rectangle((0, 0.80), 1, 0.20, transform=ax.transAxes, facecolor='#003366', zorder=1)
            ax.add_patch(banner)
            
            ax.text(0.03, 0.92, "SALUD EN CASA — SURA", fontsize=15, fontweight='bold', color='white', zorder=2)
            ax.text(0.03, 0.85, f"CRONOGRAMA PERSONAL DE TURNOS | Ciclos 20 y 21", fontsize=10, color='#B0C4DE', zorder=2)
            
            if os.path.exists(PATH_LOGO):
                try:
                    img_logo = plt.imread(PATH_LOGO)
                    fig.figimage(img_logo, xo=900, yo=480, zorder=3, origin='upper')
                except Exception:
                    pass

            ax.text(0.03, 0.70, f"Colaborador: {emp_portal}", fontsize=12, fontweight='bold', color='#003366')
            ax.text(0.03, 0.65, f"Rol: {r20['ROL'][:50]}", fontsize=9, color='#555555')
            
            colors_map = {
                "T1": "#D1ECF1", "T1E": "#99CCFF", "T2": "#FFE8D6", "T2E": "#FFCC80",
                "T3": "#F8D7DA", "T4": "#FFF3CD", "D": "#D4EDDA", "INC": "#F5C6CB", "V": "#B2EBF2", "CAL": "#E1BEE7"
            }
            
            all_cols = c20_cols + c21_cols
            all_vals = [r20[c] for c in c20_cols] + [r21[c] for c in c21_cols]
            
            for i, (d_txt, s_val) in enumerate(zip(all_cols, all_vals)):
                row = i // 7
                col = i % 7
                x = 0.03 + col * 0.138
                y = 0.35 - row * 0.28
                bg_col = colors_map.get(s_val, "#FFFFFF")
                
                box = patches.FancyBboxPatch((x, y), 0.13, 0.22, boxstyle="round,pad=0.02", facecolor=bg_col, edgecolor="#003366", linewidth=1.2, zorder=2)
                ax.add_patch(box)
                
                ax.text(x + 0.065, y + 0.15, d_txt, fontsize=9, fontweight='bold', ha='center', color="#333333", zorder=3)
                ax.text(x + 0.065, y + 0.07, s_val, fontsize=15, fontweight='bold', ha='center', color="#111111", zorder=3)

            buf_img = io.BytesIO()
            plt.savefig(buf_img, format='png', bbox_inches='tight', dpi=200)
            buf_img.seek(0)
            plt.close()
            
            st.download_button(
                label=f"⬇️ Descargar Tarjeta de Calendario Corporativa (PNG)",
                data=buf_img.getvalue(),
                file_name=f"Calendario_SaludEnCasa_SURA_{emp_portal.replace(' ', '_')}.png",
                mime="image/png"
            )
    else:
        st.info("Aún no se ha generado ningún cuadro de turnos en el sistema. Comunícate con tu coordinador.")
    st.stop()

# -------------------------------------------------------------
# VISTA 2: PANEL ADMINISTRATIVO PRINCIPAL (PROTEGIDO)
# -------------------------------------------------------------
c_logo_a, c_title_a = st.columns([1, 5])
with c_logo_a:
    if os.path.exists(PATH_LOGO):
        st.image(PATH_LOGO, use_container_width=True)
with c_title_a:
    st.title("🏢 Panel Administrativo — Salud en Casa SURA")

st.markdown("Gestión operativa con cumplimiento de 84 horas, control de vacaciones hábiles (Lunes a Sábado) y auditoría.")

# 1. RADICACIÓN DE NOVEDADES
st.subheader("📋 1. Radicación de Novedades y Solicitudes Legales")
with st.container(border=True):
    c_n1, c_n2, c_n3, c_n4 = st.columns(4)
    tipo_sol = c_n1.selectbox("Tipo de Novedad:", [
        "Vacaciones", "Incapacidad", "Calamidad Doméstica", "Petición de Día Libre", "Preferencia de Turno"
    ])
    emp_sol = c_n2.selectbox("Colaborador:", ["(Seleccionar)"] + nombres_emp)
    destino_ciclo = c_n3.selectbox("Aplicar para:", ["Ciclos Actuales (20 y 21)", "Próximos Ciclos (22 y 23 en adelante)"])

    if tipo_sol in ["Vacaciones", "Incapacidad", "Calamidad Doméstica"]:
        d_desde = c_n4.date_input("Desde:", datetime.date(2026, 10, 1), key="nov_d1")
        d_hasta = c_n4.date_input("Hasta:", datetime.date(2026, 10, 9), key="nov_d2")
        
        if tipo_sol == "Vacaciones":
            dias_hab_vac = contar_dias_habiles_vacaciones(d_desde, d_hasta)
            st.caption(f"ℹ️ **Días Legales (Lunes a Sábado):** {dias_hab_vac} días hábiles computados (domingos y festivos no descuentan de los días legales).")

        if st.button("➕ Registrar y Aprobar Novedad"):
            if emp_sol != "(Seleccionar)":
                cod = "V" if tipo_sol == "Vacaciones" else ("INC" if tipo_sol == "Incapacidad" else "CAL")
                st.session_state.novedades_list.append({
                    "tipo": cod, "emp": emp_sol, "desde": d_desde, "hasta": d_hasta, "destino": destino_ciclo
                })
                
                if tipo_sol == "Vacaciones":
                    d_h = contar_dias_habiles_vacaciones(d_desde, d_hasta)
                    reg_vac = st.session_state.vacaciones_db.get(emp_sol, {"disfrutados": 0, "total_ley": 15})
                    reg_vac["disfrutados"] += d_h
                    st.session_state.vacaciones_db[emp_sol] = reg_vac
                    guardar_json(PATH_VACACIONES, st.session_state.vacaciones_db)

                registrar_auditoria(f"Radicación {tipo_sol} ({destino_ciclo})", f"{emp_sol} ({cod}) del {d_desde} al {d_hasta}", responsable_activo)
                st.success(f"{tipo_sol} registrada y aprobada.")
                st.rerun()

    elif tipo_sol == "Petición de Día Libre":
        d_libre = c_n4.date_input("Día Solicitado:", datetime.date(2026, 9, 25), key="nov_dl")
        if st.button("➕ Registrar y Aprobar Día Libre"):
            if emp_sol != "(Seleccionar)":
                st.session_state.novedades_list.append({
                    "tipo": "D", "emp": emp_sol, "desde": d_libre, "hasta": d_libre, "destino": destino_ciclo
                })
                registrar_auditoria(f"Petición Día Libre ({destino_ciclo})", f"{emp_sol} descanso solicitado el {d_libre}", responsable_activo)
                st.success("Día libre registrado y aprobado.")
                st.rerun()

    elif tipo_sol == "Preferencia de Turno":
        pref_val = c_n4.radio("Preferencia:", ["Más mañanas (T1)", "Más tardes (T2)"], horizontal=True)
        if st.button("➕ Registrar Preferencia"):
            if emp_sol != "(Seleccionar)":
                cod_pref = "PREF_T1" if "T1" in pref_val else "PREF_T2"
                st.session_state.novedades_list.append({"tipo": cod_pref, "emp": emp_sol, "destino": destino_ciclo})
                registrar_auditoria(f"Preferencia de Turno ({destino_ciclo})", f"{emp_sol} solicita {pref_val}", responsable_activo)
                st.success("Preferencia registrada con éxito.")
                st.rerun()

    if st.session_state.novedades_list:
        st.markdown("##### Novedades en Cola:")
        cols_nov = st.columns(min(4, len(st.session_state.novedades_list)))
        for i_n, nov_item in enumerate(st.session_state.novedades_list):
            with cols_nov[i_n % len(cols_nov)]:
                d_txt = f"{nov_item['desde'].strftime('%d/%m')} al {nov_item['hasta'].strftime('%d/%m')}" if "desde" in nov_item else nov_item['tipo']
                st.info(f"**{nov_item['emp'][:18]}**\n\n`{nov_item['tipo']}` | {nov_item['destino'][:12]}..")
        if st.button("🗑️ Limpiar Cola de Novedades"):
            st.session_state.novedades_list = []
            st.rerun()

# Motor OR-Tools
def ejecutar_optimizador(df_p, history, fecha_inicio_dt, novedades, filtrar_ciclo="Ciclos Actuales (20 y 21)"):
    employees = []
    for _, row in df_p.iterrows():
        emp = str(row['EMPLEADO']).strip()
        rol = str(row['ROL']).strip() if pd.notna(row['ROL']) else ''
        res = str(row['RESTRICCION']).strip() if pd.notna(row['RESTRICCION']) else ''
        employees.append({
            'nombre': emp, 'rol': rol, 'restriccion': res,
            'has_pas': "PAS" in rol,
            'hist': history.get(emp, {'d13_sab': 'D', 'd14_dom': 'D'})
        })

    N_DAYS = 28
    N_EMP = len(employees)
    dias_28 = pd.date_range(fecha_inicio_dt, periods=28)
    SHIFTS = ["D", "T1", "T2", "T3", "T4", "T1E", "T2E", "INC", "V", "CAL"]

    model = cp_model.CpModel()
    x = {}
    for i in range(N_EMP):
        for d in range(N_DAYS):
            for s in SHIFTS:
                x[(i, d, s)] = model.NewBoolVar(f"x_{i}_{d}_{s}")
        for d in range(N_DAYS):
            model.AddExactlyOne(x[(i, d, s)] for s in SHIFTS)

    for d in range(N_DAYS):
        model.Add(sum(x[(i, d, "T3")] for i in range(N_EMP)) == 3)
        model.Add(sum(x[(i, d, "T4")] for i in range(N_EMP)) == 1)
        model.Add(sum(x[(i, d, "T1E")] for i in range(N_EMP)) <= 3)
        model.Add(sum(x[(i, d, "T2E")] for i in range(N_EMP)) <= 3)
        model.Add(sum(x[(i, d, "T1E")] + x[(i, d, "T2E")] for i in range(N_EMP)) <= 5)
        model.Add(sum(x[(i, d, "T1E")] for i in range(N_EMP)) >= sum(x[(i, d, "T2E")] for i in range(N_EMP)))
        model.Add(sum(x[(i, d, "T1")] + x[(i, d, "T1E")] for i in range(N_EMP)) >= 9)
        model.Add(sum(x[(i, d, "T1")] + x[(i, d, "T1E")] for i in range(N_EMP)) <= 11)
        model.Add(sum(x[(i, d, "T2")] + x[(i, d, "T2E")] for i in range(N_EMP)) >= 9)
        model.Add(sum(x[(i, d, "T2")] + x[(i, d, "T2E")] for i in range(N_EMP)) <= 11)

    pref_t1_vars, pref_t2_vars = [], []

    for i, emp in enumerate(employees):
        name, res, hist, has_pas = emp['nombre'], emp['restriccion'], emp['hist'], emp['has_pas']

        dias_bloqueados = {}
        tiene_pref = None
        for nov in novedades:
            match_destino = (filtrar_ciclo in nov.get("destino", "Ciclos Actuales"))
            if match_destino and nov["emp"].lower() in name.lower():
                if "desde" in nov:
                    curr_dt = nov["desde"]
                    while curr_dt <= nov["hasta"]:
                        if not is_festivo_or_domingo(curr_dt):
                            for d_idx, dt_curr in enumerate(dias_28):
                                if dt_curr.date() == curr_dt:
                                    dias_bloqueados[d_idx] = nov["tipo"]
                        curr_dt += datetime.timedelta(days=1)
                else:
                    tiene_pref = nov["tipo"]

        for d in range(N_DAYS):
            if "LISETTY" in name or "INCAPACITADA" in res or dias_bloqueados.get(d) == 'INC':
                model.Add(x[(i, d, "INC")] == 1)
            else:
                model.Add(x[(i, d, "INC")] == 0)

            if dias_bloqueados.get(d) == 'V':
                model.Add(x[(i, d, "V")] == 1)
            else:
                model.Add(x[(i, d, "V")] == 0)

            if dias_bloqueados.get(d) == 'CAL':
                model.Add(x[(i, d, "CAL")] == 1)
            else:
                model.Add(x[(i, d, "CAL")] == 0)

            if dias_bloqueados.get(d) == 'D':
                model.Add(x[(i, d, "D")] == 1)

        if "NO HACE T3" in res:
            for d in range(N_DAYS): model.Add(x[(i, d, "T3")] == 0)

        if not has_pas:
            for d in range(N_DAYS): model.Add(x[(i, d, "T4")] == 0)

        if "LUISA FERNANDA CUARTAS" in name:
            for d in range(N_DAYS):
                for s in ["T2", "T3", "T4", "T2E"]: model.Add(x[(i, d, s)] == 0)

        if "SHARON" in name:
            for d in range(N_DAYS):
                if dias_28[d].weekday() < 5:
                    for s in ["T2", "T3", "T4", "T1E", "T2E"]: model.Add(x[(i, d, s)] == 0)
                else:
                    for s in ["T3", "T4"]: model.Add(x[(i, d, s)] == 0)

        for d_range in [range(0, 14), range(14, 28)]:
            dias_nov = sum(1 for d in d_range if dias_bloqueados.get(d) in ['V', 'CAL'])
            tot = sum(x[(i, d, s)] * SHIFT_HOURS[s] for d in d_range for s in SHIFTS)
            ext_shifts = sum(x[(i, d, "T1E")] + x[(i, d, "T2E")] for d in d_range)
            
            if dias_nov >= 14 or "LISETTY" in name:
                model.Add(tot == 0)
                model.Add(ext_shifts == 0)
            elif "LUISA FERNANDA CUARTAS" in name:
                model.Add(tot == 42)
                model.Add(ext_shifts <= 1)
            elif dias_nov == 0:
                model.Add(tot == 84)
                model.Add(ext_shifts == 2)
            else:
                dias_activos = 14 - dias_nov
                descansos_ley = round(dias_activos * 4 / 14)
                turnos_a_laborar = dias_activos - descansos_ley
                max_ext_permitidos = 1 if dias_activos < 7 else 2
                
                min_h = turnos_a_laborar * 8
                max_h = turnos_a_laborar * 8 + (max_ext_permitidos * 2)
                model.Add(tot >= min_h)
                model.Add(tot <= max_h)
                model.Add(ext_shifts <= max_ext_permitidos)

        for d in range(N_DAYS - 1):
            for s_tarde in ["T2", "T2E"]:
                for s_man in ["T1", "T1E", "T4"]:
                    model.Add(x[(i, d, s_tarde)] + x[(i, d + 1, s_man)] <= 1)

        if "SHARON" not in name and "LUISA FERNANDA CUARTAS" not in name:
            for d in range(N_DAYS - 3):
                model.Add(sum(x[(i, d + k, "T1")] + x[(i, d + k, "T1E")] for k in range(4)) <= 3)

        for d in range(N_DAYS - 5):
            model.Add(sum(x[(i, d + k, s)] for k in range(6) for s in WORKING_SHIFTS) <= 5)

        for d in range(N_DAYS - 2):
            model.Add(x[(i, d, "T3")] + x[(i, d + 1, "T3")] + x[(i, d + 2, "T3")] <= 2)

        for d in range(N_DAYS - 1):
            model.Add(x[(i, d + 1, "T3")] + x[(i, d + 1, "D")] >= x[(i, d, "T3")])

        for d in range(N_DAYS - 3):
            dos_n = model.NewBoolVar(f"dn_{i}_{d}")
            model.Add(x[(i, d, "T3")] + x[(i, d + 1, "T3")] == 2).OnlyEnforceIf(dos_n)
            model.Add(x[(i, d, "T3")] + x[(i, d + 1, "T3")] < 2).OnlyEnforceIf(dos_n.Not())
            model.Add(x[(i, d + 2, "D")] == 1).OnlyEnforceIf(dos_n)
            model.Add(x[(i, d + 3, "D")] == 1).OnlyEnforceIf(dos_n)

        if "T3" in hist.get('d14_dom', 'D'):
            model.Add(x[(i, 0, "D")] == 1)
            if "T3" in hist.get('d13_sab', 'D'):
                model.Add(x[(i, 1, "D")] == 1)
        if "T2" in hist.get('d14_dom', 'D'):
            for s_man in ["T1", "T1E", "T4"]: model.Add(x[(i, 0, s_man)] == 0)

        if tiene_pref == "PREF_T1":
            for d_range in [range(0, 14), range(14, 28)]:
                model.Add(sum(x[(i, d, "T1")] + x[(i, d, "T1E")] for d in d_range) <= 5)
            for d in range(N_DAYS): pref_t1_vars.append(x[(i, d, "T1")] + x[(i, d, "T1E")])
        elif tiene_pref == "PREF_T2":
            for d_range in [range(0, 14), range(14, 28)]:
                model.Add(sum(x[(i, d, "T2")] + x[(i, d, "T2E")] for d in d_range) <= 5)
            for d in range(N_DAYS): pref_t2_vars.append(x[(i, d, "T2")] + x[(i, d, "T2E")])

    weekends = [(5, 6), (12, 13), (19, 20), (26, 27)]
    we_vars = []
    for i in range(N_EMP):
        if "LISETTY" in employees[i]['nombre']: continue
        for w_idx, (sab, dom) in enumerate(weekends):
            w_var = model.NewBoolVar(f"we_{i}_{w_idx}")
            model.Add(x[(i, sab, "D")] + x[(i, dom, "D")] == 2).OnlyEnforceIf(w_var)
            model.Add(x[(i, sab, "D")] + x[(i, dom, "D")] < 2).OnlyEnforceIf(w_var.Not())
            we_vars.append(w_var)

    for i, emp in enumerate(employees):
        if "NO HACE T3" not in emp['restriccion'] and "LISETTY" not in emp['nombre']:
            model.Add(sum(x[(i, d, "T3")] for d in range(N_DAYS)) <= 4)

    total_t1e = sum(x[(i, d, "T1E")] for i in range(N_EMP) for d in range(N_DAYS))
    total_t2e = sum(x[(i, d, "T2E")] for i in range(N_EMP) for d in range(N_DAYS))
    
    model.Maximize(sum(we_vars) * 10 + total_t1e - total_t2e + sum(pref_t1_vars)*3 + sum(pref_t2_vars)*3)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 45.0
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    return solver, x, employees, dias_28, SHIFTS, status

c_gen1, c_gen2 = st.columns(2)
with c_gen1:
    btn_gen = st.button("🚀 Optimizar y Generar Cuadro Base", type="primary")
with c_gen2:
    btn_prox = st.button("🔄 Generar Próximos 2 Ciclos (Continuidad Automática)")

if btn_gen or btn_prox:
    if os.path.exists(PATH_PERSONAL) and os.path.exists(PATH_HISTORIAL):
        with st.spinner("Resolviendo optimización matemática y limpiando novedades..."):
            try:
                history = {}
                filtro_ciclo_meta = "Ciclos Actuales (20 y 21)"
                if btn_prox and st.session_state.df_c21 is not None:
                    fecha_inicio = st.session_state.dias_28[-1].date() + datetime.timedelta(days=1)
                    cols_c21 = [c for c in st.session_state.df_c21.columns if c not in ['EMPLEADO', 'ROL']]
                    d_penult = cols_c21[-2]
                    d_ult = cols_c21[-1]
                    for _, r_emp in st.session_state.df_c21.iterrows():
                        emp_n = r_emp['EMPLEADO']
                        history[emp_n] = {
                            'd13_sab': str(r_emp[d_penult]).strip().upper(),
                            'd14_dom': str(r_emp[d_ult]).strip().upper()
                        }
                    filtro_ciclo_meta = "Próximos Ciclos (22 y 23 en adelante)"
                else:
                    df_gp = pd.read_excel(PATH_HISTORIAL, sheet_name='GESTOR-PLANIFICACION')
                    for r in range(13, len(df_gp)):
                        emp = df_gp.iloc[r, 0]
                        if pd.notna(emp) and isinstance(emp, str) and not emp.startswith("TOTAL") and not emp.startswith("CUMP"):
                            history[emp.strip()] = {
                                'd13_sab': str(df_gp.iloc[r, 98]).strip(),
                                'd14_dom': str(df_gp.iloc[r, 99]).strip()
                            }

                solver, x, employees, dias_28, SHIFTS, status = ejecutar_optimizador(
                    df_p_global, history, fecha_inicio, st.session_state.novedades_list, filtrar_ciclo=filtro_ciclo_meta
                )

                if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    st.session_state.dias_28 = dias_28
                    c20_cols = [f"{dias_es[d.weekday()]} {d.strftime('%d/%m')}" for d in dias_28[:14]]
                    c21_cols = [f"{dias_es[d.weekday()]} {d.strftime('%d/%m')}" for d in dias_28[14:]]

                    data20, data21 = [], []
                    for i, emp in enumerate(employees):
                        s20 = [SHIFTS[[solver.Value(x[(i, d, s)]) for s in SHIFTS].index(1)] for d in range(14)]
                        s21 = [SHIFTS[[solver.Value(x[(i, d, s)]) for s in SHIFTS].index(1)] for d in range(14, 28)]
                        data20.append({'EMPLEADO': emp['nombre'], 'ROL': emp['rol'], **{c20_cols[k]: s20[k] for k in range(14)}})
                        data21.append({'EMPLEADO': emp['nombre'], 'ROL': emp['rol'], **{c21_cols[k]: s21[k] for k in range(14)}})

                    st.session_state.df_c20 = pd.DataFrame(data20)
                    st.session_state.df_c21 = pd.DataFrame(data21)

                    with open(PATH_CUADRO_ACTUAL, "w", encoding="utf-8") as f:
                        json.dump({
                            "fecha_inicio": fecha_inicio.strftime('%Y-%m-%d'),
                            "c20": data20, "c21": data21
                        }, f, ensure_ascii=False)

                    registrar_auditoria("Generación de Cuadro", f"Cuadro generado 28 días desde {fecha_inicio}", responsable_activo)
                    st.session_state.novedades_list = []
                    st.success("¡Cuadro generado exitosamente!")
                    st.rerun()
                else:
                    st.error("No se encontró solución con los parámetros actuales.")
            except Exception as e:
                st.error(f"Error procesando archivos: {e}")
    else:
        st.warning("Asegúrate de que el catálogo de personal y el historial estén cargados en la barra lateral.")

# Módulo de Intercambio de Turnos
if st.session_state.df_c20 is not None:
    st.markdown("---")
    st.subheader("🔁 2. Intercambio de Turnos entre Compañeros con Validación Estricta")
    
    with st.container(border=True):
        col_sw1, col_sw2 = st.columns([1, 2])
        ciclo_sw = col_sw1.radio("Ciclo a Modificar:", ["Ciclo 20", "Ciclo 21"], horizontal=True)
        df_target = st.session_state.df_c20 if ciclo_sw == "Ciclo 20" else st.session_state.df_c21
        cols_ciclo = [c for c in df_target.columns if c not in ['EMPLEADO', 'ROL']]

        col_u1, col_u2, col_u3 = st.columns(3)
        colab_1 = col_u1.selectbox("Colaborador 1:", df_target['EMPLEADO'].tolist(), index=0, key="c1_sw")
        colab_2 = col_u2.selectbox("Colaborador 2:", df_target['EMPLEADO'].tolist(), index=1, key="c2_sw")
        dia_cambio = col_u3.selectbox("Día del Turno a Cambiar (Día 1):", cols_ciclo, key="d1_sw")

        t1_d1 = str(df_target.loc[df_target['EMPLEADO'] == colab_1, dia_cambio].values[0]).strip().upper()
        t2_d1 = str(df_target.loc[df_target['EMPLEADO'] == colab_2, dia_cambio].values[0]).strip().upper()
        h1_d1 = SHIFT_HOURS.get(t1_d1, 0)
        h2_d1 = SHIFT_HOURS.get(t2_d1, 0)

        st.info(f"📍 **Estado el {dia_cambio}:** {colab_1} tiene **{t1_d1} ({h1_d1}h)**  ⇄  {colab_2} tiene **{t2_d1} ({h2_d1}h)**")

        es_desbalanceado = (h1_d1 != h2_d1)
        dia_comp = None

        if es_desbalanceado:
            st.error(f"⛔ **BLOQUEO LEGAL:** El cambio del {dia_cambio} descuadra las horas. Selecciona el día de compensación recíproca:")
            dias_devolucion_disponibles = [d for d in cols_ciclo if d != dia_cambio]
            dia_comp = st.selectbox("Día de Devolución Recíproca:", ["(Seleccionar día)"] + dias_devolucion_disponibles)

        if st.button("🔒 Validar y Aplicar Cambio Legal de Turno", type="primary"):
            errores = []
            if colab_1 == colab_2: errores.append("Debes seleccionar dos colaboradores distintos.")
            if es_desbalanceado and (dia_comp is None or dia_comp == "(Seleccionar día)"):
                errores.append("BLOQUEADO: Falta el día de devolución para conservar las 84 horas.")

            if not errores:
                df_target.loc[df_target['EMPLEADO'] == colab_1, dia_cambio] = t2_d1
                df_target.loc[df_target['EMPLEADO'] == colab_2, dia_cambio] = t1_d1

                detalle_aud = f"Canje el {dia_cambio} aprobado por {responsable_activo}: {colab_1} pasa a {t2_d1}, {colab_2} pasa a {t1_d1}."
                if es_desbalanceado and dia_comp != "(Seleccionar día)":
                    t1_d2 = str(df_target.loc[df_target['EMPLEADO'] == colab_1, dia_comp].values[0]).strip().upper()
                    t2_d2 = str(df_target.loc[df_target['EMPLEADO'] == colab_2, dia_comp].values[0]).strip().upper()
                    df_target.loc[df_target['EMPLEADO'] == colab_1, dia_comp] = t2_d2
                    df_target.loc[df_target['EMPLEADO'] == colab_2, dia_comp] = t1_d2
                    detalle_aud += f" Compensado el {dia_comp}."

                registrar_auditoria("Cambio de Turno Aprobado", detalle_aud, responsable_activo)

                with open(PATH_CUADRO_ACTUAL, "w", encoding="utf-8") as f:
                    json.dump({
                        "fecha_inicio": fecha_inicio.strftime('%Y-%m-%d'),
                        "c20": st.session_state.df_c20.to_dict(orient="records"),
                        "c21": st.session_state.df_c21.to_dict(orient="records")
                    }, f, ensure_ascii=False)

                st.success("¡Cambio aplicado con 84 horas intactas!")
                st.rerun()
            else:
                for err in errores: st.error(err)

    # Planillas en Vivo y Pestañas Admin
    st.markdown("---")
    st.subheader("📊 3. Planillas de Turnos, Contadores, Nómina y Auditoría")

    dias_28 = st.session_state.dias_28
    c20_dates = dias_28[:14]
    c21_dates = dias_28[14:]
    c20_cols = [c for c in st.session_state.df_c20.columns if c not in ['EMPLEADO', 'ROL']]
    c21_cols = [c for c in st.session_state.df_c21.columns if c not in ['EMPLEADO', 'ROL']]

    def calcular_liquidacion_completa(df_turnos, cols, dates):
        salarios_db = st.session_state.salarios
        vac_db = st.session_state.vacaciones_db
        filas = []
        for _, row in df_turnos.iterrows():
            emp = row['EMPLEADO']
            rol = row['ROL']
            s_list = [str(row[c]).strip().upper() for c in cols]
            horas = sum(SHIFT_HOURS.get(s, 0) for s in s_list)
            ext = sum(1 for s in s_list if s in ["T1E", "T2E"])

            rn, dom_d, dom_n = 0, 0, 0
            for s, dt in zip(s_list, dates):
                if s in ['D', 'INC', 'V', 'CAL', '']: continue
                is_fest = is_festivo_or_domingo(dt)
                if not is_fest:
                    if s in ['T2', 'T2E']: rn += 3
                    elif s == 'T3': rn += 8
                else:
                    if s in ['T1', 'T4']: dom_d += 8
                    elif s == 'T1E': dom_d += 10
                    elif s == 'T2': dom_d += 5; dom_n += 3
                    elif s == 'T2E': dom_d += 7; dom_n += 3
                    elif s == 'T3': dom_n += 8

            sal_base = salarios_db.get(emp, 0)
            vho = sal_base / 210 if sal_base > 0 else 0
            rec_rn = rn * vho * 0.35
            rec_dom_d = dom_d * vho * 0.90
            rec_dom_n = dom_n * vho * 1.25
            basico_q = sal_base / 2
            total_dev = basico_q + rec_rn + rec_dom_d + rec_dom_n

            reg_v = vac_db.get(emp, {"acumulados_iniciales": 15, "disfrutados": 0})
            disf = reg_v["disfrutados"]
            pend = max(0, reg_v["total_ley"] - disf)
            val_vac = (sal_base / 30) * disf if sal_base > 0 else 0

            filas.append({
                "EMPLEADO": emp, "ROL": rol, "SALARIO BÁSICO": sal_base,
                "HORAS": horas, "EXT": ext,
                "VHO": round(vho, 2), "HORAS RN (35%)": rn, "VALOR RN": round(rec_rn),
                "DOM/FEST DIU (90%)": dom_d, "VALOR DOM DIU": round(rec_dom_d),
                "DOM/FEST NOC (125%)": dom_n, "VALOR DOM NOC": round(rec_dom_n),
                "BÁSICO QUINCENAL": round(basico_q), "TOTAL DEVENGADO": round(total_dev),
                "VAC. DISFRUTADAS (DÍAS)": disf, "VAC. PENDIENTES": pend, "VALOR VACACIONES": round(val_vac)
            })
        return pd.DataFrame(filas)

    def calcular_contadores(df_turnos, cols):
        filas = []
        specs = [
            ("Personal Mañana (T1/T1E)", lambda s: s in ["T1", "T1E"]),
            ("  ↳ En Turno Extendido (T1E)", lambda s: s == "T1E"),
            ("Apoyo Póliza (T4 - Solo PAS)", lambda s: s == "T4"),
            ("Personal Tarde (T2/T2E)", lambda s: s in ["T2", "T2E"]),
            ("  ↳ En Turno Extendido (T2E)", lambda s: s == "T2E"),
            ("Personal Noche (T3)", lambda s: s == "T3"),
            ("TOTAL TRABAJANDO", lambda s: s in WORKING_SHIFTS),
            ("Total Extendidos Día", lambda s: s in ["T1E", "T2E"]),
            ("Personal en Descanso (D)", lambda s: s == "D"),
            ("Incapacidad Médica (INC)", lambda s: s == "INC"),
            ("Vacaciones (V)", lambda s: s == "V"),
            ("Calamidad Doméstica (CAL)", lambda s: s == "CAL"),
            ("TOTAL PLANTA", lambda s: True)
        ]
        for lbl, fn in specs:
            r = {"Métrica / Turno": lbl}
            for c in cols: r[c] = sum(fn(str(v).strip().upper()) for v in df_turnos[c])
            filas.append(r)
        return pd.DataFrame(filas)

    tab_c20, tab_c21, tab_sal, tab_vac, tab_aud, tab_cal = st.tabs([
        "📌 CICLO 20", "📌 CICLO 21", "💵 Maestro de Salarios Base", "🌴 Control Vacaciones", "📜 Auditoría", "📅 Calendario Personal"
    ])

    with tab_c20:
        ed_c20 = st.data_editor(st.session_state.df_c20, key="editor_live_c20", use_container_width=True)
        st.session_state.df_c20 = ed_c20
        st.markdown("##### 📊 Contador Diario")
        st.dataframe(calcular_contadores(ed_c20, c20_cols), use_container_width=True)
        st.markdown("##### 💰 Resumen de Nómina")
        st.dataframe(calcular_liquidacion_completa(ed_c20, c20_cols, c20_dates), use_container_width=True)

    with tab_c21:
        ed_c21 = st.data_editor(st.session_state.df_c21, key="editor_live_c21", use_container_width=True)
        st.session_state.df_c21 = ed_c21
        st.markdown("##### 📊 Contador Diario")
        st.dataframe(calcular_contadores(ed_c21, c21_cols), use_container_width=True)
        st.markdown("##### 💰 Resumen de Nómina")
        st.dataframe(calcular_liquidacion_completa(ed_c21, c21_cols, c21_dates), use_container_width=True)

    with tab_sal:
        st.markdown("#### 💵 Maestro de Salarios Básicos Mensuales")
        salarios_actuales = st.session_state.salarios
        data_sal_editor = [{"EMPLEADO": emp_n, "SALARIO BÁSICO MENSUAL": float(salarios_actuales.get(emp_n, 0))} for emp_n in nombres_emp]
        sal_editado = st.data_editor(pd.DataFrame(data_sal_editor), use_container_width=True, key="sal_editor_grid_fixed")

        if st.button("💾 Guardar Salarios de Forma Permanente"):
            nuevo_dict = {r_s['EMPLEADO']: float(r_s['SALARIO BÁSICO MENSUAL']) for _, r_s in sal_editado.iterrows()}
            guardar_salarios(nuevo_dict)
            st.session_state.salarios = nuevo_dict
            registrar_auditoria("Actualización Salarial", "Se actualizaron los salarios básicos.", responsable_activo)
            st.success("¡Salarios guardados!")
            st.rerun()

    with tab_vac:
        st.markdown("#### 🌴 Control y Saldo de Días de Vacaciones por Colaborador")
        vac_db_current = st.session_state.vacaciones_db
        data_vac_editor = []
        for emp_n in nombres_emp:
            v_info = vac_db_current.get(emp_n, {"acumulados_iniciales": 15, "disfrutados": 0})
            data_vac_editor.append({
                "EMPLEADO": emp_n,
                "ACUMULADOS INICIALES (SALDO INICIAL)": float(v_info.get("acumulados_iniciales", 15)),
                "DISFRUTADOS (HÁBILES)": float(v_info.get("disfrutados", 0))
            })

        vac_edited = st.data_editor(pd.DataFrame(data_vac_editor), use_container_width=True, key="vac_editor_grid_custom")

        if st.button("💾 Guardar y Actualizar Saldos de Vacaciones"):
            new_vac_db = {}
            for _, r_v in vac_edited.iterrows():
                e_name = r_v['EMPLEADO']
                ac_init = float(r_v['ACUMULADOS INICIALES (SALDO INICIAL)'])
                disf = float(r_v['DISFRUTADOS (HÁBILES)'])
                new_vac_db[e_name] = {
                    "acumulados_iniciales": ac_init,
                    "disfrutados": disf
                }
            st.session_state.vacaciones_db = new_vac_db
            guardar_json(PATH_VACACIONES, new_vac_db)
            registrar_auditoria("Actualización Vacaciones", "Se actualizaron los saldos iniciales y días de vacaciones.", responsable_activo)
            st.success("¡Saldos de vacaciones guardados y actualizados correctamente!")
            st.rerun()

        st.markdown("##### 📊 Resumen General de Vacaciones y Liquidación")
        resumen_vac_rows = []
        for emp_n in nombres_emp:
            v_info = st.session_state.vacaciones_db.get(emp_n, {"acumulados_iniciales": 15, "disfrutados": 0})
            ac_init = v_info.get("acumulados_iniciales", 15)
            disf = v_info.get("disfrutados", 0)
            pend = max(0, ac_init - disf)
            sal = st.session_state.salarios.get(emp_n, 0)
            val_liq = (sal / 30) * disf if sal > 0 else 0
            resumen_vac_rows.append({
                "EMPLEADO": emp_n,
                "SALDO INICIAL ACUMULADO": ac_init,
                "DISFRUTADOS (LUN-SÁB)": disf,
                "SALDO PENDIENTE": pend,
                "VALOR LIQUIDADO": round(val_liq)
            })
        st.dataframe(pd.DataFrame(resumen_vac_rows), use_container_width=True)

    with tab_aud:
        st.markdown("#### 📜 Libro de Registro y Auditoría (Gestión y Edición)")
        st.caption("Puedes revisar, editar o eliminar registros de auditoría anteriores.")
        df_aud = obtener_auditoria()
        if not df_aud.empty:
            edited_aud = st.data_editor(df_aud, use_container_width=True, key="aud_editor_crud")
            if st.button("💾 Guardar Cambios en Auditoría"):
                edited_aud.to_json(PATH_AUDITORIA, orient="records", force_ascii=False, indent=2)
                st.success("¡Auditoría actualizada correctamente!")
                st.rerun()
        else:
            st.info("Aún no hay registros de auditoría.")

    with tab_cal:
        st.markdown("#### 📅 Generador de Tarjeta de Calendario Personal (SURA)")
        emp_cal = st.selectbox("Seleccionar Colaborador:", st.session_state.df_c20['EMPLEADO'].tolist(), key="sel_emp_cal")
        if emp_cal:
            r20 = st.session_state.df_c20.loc[st.session_state.df_c20['EMPLEADO'] == emp_cal].iloc[0]
            r21 = st.session_state.df_c21.loc[st.session_state.df_c21['EMPLEADO'] == emp_cal].iloc[0]
            
            df_cal_emp = pd.DataFrame({
                "Fecha / Día": c20_cols + c21_cols,
                "Turno Asignado": [r20[c] for c in c20_cols] + [r21[c] for c in c21_cols]
            })
            st.dataframe(df_cal_emp, use_container_width=True)
            
            fig, ax = plt.subplots(figsize=(14, 6))
            ax.axis('off')
            
            bg = patches.Rectangle((0, 0), 1, 1, transform=ax.transAxes, facecolor='#F8F9FA', edgecolor='#003366', linewidth=2, zorder=0)
            ax.add_patch(bg)
            
            banner = patches.Rectangle((0, 0.82), 1, 0.18, transform=ax.transAxes, facecolor='#003366', zorder=1)
            ax.add_patch(banner)
            
            ax.text(0.03, 0.92, "SURA — Gestión de Turnos Operativos", fontsize=14, fontweight='bold', color='white', zorder=2)
            ax.text(0.03, 0.86, f"CRONOGRAMA PERSONAL DE TURNOS | Ciclos 20 y 21", fontsize=10, color='#B0C4DE', zorder=2)
            
            ax.text(0.03, 0.73, f"Colaborador: {emp_cal}", fontsize=12, fontweight='bold', color='#003366')
            ax.text(0.03, 0.68, f"Rol: {r20['ROL'][:50]}", fontsize=9, color='#555555')
            
            colors_map = {
                "T1": "#D1ECF1", "T1E": "#99CCFF", "T2": "#FFE8D6", "T2E": "#FFCC80",
                "T3": "#F8D7DA", "T4": "#FFF3CD", "D": "#D4EDDA", "INC": "#F5C6CB", "V": "#B2EBF2", "CAL": "#E1BEE7"
            }
            
            all_cols = c20_cols + c21_cols
            all_vals = [r20[c] for c in c20_cols] + [r21[c] for c in c21_cols]
            
            for i, (d_txt, s_val) in enumerate(zip(all_cols, all_vals)):
                row = i // 7
                col = i % 7
                x = 0.03 + col * 0.138
                y = 0.38 - row * 0.28
                bg_col = colors_map.get(s_val, "#FFFFFF")
                
                box = patches.FancyBboxPatch((x, y), 0.13, 0.22, boxstyle="round,pad=0.02", facecolor=bg_col, edgecolor="#003366", linewidth=1.2, zorder=2)
                ax.add_patch(box)
                
                ax.text(x + 0.065, y + 0.15, d_txt, fontsize=9, fontweight='bold', ha='center', color="#333333", zorder=3)
                ax.text(x + 0.065, y + 0.07, s_val, fontsize=15, fontweight='bold', ha='center', color="#111111", zorder=3)

            buf_img = io.BytesIO()
            plt.savefig(buf_img, format='png', bbox_inches='tight', dpi=200)
            buf_img.seek(0)
            plt.close()
            
            st.download_button(
                label=f"⬇️ Descargar Tarjeta de Calendario de {emp_cal[:15]} (Imagen PNG con Logo SURA)",
                data=buf_img.getvalue(),
                file_name=f"Calendario_SURA_{emp_cal.replace(' ', '_')}.png",
                mime="image/png"
            )

    # -------------------------------------------------------------
    # 7. EXPORTADOR EXCEL INTEGRAL (2 OPCIONES)
    # -------------------------------------------------------------
    st.markdown("---")
    st.subheader("📥 Centro de Descargas (Opciones de Exportación)")
    
    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("##### Opción 1: Libro Maestro Completo (Administración)")
        st.caption("Incluye Planilla de Turnos, contadores, liquidación de nómina con salarios, control de vacaciones y auditoría.")
        
        def exportar_excel_maestro():
            wb = openpyxl.Workbook()
            border_thin = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
                                 top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            shift_colors = {
                "T1":  {"fill": "D1ECF1", "font": "0C5460"}, "T1E": {"fill": "99CCFF", "font": "002D62"},
                "T2":  {"fill": "FFE8D6", "font": "7C3701"}, "T2E": {"fill": "FFCC80", "font": "BF360C"},
                "T3":  {"fill": "F8D7DA", "font": "721C24"}, "T4":  {"fill": "FFF3CD", "font": "856404"},
                "D":   {"fill": "D4EDDA", "font": "155724"}, "INC": {"fill": "F5C6CB", "font": "721C24"},
                "V":   {"fill": "B2EBF2", "font": "006064"}, "CAL": {"fill": "E1BEE7", "font": "4A148C"}
            }

            def escribir_ciclo(ws, df_datos, dates, c_title):
                ws.merge_cells('A1:R1')
                ws['A1'] = f"CUADRO DE TURNOS Y OCUPACIÓN - {c_title}"
                ws['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
                ws['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

                cols = [c for c in df_datos.columns if c not in ['EMPLEADO', 'ROL']]
                headers = ["EMPLEADO", "ROL"] + cols + ["HORAS", "EXT"]
                for col_num, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_num, value=h_text)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if 3 <= col_num <= 16 and is_festivo_or_domingo(dates[col_num - 3]):
                        cell.fill = PatternFill(start_color="E6B8B8", end_color="E6B8B8", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="900C3F")
                    else:
                        cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
                    cell.border = border_thin

                curr_r = 4
                for _, r_data in df_datos.iterrows():
                    ws.cell(row=curr_r, column=1, value=r_data['EMPLEADO']).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=curr_r, column=1).border = border_thin
                    ws.cell(row=curr_r, column=2, value=r_data['ROL']).font = Font(name="Calibri", size=9, color="555555")
                    ws.cell(row=curr_r, column=2).border = border_thin

                    s_list = [str(r_data[c]).strip().upper() for c in cols]
                    for d_i, s_val in enumerate(s_list):
                        c_cell = ws.cell(row=curr_r, column=3 + d_i, value=s_val)
                        c_cell.alignment = Alignment(horizontal="center", vertical="center")
                        c_cell.border = border_thin
                        if s_val in shift_colors:
                            c_cell.fill = PatternFill(start_color=shift_colors[s_val]["fill"], end_color=shift_colors[s_val]["fill"], fill_type="solid")
                            c_cell.font = Font(name="Calibri", size=10, bold=True, color=shift_colors[s_val]["font"])

                    ws.cell(row=curr_r, column=17, value=sum(SHIFT_HOURS.get(s, 0) for s in s_list)).font = Font(name="Calibri", size=10, bold=True, color="1F497D")
                    ws.cell(row=curr_r, column=17).fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
                    ws.cell(row=curr_r, column=17).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=curr_r, column=17).border = border_thin

                    ws.cell(row=curr_r, column=18, value=sum(1 for s in s_list if s in ["T1E", "T2E"])).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=curr_r, column=18).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=curr_r, column=18).border = border_thin
                    curr_r += 1

                curr_r += 1
                ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=18)
                ws.cell(row=curr_r, column=1, value="📊 CONTADOR Y CONTROL DIARIO DE OCUPACIÓN").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                ws.cell(row=curr_r, column=1).fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                ws.cell(row=curr_r, column=1).alignment = Alignment(horizontal="center", vertical="center")
                curr_r += 1

                df_cnt_excel = calcular_contadores(df_datos, cols)
                for _, r_cnt in df_cnt_excel.iterrows():
                    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=2)
                    lbl_c = ws.cell(row=curr_r, column=1, value=r_cnt['Métrica / Turno'])
                    lbl_c.font = Font(name="Calibri", size=9, bold=True)
                    lbl_c.border = border_thin
                    ws.cell(row=curr_r, column=2).border = border_thin
                    for d_i, col_name in enumerate(cols):
                        v_cell = ws.cell(row=curr_r, column=3 + d_i, value=r_cnt[col_name])
                        v_cell.font = Font(name="Calibri", size=9, bold=True)
                        v_cell.alignment = Alignment(horizontal="center", vertical="center")
                        v_cell.border = border_thin
                    ws.cell(row=curr_r, column=17).border = border_thin
                    ws.cell(row=curr_r, column=18).border = border_thin
                    curr_r += 1

                ws.column_dimensions['A'].width = 30
                ws.column_dimensions['B'].width = 24
                for d_i in range(14): ws.column_dimensions[get_column_letter(3 + d_i)].width = 11

            def escribir_nomina(ws_pay, df_liq, c_title):
                ws_pay.merge_cells('A1:O1')
                ws_pay['A1'] = f"LIQUIDACIÓN DE NÓMINA Y VACACIONES - {c_title}"
                ws_pay['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
                ws_pay['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                ws_pay['A1'].alignment = Alignment(horizontal="center", vertical="center")

                headers = ["EMPLEADO", "ROL", "SALARIO BÁSICO", "VALOR HORA", "HORAS RN (35%)", "VALOR RN",
                           "DOM/FEST DIU (90%)", "VALOR DOM DIU", "DOM/FEST NOC (125%)", "VALOR DOM NOC",
                           "BÁSICO QUINCENAL", "TOTAL DEVENGADO", "VAC. DISFRUTADAS", "VAC. PENDIENTES", "VALOR VACACIONES"]
                for col_num, h_text in enumerate(headers, 1):
                    cell = ws_pay.cell(row=3, column=col_num, value=h_text)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                    cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
                    cell.border = border_thin

                for r_i, r_val in df_liq.iterrows():
                    row_idx = 4 + r_i
                    ws_pay.cell(row=row_idx, column=1, value=r_val['EMPLEADO']).font = Font(name="Calibri", size=10, bold=True)
                    ws_pay.cell(row=row_idx, column=2, value=r_val['ROL']).font = Font(name="Calibri", size=9, color="555555")

                    cols_liq = ["SALARIO BÁSICO", "VHO", "HORAS RN (35%)", "VALOR RN", "DOM/FEST DIU (90%)",
                                "VALOR DOM DIU", "DOM/FEST NOC (125%)", "VALOR DOM NOC", "BÁSICO QUINCENAL", "TOTAL DEVENGADO",
                                "VAC. DISFRUTADAS (DÍAS)", "VAC. PENDIENTES", "VALOR VACACIONES"]
                    for c_k, col_name in enumerate(cols_liq, 3):
                        c_cell = ws_pay.cell(row=row_idx, column=c_k, value=r_val[col_name])
                        c_cell.border = border_thin
                        if "SALARIO" in col_name or "VALOR" in col_name or "BÁSICO" in col_name or "TOTAL" in col_name:
                            c_cell.number_format = '$#,##0'
                        if col_name == "TOTAL DEVENGADO":
                            c_cell.font = Font(name="Calibri", size=10, bold=True, color="155724")
                            c_cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")

                ws_pay.column_dimensions['A'].width = 30
                ws_pay.column_dimensions['B'].width = 24
                ws_pay.column_dimensions['C'].width = 18

            ws1 = wb.active
            ws1.title = "CICLO 20"
            escribir_ciclo(ws1, st.session_state.df_c20, c20_dates, "CICLO 20")
            escribir_nomina(wb.create_sheet(title="LIQ NÓMINA C20"), calcular_liquidacion_completa(st.session_state.df_c20, c20_cols, c20_dates), "CICLO 20")

            ws2 = wb.create_sheet(title="CICLO 21")
            escribir_ciclo(ws2, st.session_state.df_c21, c21_dates, "CICLO 21")
            escribir_nomina(wb.create_sheet(title="LIQ NÓMINA C21"), calcular_liquidacion_completa(st.session_state.df_c21, c21_cols, c21_dates), "CICLO 21")

            ws_aud = wb.create_sheet(title="REGISTRO AUDITORÍA")
            df_aud_exp = obtener_auditoria()
            if not df_aud_exp.empty:
                for c_i, col_name in enumerate(df_aud_exp.columns, 1):
                    ws_aud.cell(row=1, column=c_i, value=col_name).font = Font(bold=True)
                for r_i, r_val in df_aud_exp.iterrows():
                    for c_i, val in enumerate(r_val, 1):
                        ws_aud.cell(row=2 + r_i, column=c_i, value=str(val))

            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        st.download_button(
            label="⬇️ Descargar Libro Maestro (Admin con Nómina)",
            data=exportar_excel_maestro(),
            file_name="LIBRO_MAESTRO_ADMINISTRATIVO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_d2:
        st.markdown("##### Opción 2: Cuadro Exclusivo para Empleados")
        st.caption("Contiene únicamente las planillas de turnos de ambos ciclos con sus contadores al pie, sin salarios ni datos de nómina.")
        
        def exportar_excel_empleados():
            wb = openpyxl.Workbook()
            border_thin = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
                                 top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            shift_colors = {
                "T1":  {"fill": "D1ECF1", "font": "0C5460"}, "T1E": {"fill": "99CCFF", "font": "002D62"},
                "T2":  {"fill": "FFE8D6", "font": "7C3701"}, "T2E": {"fill": "FFCC80", "font": "BF360C"},
                "T3":  {"fill": "F8D7DA", "font": "721C24"}, "T4":  {"fill": "FFF3CD", "font": "856404"},
                "D":   {"fill": "D4EDDA", "font": "155724"}, "INC": {"fill": "F5C6CB", "font": "721C24"},
                "V":   {"fill": "B2EBF2", "font": "006064"}, "CAL": {"fill": "E1BEE7", "font": "4A148C"}
            }

            def escribir_ciclo_emp(ws, df_datos, dates, c_title):
                ws.merge_cells('A1:R1')
                ws['A1'] = f"CUADRO DE TURNOS - {c_title}"
                ws['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
                ws['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

                cols = [c for c in df_datos.columns if c not in ['EMPLEADO', 'ROL']]
                headers = ["EMPLEADO", "ROL"] + cols + ["HORAS", "EXT"]
                for col_num, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_num, value=h_text)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if 3 <= col_num <= 16 and is_festivo_or_domingo(dates[col_num - 3]):
                        cell.fill = PatternFill(start_color="E6B8B8", end_color="E6B8B8", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="900C3F")
                    else:
                        cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
                    cell.border = border_thin

                curr_r = 4
                for _, r_data in df_datos.iterrows():
                    ws.cell(row=curr_r, column=1, value=r_data['EMPLEADO']).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=curr_r, column=1).border = border_thin
                    ws.cell(row=curr_r, column=2, value=r_data['ROL']).font = Font(name="Calibri", size=9, color="555555")
                    ws.cell(row=curr_r, column=2).border = border_thin

                    s_list = [str(r_data[c]).strip().upper() for c in cols]
                    for d_i, s_val in enumerate(s_list):
                        c_cell = ws.cell(row=curr_r, column=3 + d_i, value=s_val)
                        c_cell.alignment = Alignment(horizontal="center", vertical="center")
                        c_cell.border = border_thin
                        if s_val in shift_colors:
                            c_cell.fill = PatternFill(start_color=shift_colors[s_val]["fill"], end_color=shift_colors[s_val]["fill"], fill_type="solid")
                            c_cell.font = Font(name="Calibri", size=10, bold=True, color=shift_colors[s_val]["font"])

                    ws.cell(row=curr_r, column=17, value=sum(SHIFT_HOURS.get(s, 0) for s in s_list)).font = Font(name="Calibri", size=10, bold=True, color="1F497D")
                    ws.cell(row=curr_r, column=17).fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
                    ws.cell(row=curr_r, column=17).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=curr_r, column=17).border = border_thin

                    ws.cell(row=curr_r, column=18, value=sum(1 for s in s_list if s in ["T1E", "T2E"])).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=curr_r, column=18).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=curr_r, column=18).border = border_thin
                    curr_r += 1

                curr_r += 1
                ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=18)
                ws.cell(row=curr_r, column=1, value="📊 CONTADOR Y CONTROL DIARIO DE OCUPACIÓN").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                ws.cell(row=curr_r, column=1).fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                ws.cell(row=curr_r, column=1).alignment = Alignment(horizontal="center", vertical="center")
                curr_r += 1

                df_cnt_emp = calcular_contadores(df_datos, cols)
                for _, r_cnt in df_cnt_emp.iterrows():
                    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=2)
                    lbl_c = ws.cell(row=curr_r, column=1, value=r_cnt['Métrica / Turno'])
                    lbl_c.font = Font(name="Calibri", size=9, bold=True)
                    lbl_c.border = border_thin
                    ws.cell(row=curr_r, column=2).border = border_thin
                    for d_i, col_name in enumerate(cols):
                        v_cell = ws.cell(row=curr_r, column=3 + d_i, value=r_cnt[col_name])
                        v_cell.font = Font(name="Calibri", size=9, bold=True)
                        v_cell.alignment = Alignment(horizontal="center", vertical="center")
                        v_cell.border = border_thin
                    ws.cell(row=curr_r, column=17).border = border_thin
                    ws.cell(row=curr_r, column=18).border = border_thin
                    curr_r += 1

                ws.column_dimensions['A'].width = 30
                ws.column_dimensions['B'].width = 24
                for d_i in range(14): ws.column_dimensions[get_column_letter(3 + d_i)].width = 11

            ws1 = wb.active
            ws1.title = "CICLO 20"
            escribir_ciclo_emp(ws1, st.session_state.df_c20, c20_dates, "CICLO 20")

            ws2 = wb.create_sheet(title="CICLO 21")
            escribir_ciclo_emp(ws2, st.session_state.df_c21, c21_dates, "CICLO 21")

            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        st.download_button(
            label="⬇️ Descargar Cuadro para Empleados (Sin Nómina)",
            data=exportar_excel_empleados(),
            file_name="CUADRO_TURNOS_PUBLICO_EMPLEADOS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )