
import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from vrp_solver import solve_vrp_data, format_solution, generate_folium_map
import io
import os
import math

st.set_page_config(page_title="Gestión de Rutas - Sodexo", layout="wide", page_icon="🚛")

# --- LOGIN CONFIG ---
USUARIO_VALIDO = "usuario_prueba1"
PASSWORD_VALIDO = "prueba123"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:

    st.markdown("<br><br>", unsafe_allow_html=True)

    # Logo centrado
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)

        st.markdown("### Acceso al Sistema")

        usuario = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")

        if st.button("Iniciar Sesión", use_container_width=True):
            if usuario == USUARIO_VALIDO and password == PASSWORD_VALIDO:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")

    st.stop()


# --- CUSTOM CSS (SODEXO BRANDING) ---
st.markdown("""
    <style>
    /* Main Background */
    .stApp {
        background-color: #F4F6F8;
    }
    
    /* Headers - Sodexo Navy */
    h1, h2, h3, h4, h5, h6 {
        color: #262262 !important;
        font-family: 'Segoe UI', sans-serif;
    }
    
    /* Buttons - Sodexo Red */
    .stButton button {
        background-color: #EF4044 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
    }
    .stButton button:hover {
        background-color: #D12F33 !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.2) !important;
    }
    
    /* Metrics Styles */
    div[data-testid="stMetricValue"] {
        color: #EF4044 !important;
        font-weight: bold;
    }
    div[data-testid="stMetricLabel"] {
        color: #5D5D5D !important;
    }
    
    /* Cards/Containers (Simulated with standard Streamlit containers, but we can style markers) */
    div[data-testid="stExpander"] {
        border-color: #E0E0E0 !important;
        border-radius: 8px !important;
        background-color: white !important;
    }
    
    /* Tables */
    thead tr th:first-child {display:none}
    tbody th {display:none}
    
    /* INPUTS & MENUS - Gray Background, Blue Text */
    .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea {
        background-color: #E0E4E8 !important;
        color: #262262 !important;
        border-radius: 5px;
        border: 1px solid #B0B0B0 !important;
    }
    
    /* Dropdown Menu Container - TARGET POP OVERS */
    div[data-baseweb="popover"], div[data-baseweb="popover"] > div, ul[data-baseweb="menu"] {
        background-color: #E0E4E8 !important;
    }
    
    /* Dropdown Options Text */
    li[data-baseweb="option"] span, li[data-baseweb="option"] div, .stSelectbox [data-baseweb="menu"] li {
        color: #262262 !important;
    }
    
    /* Selected Option Background (Hover) */
    li[data-baseweb="option"]:hover, li[data-baseweb="option"][aria-selected="true"] {
        background-color: #CCD3D9 !important;
    }

    /* Selected Text in Input Box */
    div[data-baseweb="select"] span {
        color: #262262 !important;
    }
    
    /* Remove white/dark default backgrounds on the list container */
    .stSelectbox ul { 
        background-color: #E0E4E8 !important; 
    }
    
    /* HEADERS - FORCE BLUE (Fix for White Text Issue) */
    h1, h2, h3, h4, h5, h6, 
    .stHeadingContainer h1, .stHeadingContainer h2, .stHeadingContainer h3,
    div[data-testid="stMarkdownContainer"] h1, div[data-testid="stMarkdownContainer"] h2, div[data-testid="stMarkdownContainer"] h3,
    div[data-testid="stMarkdownContainer"] p, 
    .stMarkdown label p {
        color: #262262 !important;
        font-family: 'Segoe UI', sans-serif;
    }
    
    /* Ensure widgets label texts are also blue */
    .stSelectbox label, .stTextInput label, .stNumberInput label {
        color: #262262 !important;
    }
    
    /* Metric Labels */
    div[data-testid="stMetricLabel"] {
        color: #262262 !important;
    }
    
    /* DATAFRAMES / TABLES - Force Gray Background */
    [data-testid="stDataFrame"], [data-testid="stTable"] {
        background-color: #E0E4E8 !important;
    }
    
    </style>
""", unsafe_allow_html=True)

# --- SESSION STATE INITIALIZATION ---
if 'stage' not in st.session_state:
    st.session_state.stage = 'input_tickets' # input_tickets, fleet_config, results
if 'daily_tickets' not in st.session_state:
    st.session_state.daily_tickets = [] # List of dicts
if 'master_db' not in st.session_state:
    st.session_state.master_db = None
if 'optimization_result' not in st.session_state:
    st.session_state.optimization_result = None
if 'vehicles_config' not in st.session_state:
    st.session_state.vehicles_config = [] # List of dicts representing configured vehicles

# --- CONSTANTS ---
MASTER_FILE_PATH = "VRP_Spreadsheet_Solver_v3.8 14.05.xlsm"
TECHNICIAN_FILE_PATH = "PLANILLA FIX_FY25 - RUTOGRAMA.xlsx"

import shutil

# --- HELPER FUNCTIONS ---
@st.cache_data
def load_technicians(path):
    if not os.path.exists(path):
        return None
    try:
        # Read the 'FIX_OPE' sheet
        df = pd.read_excel(path, sheet_name='FIX_OPE')
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Expected columns: 'NOMBRES', 'Familia 1', 'Familia 2 '
        # We need to standardize 'Familia 2 ' to 'Familia 2'
        df = df.rename(columns=lambda x: x.strip())
        
        # Fill NA
        df['Familia 1'] = df['Familia 1'].fillna('')
        df['Familia 2'] = df['Familia 2'].fillna('')
        
        return df
    except Exception as e:
        st.error(f"Error al cargar técnicos: {e}")
        return None
@st.cache_data
def load_master_db(path):
    if not os.path.exists(path):
        return None
    try:
        # Try reading directly
        df = pd.read_excel(path, sheet_name='1 ubicaciones')
        df.columns = df.columns.str.strip()
        # FORCE NUMERIC
        for col in ['Latitud (y)', 'Longitud (x)']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df
    except PermissionError:
        # File might be open. Copy to temp and read.
        try:
            temp_path = "temp_master_copy.xlsm"
            shutil.copy2(path, temp_path)
            df = pd.read_excel(temp_path, sheet_name='1 ubicaciones')
            df.columns = df.columns.str.strip()
            # FORCE NUMERIC
            for col in ['Latitud (y)', 'Longitud (x)']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            return df
        except Exception as e:
            st.error(f"El archivo parece estar abierto y no se pudo copiar. Por favor ciérrelo. Error: {e}")
            return None
    except Exception as e:
        st.error(f"Error al cargar la base de datos maestra: {e}")
        return None

def style_dataframe(df):
    """Applies Sodexo branding to DataFrames"""
    return df.style.set_properties(**{
        'background-color': '#E0E4E8',
        'color': '#262262',
        'border-color': '#FFFFFF'
    }).set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#CCD3D9'), ('color', '#262262'), ('font-weight', 'bold')]}
    ])

def reset_app():
    st.session_state.stage = 'input_tickets'
    st.session_state.daily_tickets = []
    st.session_state.optimization_result = None

# --- APP HEADER ---
# --- APP HEADER ---
# 1. Logo (Left Aligned)
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    # Fallback text
    st.markdown("<h1 style='color: #262262; font-size: 40px; margin:0; padding:0;'>SODEXO</h1>", unsafe_allow_html=True)

# 2. Title (Centered and Below Logo)
st.markdown("<h1 style='text-align: center; color: #262262; padding-top: 10px;'>Planificador de Rutas</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #5D5D5D; font-weight: bold;'>Gestión Inteligente de Flota y Entregas - Sodexo Perú</p>", unsafe_allow_html=True)

if st.button("🔄 Reiniciar Aplicación"):
    reset_app()
    st.rerun()

if st.button("🔐 Cerrar Sesión"):
    st.session_state.authenticated = False
    st.rerun()

# --- LOAD DATABASE (ONCE) ---
if st.session_state.master_db is None:
    df_db = load_master_db(MASTER_FILE_PATH)
    if df_db is not None:
        st.session_state.master_db = df_db
        st.success("✅ Base de Datos de Oficinas cargada correctamente.")
    else:
        st.error(f"❌ No se encontró el archivo maestro en: {MASTER_FILE_PATH}")
        uploaded = st.file_uploader("Por favor cargue el archivo 'VRP_Spreadsheet_Solver_v3.8 14.05.xlsm' manualmente:", type=["xlsx", "xlsm"])
        if uploaded:
            st.session_state.master_db = pd.read_excel(uploaded, sheet_name='1 ubicaciones')
            st.session_state.master_db.columns = st.session_state.master_db.columns.str.strip()
            st.rerun()
        else:
            st.stop()

# --- STAGE 1: INGRESO DE TICKETS ---
if st.session_state.stage == 'input_tickets':
    st.header("1️⃣ Ingreso de Tickets del Día")
    
    col_input, col_table = st.columns([1, 2])
    
    with col_input:
        tab_manual, tab_import = st.tabs(["📝 Manual", "📂 Importar Excel"])
        
        with tab_manual:
            st.subheader("Nuevo Ticket Individual")

            # --- CLIENT FILTER ---
            if 'Habla a' in st.session_state.master_db.columns:
                clients = sorted(st.session_state.master_db['Habla a'].astype(str).unique().tolist())
                selected_client = st.selectbox("Filtrar por Cliente", options=["Todos"] + clients)
            else:
                options_clients = ["Todos"]
                selected_client = st.selectbox("Filtrar por Cliente", options=options_clients, disabled=True)

            # Filter Options
            if selected_client != "Todos":
                filtered_db = st.session_state.master_db[st.session_state.master_db['Habla a'].astype(str) == selected_client]
                office_options = filtered_db['Nombre'].unique().tolist()
            else:
                office_options = st.session_state.master_db['Nombre'].unique().tolist()

            # --- MANUAL ENTRY FORM ---
            with st.form("ticket_form", clear_on_submit=True):
                selected_office = st.selectbox("Seleccionar Oficina", options=sorted(office_options))
                
                ticket_id = st.text_input("Nro Ticket (ID)")
                familia = st.text_input("Familia / Especialidad")
                
                add_btn = st.form_submit_button("➕ Agregar a la Lista")
                
                if add_btn:
                    if not ticket_id:
                        st.warning("Por favor ingrese un número de ticket.")
                    else:
                        # Find coords for selected office
                        office_data = st.session_state.master_db[st.session_state.master_db['Nombre'] == selected_office].iloc[0]
                        
                        st.session_state.daily_tickets.append({
                            "Nombre": selected_office,
                            "Habla a": office_data.get('Habla a', ''),
                            "Ticket": ticket_id,
                            "Familia": familia,
                            "Latitud (y)": office_data['Latitud (y)'],
                            "Longitud (x)": office_data['Longitud (x)'],
                            "Importe de la entrega": 1 
                        })
                        st.toast(f"Ticket {ticket_id} agregado!", icon="👍")
        
        # --- TAB IMPORT ---
        with tab_import:
            st.subheader("Carga Masiva")
            uploaded_tickets = st.file_uploader("Subir Excel (Columnas: Oficina, Ticket, Familia)", type=["xlsx", "xls", "csv"])
            
            if uploaded_tickets:
                try:
                    if uploaded_tickets.name.endswith('.csv'):
                        df_upload = pd.read_csv(uploaded_tickets)
                    else:
                        df_upload = pd.read_excel(uploaded_tickets)
                    
                    st.write("Vista Previa:", df_upload.head(3))
                    
                    if st.button("Procesar Archivo"):
                        # Column Mapping Logic
                        # We need 'Oficina' (to match Master DB), 'Ticket', 'Familia'
                        cols = df_upload.columns.str.lower()
                        
                        # Guess column names
                        col_oficina = next((c for c in df_upload.columns if 'oficina' in c.lower() or 'nombre' in c.lower()), None)
                        col_ticket = next((c for c in df_upload.columns if 'ticket' in c.lower() or 'numero' in c.lower()), None)
                        col_familia = next((c for c in df_upload.columns if 'familia' in c.lower()), None)
                        
                        if not col_oficina:
                            st.error("No se encontró columna para 'Oficina' o 'Nombre'.")
                        else:
                            success_count = 0
                            fail_count = 0
                            
                            for _, row in df_upload.iterrows():
                                office_name = row[col_oficina]
                                # Match with Master DB
                                match = st.session_state.master_db[st.session_state.master_db['Nombre'] == office_name]
                                
                                if not match.empty:
                                    office_data = match.iloc[0]
                                    st.session_state.daily_tickets.append({
                                        "Nombre": office_name,
                                        "Habla a": office_data.get('Habla a', ''),
                                        "Ticket": row[col_ticket] if col_ticket else "N/A",
                                        "Familia": row[col_familia] if col_familia else "General",
                                        "Latitud (y)": office_data['Latitud (y)'],
                                        "Longitud (x)": office_data['Longitud (x)'],
                                        "Importe de la entrega": 1
                                    })
                                    success_count += 1
                                else:
                                    fail_count += 1
                                    
                            st.success(f"Procesado: {success_count} tickets agregados.")
                            if fail_count > 0:
                                st.warning(f"{fail_count} oficinas no encontradas en la Base Maestra.")
                                
                except Exception as e:
                    st.error(f"Error procesando: {e}")

    with col_table:
        st.subheader("📋 Lista de Pendientes")
        if st.session_state.daily_tickets:
            df_display = pd.DataFrame(st.session_state.daily_tickets)
            st.dataframe(style_dataframe(df_display[['Nombre', 'Ticket', 'Familia']]), use_container_width=True)
            
            if st.button("✅ Confirmar y Configurar Flota", type="primary"):
                st.session_state.stage = 'fleet_config'
                st.rerun()
        else:
            st.info("Agregue tickets para continuar.")

# --- STAGE 2: CONFIGURACION DE FLOTA ---
elif st.session_state.stage == 'fleet_config':
    st.header("2️⃣ Configuración de Cuadrillas y Flota")
    
    # Load Technicians
    df_tech = load_technicians(TECHNICIAN_FILE_PATH)
    if df_tech is None:
        st.error(f"No se pudo cargar el archivo de técnicos: {TECHNICIAN_FILE_PATH}")
        st.stop()
        
    all_technicians = df_tech['NOMBRES'].unique().tolist()
    
    # Init vehicle config in session if empty
    if not st.session_state.vehicles_config:
        # Default: 1 Car to start
        st.session_state.vehicles_config = [{'type': 'Auto', 'members': [], 'id': 0}]

    col_builder, col_preview = st.columns([2, 1])

    with col_builder:
        st.subheader("Armado de Cuadrillas")
        
        # --- CREW MANAGEMENT ---
        for i, veh in enumerate(st.session_state.vehicles_config):
            v_type = veh['type']
            v_icon = "🚙" if v_type == "Auto" else "🚶"
            
            with st.expander(f"{v_icon} {v_type} #{i+1}", expanded=True):
                # Member Selection
                current_members = veh['members']
                
                # Filter out technicians already assigned to OTHER vehicles
                assigned_others = [m for v in st.session_state.vehicles_config if v != veh for m in v['members']]
                available_techs = [t for t in all_technicians if t not in assigned_others]
                # Add back current members to be selectable
                options = sorted(available_techs + current_members)
                
                selected_members = st.multiselect(
                    f"Integrantes ({v_type})", 
                    options=options,
                    default=current_members,
                    key=f"veh_{i}_members"
                )
                
                # Update session state with selection
                # We need to do this carefully to avoid infinite reruns or state sync issues
                # But treating it as a standard input works if we update the list ref
                veh['members'] = selected_members
                
                # Validation & Skills Display
                if v_type == 'Auto':
                    if len(selected_members) < 2:
                        st.error("⚠️ Mínimo 2 personas requeridas en Auto.")
                    elif len(selected_members) > 4:
                        st.error("⚠️ Máximo 4 personas permitidas en Auto.")
                
                # Show Capabilities
                if selected_members:
                    skills = set()
                    for m in selected_members:
                        tech_data = df_tech[df_tech['NOMBRES'] == m].iloc[0]
                        if tech_data['Familia 1']: skills.add(tech_data['Familia 1'])
                        if tech_data['Familia 2']: skills.add(tech_data['Familia 2'])
                    
                    st.info(f"🛠️ Especialidades: {', '.join(sorted(skills))}")
                    veh['skills'] = list(skills)
                else:
                    veh['skills'] = []
                
                # Remove Button
                if st.button("🗑️ Eliminar Vehículo", key=f"del_{i}"):
                    st.session_state.vehicles_config.pop(i)
                    st.rerun()

        # Add Buttons
        c_add1, c_add2 = st.columns(2)
        if c_add1.button("➕ Agregar Auto"):
            st.session_state.vehicles_config.append({'type': 'Auto', 'members': [], 'id': len(st.session_state.vehicles_config)})
            st.rerun()
        if c_add2.button("➕ Agregar Caminante"):
            st.session_state.vehicles_config.append({'type': 'Walker', 'members': [], 'id': len(st.session_state.vehicles_config)})
            st.rerun()

    with col_preview:
        st.subheader("Resumen de Flota")
        n_cars = len([v for v in st.session_state.vehicles_config if v['type'] == 'Auto'])
        n_walkers = len([v for v in st.session_state.vehicles_config if v['type'] == 'Walker'])
        st.metric("Autos", n_cars)
        st.metric("Caminantes", n_walkers)
        
        # Configuration Parameters
        st.markdown("---")
        st.markdown("**Parámetros Globales**")
        max_capacity = st.number_input("Capacidad Max (Tickets)", min_value=1, value=100)
        service_time = st.number_input("Tiempo Servicio (min)", min_value=1, value=15)
        max_work_hours = st.number_input("Jornada Maxima (horas)", min_value=4, value=12)
        
        trafico = st.selectbox("Tráfico", ["Normal", "Pesado", "Ligero"])
        tf_factor = 1.0
        if trafico == "Pesado": tf_factor = 1.6
        if trafico == "Ligero": tf_factor = 0.8

        st.markdown("---")
        if st.button("🚀 Calcular Rutas", type="primary"):
            # Validation
            valid_config = True
            for v in st.session_state.vehicles_config:
                if v['type'] == 'Auto':
                    if len(v['members']) < 2 or len(v['members']) > 4:
                        valid_config = False
                        st.toast(f"Error en Auto: Deben ser 2-4 personas.", icon="❌")
            
            st.toast("Iniciando validación...", icon="🕵️")
            print("DEBUG: Validation started")
            
            if not valid_config:
                st.error("Por favor corrija la configuración de las cuadrillas.")
            elif not st.session_state.vehicles_config:
                st.error("Debe agregar al menos 1 vehículo.")
            else:
                # Prepare Data
                df_raw = pd.DataFrame(st.session_state.daily_tickets)
                
                # Aggregation
                df_grouped = df_raw.groupby(['Nombre', 'Latitud (y)', 'Longitud (x)', 'Habla a']).agg({
                    'Importe de la entrega': 'sum',
                    'Ticket': lambda x: ', '.join(x.astype(str)),
                    'Familia': lambda x: sorted(list(set(x)))[0] # Take first family or dominate one? For now assume 1 family per ticket or node
                }).reset_index()

                # --- INSERT DEPOT ---
                depot_row = pd.DataFrame([{
                    'Nombre': 'La Rambla San Borja (Punto de Partida)', 
                    'Latitud (y)': -12.0884681, 
                    'Longitud (x)': -77.0061123, 
                    'Habla a': 'SODEXO',
                    'Importe de la entrega': 0,
                    'Ticket': 'Inicio',
                    'Familia': 'Base'
                }])
                
                df_final_for_solver = pd.concat([depot_row, df_grouped], ignore_index=True)
                
                
                st.toast("Calculando rutas... (Puede tardar unos segundos)", icon="⏳")
                print("DEBUG: Calling solve_vrp_data...")
                
                # CALL SOLVER WITH NEW SIGNATURE
                try:
                    solution, routing, manager, data, df_cleaned = solve_vrp_data(
                        df_final_for_solver, 
                        st.session_state.vehicles_config, # PASS LIST INSTEAD OF COUNTS
                        max_capacity,
                        traffic_factor=tf_factor,
                        service_time_per_ticket_mins=service_time,
                        max_work_hours=max_work_hours
                    )
                    print(f"DEBUG: Solver returned. Solution object: {solution}")
                except Exception as e:
                    print(f"CRITICAL ERROR IN SOLVER: {e}")
                    st.error(f"Error crítico en el optimizador: {e}")
                    solution = None
                
                if solution:
                    st.success("✅ Solución Encontrada!")
                    st.session_state.optimization_result = (solution, routing, manager, data, df_cleaned)
                    st.session_state.stage = 'results'
                    st.rerun()
                else:
                    st.error("❌ No se encontró solución. Verifique las especialidades de las cuadrillas vs los tickets.")

# --- STAGE 3: RESULTADOS ---
elif st.session_state.stage == 'results':
    st.header("3️⃣ Rutas Optimizadas")
    
    # Handle Unpacking safely
    res_tuple = st.session_state.optimization_result
    if len(res_tuple) >= 5:
        # Standard signature for new mixed code
        solution, routing, manager, data, df_loc = res_tuple[:5]
    else:
        # Fallback
        st.error("Estado incompatible. Por favor reinicie.")
        st.stop()
        
    results, route_maps_data, total_duration, total_load = format_solution(data, manager, routing, solution, df_loc)
    
    # 1. Metrics
    c1, c2, c3 = st.columns(3)
    
    # Filter valid routes for metrics
    served_results = [r for r in results if r['VehicleID'] != -1]
    dropped_results = [r for r in results if r['VehicleID'] == -1]
    
    num_active_routes = len([r for r in route_maps_data if r['load'] > 0])
    
    # Count Cars vs Walkers used
    used_cars = len([r for r in route_maps_data if r['load'] > 0 and r.get('vehicle_type') == 'Auto'])
    used_walkers = len([r for r in route_maps_data if r['load'] > 0 and r.get('vehicle_type') == 'Walker'])
    
    # Calculate Average Duration instead of Total
    if num_active_routes > 0:
        avg_duration = total_duration / num_active_routes
        hours = int(avg_duration / 3600)
        mins = int((avg_duration % 3600) / 60)
    else:
        hours, mins = 0, 0
    
    c1.metric("Tiempo Promedio", f"{hours}h {mins}m")
    c2.metric("Tickets Atendidos", total_load)
    c3.metric("Recursos", f"{used_cars} Autos / {used_walkers} a Pie")
    
    # WARNING FOR DROPPED NODES
    if dropped_results:
        # Analyze WHY they were dropped
        missing_skills_msg = []
        
        # Get active vehicle skills
        active_skills = set()
        for v in st.session_state.vehicles_config:
            active_skills.update(v.get('skills', []))
            
        for r in dropped_results:
            # Re-derive required skill from original DF (inefficient but safe)
            # We need to find the node index in df_loc to get 'Familia'
            node_idx = r['NodeID']
            raw_family = str(df_loc.iloc[node_idx]['Familia']).strip().upper()
            
            req_skill = 'Multiusos'
            if 'GASFITERIA' in raw_family: req_skill = 'Gasfitero'
            elif 'ELECTRICIDAD' in raw_family: req_skill = 'Electricista'
            elif 'PINTURA' in raw_family: req_skill = 'Pintor'
            elif 'CERRAJERIA' in raw_family: req_skill = 'Cerrajero'
            
            if req_skill not in active_skills:
                msg = f"Ticket '{r['LocationName']}' no asignado. Requiere: {req_skill} (No disponible en flota)"
                if msg not in missing_skills_msg:
                    missing_skills_msg.append(msg)
            else:
                # If skill exists but still dropped, it's likely capacity or time (though time is infinite now)
                # Or maybe the vehicle with the skill is too full?
                pass

        st.warning(f"⚠️ **{len(dropped_results)} Tickets No Asignados**")
        
        if missing_skills_msg:
            st.error("⛔ **Razón Principal:** Los técnicos asignados a la ruta no tienen la especialidad requerida.")
            for m in missing_skills_msg:
                st.write(f"- {m}")
        else:
             st.info("Posible causa: Capacidad de vehículos excedida.")

        with st.expander("Ver Detalle de Tickets No Atendidos"):
            df_dropped = pd.DataFrame(dropped_results)
            st.dataframe(df_dropped[['LocationName', 'Client', 'Status']])
    
    # 2. Map
    st.subheader("🗺️ Visualización de Rutas")
    m_result = generate_folium_map(df_loc, route_maps_data)
    st_folium(m_result, height=500, width="100%")
    
    # 3. Tables & Export
    st.subheader("📝 Detalle de Visitas (Consolidado)")
    df_res = pd.DataFrame(results)
    
    # Sort by Vehicle and Order
    df_res = df_res.sort_values(by=['VehicleID', 'OrderInRoute'])
    st.dataframe(style_dataframe(df_res[['VehicleID', 'VehicleType', 'OrderInRoute', 'LocationName', 'Client', 'AccumulatedDuration_Mins', 'Status']]), use_container_width=True)
    
    # --- INDIVIDUAL ITINERARIES ---
    st.markdown("---")
    st.subheader("🚚 Itinerarios por Recurso")
    
    unique_vehicles = sorted([v for v in df_res['VehicleID'].unique() if v != -1])
    
    # We need to link VehicleID back to our vehicle_config to show members
    # data['vehicle_types'] gave us types, but we might want the original config object
    # If the order is preserved (0 to N), we can map directly by index.
    
    # cols = st.columns(len(unique_vehicles)) if len(unique_vehicles) < 4 else [st.container() for _ in range(len(unique_vehicles))]
    
    for i, vid in enumerate(unique_vehicles):
        vehicle_route = df_res[df_res['VehicleID'] == vid]
        
        # Get Vehicle Config Info
        v_config = st.session_state.vehicles_config[vid] if vid < len(st.session_state.vehicles_config) else {}
        v_members = v_config.get('members', [])
        v_skills = v_config.get('skills', [])
        
        v_type = vehicle_route.iloc[0]['VehicleType'] if 'VehicleType' in vehicle_route.columns else "Auto"
        icon = "🚙" if v_type == "Auto" else "🚶"
        
        title_str = f"{icon} {v_type} #{vid + 1} ({len(vehicle_route)} paradas)"
            
        with st.expander(title_str, expanded=True):
            st.markdown(f"**Integrantes:** {', '.join(v_members)}")
            if v_skills:
                st.markdown(f"**Especialidades:** {', '.join(v_skills)}")
            
            st.write(f"**Tiempo Estimado:** {vehicle_route['AccumulatedDuration_Mins'].max()} min")
            
            # VISUAL FLOW SEQUENCE & MAPS LINK
            path_steps = []
            
            # --- GOOGLE MAPS LINK GENERATION (Official API) ---
            # Origin & Dest: La Rambla San Borja
            depot_coords = "-12.0884681,-77.0061123"
            
            # Collect Waypoints
            waypoints_list = []
            
            for _, row in vehicle_route.iterrows():
                loc_str = row['LocationName']
                if row['Client']:
                    loc_str += f" ({row['Client']})"
                path_steps.append(loc_str)
                
                if 'Latitude' in row and 'Longitude' in row and pd.notnull(row['Latitude']) and pd.notnull(row['Longitude']):
                    waypoints_list.append(f"{row['Latitude']},{row['Longitude']}")
            
            flow_str = " ➝ ".join(path_steps)
            st.info(f"**Secuencia de Visita:**\n\n🏁 Inicio ➝ {flow_str} ➝ 🏁 Fin")
            
            if waypoints_list:
                # Construct URL
                # Parameters
                origin = depot_coords
                destination = depot_coords # Loop back to start
                waypoints = "|".join(waypoints_list)
                
                tm = "walking" if v_type == "Walker" else "driving"
                
                maps_url = (
                    f"https://www.google.com/maps/dir/?api=1"
                    f"&origin={origin}"
                    f"&destination={destination}"
                    f"&waypoints={waypoints}"
                    f"&travelmode={tm}"
                )
                
                st.link_button("🗺️ Abrir Ruta en Google Maps", maps_url)
            
            # Create a nice timeline list
            for _, stop in vehicle_route.iterrows():
                client_str = f" ({stop['Client']})" if stop['Client'] else ""
                st.markdown(f"**{stop['OrderInRoute']}. {stop['LocationName']}{client_str}**")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_res.to_excel(writer, index=False, sheet_name='Hoja de Ruta')
    
    c_down, c_reset = st.columns([1, 4])
    c_down.download_button("📥 Descargar Excel", buffer.getvalue(), "rutas_sodexo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    if c_reset.button("🔄 Nueva Planificación"):
        reset_app()
        st.rerun()



