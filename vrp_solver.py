
import pandas as pd
import numpy as np
import math
import folium
import os
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp
from datetime import datetime

# --- CONFIGURATION ---
INPUT_FILE = "VRP_Spreadsheet_Solver_v3.8 14.05.xlsm"
OUTPUT_MAP = "mapa_rutas_peru.html"
OUTPUT_EXCEL = "solucion_rutas.xlsx"

def haversine(lat1, lon1, lat2, lon2):
    """Calculates the great circle distance in km between two points."""
    R = 6371  # Earth radius in km
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    
    a = math.sin(dphi / 2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def create_distance_matrix(locations):
    """
    Creates a distance matrix (meters) from locations.
    """
    print(f"Calculating distance matrix...")
    n = len(locations)
    coords = locations[['Latitud (y)', 'Longitud (x)']].values
    
    full_matrix = []
    for from_node in range(n):
        row = []
        for to_node in range(n):
            if from_node == to_node:
                row.append(0)
            else:
                dist_km = haversine(coords[from_node][0], coords[from_node][1], 
                                  coords[to_node][0], coords[to_node][1])
                dist_m = int(dist_km * 1000)
                row.append(dist_m)
        full_matrix.append(row)
        
    return full_matrix

def solve_vrp_data(df_loc, vehicles_config, vehicle_capacity, start_node_index=0, max_seconds=30, traffic_factor=1.0, 
                   service_time_per_ticket_mins=15, max_work_hours=12):
    """
    Solves VRP for Mixed Fleet (Cars + Walkers) with Skill Matching.
    vehicles_config: List of dicts [{'type': 'Auto', 'members': [], 'skills': [], 'id': 0}, ...]
    """
    num_locations = len(df_loc)
    num_vehicles = len(vehicles_config)
    print(f"Solving for {num_locations} locations. Vehicles: {num_vehicles}")

    # CLEANING: Ensure coordinates are numeric
    for col in ['Latitud (y)', 'Longitud (x)']:
        if col in df_loc.columns:
            df_loc[col] = pd.to_numeric(df_loc[col], errors='coerce')
            
    # Remove NaN coordinates just in case
    df_loc = df_loc.dropna(subset=['Latitud (y)', 'Longitud (x)'])
    # Re-index to ensure continuity
    df_loc = df_loc.reset_index(drop=True)
    num_locations = len(df_loc) 

    # Build Vehicle Config Arrays
    vehicle_capacities = [vehicle_capacity] * num_vehicles # Or customize per vehicle if needed
    starts = [start_node_index] * num_vehicles
    ends = [start_node_index] * num_vehicles
    
    vehicle_types = [v['type'] for v in vehicles_config]
    vehicle_skills = [set(v.get('skills', [])) for v in vehicles_config]
    
    # SETUP DATA MODEL
    data = {}
    
    # 1. DISTANCE MATRIX (Meters)
    data['distance_matrix'] = create_distance_matrix(df_loc)
    
    # Demands
    if 'Importe de la entrega' in df_loc.columns:
        demands = df_loc['Importe de la entrega'].fillna(0).astype(int).tolist()
    elif 'Tickets' in df_loc.columns:
        demands = df_loc['Tickets'].fillna(0).astype(int).tolist()
    else:
        demands = [0] * num_locations 
        
    data['demands'] = demands
    data['vehicle_capacities'] = vehicle_capacities
    data['num_vehicles'] = num_vehicles
    data['starts'] = starts
    data['ends'] = ends
    data['vehicle_types'] = vehicle_types # Store for later formatting
    
    # OR-TOOLS SETUP
    manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                           data['num_vehicles'],
                                           data['starts'],
                                           data['ends'])
    
    routing = pywrapcp.RoutingModel(manager)
    
    # --- SKILL MATCHING CONSTRAINTS ---
    # For each node (except depot), check if it requires a skill (Familia)
    # If so, only allow vehicles that have that skill.
    
    if 'Familia' in df_loc.columns:
        for node_idx in range(1, num_locations): # Skip Depot (0)
            # NORMALIZE SKILL (User Request)
            raw_skill = str(df_loc.iloc[node_idx]['Familia']).strip().upper()
            
            # Map Family -> Technician Spec
            if 'GASFITERIA' in raw_skill:
                required_skill = 'Gasfitero'
            elif 'ELECTRICIDAD' in raw_skill:
                required_skill = 'Electricista'
            elif 'PINTURA' in raw_skill:
                required_skill = 'Pintor'
            elif 'CERRAJERIA' in raw_skill:
                required_skill = 'Cerrajero'
            else:
                # Fallback for everything else (Aire Acondicionado, Civil, etc.)
                required_skill = 'Multiusos'
            
            # If ticket has a specific family (not General/Base)
            # multiosos is also a skill now, so we always check unless it's the depot/base logic which is handled by loop range
            
            allowed_vehicles = []
            for v_idx, v_skills in enumerate(vehicle_skills):
                # Check if vehicle has the required skill
                # Logic: One of the crew members must have the skill.
                # v_skills is a set/list of strings like ['Electricista', 'Multiusos']
                
                # Check if required skill is present
                if required_skill in v_skills:
                    allowed_vehicles.append(v_idx)
                
                # SPECIAL CASE: If required is 'Multiusos', and vehicle has 'Multiusos', acceptable.
                # But what if a 'Gasfitero' can also do 'Multiusos' work? 
                # User said: "Cualquier otra especialidad sera considerada como Multiusos" -> This implies the TICKET is Multiusos.
                # Does a specialized 'Gasfitero' count as 'Multiusos'? 
                # Usually yes, but technically we look for the string 'Multiusos' in their skill list.
                # If the technician data only says "Gasfitero", they might NOT be "Multiusos".
                # Safest bet: Strict matching against the skill list. 
                # If a Gasfitero can do Multiusos, they should have "Multiusos" in their skill list in the Excel.
            
            allowed_vehicles = [int(v) for v in allowed_vehicles] # Ensure standard python ints
            
            if not allowed_vehicles:
                print(f"WARNING: Node {node_idx} ('{df_loc.iloc[node_idx]['Nombre']}') requires '{required_skill}' (Raw: {raw_skill}) but no vehicle has it.")
                # Force empty allowed -> Unassigned. SetValues([]) might fail or make it impossible
                # For unassigned, we usually remove the node or use disjunction.
                # But if we want to force failure if not possible:
                # try:
                #     routing.VehicleVar(manager.NodeToIndex(node_idx)).SetValues([-1]) # Force unperformed?
                # except:
                #     pass
                print(f"DEBUG: Constraint active for Node {node_idx}. No vehicles allowed -> Dropped.")
                # We let it be empty, which likely forces a drop via Disjunctions (if enabled) or fails.
                # With Disjunctions enabled, this node will be dropped.
                try:
                    routing.VehicleVar(manager.NodeToIndex(node_idx)).SetValues([-1]) # Force drop
                except:
                    pass
            else:
                # WORKAROUND: SetAllowedVehiclesForIndex fails with TypeError on some versions
                try:
                    routing.VehicleVar(manager.NodeToIndex(node_idx)).SetValues(allowed_vehicles)
                    print(f"DEBUG: Constraint active for Node {node_idx}. Req: {required_skill}. Allowed V: {allowed_vehicles}")
                except Exception as e:
                    print(f"Error setting allowed vehicles for node {node_idx}: {e}")
    
    # --- DEFINE SPEEDS ---
    # Car Speed (m/s)
    speed_car_kmh = 30 / traffic_factor
    speed_car_ms = speed_car_kmh * (1000 / 3600)
    
    # Walker Speed (m/s)
    speed_walker_kmh = 5.0
    speed_walker_ms = speed_walker_kmh * (1000 / 3600)
    
    service_time_seconds = service_time_per_ticket_mins * 60

    # --- CALLBACKS ---
    def car_time_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        # Service time
        service = data['demands'][from_node] * service_time_seconds
        # Travel time
        dist = data['distance_matrix'][from_node][to_node]
        travel = int(dist / speed_car_ms)
        return travel + service

    def walker_time_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        # Service time
        service = data['demands'][from_node] * service_time_seconds
        # Travel time
        dist = data['distance_matrix'][from_node][to_node]
        travel = int(dist / speed_walker_ms)
        return travel + service

    car_evaluator_index = routing.RegisterTransitCallback(car_time_callback)
    walker_evaluator_index = routing.RegisterTransitCallback(walker_time_callback)
    
    # Assign Callbacks to Vehicles
    for i in range(num_vehicles):
        if vehicle_types[i] == 'Auto':
            routing.SetArcCostEvaluatorOfVehicle(car_evaluator_index, i)
        else:
            routing.SetArcCostEvaluatorOfVehicle(walker_evaluator_index, i)

    # --- DIMENSIONS ---
    
    # Capacity
    def demand_callback(from_index):
        from_node = manager.IndexToNode(from_index)
        return data['demands'][from_node]

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(
        demand_callback_index,
        0,  
        data['vehicle_capacities'],
        True,
        'Capacity')
    
    # Time Dimension
    # We need a generic callback for the dimension if we want to retrieve values, 
    # BUT AddDimensionWithVehicleTransits allows per-vehicle transit!
    
    # Create vector of evaluator indices for all vehicles
    transit_evaluator_indices = []
    for i in range(num_vehicles):
        if vehicle_types[i] == 'Auto':
            transit_evaluator_indices.append(car_evaluator_index)
        else:
            transit_evaluator_indices.append(walker_evaluator_index)

    # Relaxed horizon to ensure solution is always found (user request)
    # We use a very large number (e.g., 30 days) as the hard limit/horizon.
    # The 'max_work_hours' parameter can still be used for soft limits or cost coefficients if needed, 
    # but for "always calculate", we remove the hard cap.
    horizon_seconds = 30 * 24 * 3600 # 30 days
    
    # Use AddDimensionWithVehicleTransits to correctly track time per vehicle type
    routing.AddDimensionWithVehicleTransits(
        transit_evaluator_indices,
        horizon_seconds, # Slack
        horizon_seconds, # Horizon - NOW INFINITE (practically)
        False, # Start at zero? No, accumulate
        'Time')
    
    # Global Span Cost
    time_dimension = routing.GetDimensionOrDie('Time')
    time_dimension.SetGlobalSpanCostCoefficient(100)
    
    # Optional: Add Soft Upper Bound based on the user's "max_work_hours" preference?
    # This would prioritize keeping routes within the limit but allow exceeding it if necessary.
    # user said "remove restriction", so let's just penalty it.
    soft_limit_seconds = max_work_hours * 3600
    for i in range(num_vehicles):
        time_dimension.SetCumulVarSoftUpperBound(routing.End(i), soft_limit_seconds, 1000)

    # --- DISJUNCTIONS (Allow dropping nodes) ---
    # Ensure a solution is always found by allowing nodes to be skipped (with a high penalty)
    penalty = 1000000
    for node_idx in range(1, num_locations):
        routing.AddDisjunction([manager.NodeToIndex(node_idx)], penalty)

    # SOLVE
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
    search_parameters.local_search_metaheuristic = (
        routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH)
    search_parameters.time_limit.seconds = max_seconds
    
    solution = routing.SolveWithParameters(search_parameters)
    return solution, routing, manager, data, df_loc

def format_solution(data, manager, routing, solution, df_loc):
    """Formats solution into standard structure for display/export"""
    total_duration = 0
    total_load = 0
    
    results = [] 
    route_maps_data = [] # For plotting
    
    # Track visited nodes to identify dropped ones
    visited_nodes = set()

    for vehicle_id in range(data['num_vehicles']):
        index = routing.Start(vehicle_id)
        route_duration = 0
        route_load = 0
        route_nodes = []
        route_coords = []
        
        while not routing.IsEnd(index):
            node_index = manager.IndexToNode(index)
            # Mark as visited (except start if it acts as depot loop)
            if node_index != 0:
                visited_nodes.add(node_index)

            route_load += data['demands'][node_index]
            
            previous_index = index
            index = solution.Value(routing.NextVar(index))
            
            # Cost is roughly seconds now
            duration = routing.GetArcCostForVehicle(previous_index, index, vehicle_id)
            route_duration += duration
            
            # Data collection
            if previous_index != routing.Start(vehicle_id):
                loc_name = df_loc.iloc[node_index]['Nombre'] if 'Nombre' in df_loc.columns else f"Loc {node_index}"
                client_name = df_loc.iloc[node_index]['Habla a'] if 'Habla a' in df_loc.columns else ""
                lat = df_loc.iloc[node_index]['Latitud (y)']
                lon = df_loc.iloc[node_index]['Longitud (x)']
                
                # Identify Vehicle Type
                v_type = data['vehicle_types'][vehicle_id]
                
                results.append({
                    'VehicleID': vehicle_id,
                    'VehicleType': v_type,
                    'NodeID': node_index,
                    'LocationName': loc_name,
                    'Client': client_name,
                    'Latitude': lat,
                    'Longitude': lon,
                    'Load': route_load,
                    'OrderInRoute': len(route_nodes), 
                    'AccumulatedDuration_Mins': int(route_duration / 60),
                    'Status': 'Serviced'
                })

            route_nodes.append(node_index)
            route_coords.append((df_loc.iloc[node_index]['Latitud (y)'], df_loc.iloc[node_index]['Longitud (x)']))
            
        # Return to depot (visual)
        node_index = manager.IndexToNode(index) 
        lat = df_loc.iloc[node_index]['Latitud (y)']
        lon = df_loc.iloc[node_index]['Longitud (x)']
        route_coords.append((lat, lon))
        
        total_duration += route_duration
        total_load += route_load
        
        route_maps_data.append({
            'vehicle_id': vehicle_id,
            'vehicle_type': data['vehicle_types'][vehicle_id],
            'coords': route_coords,
            'load': route_load,
            'duration_s': route_duration
        })
        
    # --- CHECK DROPPED NODES ---
    num_locations = len(df_loc)
    for node_idx in range(1, num_locations):
        if node_idx not in visited_nodes:
            loc_name = df_loc.iloc[node_idx]['Nombre'] if 'Nombre' in df_loc.columns else f"Loc {node_idx}"
            client_name = df_loc.iloc[node_idx]['Habla a'] if 'Habla a' in df_loc.columns else ""
            lat = df_loc.iloc[node_idx]['Latitud (y)']
            lon = df_loc.iloc[node_index]['Longitud (x)']
            
            results.append({
                'VehicleID': -1,
                'VehicleType': 'None',
                'NodeID': node_idx,
                'LocationName': loc_name,
                'Client': client_name,
                'Latitude': lat,
                'Longitude': lon,
                'Load': 0,
                'OrderInRoute': -1,
                'AccumulatedDuration_Mins': 0,
                'Status': 'Dropped'
            })
            
            # Add a dummy "Dropped" route for visual feedback if needed, 
            # currently we just list them in results.
        
    return results, route_maps_data, total_duration, total_load

def generate_folium_map(df_loc, route_maps_data):
    """Generates Folium map object"""
    center_lat = df_loc.iloc[0]['Latitud (y)']
    center_lon = df_loc.iloc[0]['Longitud (x)']
    m = folium.Map(location=[center_lat, center_lon], zoom_start=6)
    
    colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'lightred', 'beige', 
              'darkblue', 'darkgreen', 'cadetblue', 'darkpurple', 'pink', 'gray', 'black']
    
    # Plot routes
    for r_data in route_maps_data:
        vid = r_data['vehicle_id']
        coords = r_data['coords']
        color = colors[vid % len(colors)]
        
        folium.PolyLine(coords, color=color, weight=3, opacity=0.8).add_to(m)
        
        # Plot markers (except last one which is return to depot, redundant for markers)
        # Actually loop through coordinates to place markers
        for i, (lat, lon) in enumerate(coords[:-1]): 
            # We don't have the node name easily here without re-querying df_loc or passing it down
            # Simplification: Just put a dot
            folium.CircleMarker(
                location=[lat, lon],
                radius=3,
                color=color,
                fill=True,
                fill_color=color,
                popup=f"V{vid}"
            ).add_to(m)

    return m

def solve_vrp_file():
    """Legacy function to run from file as before"""
    print(f"Reading file: {INPUT_FILE}")
    try:
        df_loc = pd.read_excel(INPUT_FILE, sheet_name='1 ubicaciones')
        df_loc.columns = df_loc.columns.str.strip()
        
        # --- LIMPIEZA DE DATOS ---
        print(f"Filas originales: {len(df_loc)}")
        
        # 1. Convertir a numérico forzando errores a NaN
        for col in ['Latitud (y)', 'Longitud (x)']:
            if col in df_loc.columns:
                df_loc[col] = pd.to_numeric(df_loc[col], errors='coerce')
        
        # 2. Eliminar filas con NaN en coordenadas
        valid_rows = df_loc.dropna(subset=['Latitud (y)', 'Longitud (x)'])
        dropped = len(df_loc) - len(valid_rows)
        if dropped > 0:
            print(f"Advertencia: Se eliminaron {dropped} filas con coordenadas inválidas (texto o vacío).")
        df_loc = valid_rows.copy()
        
        # 3. Filtrar coordenadas fuera de rango (Perú aprox: Lat -20 a 0, Lon -82 a -68)
        # Esto ayuda a filtrar errores como -120 o -700
        mask_peru = (
            (df_loc['Latitud (y)'] > -20) & (df_loc['Latitud (y)'] < 0) &
            (df_loc['Longitud (x)'] > -85) & (df_loc['Longitud (x)'] < -65)
        )
        out_of_bounds = df_loc[~mask_peru]
        if not out_of_bounds.empty:
            print(f"Advertencia: Se eliminaron {len(out_of_bounds)} filas con coordenadas fuera de Perú:")
            # print(out_of_bounds[['Nombre', 'Latitud (y)', 'Longitud (x)']].head())
            df_loc = df_loc[mask_peru].copy()

        print(f"Filas válidas para optimizar: {len(df_loc)}")
        # -------------------------
        
        # --- NUEVO REQUERIMIENTO: FIJAR DEPOT EN LA RAMBLA SAN BORJA ---
        # Coordenadas: -12.0884681,-77.0061123
        depot_row = pd.DataFrame([{
            'Nombre': 'La Rambla San Borja',
            'Habla a': 'SODEXO', 
            'Latitud (y)': -12.0884681,
            'Longitud (x)': -77.0061123,
            'Importe de la entrega': 0,
            'Tickets': 0
        }])
        
        # Concatenar al inicio
        df_loc = pd.concat([depot_row, df_loc], ignore_index=True)
        print("Depot fijado en: La Rambla San Borja (-12.0884681, -77.0061123)")
        # -------------------------------------------------------------
        
        # vehicles config from file... simplified usage for now or reimplement full read
        # For backward compatibility, I'll attempt to minimally reconstruct the vehicle logic
        # OR just use the new generic solver with defaults if that suffices?
        # The user's original request implies moving AWAY from the hardcoded file, but keeping the CLI working is nice.
        
        # Let's read the vehicles sheet to get count/cap
        df_veh = pd.read_excel(INPUT_FILE, sheet_name='3.Vehículos')
        df_veh.columns = df_veh.columns.str.strip()
        
        # Simple extraction of total vehicles and max capacity for the generic solver
        total_vehicles = 0
        max_cap = 0
        for _, row in df_veh.iterrows():
            count = int(row.get('Numero de vehiculos', 1))
            cap = int(row.get('Capacidad', 100))
            total_vehicles += count
            max_cap = max(max_cap, cap)
            
        if total_vehicles == 0: total_vehicles = 5
        if max_cap == 0: max_cap = 100
        
        solution, routing, manager, data, df_loc_solved = solve_vrp_data(df_loc, total_vehicles, max_cap)
        
        if solution:
            results, route_data, dist, load = format_solution(data, manager, routing, solution, df_loc_solved)
            
            print(f"Total Distance: {dist}m")
            print(f"Total Load: {load}")
            
            # Save excel
            pd.DataFrame(results).to_excel(OUTPUT_EXCEL, index=False)
            print(f"Saved {OUTPUT_EXCEL}")
            
            # Save map
            m = generate_folium_map(df_loc, route_data)
            m.save(OUTPUT_MAP)
            print(f"Saved {OUTPUT_MAP}")
        else:
            print("No solution found.")
            
    except Exception as e:
        print(f"Error executing legacy mode: {e}")

if __name__ == '__main__':
    solve_vrp_file()
