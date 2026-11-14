
from pathlib import Path
import math
import pandas as pd

# ==========================================================
# 1. CONFIGURACION
# ==========================================================

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "optimización.xlsx"
SALIDA = CARPETA / "OPTIMIZACION_SEMANA.xlsx"

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

MAX_CAMIONES_DIA = 12
JORNADA_MAX_MIN = 14 * 60

CARGA_MIN = 60
DESCARGA_MIN = 60

VELOCIDAD = 80
FACTOR_CARRETERA = 1.2

# Capacidades provisionales
CAPACIDADES = {
    "Destete": {3: 1800, 4: 2400},
    "Salida transición": {3: 900, 4: 1200},
    "Recría": {3: 900, 4: 1200},
}

CAMIONES = (
    [{"Camion": f"C{i:02d}", "Pisos": 4} for i in range(1, 5)]
    + [{"Camion": f"C{i:02d}", "Pisos": 3} for i in range(5, 16)]
)

CONDUCTORES = [f"Conductor_{i:02d}" for i in range(1, 14)]

MAPA_BIO = {
    "MULTIPLICADOR": "00 Multiplicación",
    "CUARENTENA": "00 Multiplicación",
    "RECRIA": "00 Recrías",
    "SITIO_1": "0 Negativa",
    "SITIO_2": "1 Estable",
    "SITIO_3_CEBO": "2/3 Positiva",
}

# ==========================================================
# 2. MATRIZ DE NOCHES COMPLETAS
# ==========================================================

NOCHES = {
    ("00 Multiplicación", "00 Multiplicación"): 0,
    ("00 Multiplicación", "00 Recrías"): 0,
    ("00 Multiplicación", "0 Negativa"): 0,
    ("00 Multiplicación", "1 Estable"): 0,
    ("00 Multiplicación", "2/3 Positiva"): 0,

    ("0 Negativa", "00 Multiplicación"): 1,
    ("0 Negativa", "00 Recrías"): 0,
    ("0 Negativa", "0 Negativa"): 0,
    ("0 Negativa", "1 Estable"): 0,
    ("0 Negativa", "2/3 Positiva"): 0,

    ("1 Estable", "00 Multiplicación"): 2,
    ("1 Estable", "00 Recrías"): 1,
    ("1 Estable", "0 Negativa"): 1,
    ("1 Estable", "1 Estable"): 0,
    ("1 Estable", "2/3 Positiva"): 0,

    ("2/3 Positiva", "00 Multiplicación"): 3,
    ("2/3 Positiva", "00 Recrías"): 1,
    ("2/3 Positiva", "0 Negativa"): 2,
    ("2/3 Positiva", "1 Estable"): 1,
    ("2/3 Positiva", "2/3 Positiva"): 0,
}

# ==========================================================
# 3. FUNCIONES
# ==========================================================

def numero(x):
    return pd.to_numeric(x, errors="coerce")


def categoria(tipo):
    if pd.isna(tipo):
        return None
    return MAPA_BIO.get(str(tipo).strip().upper())


def haversine(lat1, lon1, lat2, lon2):
    valores = [lat1, lon1, lat2, lon2]

    if any(pd.isna(x) for x in valores):
        return 0.0

    r = 6371

    lat1, lon1, lat2, lon2 = map(
        math.radians, valores
    )

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2)
        * math.sin(dlon / 2) ** 2
    )

    return 2 * r * math.asin(math.sqrt(a))


def noches_disponibles(dia_anterior, dia_actual):
    if dia_anterior is None:
        return None

    return max(0, dia_actual - dia_anterior - 1)


def bioseguridad_permitida(ultima_categoria, nueva_categoria,
                           dia_anterior, dia_actual):

    if ultima_categoria is None:
        return True

    clave = (ultima_categoria, nueva_categoria)

    # Una transición que no aparece en la matriz
    # no se considera autorizada.
    if clave not in NOCHES:
        return False

    noches_necesarias = NOCHES[clave]

    noches_reales = noches_disponibles(
        dia_anterior, dia_actual
    )

    return noches_reales >= noches_necesarias


def obtener_distancia(granjas, origen, destino):
    if origen not in granjas or destino not in granjas:
        return None

    a = granjas[origen]
    b = granjas[destino]

    distancia = haversine(
        a["Latitud_FICTICIA"],
        a["Longitud_FICTICIA"],
        b["Latitud_FICTICIA"],
        b["Longitud_FICTICIA"]
    )

    return distancia * FACTOR_CARRETERA


# ==========================================================
# 4. LEER EXCEL
# ==========================================================

if not ARCHIVO.exists():
    raise FileNotFoundError(
        f"No se encuentra el archivo {ARCHIVO}"
    )

hojas = pd.read_excel(ARCHIVO, sheet_name=None)

for nombre in [
    "Granjas", "Destetes",
    "Salidas-transición", "Recrias"
]:
    if nombre not in hojas:
        raise ValueError(f"Falta la hoja {nombre}")


granjas_df = hojas["Granjas"].copy()

granjas_df["Granja_ID"] = (
    granjas_df["Granja_ID"].astype(str).str.strip()
)

granjas = granjas_df.set_index("Granja_ID").to_dict("index")

# ==========================================================
# 5. UNIFICAR MOVIMIENTOS
# ==========================================================

movimientos = []

for _, f in hojas["Destetes"].iterrows():
    movimientos.append({
        "Semana": f.get("Semana"),
        "Origen_ID": f.get("Sitio1_ID"),
        "Origen_tipo": "SITIO_1",
        "Destino_ID": f.get("Destino_ID"),
        "Destino_tipo": f.get("Destino_tipo"),
        "Cantidad": f.get("Lechones"),
        "Operacion": "Destete"
    })

for _, f in hojas["Salidas-transición"].iterrows():
    movimientos.append({
        "Semana": f.get("Semana"),
        "Origen_ID": f.get("Origen_ID"),
        "Origen_tipo": f.get("Origen_tipo"),
        "Destino_ID": f.get("Cebo_destino_ID"),
        "Destino_tipo": "SITIO_3_CEBO",
        "Cantidad": f.get("Cantidad"),
        "Operacion": "Salida transición"
    })

for _, f in hojas["Recrias"].iterrows():
    movimientos.append({
        "Semana": f.get("Semana"),
        "Origen_ID": f.get("Origen_ID"),
        "Origen_tipo": f.get("Origen_tipo"),
        "Destino_ID": f.get("Destino_ID"),
        "Destino_tipo": f.get("Destino_tipo"),
        "Cantidad": f.get("Animales_aprox"),
        "Operacion": "Recría"
    })

df = pd.DataFrame(movimientos)

# ==========================================================
# 6. PREPARAR CARGAS
# ==========================================================

tareas = []
problemas = []

for i, f in df.iterrows():

    origen = str(f["Origen_ID"]).strip()
    destino = str(f["Destino_ID"]).strip()

    cat_origen = categoria(f["Origen_tipo"])
    cat_destino = categoria(f["Destino_tipo"])

    cantidad = numero(f["Cantidad"])

    if pd.isna(cantidad) or cantidad <= 0:
        problemas.append({
            "ID_movimiento": i + 1,
            "Origen": origen,
            "Destino": destino,
            "Cantidad": f["Cantidad"],
            "Motivo": "Cantidad no válida"
        })
        continue

    if origen not in granjas or destino not in granjas:
        problemas.append({
            "ID_movimiento": i + 1,
            "Origen": origen,
            "Destino": destino,
            "Cantidad": cantidad,
            "Motivo": "Granja no encontrada"
        })
        continue

    if cat_origen is None or cat_destino is None:
        problemas.append({
            "ID_movimiento": i + 1,
            "Origen": origen,
            "Destino": destino,
            "Cantidad": cantidad,
            "Motivo": "Categoría sanitaria desconocida"
        })
        continue

    operacion = f["Operacion"]
    capacidad_maxima = max(CAPACIDADES[operacion].values())

    restante = int(cantidad)
    numero_carga = 1

    while restante > 0:

        carga = min(restante, capacidad_maxima)
        restante -= carga

        tareas.append({
            "ID_movimiento": i + 1,
            "Semana": f["Semana"],
            "Origen": origen,
            "Destino": destino,
            "Origen_tipo": f["Origen_tipo"],
            "Destino_tipo": f["Destino_tipo"],
            "Categoria_origen": cat_origen,
            "Categoria_destino": cat_destino,
            "Operacion": operacion,
            "Cantidad_original": int(cantidad),
            "Cantidad_carga": carga,
            "Numero_carga": numero_carga,
        })

        numero_carga += 1

# Ordenar primero las cargas más grandes
tareas.sort(
    key=lambda x: x["Cantidad_carga"],
    reverse=True
)

# ==========================================================
# 7. ESTADO DE CADA CAMION
# ==========================================================

estado_camiones = {}

for c in CAMIONES:
    estado_camiones[c["Camion"]] = {
        "Pisos": c["Pisos"],
        "Ultima_categoria": None,
        "Ultimo_dia": None,
        "Ultimo_destino": None,
        "Minutos_dia": 0,
        "Dia_actual": None,
        "Conductor": None,
        "Usado_hoy": False,
    }

asignaciones = []
no_asignados = []

pendientes = tareas.copy()

# ==========================================================
# 8. PLANIFICACION DE LUNES A VIERNES
# ==========================================================

for dia_num, dia_nombre in enumerate(DIAS):

    camiones_activos = set()
    conductores_usados = set()

    # Reiniciar contador diario
    for estado in estado_camiones.values():
        estado["Minutos_dia"] = 0
        estado["Dia_actual"] = dia_num
        estado["Conductor"] = None
        estado["Usado_hoy"] = False

    nuevas_pendientes = []

    for tarea in pendientes:

        origen = tarea["Origen"]
        destino = tarea["Destino"]

        cat_origen = tarea["Categoria_origen"]
        cat_destino = tarea["Categoria_destino"]

        operacion = tarea["Operacion"]
        cantidad = tarea["Cantidad_carga"]

        candidatos = []

        for camion, estado in estado_camiones.items():

            pisos = estado["Pisos"]

            capacidad = CAPACIDADES[operacion][pisos]

            if cantidad > capacidad:
                continue

            if not bioseguridad_permitida(
                estado["Ultima_categoria"],
                cat_origen,
                estado["Ultimo_dia"],
                dia_num
            ):
                continue

            # No superar 12 camiones simultáneos
            if (
                camion not in camiones_activos
                and len(camiones_activos) >= MAX_CAMIONES_DIA
            ):
                continue

            distancia_carga = obtener_distancia(
                granjas, origen, destino
            )

            distancia_vacio = 0

            if estado["Ultimo_destino"] is not None:
                distancia_vacio = obtener_distancia(
                    granjas,
                    estado["Ultimo_destino"],
                    origen
                )

            if distancia_carga is None or distancia_vacio is None:
                continue

            km_totales = distancia_carga + distancia_vacio

            minutos = (
                km_totales / VELOCIDAD * 60
                + CARGA_MIN
                + DESCARGA_MIN
            )

            if (
                estado["Minutos_dia"] + minutos
                > JORNADA_MAX_MIN
            ):
                continue

            # Preferir camiones de 3 pisos cuando la carga
            # cabe en ellos y reservar los de 4 pisos.
            preferencia = 0 if (
                pisos == 3 and cantidad <=
                CAPACIDADES[operacion][3]
            ) else 1

            puntuacion = (
                km_totales
                + estado["Minutos_dia"] * 0.02
                + preferencia * 5
            )

            candidatos.append((
                puntuacion,
                camion,
                minutos,
                km_totales,
                distancia_vacio
            ))

        if not candidatos:
            nuevas_pendientes.append(tarea)
            continue

        candidatos.sort(key=lambda x: x[0])

        _, camion, minutos, km, km_vacio = candidatos[0]

        estado = estado_camiones[camion]

        # Asignar conductor al camión durante ese día
        if estado["Conductor"] is None:

            disponibles = [
                c for c in CONDUCTORES
                if c not in conductores_usados
            ]

            if not disponibles:
                nuevas_pendientes.append(tarea)
                continue

            conductor = disponibles[0]
            estado["Conductor"] = conductor
            conductores_usados.add(conductor)

        conductor = estado["Conductor"]

        estado["Minutos_dia"] += minutos
        estado["Ultima_categoria"] = cat_destino
        estado["Ultimo_dia"] = dia_num
        estado["Ultimo_destino"] = destino

        estado["Usado_hoy"] = True
        camiones_activos.add(camion)

        asignaciones.append({
            "ID_movimiento": tarea["ID_movimiento"],
            "Carga": tarea["Numero_carga"],
            "Semana": tarea["Semana"],
            "Dia": dia_nombre,
            "Origen": origen,
            "Destino": destino,
            "Operacion": operacion,
            "Cantidad_animales": cantidad,
            "Categoria_origen": cat_origen,
            "Categoria_destino": cat_destino,
            "Camion": camion,
            "Pisos": estado["Pisos"],
            "Conductor": conductor,
            "Km_aprox_incluyendo_vacio": round(km, 2),
            "Km_vacio_aprox": round(km_vacio, 2),
            "Tiempo_estimado_min": round(minutos, 1),
            "Estado": "ASIGNADO PROVISIONAL",
        })

    pendientes = nuevas_pendientes

# ==========================================================
# 9. MOVIMIENTOS QUE NO SE HAN PODIDO ASIGNAR
# ==========================================================

for tarea in pendientes:
    no_asignados.append({
        "ID_movimiento": tarea["ID_movimiento"],
        "Origen": tarea["Origen"],
        "Destino": tarea["Destino"],
        "Operacion": tarea["Operacion"],
        "Cantidad_animales": tarea["Cantidad_carga"],
        "Categoria_origen": tarea["Categoria_origen"],
        "Categoria_destino": tarea["Categoria_destino"],
        "Motivo": (
            "Sin hueco por capacidad, jornada o bioseguridad. "
            "Revisar historial del camión y planificación."
        )
    })

# ==========================================================
# 10. RESULTADOS
# ==========================================================

df_asignaciones = pd.DataFrame(asignaciones)
df_no_asignados = pd.DataFrame(no_asignados)
df_problemas = pd.DataFrame(problemas)

if not df_asignaciones.empty:

    resumen_camiones = (
        df_asignaciones
        .groupby(["Camion", "Pisos"], as_index=False)
        .agg(
            Viajes=("Camion", "count"),
            Animales=("Cantidad_animales", "sum"),
            Km_aprox=("Km_aprox_incluyendo_vacio", "sum"),
            Dias_trabajados=("Dia", "nunique")
        )
    )

    resumen_dias = (
        df_asignaciones
        .groupby("Dia", as_index=False)
        .agg(
            Viajes=("Camion", "count"),
            Camiones=("Camion", "nunique"),
            Animales=("Cantidad_animales", "sum"),
            Km_aprox=("Km_aprox_incluyendo_vacio", "sum")
        )
    )

else:
    resumen_camiones = pd.DataFrame()
    resumen_dias = pd.DataFrame()

resumen_general = pd.DataFrame([
    ["Movimientos originales", len(df)],
    ["Cargas creadas", len(tareas)],
    ["Cargas asignadas", len(asignaciones)],
    ["Cargas sin asignar", len(no_asignados)],
    ["Problemas de datos", len(problemas)],
    ["Camiones disponibles", len(CAMIONES)],
    ["Máximo de camiones simultáneos", MAX_CAMIONES_DIA],
], columns=["Concepto", "Resultado"])

historial = []

for camion, estado in estado_camiones.items():
    historial.append({
        "Camion": camion,
        "Pisos": estado["Pisos"],
        "Ultima_categoria_semana": estado["Ultima_categoria"],
        "Ultimo_dia_semana": (
            DIAS[estado["Ultimo_dia"]]
            if estado["Ultimo_dia"] is not None else None
        ),
        "Ultimo_destino": estado["Ultimo_destino"],
    })

df_historial = pd.DataFrame(historial)

# ==========================================================
# 11. EXPORTAR EXCEL
# ==========================================================

with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:

    df_asignaciones.to_excel(
        writer, sheet_name="Asignaciones", index=False
    )

    df_no_asignados.to_excel(
        writer, sheet_name="No_asignados", index=False
    )

    df_problemas.to_excel(
        writer, sheet_name="Errores_datos", index=False
    )

    resumen_general.to_excel(
        writer, sheet_name="Resumen", index=False
    )

    resumen_camiones.to_excel(
        writer, sheet_name="Por_camion", index=False
    )

    resumen_dias.to_excel(
        writer, sheet_name="Por_dia", index=False
    )

    df_historial.to_excel(
        writer, sheet_name="Historial_camiones", index=False
    )

    granjas_df.to_excel(
        writer, sheet_name="Granjas", index=False
    )

print("\n======================================")
print("OPTIMIZACION TERMINADA")
print("======================================")
print("Movimientos originales:", len(df))
print("Cargas creadas:", len(tareas))
print("Cargas asignadas:", len(asignaciones))
print("Cargas sin asignar:", len(no_asignados))
print("Problemas de datos:", len(problemas))
print("======================================")
print("Archivo:", SALIDA)
print("======================================")

