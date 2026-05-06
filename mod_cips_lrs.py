# mod_cips_lrs.py
import pandas as pd
import numpy as np
import geopandas as gpd
import os
import matplotlib.pyplot as plt
import simplekml


from pyproj import Transformer
from sklearn.linear_model import LinearRegression
from shapely.ops import linemerge
from shapely.geometry import LineString, MultiLineString

import os
import numpy as np
import pandas as pd
from shapely.geometry import Point, LineString
import shapefile  # pyshp


# =========================================
# LEER LINEA DESDE SHAPEFILE
# =========================================

def leer_linea_shp(ruta_shp):

    sf = shapefile.Reader(ruta_shp)
    shapes = sf.shapes()

    # tomar la primera geometría (línea)
    puntos = shapes[0].points

    return LineString(puntos)


# =========================================
# DETECTAR SHAPEFILE CORRECTO
# =========================================

def detectar_ducto(carpeta_shp, df_cips, sample=50):

    print("Detectando ducto automáticamente...")

    # tomar muestra de puntos CIPS
    df_sample = df_cips.sample(min(sample, len(df_cips)))

    puntos_cips = [
        Point(xy) for xy in zip(df_sample["Long_corr"], df_sample["Lat_corr"])
    ]

    mejor_shp = None
    mejor_score = float("inf")

    for archivo in os.listdir(carpeta_shp):

        if not archivo.endswith(".shp"):
            continue

        ruta_shp = os.path.join(carpeta_shp, archivo)

        try:

            linea = leer_linea_shp(ruta_shp)

            # calcular distancia promedio
            distancias = [
                punto.distance(linea) for punto in puntos_cips
            ]

            score = np.mean(distancias)

            print(f"{archivo} → distancia promedio: {score}")

            if score < mejor_score:
                mejor_score = score
                mejor_shp = ruta_shp

        except Exception as e:
            print(f"Error en {archivo}: {e}")

    print("\nDUCTO DETECTADO:", mejor_shp)

    return mejor_shp

def exportar_kmz_comparacion(gdf, carpeta):

    salida_kmz = os.path.join(carpeta, "CIPS_COMPARACION_COORDENADAS.kmz")

    kml = simplekml.Kml()

    # =========================
    # CARPETA GPS ORIGINAL
    # =========================

    carpeta_gps = kml.newfolder(name="GPS Original")

    for _, row in gdf.iterrows():

        pnt = carpeta_gps.newpoint()

        pnt.coords = [(row["Long"], row["Lat"])]

        pnt.name = f"PK {round(row['PK_real_m'],1)} m"

        pnt.description = f"""
        PK Equipo: {row['PK_equipo']}
        OFF (mV): {row['Off_mV_limpio']}
        Estado: {row['Estado_CP']}
        """

        pnt.style.iconstyle.color = simplekml.Color.red
        pnt.style.iconstyle.scale = 0.7


    # =========================
    # CARPETA CORREGIDOS
    # =========================

    carpeta_corr = kml.newfolder(name="GPS Corregido (sobre ducto)")

    for _, row in gdf.iterrows():

        pnt = carpeta_corr.newpoint()

        pnt.coords = [(row["Long_corr"], row["Lat_corr"])]

        pnt.name = f"PK {round(row['PK_real_m'],1)} m"

        pnt.description = f"""
        PK Real: {round(row['PK_real_m'],1)} m
        OFF (mV): {row['Off_mV_limpio']}
        Estado: {row['Estado_CP']}
        """

        pnt.style.iconstyle.color = simplekml.Color.green
        pnt.style.iconstyle.scale = 0.7


    # =========================
    # LINEA DE DESVIACIÓN
    # =========================

    carpeta_lineas = kml.newfolder(name="Corrección GPS")

    for _, row in gdf.iterrows():

        line = carpeta_lineas.newlinestring()

        line.coords = [
            (row["Long"], row["Lat"]),
            (row["Long_corr"], row["Lat_corr"])
        ]

        line.style.linestyle.color = simplekml.Color.yellow
        line.style.linestyle.width = 2


    kml.savekmz(salida_kmz)

    print("KMZ generado:", salida_kmz)

def ejecutar_cips_lrs(carpeta, archivo_unificado, shp_path):
    # Mensaje en consola para confirmar inicio del módulo
    print("=== SCRIPT INICIADO (mod_cips_lrs) ===")

    # =========================
    # CONFIGURACIÓN GENERAL
    # =========================

    BUFFER_M = 1                # Buffer de referencia (m) para control geométrico (ya no condiciona el snap)
    CRITERIO_OK = -850          # Criterio AMPP/NACE de protección adecuada (mV)
    CRITERIO_MARGINAL = -1200   # Umbral de sobreprotección (mV)

    # =========================
    # 1. CARGA Y NORMALIZACIÓN DE DATOS
    # =========================

    # Abrir el archivo Excel unificado
    xls = pd.ExcelFile(archivo_unificado)

    # Leer hoja principal de inspección CIPS
    df = pd.read_excel(xls, sheet_name="Survey Data")

    # Leer hoja DCP (se conserva intacta para exportar)
    df_dcp = pd.read_excel(xls, sheet_name="DCP Data")

    # Renombrar columnas a nombres normalizados para el procesamiento
    df = df.rename(columns={
        "Dist From Start": "PK_equipo",          # PK reportado por el equipo
        "On Voltage": "On_V",                    # Potencial ON en Voltios
        "Off Voltage": "Off_V",                  # Potencial OFF en Voltios
        "Latitude": "Lat",                       # Latitud GPS
        "Longitude": "Long",                     # Longitud GPS
        "Comment": "Comentario",                 # Comentarios del APP
        "DCP/Feature/DCVG Anomaly": "Anomalia"   # Anomalía asociada
    })
    # =========================
    # LIMPIEZA SEGURA DE COORDENADAS
    # =========================

    def limpiar_coord(valor):
        try:
            # limpia espacios
            valor = str(valor).strip()
            
            # valores basura comunes
            if valor in ["", ".", "-", "None", "nan"]:
                return np.nan
            
            return float(valor)
        
        except:
            return np.nan

    df["Lat"] = df["Lat"].apply(limpiar_coord)
    df["Long"] = df["Long"].apply(limpiar_coord)
    # =========================
    # 2. INTERPOLACIÓN DE COORDENADAS GPS FALTANTES
    # =========================

    # Para Lat y Long, se interpolan valores faltantes usando PK_equipo
    for coord in ["Lat", "Long"]:
        mask = df[coord].isna()                  # Identificar filas sin coordenada
        if mask.any():
            modelo = LinearRegression()           # Modelo de regresión lineal
            modelo.fit(
                df.loc[~mask, ["PK_equipo"]],    # PK donde sí hay coordenada
                df.loc[~mask, coord]              # Coordenadas válidas
            )
            # Predecir coordenadas faltantes
            df.loc[mask, coord] = modelo.predict(
                df.loc[mask, ["PK_equipo"]]
            )

    # =========================
    # 3. CONVERSIÓN A SISTEMA DE COORDENADAS MÉTRICO (GIS)
    # =========================

    # Transformador de WGS84 (lat/long) a Web Mercator (metros)
    t = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)

    # Conversión de coordenadas
    df["X"], df["Y"] = t.transform(
        df["Long"].values,
        df["Lat"].values
    )

    # Crear GeoDataFrame con puntos CIPS
    gdf = gpd.GeoDataFrame(
        df,
        geometry=gpd.points_from_xy(df.X, df.Y),
        crs=3857
    )

    # =========================
    # 4. CARGA Y PREPARACIÓN DEL DUCTO
    # =========================

    # Buscar automáticamente el shapefile en la carpeta
    ducto = gpd.read_file(shp_path)

    # Asignar CRS si el shapefile no lo tiene definido
    if ducto.crs is None:
        ducto = ducto.set_crs(epsg=4326)

    # Reproyectar ducto a sistema métrico
    ducto = ducto.to_crs(3857)

    # Extraer solo geometrías LineString válidas
    lineas_simples = []
    for geom in ducto.geometry:
        if geom is None:
            continue
        if isinstance(geom, LineString):
            lineas_simples.append(geom)
        elif isinstance(geom, MultiLineString):
            for parte in geom.geoms:
                if isinstance(parte, LineString):
                    lineas_simples.append(parte)

    # Validar que existan líneas válidas
    if len(lineas_simples) == 0:
        raise ValueError("No se encontraron LineString válidas en el shapefile")

    # Unificar todas las líneas en una sola geometría
    merged = linemerge(lineas_simples)

    # Si aún es MultiLineString, concatenar coordenadas
    if isinstance(merged, MultiLineString):
        coords = []
        for ls in merged.geoms:
            coords.extend(list(ls.coords))
        linea = LineString(coords)
    else:
        linea = merged

    # Información de control
    print("✔ Ducto preparado para LRS")
    print("Tipo geometría final:", linea.geom_type)
    print("Longitud ducto (m):", round(linea.length, 2))

    # =========================
    # 5. SNAP DE PUNTOS CIPS A LA TRAZA
    # =========================

    # Proyectar cada punto CIPS sobre la línea del ducto
    gdf["geom_snap"] = gdf.geometry.apply(
        lambda p: linea.interpolate(linea.project(p))
    )

    # Calcular distancia del punto original a la traza
    gdf["Dist_traza_m"] = gdf.geometry.distance(gdf["geom_snap"])

    # Forzar la geometría final a la posición sobre el ducto
    gdf["geometry"] = gdf["geom_snap"]

    # =========================
    # 6. PK GEOMÉTRICO (LRS)
    # =========================

    # Calcular PK real basado en la proyección sobre la línea
    gdf["PK_geom_m"] = gdf.geometry.apply(
        lambda p: linea.project(p)
    )

    # =========================
    # 7. AUTO-DETECCIÓN DE SENTIDO DEL PK
    # =========================

    # Comparar PK del equipo vs PK geométrico
    df_pk = gdf[["PK_equipo", "PK_geom_m"]].dropna()
    corr = df_pk["PK_equipo"].corr(df_pk["PK_geom_m"])
    print("Correlación PK_equipo vs PK_geom_m:", round(corr, 3))

    # Si la correlación es negativa, el PK está invertido
    if corr < 0:
        print("⚠ PK invertido detectado. Corrigiendo automáticamente...")
        gdf["PK_geom_m"] = linea.length - gdf["PK_geom_m"]
    else:
        print("✔ Sentido del PK correcto.")

    # Normalizar PK desde 0
    gdf["PK_real_m"] = gdf["PK_geom_m"] - gdf["PK_geom_m"].min()
    gdf["PK_real_km"] = gdf["PK_real_m"] / 1000

    # Ordenar espacialmente los registros
    gdf = gdf.sort_values("PK_geom_m").reset_index(drop=True)

    # =========================
    # 8. LATITUD Y LONGITUD CORREGIDAS
    # =========================

    # Transformador inverso a WGS84
    t_back = Transformer.from_crs("EPSG:3857", "EPSG:4326", always_xy=True)

    # Obtener coordenadas corregidas sobre el ducto
    gdf["Long_corr"], gdf["Lat_corr"] = t_back.transform(
        gdf.geometry.x.values,
        gdf.geometry.y.values
    )

    # =========================
    # 9. CÁLCULOS CIPS
    # =========================

    # Convertir potenciales a milivoltios
    gdf["On_mV"] = gdf["On_V"]*1000
    gdf["Off_mV"] = gdf["Off_V"]*1000

    # Calcular IR Drop original
    gdf["IR_Drop_mV"] = gdf["On_mV"] - gdf["Off_mV"]

    # =========================
    # 10. LIMPIEZA DE OUTLIERS (ON / OFF)
    # =========================

    WINDOW = 15        # Tamaño de ventana espacial
    UMBRAL_MV = 250    # Umbral de diferencia aceptable

    # ---- OFF ----
    off_mediana = gdf["Off_mV"].rolling(WINDOW, center=True).median()
    off_delta = abs(gdf["Off_mV"] - off_mediana)
    off_outlier = off_delta > UMBRAL_MV

    gdf["Off_mV_limpio"] = gdf["Off_mV"]
    gdf.loc[off_outlier, "Off_mV_limpio"] = off_mediana[off_outlier]

    # ---- ON ----
    on_mediana = gdf["On_mV"].rolling(WINDOW, center=True).median()
    on_delta = abs(gdf["On_mV"] - on_mediana)
    on_outlier = on_delta > UMBRAL_MV

    gdf["On_mV_limpio"] = gdf["On_mV"]
    gdf.loc[on_outlier, "On_mV_limpio"] = on_mediana[on_outlier]

    # IR Drop limpio
    gdf["IR_Drop_mV_limpio"] = gdf["On_mV_limpio"] - gdf["Off_mV_limpio"]

    # =========================
    # 11. VALIDACIÓN AMPP / NACE
    # =========================

    # Clasificación del estado de protección
    def validar_cips(off):
        if off <= CRITERIO_MARGINAL:
            return "SOBREPROTEGIDO"
        elif off <= CRITERIO_OK:
            return "PROTEGIDO"
        else:
            return "DESPROTEGIDO"

    gdf["Estado_CP"] = gdf["Off_mV_limpio"].apply(validar_cips)

    # =========================
    # 12. LIMPIEZA FINAL DE COLUMNAS
    # =========================

    COLUMNAS_ELIMINAR = [
        "X", "Y", "geometry", "geom_snap", "Off Time",
        "Fix Quality", "GPS Type", "Sats In Use",
        "PDOP", "HDOP", "VDOP", "Fix Time",
        "Dist_traza_m", "PK_real_km",
        "On_mV", "Off_mV", "IR_Drop_mV"
    ]

    # Eliminar columnas técnicas no requeridas en el entregable
    gdf_final = gdf.drop(columns=COLUMNAS_ELIMINAR, errors="ignore")

    # =========================
    # 13. EXPORTACIÓN FINAL
    # =========================

    # Ruta del archivo de salida
    salida = os.path.join(carpeta, "CIPS_VALIDADO_FINAL.xlsx")

    # Exportar ambas hojas al Excel final
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        gdf_final.to_excel(writer, sheet_name="Survey Data", index=False)
        df_dcp.to_excel(writer, sheet_name="DCP Data", index=False)

    print("===================================")
    print("✔ ARCHIVO EXCEL GENERADO")
    print("✔ Ruta:", salida)
    print("✔ Filas exportadas:", len(gdf))
    print("===================================")

    return salida

    