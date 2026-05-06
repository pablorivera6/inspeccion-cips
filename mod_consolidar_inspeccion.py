import os
import pandas as pd
import numpy as np
from scipy.spatial import cKDTree
from pyproj import Transformer
import simplekml


# =========================================
# SEPARAR LATITUD Y LONGITUD DESDE GPS
# =========================================

def separar_gps(df):

    lat = []
    lon = []

    for val in df["Localizacion GPS"].astype(str):

        if "," in val:

            partes = val.split(",")

            try:
                lat.append(float(partes[0].strip()))
                lon.append(float(partes[1].strip()))
            except:
                lat.append(np.nan)
                lon.append(np.nan)

        else:
            lat.append(np.nan)
            lon.append(np.nan)

    df["Latitud"] = lat
    df["Longitud"] = lon

    return df


# =========================================
# EXPORTAR KMZ
# =========================================

def exportar_kmz(df_cips, df_campo, match, carpeta):

    kml = simplekml.Kml()

    # CIPS
    folder_cips = kml.newfolder(name="CIPS")

    for _, row in df_cips.iterrows():

        pnt = folder_cips.newpoint()

        pnt.coords = [(row["Long_corr"], row["Lat_corr"])]

        pnt.name = f"PK {round(row['PK_real_m'],1)}"

        pnt.style.iconstyle.color = simplekml.Color.green
        pnt.style.iconstyle.scale = 0.7


    # CAMPO
    folder_campo = kml.newfolder(name="Campo")

    for _, row in df_campo.iterrows():

        if pd.isna(row["Latitud"]):
            continue

        pnt = folder_campo.newpoint()

        pnt.coords = [(row["Longitud"], row["Latitud"])]

        pnt.style.iconstyle.color = simplekml.Color.blue
        pnt.style.iconstyle.scale = 0.7


    # LINEAS MATCH
    folder_line = kml.newfolder(name="Correlacion")

    for _, row in match.iterrows():

        cips = df_cips.iloc[int(row["idx_cips"])]

        line = folder_line.newlinestring()

        line.coords = [
            (row["Longitud"], row["Latitud"]),
            (cips["Long_corr"], cips["Lat_corr"])
        ]

        line.style.linestyle.color = simplekml.Color.yellow
        line.style.linestyle.width = 2


    salida = os.path.join(carpeta, "CIPS_CONSOLIDADO.kmz")

    kml.savekmz(salida)

    print("KMZ generado:", salida)


# =========================================
# CONSOLIDAR INSPECCION
# =========================================

def consolidar_inspeccion(
        carpeta,
        archivo_cips="CIPS_VALIDADO_FINAL.xlsx",
        archivo_campo=None,
        tolerancia_m=20
):

    print("=== CONSOLIDANDO INSPECCION ===")

    ruta_cips = os.path.join(carpeta, archivo_cips)

    if os.path.isabs(archivo_campo):
        ruta_campo = archivo_campo
    else:
        ruta_campo = os.path.join(carpeta, archivo_campo)

    df_cips = pd.read_excel(ruta_cips, sheet_name="Survey Data")
    df_campo = pd.read_excel(ruta_campo)

    # =========================================
    # EXTRAER GPS
    # =========================================

    df_campo = separar_gps(df_campo)

    df_campo_validos = df_campo.dropna(subset=["Latitud","Longitud"])
    df_campo_sin_gps = df_campo[df_campo[["Latitud","Longitud"]].isna().any(axis=1)]

    print("Campo total:", len(df_campo))
    print("GPS valido:", len(df_campo_validos))
    print("Sin GPS:", len(df_campo_sin_gps))

    # =========================================
    # CONVERTIR A METROS
    # =========================================

    transformer = Transformer.from_crs(
        "EPSG:4326",
        "EPSG:3857",
        always_xy=True
    )

    x_cips, y_cips = transformer.transform(
        df_cips["Long_corr"].values,
        df_cips["Lat_corr"].values
    )

    x_campo, y_campo = transformer.transform(
        df_campo_validos["Longitud"].values,
        df_campo_validos["Latitud"].values
    )

    coords_cips = np.column_stack((x_cips, y_cips))
    coords_campo = np.column_stack((x_campo, y_campo))

    # =========================================
    # MATCH ESPACIAL
    # =========================================

    tree = cKDTree(coords_cips)

    dist, idx = tree.query(coords_campo, k=1)

    df_campo_validos["idx_cips"] = idx
    df_campo_validos["distancia_match_m"] = dist

    match = df_campo_validos[df_campo_validos["distancia_match_m"] <= tolerancia_m]
    no_match = df_campo_validos[df_campo_validos["distancia_match_m"] > tolerancia_m]

    # =========================================
    # CREAR TABLA MATCH
    # =========================================

    df_match = match.merge(
        df_cips,
        left_on="idx_cips",
        right_index=True,
        how="left",
        suffixes=("_campo","_cips")
    )

    # =========================================
    # CIPS SIN MATCH
    # =========================================

    cips_match_ids = match["idx_cips"].unique()

    df_cips_solo = df_cips.drop(index=cips_match_ids)

    # =========================================
    # CAMPO SIN MATCH
    # =========================================

    df_campo_solo = no_match.copy()

    # agregar registros sin GPS
    df_campo_solo = pd.concat([df_campo_solo, df_campo_sin_gps])

    # =========================================
    # UNIFICAR TODO
    # =========================================

    df_final = pd.concat([
        df_match,
        df_cips_solo,
        df_campo_solo
    ], ignore_index=True)

    # =========================================
    # ORDENAR POR PK
    # =========================================

    if "PK_real_m" in df_final.columns:
        df_final = df_final.sort_values("PK_real_m")

    # =========================================
    # EXPORTAR
    # =========================================

    salida = os.path.join(carpeta, "CIPS_CONSOLIDADO_FINAL.xlsx")

    df_final.to_excel(salida, index=False)

    print("Archivo generado:", salida)

    print("Filas CIPS:", len(df_cips))
    print("Filas Campo:", len(df_campo))
    print("Filas Final:", len(df_final))

    # =========================================
    # KMZ
    # =========================================

    exportar_kmz(df_cips, df_campo_validos, match, carpeta)

    return salida