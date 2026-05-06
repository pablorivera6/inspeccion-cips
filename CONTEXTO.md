# PCC Integrity — Dashboard de Inspecciones Catódicas

## ¿Qué es esta app?

Dashboard Streamlit para visualizar datos de inspecciones de **Protección Catódica de Corriente Impresa (PAP y DCVG)** de la empresa **PCC Integrity**.  
Los datos vienen de la app móvil **FastField**, exportados como archivos Excel (`.xlsx`).  
No hay conexión a base de datos — el usuario sube los archivos directamente en la interfaz.

## Cómo correr

```bash
cd /ruta/a/pcc-dashboard
pip install -r requirements.txt
streamlit run app.py
```

## Estructura de archivos

```
pcc-dashboard/
├── app.py              ← App completa (toda la lógica en un solo archivo)
├── requirements.txt    ← streamlit, pandas, plotly, msal, requests
└── .streamlit/
    └── secrets.toml    ← Credenciales SharePoint (no usadas aún — ver nota)
```

---

## Estructura del Excel de FastField

Cada inspección exportada genera un `.xlsx` con estas hojas:

| Hoja | Contenido |
|------|-----------|
| `Root` | Metadata: inspector, cargo, fecha, nombre del formulario |
| `subform_1` | Puntos de medición (la hoja principal con todos los datos) |
| `multiphoto_picker_*` | Fotos adjuntas (no se usan en el dashboard) |
| `Form Meta Data` | Timestamps de inicio/fin (no se usa) |

### Hoja `Root` — columnas relevantes
| Columna | Descripción |
|---------|-------------|
| `Personal` | Nombre del inspector |
| `Cargo` | Cargo del inspector |
| `Fecha ` | Fecha de la inspección (ojo: tiene espacio al final) |
| `Form Name` | Nombre del formulario — determina si es PAP o DCVG |

**Detección automática de tipo:**  
- Si `Form Name` contiene "DCVG" → inspección DCVG  
- Cualquier otro caso → inspección PAP

---

## Columnas de `subform_1` — PAP

| Columna interna | Descripción | Notas |
|----------------|-------------|-------|
| `Abscisa` | Kilómetro del punto | Formato mixto: `340+000` o `340.458` |
| `Localizacion GPS` | Coordenadas | String `"lat,lon"` — se parsea al cargar |
| `On [mV]` | Potencial encendido | Número; 0.0 = no medido (se convierte a NaN) |
| `Off [mV]` | Potencial apagado | Igual. Base para calcular estado de protección |
| `IR ON-OFF [mV]` | Caída IR | Número |
| `Voltaje AC` | Voltaje de corriente alterna | Número |
| `Potencial Natural [mV]` | Potencial natural | Número |
| `Resistencia entre NEG1-NEG2 [ohm]` | Resistencia | Número |
| `Estado Pintura` | Estado físico | Texto: Bueno / Regular / Malo |
| `Estado Conexiones` | Estado físico | Texto |
| `Estado Verticalidad` | Estado físico | Texto |
| `Tipo mantenimiento` | Tipo | Texto: Tipo 1 / Tipo 4 / etc. |
| `Observaciones` | Texto libre | String |
| `Tramo` | Nombre del tramo / zona | Ej: "Cupiagua-Cusiana" |
| `Tipo de tramo` | Tipo | Ej: "Troncal" |

### Parseo de Abscisa
```python
# "340+000" → 340000 metros
# "340.458" → 340458 metros (se multiplica por 1000)
def parse_abscisa(val):
    if "+" in str(val):
        km, m = str(val).split("+")
        return int(km) * 1000 + int(m)
    else:
        return round(float(val) * 1000)
```

### Lógica de Estado de Protección (PAP)
```python
# Basado en el valor de Off [mV]
if off_mv < -1200:          → "Sobreprotegido"  (azul oscuro #0D47A1)
if -1200 <= off_mv <= -850: → "Protegido"        (azul #1565C0)
if off_mv > -850:           → "Sin protección"   (rojo #C62828)
if off_mv == 0 o NaN:       → "Sin medición"     (gris)

# Líneas de referencia en el gráfico: -850 mV (azul) y -1200 mV (rojo)
```

---

## Columnas de `subform_1` — DCVG

| Columna interna | Descripción |
|----------------|-------------|
| `Abscisa` | Mismo formato que PAP |
| `Localizacion GPS` | Mismo formato que PAP |
| `PotencialONmV` | Potencial ON en mV |
| `PotencialOFFmV` | Potencial OFF en mV |
| `P_REmV` | P_RE en mV |
| `OL_REmV` | OL_RE en mV |
| `PORC_IR` | Porcentaje IR (0–100) — columna principal para gráficos |
| `CaracterON_x002d_OFF` | Carácter ON-OFF (texto: CC, etc.) |
| `Clasificacion` | Clasificación del punto (Muy Pequeño, Pequeño, Mediano, etc.) |
| `Comentarios` | Texto libre |
| `Tramo` | Igual que PAP |

### Umbrales PORC_IR
```
15% → límite "Muy pequeño"  (verde)
35% → límite "Pequeño"       (naranja)
60% → límite "Mediano"       (rojo)
```

---

## Funcionalidades actuales

### Carga de archivos
- Múltiples archivos `.xlsx` simultáneos
- Detección automática de tipo (PAP / DCVG)
- Agrupación por tramo/zona en el sidebar

### Dashboard PAP (todo en una página)
1. Cabecera: tramo, inspector, cargo, fecha, total puntos
2. **Fila 1:** Tabla scrollable (Abscisa, Off mV, On mV, Observaciones) | Mapa GPS con colores por estado | Panel derecho con conteo por estado + donut chart
3. **Fila 2:** Gráfico de líneas On mV + Off mV vs Abscisa (con zona sombreada y líneas de referencia -850/-1200)
4. **Fila 3:** IR ON-OFF vs Abscisa | Voltaje AC vs Abscisa (lado a lado)
5. **Fila 4:** Tabla mediciones eléctricas completa
6. **Fila 5:** Tabla estado de infraestructura (Estado Pintura, Conexiones, Verticalidad, Tipo Mantenimiento)
7. Footer PCC

### Dashboard DCVG (todo en una página)
1. Cabecera igual al PAP
2. **Fila 1:** Barra Carácter ON-OFF | Barra Clasificación | Tabla (Carácter, Clasificación, PORC_IR) | Donut Clasificación
3. **Fila 2:** Gráfico PORC_IR vs Abscisa (con umbrales 15/35/60%) | Mapa GPS
4. **Fila 3:** Tabla completa DCVG
5. Footer PCC

### Resumen Global (cuando hay >1 archivo)
- KPIs: total archivos, tramos, puntos PAP, puntos DCVG
- Mapa global con todos los puntos (coloreados por tramo)
- Tabla de todas las inspecciones cargadas
- Barras de estado de protección por tramo (PAP)

---

## Diseño / Estética

- **Inspirado en Power BI**: fondo blanco, sidebar gris claro, layout de cuadrícula
- **Color principal:** `#8B0000` (rojo oscuro PCC)
- **Protegido:** `#1565C0` (azul)
- **Sobreprotegido:** `#0D47A1` (azul oscuro)
- **Sin protección:** `#C62828` (rojo)
- **Footer:** barra roja con logo PCC al pie de cada dashboard
- **Sin tabs:** todo visible en scroll continuo

---

## Conexión a SharePoint (pendiente / no implementada)

La app originalmente iba a conectarse a SharePoint Online de `catodica.sharepoint.com`.  
Se dejó esta funcionalidad pendiente porque la autenticación Azure AD no estaba configurada.  
Las listas de SharePoint relevantes son:
- `Subform` — puntos PAP
- `Subform DCVG` — puntos DCVG  
- `PCC_Cabecera_Inspeccion` — cabeceras de inspección

Si se quiere retomar, requiere Azure AD App Registration con permisos `Sites.Read.All` en SharePoint.  
Las credenciales van en `.streamlit/secrets.toml` bajo `[sharepoint]`.

---

## Lo que falta / ideas para continuar

- [ ] Filtro por rango de Abscisa (slider) dentro del dashboard
- [ ] Filtro por estado de protección (checkbox)
- [ ] Comparación de dos inspecciones del mismo tramo en el mismo gráfico
- [ ] Exportar el dashboard como PDF o imagen
- [ ] Tabla de resumen con estadísticas (min, max, promedio de Off mV por tramo)
- [ ] Conexión directa a SharePoint (sin subir archivo manual)
- [ ] Soporte para formulario DCVG con resistividad (tiene columnas adicionales de resistividad que no están mapeadas aún)
- [ ] Logo PCC real en el footer (imagen, no emoji)
