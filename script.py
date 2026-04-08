import pandas as pd
from datetime import datetime
import os

# ==========================================
# 📥 CONFIGURACIÓN
# ==========================================

ARCHIVO = "nuevo_formulario_eleva_t.xlsx"
HOJA_PRINCIPAL = "Sheet1"
HOJA_SECUNDARIA = "Hoja1"

CARPETA_SALIDA = "reportes"

COLUMNAS_ORIGEN = {
    "id_colaborador": "ID DE COLABORADOR DEL EMPLEADO A ENTREVISTAR",
    "pais": "SELECCIONA EL PAÍS.",
    "contacto": "Column1",
    "fecha_contacto": "Hora de inicio"
}

COLUMNAS_HOJA1 = {
    "id_colaborador": "ID COLABORADOR",
    "nombre": "COLABORADOR",
    "sucursal": "LUGAR TRABAJO"
}

COLUMNAS_DESTINO = {
    "id_colaborador": "CODIGO_COLABORADOR",
    "pais": "PAIS",
    "contacto": "CONTACTO",
    "fecha_contacto": "FECHA DE CONTACTO",
    "nombre": "NOMBRE COLABORADOR",
    "sucursal": "SUCURSAL"
}

# ==========================================
# 🔧 LIMPIAR ID
# ==========================================

def limpiar_id(valor):
    try:
        return str(int(float(valor)))
    except:
        return str(valor).strip()

# ==========================================
# 📂 LEER ARCHIVOS
# ==========================================

df = pd.read_excel(ARCHIVO, sheet_name=HOJA_PRINCIPAL)
df_hoja1 = pd.read_excel(ARCHIVO, sheet_name=HOJA_SECUNDARIA)

df.columns = df.columns.astype(str).str.strip()
df_hoja1.columns = df_hoja1.columns.astype(str).str.strip()

# ==========================================
# 🧹 FILTRAR
# ==========================================

col_id = COLUMNAS_ORIGEN["id_colaborador"]
col_pais = COLUMNAS_ORIGEN["pais"]
col_contacto = COLUMNAS_ORIGEN["contacto"]
col_fecha = COLUMNAS_ORIGEN["fecha_contacto"]

df_filtrado = df[df[col_id].notna()].copy()

# ==========================================
# 🔑 NORMALIZAR IDS
# ==========================================

df_filtrado["ID_LIMPIO"] = df_filtrado[col_id].apply(limpiar_id)
df_hoja1["ID_LIMPIO"] = df_hoja1[COLUMNAS_HOJA1["id_colaborador"]].apply(limpiar_id)

# ==========================================
# 🔄 PROCESAR
# ==========================================

resultado = []

for _, row in df_filtrado.iterrows():

    id_original = row[col_id]
    id_limpio = row["ID_LIMPIO"]

    match = df_hoja1[df_hoja1["ID_LIMPIO"] == id_limpio]

    if not match.empty:
        nombre = match.iloc[0][COLUMNAS_HOJA1["nombre"]]
        sucursal = match.iloc[0][COLUMNAS_HOJA1["sucursal"]]
    else:
        nombre = None
        sucursal = None

    if pd.isna(nombre):
        nombre = None
    if pd.isna(sucursal):
        sucursal = None

    resultado.append({
        COLUMNAS_DESTINO["id_colaborador"]: id_original,
        COLUMNAS_DESTINO["pais"]: row.get(col_pais, None),
        COLUMNAS_DESTINO["contacto"]: row.get(col_contacto, None),
        COLUMNAS_DESTINO["fecha_contacto"]: row.get(col_fecha, None),
        COLUMNAS_DESTINO["nombre"]: nombre,
        COLUMNAS_DESTINO["sucursal"]: sucursal
    })

# ==========================================
# 📊 RESULTADO
# ==========================================

df_resultado = pd.DataFrame(resultado)

# 🔥 NORMALIZAR COLUMNA ANTES DE ORDENAR
col_sort = COLUMNAS_DESTINO["id_colaborador"]

df_resultado[col_sort] = (
    df_resultado[col_sort]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

# 🔥 ORDENAR (YA SIN ERROR)
df_resultado = df_resultado.sort_values(by=col_sort)

# ==========================================
# 💾 EXPORTAR
# ==========================================

os.makedirs(CARPETA_SALIDA, exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
ruta = f"{CARPETA_SALIDA}/reporte_{timestamp}.xlsx"

df_resultado.to_excel(ruta, index=False)

print("\n📄 Resultado generado:")
print(df_resultado.head(10))

print(f"\n✅ Reporte guardado en: {ruta}")
