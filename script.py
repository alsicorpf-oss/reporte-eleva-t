import pandas as pd

# ==========================================
# 📥 CONFIGURACIÓN DE ENTRADA
# ==========================================

ARCHIVO = "nuevo_formulario_eleva_t.xlsx"
HOJA_PRINCIPAL = "Sheet1"
HOJA_SECUNDARIA = "Hoja1"

# Columnas en el Excel original (FUENTE)
COLUMNAS_ORIGEN = {
    "id_colaborador": "ID DE COLABORADOR DEL EMPLEADO A ENTREVISTAR",
    "pais": "SELECCIONA EL PAÍS.",
    "contacto": "Column1",
    "fecha_contacto": "Hora de inicio"
}

# Columnas desde Hoja1
COLUMNAS_HOJA1 = {
    "id_colaborador": "ID COLABORADOR",
    "nombre": "COLABORADOR",
    "sucursal": "LUGAR TRABAJO"
}

# Columnas del nuevo reporte (DESTINO)
COLUMNAS_DESTINO = {
    "id_colaborador": "CODIGO_COLABORADOR",
    "pais": "PAIS",
    "contacto": "CONTACTO",
    "fecha_contacto": "FECHA DE CONTACTO",
    "nombre": "NOMBRE COLABORADOR",
    "sucursal": "SUCURSAL"
}

# ==========================================
# 📂 LEER ARCHIVOS
# ==========================================

df = pd.read_excel(ARCHIVO, sheet_name=HOJA_PRINCIPAL)
df_hoja1 = pd.read_excel(ARCHIVO, sheet_name=HOJA_SECUNDARIA)

# Limpiar columnas
df.columns = df.columns.astype(str).str.strip()
df_hoja1.columns = df_hoja1.columns.astype(str).str.strip()

# ==========================================
# 🔍 VALIDAR COLUMNAS
# ==========================================

for col in COLUMNAS_ORIGEN.values():
    if col not in df.columns:
        raise Exception(f"❌ No se encontró la columna: {col}")

for col in COLUMNAS_HOJA1.values():
    if col not in df_hoja1.columns:
        raise Exception(f"❌ No se encontró la columna en Hoja1: {col}")

# ==========================================
# 🧹 FILTRAR DATOS PRINCIPALES
# ==========================================

col_id = COLUMNAS_ORIGEN["id_colaborador"]
col_pais = COLUMNAS_ORIGEN["pais"]
col_contacto = COLUMNAS_ORIGEN["contacto"]
col_fecha = COLUMNAS_ORIGEN["fecha_contacto"]

df_filtrado = df[[col_id, col_pais, col_contacto, col_fecha]].dropna()

# ==========================================
# 🔗 PREPARAR JOIN CON HOJA1
# ==========================================

col_id_h1 = COLUMNAS_HOJA1["id_colaborador"]

# Normalizar tipos (CLAVE 🔥)
df_filtrado[col_id] = df_filtrado[col_id].astype(str).str.strip()
df_hoja1[col_id_h1] = df_hoja1[col_id_h1].astype(str).str.strip()

# ==========================================
# 🔄 PROCESAR DATOS (JOIN)
# ==========================================

resultado = []

for _, row in df_filtrado.iterrows():

    id_actual = row[col_id]

    # Buscar en Hoja1
    match = df_hoja1[df_hoja1[col_id_h1] == id_actual]

    if not match.empty:
        nombre = match.iloc[0][COLUMNAS_HOJA1["nombre"]]
        sucursal = match.iloc[0][COLUMNAS_HOJA1["sucursal"]]
    else:
        nombre = None
        sucursal = None

    resultado.append({
        COLUMNAS_DESTINO["id_colaborador"]: id_actual,
        COLUMNAS_DESTINO["pais"]: row[col_pais],
        COLUMNAS_DESTINO["contacto"]: row[col_contacto],
        COLUMNAS_DESTINO["fecha_contacto"]: row[col_fecha],
        COLUMNAS_DESTINO["nombre"]: nombre,
        COLUMNAS_DESTINO["sucursal"]: sucursal
    })

# ==========================================
# 📊 CREAR DATAFRAME FINAL
# ==========================================

df_resultado = pd.DataFrame(resultado)

# Ordenar
df_resultado = df_resultado.sort_values(
    by=COLUMNAS_DESTINO["id_colaborador"]
)

# ==========================================
# 💾 EXPORTAR
# ==========================================

OUTPUT = "reporte_colaboradores.xlsx"
df_resultado.to_excel(OUTPUT, index=False)

print("\n📄 Resultado generado:")
print(df_resultado.head(10))

print(f"\n✅ Reporte generado: {OUTPUT}")
