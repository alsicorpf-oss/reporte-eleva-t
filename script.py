import pandas as pd

# ==========================================
# 📥 CONFIGURACIÓN DE ENTRADA
# ==========================================

ARCHIVO = "nuevo_formulario_eleva_t.xlsx"
HOJA = "Sheet1"

# Columnas en el Excel original (FUENTE)
COLUMNAS_ORIGEN = {
    "id_colaborador": "ID DE COLABORADOR DEL EMPLEADO A ENTREVISTAR",
    "pais": "SELECCIONA EL PAÍS."
}

# Columnas del nuevo reporte (DESTINO)
COLUMNAS_DESTINO = {
    "id_colaborador": "CODIGO_COLABORADOR",
    "pais": "PAIS"
}

# ==========================================
# 📂 LEER ARCHIVO
# ==========================================

print("📂 Leyendo archivo...")

df = pd.read_excel(ARCHIVO, sheet_name=HOJA)

print("✅ Archivo leído correctamente\n")

# Limpiar nombres de columnas
df.columns = df.columns.astype(str).str.strip()

# ==========================================
# 🔍 VALIDAR COLUMNAS
# ==========================================

for key, col in COLUMNAS_ORIGEN.items():
    if col not in df.columns:
        raise Exception(f"❌ No se encontró la columna: {col}")

print("✅ Columnas validadas correctamente\n")

# ==========================================
# 🧹 FILTRAR SOLO COLUMNAS NECESARIAS
# ==========================================

col_id = COLUMNAS_ORIGEN["id_colaborador"]
col_pais = COLUMNAS_ORIGEN["pais"]

df_filtrado = df[[col_id, col_pais]].dropna()

# ==========================================
# 🔄 PROCESAR TODAS LAS FILAS (SIN ELIMINAR DUPLICADOS)
# ==========================================

print("🔄 Procesando datos (sin eliminar duplicados)...")

resultado = []

for _, row in df_filtrado.iterrows():
    resultado.append({
        COLUMNAS_DESTINO["id_colaborador"]: row[col_id],
        COLUMNAS_DESTINO["pais"]: row[col_pais]
    })

# ==========================================
# 📊 CREAR DATAFRAME FINAL
# ==========================================

df_resultado = pd.DataFrame(resultado)

# 🔥 ORDENAR POR CÓDIGO DE COLABORADOR
df_resultado = df_resultado.sort_values(
    by=COLUMNAS_DESTINO["id_colaborador"]
)

print("\n📄 Resultado generado (ordenado):")
print(df_resultado.head(10))

# ==========================================
# 💾 EXPORTAR A EXCEL
# ==========================================

OUTPUT = "reporte_colaboradores.xlsx"

df_resultado.to_excel(OUTPUT, index=False)

print(f"\n✅ Reporte generado: {OUTPUT}")
