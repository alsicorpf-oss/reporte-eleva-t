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
# 🧹 LIMPIAR DATOS
# ==========================================

col_id = COLUMNAS_ORIGEN["id_colaborador"]
col_pais = COLUMNAS_ORIGEN["pais"]

# Eliminar nulos
df = df[[col_id, col_pais]].dropna()

# ==========================================
# 🔄 AGRUPAR POR COLABORADOR
# ==========================================

print("🔄 Procesando datos...")

resultado = []

# Obtener IDs únicos
ids_unicos = df[col_id].unique()

for id_colab in ids_unicos:
    registros = df[df[col_id] == id_colab]

    # Tomar el primer valor de país (puedes cambiar lógica después)
    pais = registros[col_pais].iloc[0]

    resultado.append({
        COLUMNAS_DESTINO["id_colaborador"]: id_colab,
        COLUMNAS_DESTINO["pais"]: pais
    })

# ==========================================
# 📊 CREAR DATAFRAME FINAL
# ==========================================

df_resultado = pd.DataFrame(resultado)

print("\n📄 Resultado generado:")
print(df_resultado.head())

# ==========================================
# 💾 EXPORTAR A EXCEL
# ==========================================

OUTPUT = "reporte_colaboradores.xlsx"

df_resultado.to_excel(OUTPUT, index=False)

print(f"\n✅ Reporte generado: {OUTPUT}")
