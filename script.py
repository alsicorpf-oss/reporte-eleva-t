import pandas as pd

# ==========================================
# 📥 CONFIGURACIÓN DE ENTRADA
# ==========================================

ARCHIVO = "nuevo_formulario_eleva_t.xlsx"
HOJA = "Sheet1"

# Columnas en el Excel original (FUENTE)
COLUMNAS_ORIGEN = {
    "id_colaborador": "ID DE COLABORADOR DEL EMPLEADO A ENTREVISTAR",
    "pais": "SELECCIONA EL PAÍS.",
    "contacto": "Column1",
    "fecha_contacto": "Hora de inicio"
}

# Columnas del nuevo reporte (DESTINO)
COLUMNAS_DESTINO = {
    "id_colaborador": "CODIGO_COLABORADOR",
    "pais": "PAIS",
    "contacto": "CONTACTO",
    "fecha_contacto": "FECHA DE CONTACTO"
}

# ==========================================
# 📂 PROCESO ORIGINAL (NO TOCAR)
# ==========================================

print("📂 Leyendo archivo principal...")

df = pd.read_excel(ARCHIVO, sheet_name=HOJA)

print("✅ Archivo leído correctamente\n")

df.columns = df.columns.astype(str).str.strip()

for key, col in COLUMNAS_ORIGEN.items():
    if col not in df.columns:
        raise Exception(f"❌ No se encontró la columna: {col}")

print("✅ Columnas validadas correctamente\n")

col_id = COLUMNAS_ORIGEN["id_colaborador"]
col_pais = COLUMNAS_ORIGEN["pais"]
col_contacto = COLUMNAS_ORIGEN["contacto"]
col_fecha = COLUMNAS_ORIGEN["fecha_contacto"]

df_filtrado = df[[col_id, col_pais, col_contacto, col_fecha]].dropna()

print("🔄 Procesando datos...")

resultado = []

for _, row in df_filtrado.iterrows():
    resultado.append({
        COLUMNAS_DESTINO["id_colaborador"]: row[col_id],
        COLUMNAS_DESTINO["pais"]: row[col_pais],
        COLUMNAS_DESTINO["contacto"]: row[col_contacto],
        COLUMNAS_DESTINO["fecha_contacto"]: row[col_fecha]
    })

df_resultado = pd.DataFrame(resultado)

df_resultado = df_resultado.sort_values(
    by=COLUMNAS_DESTINO["id_colaborador"]
)

print("\n📄 Resultado generado:")
print(df_resultado.head(10))

OUTPUT = "reporte_colaboradores.xlsx"
df_resultado.to_excel(OUTPUT, index=False)

print(f"\n✅ Reporte generado: {OUTPUT}")

# ==========================================
# 🔍 DEBUG HOJA1 (MEJORADO)
# ==========================================

print("\n==============================")
print("🔍 DEBUG HOJA1 (TODAS LAS COLUMNAS)")
print("==============================")

df_hoja1 = pd.read_excel(ARCHIVO, sheet_name="Hoja1")

df_hoja1.columns = df_hoja1.columns.astype(str).str.strip()

# 🔥 MOSTRAR TODAS LAS COLUMNAS COMO ARRAY COMPLETO
columnas = df_hoja1.columns.tolist()

print("\n📊 TOTAL COLUMNAS:", len(columnas))

print("\n📋 COLUMNAS COMPLETAS:")
print(columnas)

# 🔥 OPCIONAL: imprimir una por línea (más legible)
print("\n📌 COLUMNAS UNA POR UNA:")
for col in columnas:
    print(f"'{col}'")

# 🔍 Mostrar primeras filas
print("\n🔍 Primeras filas de Hoja1:")
print(df_hoja1.head(10))
