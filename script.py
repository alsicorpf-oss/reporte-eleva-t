import pandas as pd

archivo = "nuevo_formulario_eleva_t.xlsx"

print("📂 Leyendo archivo...")

# 🔹 Leer Sheet1
df = pd.read_excel(archivo, sheet_name="Sheet1")

print("✅ Archivo leído correctamente\n")

# 🔹 Limpiar nombres de columnas
df.columns = df.columns.astype(str).str.strip()

# 🎯 Nombre exacto de la columna
columna = "ID DE COLABORADOR DEL EMPLEADO A ENTREVISTAR"

# 🚨 Validación
if columna not in df.columns:
    print(f"\n❌ No se encontró la columna: {columna}")
else:
    print(f"\n🆔 Columna encontrada: {columna}")

    # 🔹 Obtener todos los valores (sin nulos)
    ids = df[columna].dropna()

    print("\n📋 TODOS los IDs encontrados:\n")

    for i, valor in enumerate(ids, start=1):
        print(f"{i}. {valor}")
