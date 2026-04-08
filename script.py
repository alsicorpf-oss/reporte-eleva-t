import pandas as pd

archivo = "nuevo_formulario_eleva_t.xlsx"

print("📂 Leyendo archivo...")

# 🔹 Ver hojas disponibles
xls = pd.ExcelFile(archivo)
print("\n📑 Hojas disponibles:")
print(xls.sheet_names)

# 🔹 Leer Sheet1 SIN asumir encabezados
df = pd.read_excel(archivo, sheet_name="Sheet1", header=None)

print("\n🔍 Primeras 10 filas (RAW):")
print(df.head(10))

# 🔹 Ahora intentar con header en fila 0
df2 = pd.read_excel(archivo, sheet_name="Sheet1")

print("\n📊 Columnas detectadas (modo normal):")
print(df2.columns.tolist())

# 🔹 Mostrar columnas con comillas (detecta espacios invisibles)
print("\n🧪 Columnas EXACTAS:")
for col in df2.columns:
    print(f"'{col}'")

# 🔹 Mostrar primeras filas interpretadas
print("\n📄 Data interpretada:")
print(df2.head())
