import pandas as pd

archivo = "nuevo_formulario_eleva_t.xlsx"

print("📂 Leyendo archivo...")

# Leer Excel
df = pd.read_excel(archivo)

print("✅ Archivo leído correctamente\n")

# Mostrar columnas
print("📊 Columnas encontradas:")
print(df.columns.tolist())

# Mostrar primeras filas
print("\n🔍 Primeras filas:")
print(df.head())
