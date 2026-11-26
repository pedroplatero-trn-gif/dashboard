from pathlib import Path
import pandas as pd

base = Path(__file__).parent
archivos = [
    "conversiones_octubre_2026.xlsx",
    "reporte_sucio.xlsx",
]

for nombre in archivos:
    ruta = base / nombre
    print(f"\n### {nombre}")

    if not ruta.exists():
        print(f"No se encuentra: {ruta}")
        continue

    df = pd.read_excel(ruta, sheet_name="Resultados")
    print("Columnas:")
    for columna in df.columns:
        print(f"  {columna!r}")

    print("\nFilas de Resultados:")
    print(df.head(20).to_string(index=False))
