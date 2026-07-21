import pandas as pd
import json

# archivo: (nombre_archivo_original, tipo)
archivos = {
    "becarios_agosto_2024": ("Base de datos de becarios agosto 2024_procesado.json", "json"),
    "becarios_actualizacion": ("Becarios_actualizacion_procesado.csv", "csv"),
}

for nombre, (archivo, tipo) in archivos.items():
    print(f"Procesando {archivo}...")

    if tipo == "json":
        with open(archivo, "r", encoding="utf-8") as f:
            registros = json.load(f)
        df = pd.DataFrame(registros)
    else:
        df = pd.read_csv(archivo, low_memory=False)

    salida = f"{nombre}_para_knime.csv"
    df.to_csv(salida, index=False, encoding="utf-8")
    print(f"  Exportado: {salida} ({len(df)} filas)\n")

print("Listo, ya puedes usar CSV Reader en KNIME con estos archivos.")