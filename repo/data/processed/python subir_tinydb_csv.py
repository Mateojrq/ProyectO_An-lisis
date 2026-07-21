@'
import pandas as pd
from tinydb import TinyDB
import chardet

archivos = {
    "becarios_agosto_2024": "Base de datos de becarios agosto 2024_procesado.csv",
    "becarios_actualizacion": "Becarios - actualización_procesado.csv",
}

for nombre_bd, archivo in archivos.items():
    print(f"Procesando {archivo}...")

    with open(archivo, "rb") as f:
        resultado = chardet.detect(f.read(100000))
    encoding_detectado = resultado["encoding"]
    print(f"  Encoding detectado: {encoding_detectado}")

    df = pd.read_csv(archivo, encoding=encoding_detectado, low_memory=False)
    df = df.fillna("")
    registros = df.to_dict(orient="records")

    db_file = f"{nombre_bd}_db.json"
    db = TinyDB(db_file)
    db.truncate()
    db.insert_multiple(registros)

    print(f"  Guardados {len(registros)} registros en '{db_file}'")

print("Todo listo!")
