# =====================================================
# PUNTO 7 - ANALISIS CON PYTHON
# Proyecto: Educacion Superior en Ecuador 2015-2023
# =====================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine

engine = create_engine(
    "postgresql://neondb_owner:npg_WRUJ32gYkPvZ@ep-mute-grass-auk1vp9n-pooler.c-10.us-east-1.aws.neon.tech/neondb?sslmode=require"
)

TABLA_ACADEMICA = "fact_educacion_superior"
TABLA_BECARIOS = "fact_becarios"
TABLA_MINEDUC = "fact_matricula_escolar"

df_academica = pd.read_sql(f"SELECT * FROM {TABLA_ACADEMICA}", engine)
df_becarios = pd.read_sql(f"SELECT * FROM {TABLA_BECARIOS}", engine)
df_mineduc = pd.read_sql(f"SELECT * FROM {TABLA_MINEDUC}", engine)

print("Dimensiones:")
print("Academica:", df_academica.shape)
print("Becarios:", df_becarios.shape)
print("MINEDUC:", df_mineduc.shape)

print("\n--- Estadisticas descriptivas (academica) ---")
print(df_academica[["TOTAL_ESTUDIANTES", "TOTAL_DOCENTES", "TOTAL_ARTICULOS"]].describe())

print("\n--- Nulos restantes ---")
print(df_academica.isnull().sum().sort_values(ascending=False).head(15))

plt.figure(figsize=(10, 5))
matricula_anual = df_academica.groupby("AÑO")["TOTAL_ESTUDIANTES"].sum()
matricula_anual.plot(kind="line", marker="o")
plt.title("Evolucion de la matricula de educacion superior (2015-2023)")
plt.xlabel("Año")
plt.ylabel("Total de estudiantes")
plt.grid(True)
plt.tight_layout()
plt.savefig("grafico_1_evolucion_matricula.png", dpi=150)
plt.close()

plt.figure(figsize=(8, 5))
df_academica.groupby("TIPO_FINANCIAMIENTO")["TOTAL_ESTUDIANTES"].sum().plot(kind="bar", color=["#4C72B0", "#DD8452"])
plt.title("Matricula total por tipo de financiamiento")
plt.ylabel("Total de estudiantes")
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("grafico_2_matricula_financiamiento.png", dpi=150)
plt.close()

plt.figure(figsize=(10, 6))
top_ies = df_academica.groupby("NOMBRE_IES_LIMPIO")["TOTAL_ESTUDIANTES"].sum().sort_values(ascending=False).head(10)
top_ies.plot(kind="barh")
plt.title("Top 10 instituciones con mayor matricula")
plt.xlabel("Total de estudiantes")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.savefig("grafico_3_top10_ies.png", dpi=150)
plt.close()

plt.figure(figsize=(8, 6))
numericas = df_academica[["TOTAL_ESTUDIANTES", "TOTAL_DOCENTES", "TOTAL_ARTICULOS"]].dropna()
sns.heatmap(numericas.corr(), annot=True, cmap="coolwarm", fmt=".2f")
plt.title("Matriz de correlacion - variables numericas")
plt.tight_layout()
plt.savefig("grafico_4_correlacion.png", dpi=150)
plt.close()

plt.figure(figsize=(8, 5))
sns.boxplot(x=df_academica["TOTAL_ESTUDIANTES"])
plt.title("Deteccion de outliers - Total de estudiantes por registro")
plt.tight_layout()
plt.savefig("grafico_5_outliers.png", dpi=150)
plt.close()

Q1 = df_academica["TOTAL_ESTUDIANTES"].quantile(0.25)
Q3 = df_academica["TOTAL_ESTUDIANTES"].quantile(0.75)
IQR = Q3 - Q1
limite_inf = Q1 - 1.5 * IQR
limite_sup = Q3 + 1.5 * IQR

outliers = df_academica[(df_academica["TOTAL_ESTUDIANTES"] < limite_inf) | (df_academica["TOTAL_ESTUDIANTES"] > limite_sup)]
print(f"\nOutliers detectados en TOTAL_ESTUDIANTES: {len(outliers)} de {len(df_academica)} registros")

df_academica_limpio = df_academica[(df_academica["TOTAL_ESTUDIANTES"] >= limite_inf) & (df_academica["TOTAL_ESTUDIANTES"] <= limite_sup)].copy()
df_academica_limpio["NOMBRE_IES_LIMPIO"] = df_academica_limpio["NOMBRE_IES_LIMPIO"].astype(str).str.strip().str.upper()
df_academica_limpio["AÑO"] = df_academica_limpio["AÑO"].astype(int)

print(f"\nRegistros finales tras limpieza avanzada: {len(df_academica_limpio)}")

from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

resumen_ies = df_academica_limpio.groupby("NOMBRE_IES_LIMPIO").agg(
    total_matricula=("TOTAL_ESTUDIANTES", "sum"),
    total_docentes=("TOTAL_DOCENTES", "sum"),
    total_articulos=("TOTAL_ARTICULOS", "sum")
).reset_index()

X = StandardScaler().fit_transform(resumen_ies[["total_matricula", "total_articulos"]])
kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
resumen_ies["cluster"] = kmeans.fit_predict(X)

print("\n--- Resumen de clusters de instituciones ---")
print(resumen_ies.groupby("cluster")[["total_matricula", "total_articulos"]].mean())
print(resumen_ies.groupby("cluster").size())

plt.figure(figsize=(10, 6))
sns.scatterplot(data=resumen_ies, x="total_matricula", y="total_articulos", hue="cluster", palette="Set2", s=100)
plt.title("Clustering de instituciones: matricula vs produccion cientifica")
plt.xlabel("Matricula total")
plt.ylabel("Total articulos indexados")
plt.tight_layout()
plt.savefig("grafico_6_clusters.png", dpi=150)
plt.close()

from sklearn.linear_model import LinearRegression

X_anio = matricula_anual.index.values.reshape(-1, 1)
y_matricula = matricula_anual.values

modelo = LinearRegression()
modelo.fit(X_anio, y_matricula)

print(f"\nTendencia de matricula: {modelo.coef_[0]:.0f} estudiantes por año")
print(f"Prediccion matricula 2024: {modelo.predict([[2024]])[0]:.0f}")

df_academica_limpio.to_csv("academica_limpia_final.csv", index=False)
resumen_ies.to_csv("resumen_ies_clusters.csv", index=False)

print("\nAnalisis completo! Revisa los archivos .png generados y los CSV finales.")
