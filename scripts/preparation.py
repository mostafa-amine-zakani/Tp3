"""
TP3 - Analyse des Prix et Tendances des Vols
Script de préparation (équivalent Python des étapes Power Query 1 et 2).

Il sert à :
  1. produire les fichiers nettoyés de data/clean/ (chargeables directement dans Power BI),
  2. calculer les chiffres cités dans TP3_Explications.md, pour vérifier les visuels Power BI.

Usage :  python scripts/preparation.py      (depuis la racine du dépôt)
Dépendances : pandas, openpyxl, numpy
"""
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
CLEAN = ROOT / "data" / "clean"
CLEAN.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------- Étape 1.1 : importation
flights = pd.read_excel(RAW / "Airlines_flights_data.xlsx", sheet_name="airlines_flights_data")
airlines = pd.read_excel(RAW / "Airlines_infos.xlsx", sheet_name="Sheet1")
weather = pd.read_excel(RAW / "Weather_conditions.xlsx", sheet_name="Sheet2")

# Renommage des colonnes pour des noms sans espaces (plus simples en DAX)
flights = flights.rename(columns={"index": "flight_id", "Weather conditions": "Weather_ID"})
airlines = airlines.rename(columns={"IATA_CODDE": "IATA_CODE", "Headquarters Location": "Headquarters_Location"})
weather = weather.rename(columns={"ID": "Weather_ID"})

# ---------------------------------------------------------------- Étape 1.2 : remplacement de valeurs
class_map = {"eco": "Economy", "economy": "Economy", "business": "Business", "bus": "Business"}
flights["class"] = flights["class"].str.strip().str.lower().map(class_map).fillna(flights["class"])
# stops texte -> nombre d'escales (nécessaire pour la mesure Taux Vols Directs : stops = 0)
flights["stops"] = flights["stops"].map({"zero": 0, "one": 1, "two_or_more": 2})
# remplacement des "_" pour des libellés lisibles
for c in ["departure_time", "arrival_time"]:
    flights[c] = flights[c].str.replace("_", " ")
airlines["Airline"] = airlines["Airline"].str.replace("_", " ")

# ---------------------------------------------------------------- Étape 1.3 : valeurs nulles
nulls_before = int(flights.isna().sum().sum())
flights = flights.dropna(subset=["price", "duration", "source_city", "destination_city", "class"])

# ---------------------------------------------------------------- Étape 1.4 : types
flights["Arrival_date"] = pd.to_datetime(flights["Arrival_date"]).dt.date
flights["price"] = flights["price"].astype(float)
flights["duration"] = flights["duration"].astype(float)
for c in ["source_city", "destination_city", "class"]:
    flights[c] = flights[c].astype(str)

# ---------------------------------------------------------------- Étape 1.5 : majuscules
for c in ["source_city", "destination_city"]:
    flights[c] = flights[c].str.upper()

# ---------------------------------------------------------------- Étape 2.1 / 2.2 : colonnes
flights["Prix_Categorie"] = np.select(
    [flights["price"] < 10000, flights["price"] < 20000], ["Bas", "Moyen"], default="Élevé"
)
flights["Duree_Minutes"] = (flights["duration"] * 60).round(1)
flights["Trajet"] = flights["source_city"] + " → " + flights["destination_city"]
# ordre chronologique des créneaux (tri des visuels + corrélation heure / prix)
ordre = {"Early Morning": 1, "Morning": 2, "Afternoon": 3, "Evening": 4, "Night": 5, "Late Night": 6}
flights["Ordre_Depart"] = flights["departure_time"].map(ordre)

# (Étape 2.4 : Variation_Prix est une colonne calculée DAX, créée dans Power BI -> pas exportée ici)

# ---------------------------------------------------------------- Étape 5 : anomalies (règle 3 sigma par classe)
g = flights.groupby("class")["price"]
flights["Seuil_Anomalie"] = (g.transform("mean") + 3 * g.transform("std")).round(2)
flights["Est_Anomalie"] = np.where(flights["price"] > flights["Seuil_Anomalie"], "Oui", "Non")

# ---------------------------------------------------------------- Export
flights.to_csv(CLEAN / "flight_prices.csv", index=False, encoding="utf-8")
airlines.to_csv(CLEAN / "Airlines_info.csv", index=False, encoding="utf-8")
weather.to_csv(CLEAN / "Weather_conditions.csv", index=False, encoding="utf-8")

# ---------------------------------------------------------------- Résultats (contrôle)
f = flights.merge(airlines, left_on="OP_CARRIER", right_on="IATA_CODE", how="left").merge(
    weather, on="Weather_ID", how="left"
)
pd.set_option("display.width", 160)
print("Lignes :", len(f), "| valeurs nulles avant nettoyage :", nulls_before)
print("Doublons (hors identifiant) :", int(flights.drop(columns="flight_id").duplicated().sum()))
print("\nPrix moyen global : %.0f | Durée moyenne : %.2f h | Taux vols directs : %.2f %%"
      % (f.price.mean(), f.duration.mean(), 100 * (f.stops == 0).mean()))
print("\nPrix moyen par compagnie :\n", f.groupby("Airline").price.agg(["mean", "count"]).round(0).sort_values("mean"))
print("\nPrix moyen par classe :\n", f.groupby("class").price.mean().round(0))
print("\nÉconomie - comparaison des compagnies :\n",
      f[f["class"] == "Economy"].groupby("Airline").agg(prix=("price", "mean"), duree=("duration", "mean"),
                                                         directs=("stops", lambda s: (s == 0).mean()),
                                                         vols=("price", "size")).round(2).sort_values("prix"))
print("\nPrix moyen par escales :\n", f.groupby("stops").agg(prix=("price", "mean"), duree=("duration", "mean"), n=("price", "size")).round(1))
print("\nVols par heure de départ :\n", f.groupby("departure_time").agg(vols=("price", "size"), prix=("price", "mean")).round(0))
print("\nPrix moyen par heure de départ et classe :\n", f.pivot_table(index="departure_time", columns="class", values="price", aggfunc="mean").round(0))
print("\nPrix_Categorie :\n", f.Prix_Categorie.value_counts())
print("\nVols par année d'arrivée :\n", pd.to_datetime(f.Arrival_date).dt.year.value_counts().sort_index())
print("\nPrix moyen par météo :\n", f.groupby("Weather_condition").price.mean().round(0))
f["Variation_Prix"] = f["price"] / f["price"].mean() - 1
print("\nVariation_Prix (min / max) :", round(f.Variation_Prix.min(), 3), round(f.Variation_Prix.max(), 3))
print("Corrélation heure de départ (ordre) / prix : %.3f" % f.Ordre_Depart.corr(f.price))
print("Corrélation durée / prix : %.3f | escales / prix : %.3f" % (f.duration.corr(f.price), f.stops.corr(f.price)))
print("Prix moyen par année d'arrivée :\n", f.groupby(pd.to_datetime(f.Arrival_date).dt.year).price.mean().round(0))
print("\nTop trajets (nb vols) :\n", f.Trajet.value_counts().head(5))
print("\nDurée moyenne par ville de départ :\n", f.groupby("source_city").agg(duree=("duration", "mean"), escales=("stops", "mean")).round(2))
print("\nAnomalies :\n", f.groupby("class").agg(seuil=("Seuil_Anomalie", "first"), anomalies=("Est_Anomalie", lambda s: (s == "Oui").sum())))
print("\nAnomalies par compagnie :\n", f[f.Est_Anomalie == "Oui"].Airline.value_counts())

# Clusterisation k-means (k = 4) sur classe, durée, prix standardisés - illustration de l'étape 5
X = np.column_stack([(f["class"] == "Business").astype(float), f.duration, f.price])
Xs = (X - X.mean(0)) / X.std(0)
rng = np.random.default_rng(0)
C = Xs[rng.choice(len(Xs), 4, replace=False)]
for _ in range(50):
    lab = np.argmin(((Xs[:, None, :] - C[None]) ** 2).sum(-1), axis=1)
    C = np.array([Xs[lab == k].mean(0) for k in range(4)])
f["cluster"] = lab
print("\nProfils des clusters :\n", f.groupby("cluster").agg(vols=("price", "size"), part_business=("class", lambda s: (s == "Business").mean()),
                                                           duree=("duration", "mean"), prix=("price", "mean")).round(2))
