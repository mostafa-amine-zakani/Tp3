# TP3 – Analyse des Prix et Tendances des Vols (Power BI)

Solution du TP3 du cours *Data Analysis – Data Driven Decision Making*.
👉 **Explications de chaque exercice et réponses : [TP3_Explications.md](TP3_Explications.md)**

## Contenu du dépôt

| Dossier / fichier | Contenu |
|---|---|
| `TP3_enonce.pdf` | Énoncé du TP |
| `TP3_Explications.md` | Texte explicatif : démarche, réponses, chiffres attendus |
| `data/raw/` | Les 3 fichiers Excel d'origine |
| `data/clean/` | Données déjà nettoyées (CSV UTF-8) : `flight_prices.csv`, `Airlines_info.csv`, `Weather_conditions.csv` |
| `powerquery/` | Code M (Power Query) de chaque requête, à coller dans l'Éditeur avancé |
| `dax/` | Colonne calculée, tables calculées (Dim_Date, paramètre de champs) et mesures DAX |
| `TP3.pbip`, `TP3.SemanticModel/`, `TP3.Report/` | Projet Power BI prêt à ouvrir |
| `scripts/preparation.py` | Même nettoyage en Python : régénère `data/clean/` et affiche les chiffres de contrôle |

## Ouvrir directement dans Power BI (projet `TP3.pbip`)

Le dépôt contient un **projet Power BI** prêt à ouvrir : `TP3.pbip` (modèle avec les 4 tables, les relations, la colonne `Variation_Prix`, toutes les mesures, et 4 pages de rapport : Vue générale, Vue temporelle, Vue géographique, Vue analytique).

Dans PowerShell :

```powershell
git clone https://github.com/mostafa-amine-zakani/Tp3.git C:\Tp3
Start-Process C:\Tp3\TP3.pbip
```

Puis dans Power BI Desktop : **Accueil > Actualiser** (le projet ne stocke pas les données, il lit les CSV de `C:\Tp3\data\clean`).

- Si le dépôt est cloné ailleurs que `C:\Tp3` : *Transformer les données > Modifier les paramètres* > `DossierTP3` = votre chemin, puis Actualiser.
- Si Power BI refuse d'ouvrir le `.pbip` : *Fichier > Options et paramètres > Options > Fonctionnalités en préversion* > cocher **Enregistrement du projet Power BI (.pbip)**, redémarrer Power BI.
- Pour obtenir un `.pbix` : *Fichier > Enregistrer sous* > type *Fichier Power BI (.pbix)*.
- À ajouter à la main (non générables en fichier texte de façon fiable) : signets, info-bulles personnalisées, paramètre de champs `Indicateur`, couleur conditionnelle (§5.5 de `TP3_Explications.md`).
- Le projet est généré par `scripts/generer_pbip.py`.

## Construire le fichier Power BI à la main (pas à pas)

1. **Récupérer le dépôt** : `git clone https://github.com/mostafa-amine-zakani/Tp3.git` (ou `git pull` si déjà cloné).
2. Ouvrir **Power BI Desktop** > *Accueil > Transformer les données* (ouvre Power Query).
3. **Paramètre du dossier** : *Gérer les paramètres > Nouveau paramètre* → Nom `DossierTP3`, Type *Texte*, Valeur = chemin du dossier cloné (ex. `C:\Users\moi\Documents\Tp3`, sans `\` final).
4. **Requêtes** : pour chaque fichier de `powerquery/` : *Nouvelle source > Requête vide*, *Éditeur avancé*, coller le code, renommer la requête :
   - `01_flight_prices.m` → `flight_prices` (refait toutes les étapes 1 et 2 depuis l'Excel brut)
     *ou, plus rapide,* `04_option_CSV_flight_prices.m` → `flight_prices` (charge le CSV déjà nettoyé)
   - `02_Airlines_info.m` → `Airlines_info`
   - `03_Weather_conditions.m` → `Weather_conditions`
5. *Fermer et appliquer*.
6. **Tables calculées** (*Modélisation > Nouvelle table*) : `Dim_Date` depuis `dax/02_tables_calculees.dax`, puis *Marquer comme table de dates*.
7. **Relations** (*Vue Modèle*), toutes en plusieurs-à-un, filtrage unique :
   - `flight_prices[OP_CARRIER]` → `Airlines_info[IATA_CODE]`
   - `flight_prices[Weather_ID]` → `Weather_conditions[Weather_ID]`
   - `flight_prices[Arrival_date]` → `Dim_Date[Date]`
8. **Colonne calculée** `Variation_Prix` (`dax/01_colonne_calculee.dax`) et **mesures** (`dax/03_mesures.dax`) dans la table `flight_prices` (une *Nouvelle mesure* par bloc).
9. **Tri** : `departure_time` > *Trier par colonne* > `Ordre_Depart`.
10. **Paramètre de champs** `Indicateur` : *Modélisation > Nouveau paramètre > Champs* (voir `dax/02_tables_calculees.dax`).
11. Construire les pages du rapport décrites au §5 de `TP3_Explications.md`, puis *Enregistrer sous* `TP3.pbix`.

> Remarque : la colonne `days_left` citée dans l'énoncé n'existe pas dans les données fournies ; voir `TP3_Explications.md` §1 et §4.
