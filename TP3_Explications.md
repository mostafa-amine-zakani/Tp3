# TP3 – Analyse des Prix et Tendances des Vols : explications et réponses

> Cours : *Data Analysis – Data Driven Decision Making* (K. Dahi, 2026-2027).
> Toute la démarche s'appuie sur les notions du cours : chaîne décisionnelle (ETL → Stockage → Analyses → Restitution), méthode CRISP-DM, modélisation multidimensionnelle (faits / dimensions, modèle en étoile, cardinalité), types d'analyse (descriptive, diagnostique, prédictive, prescriptive) et Power BI (Power Query, DAX, tableaux de bord).
> Les chiffres cités ont été calculés sur les données fournies (`scripts/preparation.py`) ; vos visuels Power BI doivent afficher les mêmes valeurs.

---

## 0. Démarche générale (méthode CRISP-DM du cours)

| Phase CRISP-DM | Ce qu'on fait dans ce TP |
|---|---|
| **Compréhension du métier** | Agence de voyages en ligne : comprendre ce qui fait varier le prix d'un billet (compagnie, classe, distance/durée, escales, horaires, trajet) pour conseiller les clients et fixer des offres. |
| **Compréhension des données** | 3 fichiers : une table de vols (300 153 lignes) et deux tables descriptives (compagnies, météo). Voir §1. |
| **Préparation des données** | Étapes 1 et 2 : nettoyage et transformation dans **Power Query** (= phase **ETL** de la chaîne BI). |
| **Modélisation** | « Pour la BI : créer les jointures, relations entre les tables » → **modèle en étoile** + mesures DAX (étapes 2.3, 2.4, 4). |
| **Évaluation** | Vérifier que les indicateurs répondent aux questions métier (cohérence des chiffres, §3 et §5). |
| **Déploiement** | Rapport Power BI interactif (étape 3) + documentation (ce document). |

Les étapes du TP suivent aussi la **chaîne décisionnelle** du cours : *Extract* (import des 3 fichiers) → *Transform* (nettoyage, colonnes) → *Load* (chargement dans le modèle Power BI) → *Stockage* (modèle en étoile) → *Analyses* (DAX, IA) → *Restitution* (tableaux de bord).

---

## 1. Compréhension des données

| Fichier | Table Power BI | Lignes × colonnes | Rôle |
|---|---|---|---|
| `Airlines_flights_data.xlsx` (feuille `airlines_flights_data`) | `flight_prices` | 300 153 × 13 | **Table de faits** : un vol = une ligne, mesures `price`, `duration`, `stops` |
| `Airlines_infos.xlsx` (feuille `Sheet1`) | `Airlines_info` | 6 × 4 | **Dimension Compagnie** : code, nom, siège, pays |
| `Weather_conditions.xlsx` (feuille `Sheet2`) | `Weather_conditions` | 6 × 2 | **Dimension Météo** : identifiant, libellé |

Constats de qualité (étape « évaluer la qualité » du cours) :
- **Aucune valeur nulle** dans les trois fichiers ; l'étape 1.3 est quand même appliquée (règle de nettoyage robuste si les données sont rafraîchies).
- Les classes sont déjà `Economy` / `Business` ; le remplacement `Eco → Economy` est appliqué pour garantir l'uniformité (principe d'**intégration** du Data Warehouse : « H/F » et « Homme/Femme » → un seul format).
- `stops` est en texte (`zero`, `one`, `two_or_more`) : il faut le convertir en nombre (0, 1, 2) pour calculer le **taux de vols directs**.
- La colonne `Weather conditions` contient un identifiant (1 à 6) qui fait le lien avec la table météo.
- Le code compagnie s'appelle `IATA_CODDE` (faute de frappe) dans `Airlines_infos` → renommé `IATA_CODE`.
- 200 lignes sont identiques sur toutes les colonnes sauf `index`. Comme l'énoncé définit `index` comme **identifiant unique du vol**, ce sont des observations distinctes : elles sont conservées.
- ⚠️ **La colonne `days_left` (jours restants avant le départ) citée dans l'énoncé n'existe pas dans le fichier fourni.** Les analyses qui en dépendent (mesure « Corrélation Jours-Prix », forecasting par jours restants) sont expliquées mais remplacées par l'analyse la plus proche possible (voir §4 et §5).

---

## 2. Étape 1 – Nettoyage et préparation (Power Query)

Code complet : [`powerquery/01_flight_prices.m`](powerquery/01_flight_prices.m), [`02_Airlines_info.m`](powerquery/02_Airlines_info.m), [`03_Weather_conditions.m`](powerquery/03_Weather_conditions.m).

### 1.1 Importation des trois fichiers
**Réponse** : *Accueil > Obtenir des données > Classeur Excel*, une fois par fichier, puis *Transformer les données*. Dans le code M : `Excel.Workbook(File.Contents(...))`, choix de la feuille, puis `Table.PromoteHeaders` (première ligne → en-têtes).
**Pourquoi** : c'est le *E* de l'ETL ; les données viennent de sources séparées (« distribuées, hétérogènes ») qu'on veut intégrer dans un seul modèle.

### 1.2 Remplacement de valeurs
**Réponse** : sur `class`, *Transformer > Remplacer les valeurs* : `Eco` → `Economy` (et `Bus` → `Business`), après `Rogner` + `Majuscule à chaque mot` pour éliminer espaces et casse différente. Mêmes outils pour :
- `stops` : `zero` → 0, `one` → 1, `two_or_more` → 2 ;
- `departure_time` / `arrival_time` : `Early_Morning` → `Early Morning` (libellés lisibles) ;
- `Airline` : `Air_India` → `Air India`.

**Pourquoi** : uniformiser les valeurs est le principe d'**intégration** vu en cours ; sinon « Eco » et « Economy » formeraient deux catégories différentes dans les graphiques.

### 1.3 Gestion des valeurs nulles
**Réponse** : *Accueil > Supprimer les lignes > Supprimer les lignes vides* / filtre « ≠ null » sur les colonnes indispensables (`price`, `duration`, villes, `class`). Résultat : 0 ligne supprimée (aucun null), 300 153 lignes conservées.
**Pourquoi supprimer plutôt que remplacer** : un vol sans prix ni durée ne peut pas entrer dans une moyenne ; le remplacer par une valeur inventée fausserait les indicateurs.

### 1.4 Conversion des types
**Réponse** : *Transformer > Type de données* :

| Colonne | Type |
|---|---|
| `Arrival_date` | Date |
| `source_city`, `destination_city`, `class` | Texte |
| `price`, `duration` | Nombre décimal |
| `stops`, `flight_id`, `Weather_ID` | Nombre entier |

**Pourquoi** : le type conditionne ce que Power BI peut faire (moyenne sur un nombre, hiérarchie Année/Mois sur une date, relation entre deux colonnes de même type). C'est une **métadonnée** au sens du cours (« Table SQL : nom des colonnes, types, clés »).

### 1.5 Transformation de texte
**Réponse** : sélectionner `source_city` et `destination_city` > *Transformer > Format > MAJUSCULES* (`Text.Upper`). Ex. `Delhi` → `DELHI`.

---

## 3. Étape 2 – Transformation des données

### 2.1 Colonne conditionnelle `Prix_Categorie`
**Réponse** : *Ajouter une colonne > Colonne conditionnelle* :
- Si `price` < 10000 → « Bas »
- Sinon si `price` < 20000 → « Moyen »
- Sinon → « Élevé »

Résultat : **Bas 173 653 vols (57,9 %)**, **Moyen 31 602 (10,5 %)**, **Élevé 94 898 (31,6 %)**. La catégorie « Élevé » correspond presque entièrement à la classe Business (prix moyen 52 540).

### 2.2 Colonne personnalisée `Duree_Minutes`
**Réponse** : *Ajouter une colonne > Colonne personnalisée* : `= [duration] * 60`. Ex. 2,17 h → 130,2 min.

Deux colonnes utiles aux visuels sont ajoutées de la même façon :
- `Trajet = [source_city] & " → " & [destination_city]` (vue géographique) ;
- `Ordre_Depart` (1 = Early Morning … 6 = Late Night) pour trier les créneaux dans l'ordre chronologique et calculer la corrélation heure/prix.

### 2.3 Création des relations (modèle en étoile)
**Réponse** : *Vue Modèle*, glisser les colonnes :

| De (côté *) | Vers (côté 1) | Cardinalité | Filtrage |
|---|---|---|---|
| `flight_prices[OP_CARRIER]` | `Airlines_info[IATA_CODE]` | Plusieurs à un (* → 1) | Unique |
| `flight_prices[Weather_ID]` | `Weather_conditions[Weather_ID]` | Plusieurs à un (* → 1) | Unique |
| `flight_prices[Arrival_date]` | `Dim_Date[Date]` | Plusieurs à un (* → 1) | Unique |

```
                 Airlines_info (1)
                       │
 Weather_conditions ─ flight_prices (*) ─ Dim_Date
        (1)            table de faits        (1)
```

**Explication avec le cours** :
- **Faits** = ce qu'on mesure (quantitatif) : `price`, `duration`, `stops` → `flight_prices` est la **table de faits**.
- **Dimensions** = axes d'analyse qui contextualisent les faits : compagnie, météo, temps (et classe/ville, restées dans la table de faits car ce sont des attributs simples).
- C'est un **modèle en étoile** : une table de faits au centre, des dimensions dé-normalisées **sans lien entre elles** → peu de jointures, navigation facile.
- **Cardinalité 1 → \*** : une compagnie (une ligne de `Airlines_info`) correspond à plusieurs vols (ex. Vistara → 127 859 vols), comme « un client → plusieurs commandes ».
- `Dim_Date` est la **dimension Temps** du cours (jour, mois, trimestre, année = **hiérarchie** de granularité). Code : [`dax/02_tables_calculees.dax`](dax/02_tables_calculees.dax).

### 2.4 Colonne calculée DAX `Variation_Prix`
**Réponse** (*Outils de table > Nouvelle colonne*, code dans [`dax/01_colonne_calculee.dax`](dax/01_colonne_calculee.dax)) :

```DAX
Variation_Prix =
DIVIDE (
    flight_prices[price],
    CALCULATE ( AVERAGE ( flight_prices[price] ), ALL ( flight_prices ) )
) - 1
```

**Explication** : pour chaque vol *i*, `Variation_Prix = pᵢ / p̄ − 1` où *p̄* est le prix moyen de **tous** les vols (20 890). `ALL` supprime tous les filtres pour que la moyenne soit globale ; `DIVIDE` évite une erreur de division par zéro.
Ex. : premier vol à 5 953 → 5 953 / 20 890 − 1 = **−71,5 %** (71,5 % moins cher que la moyenne). Les valeurs vont de −94,7 % à +489 %.

**Colonne calculée ou mesure ?** (notion du TP1) Une colonne calculée est évaluée **ligne par ligne** au chargement et stockée ; une mesure est calculée **à la volée** selon les filtres du visuel. `Variation_Prix` décrit chaque vol → colonne. Les indicateurs agrégés de l'étape 4 → mesures.

---

## 4. Étape 4 – DAX et mesures

Code complet : [`dax/03_mesures.dax`](dax/03_mesures.dax).

| Mesure | Formule | Valeur globale | Explication |
|---|---|---|---|
| **Prix Moyen** | `AVERAGE(flight_prices[price])` | **20 890** | Moyenne des prix dans le contexte de filtre (compagnie, classe…). |
| **Durée Moyenne** | `AVERAGE(flight_prices[duration])` | **12,22 h** | Idem sur la durée. |
| **Prix Moyen par Classe** | `CALCULATE([Prix Moyen], ALLEXCEPT(flight_prices, flight_prices[class]))` | Economy 6 572 / Business 52 540 | Prix moyen de la classe, en ignorant les autres filtres. ⚠️ La formule de l'énoncé `CALCULATE(AVERAGE(...), flight_prices[class])` n'est pas valide : le 2ᵉ argument de CALCULATE doit être un filtre (condition ou table), pas une colonne seule. `ALLEXCEPT` réalise l'intention. |
| **Écart Prix** | `[Prix Moyen] - CALCULATE([Prix Moyen], ALL(flight_prices))` | ex. Vistara +9 507, AirAsia −16 799 | Écart d'un groupe par rapport à la moyenne générale. |
| **Corrélation Jours-Prix** | *non calculable* | – | `days_left` absent du fichier fourni. La formule correcte (si la colonne existait) est donnée en commentaire dans le fichier DAX. Remplacée par **Corrélation Heure-Prix** (r = 0,021) et **Corrélation Durée-Prix** (r = 0,204), calculées avec la formule de Pearson r = cov(x,y)/(σx·σy). |
| **Taux Vols Directs** | `DIVIDE(COUNTROWS(FILTER(flight_prices, flight_prices[stops]=0)), COUNTROWS(flight_prices))` | **12,0 %** | Part des vols sans escale ; possible car `stops` a été converti en nombre (étape 1.2). |

Mesures complémentaires : `Nombre de Vols`, `Escales Moyennes` (pour le field parameter), `Nb Anomalies` (3 683), `Couleur Prix` (mise en forme conditionnelle).

---

## 5. Étape 3 – Visualisation dans Power BI

### 5.1 Vue générale du marché (analyse **descriptive** : « que s'est-il passé ? »)
- **Cartes** : `Prix Moyen` (20 890), `Durée Moyenne` (12,22 h), `Taux Vols Directs` (12 %), `Nombre de Vols` (300 153).
- **Histogramme groupé** : axe `Airlines_info[Airline]`, valeur `Prix Moyen`, couleur `Couleur Prix`.

| Compagnie | Prix moyen | Vols |
|---|---|---|
| AirAsia | 4 091 | 16 098 |
| Indigo | 5 324 | 43 120 |
| GO FIRST | 5 652 | 23 173 |
| SpiceJet | 6 179 | 9 011 |
| Air India | 23 507 | 80 892 |
| Vistara | 30 397 | 127 859 |

- **Histogramme** : axe `class`, valeur `Prix Moyen` → Business **52 540** vs Economy **6 572** (×8).
- **Jauge ou carte** : `Durée Moyenne`.

### 5.2 Vue temporelle
- **Répartition par heure de départ** : histogramme, axe `departure_time` (trié par `Ordre_Depart`), valeur `Nombre de Vols`. Pics : Morning 71 146, Early Morning 66 790, Evening 65 102 ; Late Night seulement 1 306.
- **Répartition par date d'arrivée** : courbe, axe `Dim_Date` (hiérarchie Année > Mois), valeur `Nombre de Vols`. ≈ 51 500 vols par an de 2020 à 2024, 42 777 en 2025 (données jusqu'au 31/10/2025).
- **Corrélation heure / prix** : graphique en nuage de points ou histogramme `departure_time` × `Prix Moyen` avec légende `class`, plus une carte `Corrélation Heure-Prix`.
  Lecture : r = 0,02 → **pas de lien linéaire** entre l'heure et le prix. Seul « Late Night » est nettement moins cher (9 295, et 4 785 en Economy).

### 5.3 Vue géographique
- **Matrice** : lignes `source_city`, colonnes `destination_city`, valeurs `Nombre de Vols` (mise en forme conditionnelle « échelle de couleurs ») — ou visuel **Ruban/Sankey** sur `Trajet`.
- Trajets les plus fréquents : DELHI → MUMBAI (15 289), MUMBAI → DELHI (14 809), DELHI → BANGALORE (14 012).
- Tendance durée / escales par ville de départ : KOLKATA a la durée moyenne la plus longue (13,25 h, 0,97 escale), DELHI la plus courte (11,52 h).

### 5.4 Vue analytique : compagnies sur le segment économique
Filtre de page `class = Economy`, matrice `Airline` × (`Prix Moyen`, `Durée Moyenne`, `Taux Vols Directs`, `Nombre de Vols`) :

| Compagnie | Prix moyen | Durée moy. | Vols directs |
|---|---|---|---|
| AirAsia | 4 091 | 8,9 h | 15 % |
| Indigo | 5 324 | 5,8 h | 26 % |
| GO FIRST | 5 652 | 8,8 h | 14 % |
| SpiceJet | 6 179 | 12,6 h | 27 % |
| Air India | 7 314 | 16,1 h | 7 % |
| Vistara | 7 807 | 13,4 h | 8 % |

Lecture : en Economy, **Indigo** offre le meilleur compromis (prix bas, vols les plus courts) ; Vistara et Air India restent les plus chères avec plus d'escales.

### 5.5 Outils avancés
- **Bookmarks (signets)** : construire deux groupes de visuels (compagnies / géographie). *Affichage > Signets* + *Affichage > Sélection* : masquer un groupe, *Ajouter* un signet « Vue compagnies » ; inverser, signet « Vue géographique ». *Insérer > Boutons* → action *Signet* sur chaque bouton.
- **Tooltips personnalisés** : nouvelle page, *Format > Informations sur la page > Autoriser l'utilisation comme info-bulle*, taille « Info-bulle » ; y placer les cartes `Durée Moyenne` et `Escales Moyennes`. Dans les visuels principaux : *Format > Info-bulle > Type : Page de rapport* → cette page.
- **Mise en forme conditionnelle** : sur l'histogramme par compagnie, *Format > Barres > Couleur > fx > Valeur du champ* = mesure `Couleur Prix` (vert < 10 000, orange < 20 000, rouge sinon), cohérente avec `Prix_Categorie`. Dans une table, on peut aussi utiliser la colonne `Prix_Categorie`.
- **Field Parameters** : *Modélisation > Nouveau paramètre > Champs* → `Prix Moyen`, `Durée Moyenne`, `Escales Moyennes`, nom `Indicateur`, cocher « Ajouter un segment ». Mettre `Indicateur` en valeur d'un graphique : le segment bascule entre prix, durée et escales.

---

## 6. Étape 5 – Analyse avancée et IA

### Forecasting (analyse **prédictive** : « que va-t-il se passer ? »)
L'énoncé demande de prédire le prix selon le nombre de jours avant le départ, mais **`days_left` n'est pas dans les données**. La prévision disponible dans Power BI se fait sur un axe temporel continu : courbe `Dim_Date[Date]` (niveau Mois) × `Prix Moyen`, puis *Volet Analytique > Prévision* (longueur 6 mois, intervalle de confiance 95 %).
Résultat attendu : prix moyen annuel quasi constant (20 695 à 21 103 entre 2020 et 2025) → la série est **stationnaire** (notion du cours : propriétés statistiques constantes dans le temps), la prévision est une droite horizontale autour de 20 900.

### Anomaly Detection
- **Règle statistique** (dans Power Query, colonne `Est_Anomalie`) : un vol est anormal si son prix dépasse **moyenne + 3 écarts-types de sa classe** (on compare chaque vol à sa propre classe, sinon tous les Business seraient « anormaux »).
  Seuils : Economy 17 803, Business 91 448. **3 683 vols anormaux** (3 317 Economy, 366 Business), surtout Vistara (1 849) et Air India (1 112).
- **Outil Power BI** : sur une courbe `Dim_Date` × `Prix Moyen`, *Volet Analytique > Rechercher des anomalies* ; cliquer un point marqué donne les facteurs explicatifs.

### Copilot / AI Insights : « Explique-moi pourquoi le prix moyen augmente sur certaines compagnies »
Réponse argumentée (analyse **diagnostique** : « pourquoi ? ») — visuel *Influenceurs clés* (cible `price`, facteurs `class`, `stops`, `duration`, `Airline`) ou *Arbre de décomposition* :
1. **Effet de mix de classes** : seules **Vistara** et **Air India** vendent des billets **Business** (93 487 vols au total). Leur prix moyen global (30 397 et 23 507) est tiré vers le haut par cette classe ×8 plus chère.
2. **Même en Economy** elles restent les plus chères (7 807 et 7 314) : ce sont les vols **les plus longs** (13–16 h) et avec **le moins de vols directs** (7–8 %) ; la corrélation durée-prix est positive (r = 0,20).
3. Les **escales** comptent : prix moyen 9 376 pour un vol direct contre 22 901 avec une escale.
4. La météo n'explique rien (prix moyen 20 716 à 21 011 quelle que soit la condition).

### Clusterisation (segmentation)
Dans Power BI : nuage de points `Durée Moyenne` (X) × `Prix Moyen` (Y), détails `flight` ; *… > Rechercher automatiquement des clusters* (4 clusters) ; ajouter `class` en légende pour interpréter.
Profils obtenus (k-means, 4 groupes, sur classe / durée / prix) :

| Profil | Vols | Classe | Durée moy. | Prix moy. |
|---|---|---|---|---|
| Business | 93 027 | Business | 13,8 h | 52 716 |
| Eco court-courrier | 89 696 | Economy | 5,3 h | 5 588 |
| Eco durée moyenne | 82 198 | Economy | 12,9 h | 7 078 |
| Eco très long (escales) | 35 232 | Economy | 24,3 h | 8 035 |

---

## 7. Conclusion (analyse **prescriptive** : « que faire ? »)
- Le prix dépend d'abord de la **classe**, puis de la **compagnie**, de la **durée** et des **escales** ; l'heure de départ et la météo ont peu d'effet.
- Pour un client sensible au prix : recommander AirAsia/Indigo en Economy, et les vols directs (moins chers et plus courts).
- Pour l'agence : mettre en avant les vols « Late Night » (les moins chers) et surveiller les 3 683 prix anormaux (alerte commerciale).
- Limite : sans `days_left`, l'impact de l'anticipation de la réservation sur le prix ne peut pas être mesuré ; il faudrait demander cette donnée à la source.

---

## 8. Livrables
- **Fichier .pbix** : à construire en suivant [README.md](README.md) (tout le code Power Query et DAX est fourni et prêt à coller), puis *Fichier > Enregistrer sous* `TP3.pbix`.
- **Rapport interactif** : pages 5.1 à 5.4 + signets, info-bulles, mise en forme conditionnelle et paramètre de champs.
- **Documentation** : ce fichier (méthode CRISP-DM, explication des indicateurs, règles de nettoyage — les éléments de documentation du déploiement cités dans le cours).
