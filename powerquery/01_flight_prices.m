// Requête : flight_prices (table de faits)
// Accueil > Nouvelle source > Requête vide > Éditeur avancé > coller ce code > renommer la requête "flight_prices"
let
    // ---------- Étape 1.1 : importation
    Source      = Excel.Workbook(File.Contents(DossierTP3 & "\data\raw\Airlines_flights_data.xlsx"), null, true),
    Feuille     = Source{[Item = "airlines_flights_data", Kind = "Sheet"]}[Data],
    EnTetes     = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Renommees   = Table.RenameColumns(EnTetes, {{"index", "flight_id"}, {"Weather conditions", "Weather_ID"}}),

    // ---------- Étape 1.2 : remplacement de valeurs (uniformiser les classes, les escales, les créneaux)
    ClasseTrim  = Table.TransformColumns(Renommees, {{"class", each Text.Proper(Text.Trim(_)), type text}}),
    ClasseEco   = Table.ReplaceValue(ClasseTrim, "Eco", "Economy", Replacer.ReplaceValue, {"class"}),
    ClasseBus   = Table.ReplaceValue(ClasseEco, "Bus", "Business", Replacer.ReplaceValue, {"class"}),
    Escales     = Table.TransformColumns(ClasseBus, {{"stops", each
                        if _ = "zero" then 0 else if _ = "one" then 1 else if _ = "two_or_more" then 2 else null,
                        Int64.Type}}),
    Creneaux    = Table.ReplaceValue(Escales, "_", " ", Replacer.ReplaceText, {"departure_time", "arrival_time"}),

    // ---------- Étape 1.3 : valeurs nulles -> suppression des lignes inexploitables
    SansNulls   = Table.SelectRows(Creneaux, each [price] <> null and [duration] <> null
                        and [source_city] <> null and [destination_city] <> null and [class] <> null),

    // ---------- Étape 1.4 : conversion des types
    Types       = Table.TransformColumnTypes(SansNulls, {
                        {"flight_id", Int64.Type}, {"OP_CARRIER", type text}, {"flight", type text},
                        {"source_city", type text}, {"Arrival_date", type date}, {"departure_time", type text},
                        {"arrival_time", type text}, {"destination_city", type text}, {"class", type text},
                        {"duration", type number}, {"price", type number}, {"Weather_ID", Int64.Type}}),

    // ---------- Étape 1.5 : noms de villes en majuscules
    Majuscules  = Table.TransformColumns(Types, {{"source_city", Text.Upper, type text}, {"destination_city", Text.Upper, type text}}),

    // ---------- Étape 2.1 : colonne conditionnelle Prix_Categorie
    PrixCat     = Table.AddColumn(Majuscules, "Prix_Categorie", each
                        if [price] < 10000 then "Bas" else if [price] < 20000 then "Moyen" else "Élevé", type text),

    // ---------- Étape 2.2 : colonne personnalisée Duree_Minutes
    DureeMin    = Table.AddColumn(PrixCat, "Duree_Minutes", each [duration] * 60, type number),

    // Colonnes utiles pour les visuels : trajet (vue géographique) et ordre des créneaux (tri + corrélation)
    Trajet      = Table.AddColumn(DureeMin, "Trajet", each [source_city] & " → " & [destination_city], type text),
    Ordre       = Table.AddColumn(Trajet, "Ordre_Depart", each
                        if [departure_time] = "Early Morning" then 1 else if [departure_time] = "Morning" then 2
                        else if [departure_time] = "Afternoon" then 3 else if [departure_time] = "Evening" then 4
                        else if [departure_time] = "Night" then 5 else 6, Int64.Type),

    // ---------- Étape 5 : détection d'anomalies (prix > moyenne + 3 écarts-types de sa classe)
    Stats       = Table.Group(Ordre, {"class"}, {
                        {"Moyenne", each List.Average([price]), type number},
                        {"EcartType", each List.StandardDeviation([price]), type number}}),
    Fusion      = Table.NestedJoin(Ordre, {"class"}, Stats, {"class"}, "S", JoinKind.LeftOuter),
    Developpe   = Table.ExpandTableColumn(Fusion, "S", {"Moyenne", "EcartType"}),
    Seuil       = Table.AddColumn(Developpe, "Seuil_Anomalie", each [Moyenne] + 3 * [EcartType], type number),
    Anomalie    = Table.AddColumn(Seuil, "Est_Anomalie", each if [price] > [Seuil_Anomalie] then "Oui" else "Non", type text),
    Resultat    = Table.RemoveColumns(Anomalie, {"Moyenne", "EcartType"})
in
    Resultat
