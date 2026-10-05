// OPTION RAPIDE (à la place de 01_flight_prices.m) : charger le CSV déjà nettoyé par scripts/preparation.py.
// Il contient exactement les colonnes produites par les étapes 1 et 2 (sauf Variation_Prix, créée en DAX).
let
    Source  = Csv.Document(File.Contents(DossierTP3 & "\data\clean\flight_prices.csv"),
                  [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Types   = Table.TransformColumnTypes(EnTetes, {
                  {"flight_id", Int64.Type}, {"OP_CARRIER", type text}, {"flight", type text},
                  {"source_city", type text}, {"Arrival_date", type date}, {"departure_time", type text},
                  {"stops", Int64.Type}, {"arrival_time", type text}, {"destination_city", type text},
                  {"class", type text}, {"duration", type number}, {"price", type number}, {"Weather_ID", Int64.Type},
                  {"Prix_Categorie", type text}, {"Duree_Minutes", type number}, {"Trajet", type text},
                  {"Ordre_Depart", Int64.Type}, {"Seuil_Anomalie", type number}, {"Est_Anomalie", type text}}, "en-US")
in
    Types
