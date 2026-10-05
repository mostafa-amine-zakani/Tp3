// Requête : Weather_conditions (dimension météo)
let
    Source      = Excel.Workbook(File.Contents(DossierTP3 & "\data\raw\Weather_conditions.xlsx"), null, true),
    Feuille     = Source{[Item = "Sheet2", Kind = "Sheet"]}[Data],
    EnTetes     = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Renommees   = Table.RenameColumns(EnTetes, {{"ID", "Weather_ID"}}),
    Types       = Table.TransformColumnTypes(Renommees, {{"Weather_ID", Int64.Type}, {"Weather_condition", type text}}),
    SansNulls   = Table.SelectRows(Types, each [Weather_ID] <> null),
    SansDoublons = Table.Distinct(SansNulls, {"Weather_ID"})
in
    SansDoublons
