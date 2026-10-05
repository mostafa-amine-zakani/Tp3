// Requête : Airlines_info (dimension compagnie)
let
    Source      = Excel.Workbook(File.Contents(DossierTP3 & "\data\raw\Airlines_infos.xlsx"), null, true),
    Feuille     = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data],
    EnTetes     = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Renommees   = Table.RenameColumns(EnTetes, {{"IATA_CODDE", "IATA_CODE"}, {"Headquarters Location", "Headquarters_Location"}}),
    Nettoyees   = Table.TransformColumns(Renommees, {
                        {"IATA_CODE", Text.Trim, type text},
                        {"Airline", each Text.Replace(Text.Trim(_), "_", " "), type text},
                        {"Headquarters_Location", Text.Trim, type text},
                        {"Country", Text.Trim, type text}}),
    SansNulls   = Table.SelectRows(Nettoyees, each [IATA_CODE] <> null and [IATA_CODE] <> ""),
    SansDoublons = Table.Distinct(SansNulls, {"IATA_CODE"})
in
    SansDoublons
