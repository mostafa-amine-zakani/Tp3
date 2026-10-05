"""
Génère le projet Power BI TP3.pbip (modèle sémantique model.bim + rapport au format PBIR).

Le projet charge les CSV de data/clean depuis GitHub (paramètre DossierTP3 = URL du dépôt) ;
on peut remplacer DossierTP3 par le chemin d'un clone local (ex. C:\\Tp3). Ouvrir TP3.pbip dans Power BI Desktop puis cliquer sur « Actualiser ».

Usage : python scripts/generer_pbip.py   (depuis la racine du dépôt)
"""
import json
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NS = uuid.UUID("6f1c2a52-6a0e-4d3e-9a51-0b7c9e1f7d10")


def gid(name):
    """GUID stable (le fichier généré ne change pas d'une exécution à l'autre)."""
    return str(uuid.uuid5(NS, name))


def write(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    text = obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=2)
    path.write_text(text, encoding="utf-8")


# =====================================================================================
# Modèle sémantique
# =====================================================================================
def col(name, dtype, fmt=None, summarize="none", **extra):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "lineageTag": gid("col" + name),
         "summarizeBy": summarize}
    if fmt:
        c["formatString"] = fmt
    c.update(extra)
    return c


def m(lines):
    return {"type": "m", "expression": lines}


def measure(table, name, expr, fmt):
    ms = {"name": name, "expression": expr, "lineageTag": gid(table + "m" + name)}
    if fmt:
        ms["formatString"] = fmt
    return ms


# DossierTP3 = URL GitHub (par défaut, aucun chemin local à régler) ou chemin d'un dossier local cloné
CSV = ('Csv.Document(if Text.StartsWith(DossierTP3, "http") '
       'then Web.Contents(DossierTP3 & "/data/clean/{f}") '
       'else File.Contents(DossierTP3 & "\\data\\clean\\{f}"), '
       '[Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])')

flight_prices = {
    "name": "flight_prices",
    "lineageTag": gid("flight_prices"),
    "columns": [
        col("flight_id", "int64", "0"),
        col("OP_CARRIER", "string"),
        col("flight", "string"),
        col("source_city", "string"),
        col("Arrival_date", "dateTime", "Short Date"),
        col("departure_time", "string", sortByColumn="Ordre_Depart"),
        col("stops", "int64", "0"),
        col("arrival_time", "string"),
        col("destination_city", "string"),
        col("class", "string"),
        col("duration", "double", "0.00"),
        col("price", "double", "#,0"),
        col("Weather_ID", "int64", "0"),
        col("Prix_Categorie", "string"),
        col("Duree_Minutes", "double", "#,0.0"),
        col("Trajet", "string"),
        col("Ordre_Depart", "int64", "0"),
        col("Seuil_Anomalie", "double", "#,0"),
        col("Est_Anomalie", "string"),
        {
            "type": "calculated", "name": "Variation_Prix", "dataType": "double", "isDataTypeInferred": True,
            "expression": "DIVIDE ( flight_prices[price], CALCULATE ( AVERAGE ( flight_prices[price] ), ALL ( flight_prices ) ) ) - 1",
            "formatString": "0.0%", "lineageTag": gid("colVariation_Prix"), "summarizeBy": "none",
        },
    ],
    "partitions": [{
        "name": "flight_prices", "mode": "import",
        "source": m([
            "let",
            "    Source = " + CSV.format(f="flight_prices.csv") + ",",
            "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),",
            "    Types = Table.TransformColumnTypes(EnTetes, {",
            '        {"flight_id", Int64.Type}, {"OP_CARRIER", type text}, {"flight", type text},',
            '        {"source_city", type text}, {"Arrival_date", type date}, {"departure_time", type text},',
            '        {"stops", Int64.Type}, {"arrival_time", type text}, {"destination_city", type text},',
            '        {"class", type text}, {"duration", type number}, {"price", type number}, {"Weather_ID", Int64.Type},',
            '        {"Prix_Categorie", type text}, {"Duree_Minutes", type number}, {"Trajet", type text},',
            '        {"Ordre_Depart", Int64.Type}, {"Seuil_Anomalie", type number}, {"Est_Anomalie", type text}}, "en-US")',
            "in",
            "    Types",
        ]),
    }],
    "measures": [
        measure("f", "Prix Moyen", "AVERAGE ( flight_prices[price] )", "#,0"),
        measure("f", "Durée Moyenne", "AVERAGE ( flight_prices[duration] )", "0.00"),
        measure("f", "Prix Moyen par Classe", "CALCULATE ( [Prix Moyen], ALLEXCEPT ( flight_prices, flight_prices[class] ) )", "#,0"),
        measure("f", "Écart Prix", "[Prix Moyen] - CALCULATE ( [Prix Moyen], ALL ( flight_prices ) )", "#,0"),
        measure("f", "Taux Vols Directs",
                "DIVIDE ( COUNTROWS ( FILTER ( flight_prices, flight_prices[stops] = 0 ) ), COUNTROWS ( flight_prices ) )", "0.0%"),
        measure("f", "Nombre de Vols", "COUNTROWS ( flight_prices )", "#,0"),
        measure("f", "Escales Moyennes", "AVERAGE ( flight_prices[stops] )", "0.00"),
        measure("f", "Durée Moyenne (min)", "AVERAGE ( flight_prices[Duree_Minutes] )", "#,0"),
        measure("f", "Nb Anomalies", 'CALCULATE ( COUNTROWS ( flight_prices ), flight_prices[Est_Anomalie] = "Oui" )', "#,0"),
        measure("f", "Couleur Prix",
                'SWITCH ( TRUE (), [Prix Moyen] < 10000, "#2E7D32", [Prix Moyen] < 20000, "#F9A825", "#C62828" )', ""),
        measure("f", "Prix Moyen Economy", 'CALCULATE ( [Prix Moyen], flight_prices[class] = "Economy" )', "#,0"),
        measure("f", "Durée Moyenne Economy", 'CALCULATE ( [Durée Moyenne], flight_prices[class] = "Economy" )', "0.00"),
        measure("f", "Taux Vols Directs Economy", 'CALCULATE ( [Taux Vols Directs], flight_prices[class] = "Economy" )', "0.0%"),
        measure("f", "Nombre de Vols Economy", 'CALCULATE ( [Nombre de Vols], flight_prices[class] = "Economy" )', "#,0"),
        measure("f", "Corrélation Heure-Prix", "\n".join([
            "VAR mx = AVERAGE ( flight_prices[Ordre_Depart] )",
            "VAR my = AVERAGE ( flight_prices[price] )",
            "VAR cov = SUMX ( flight_prices, ( flight_prices[Ordre_Depart] - mx ) * ( flight_prices[price] - my ) )",
            "VAR sx = SQRT ( SUMX ( flight_prices, ( flight_prices[Ordre_Depart] - mx ) ^ 2 ) )",
            "VAR sy = SQRT ( SUMX ( flight_prices, ( flight_prices[price] - my ) ^ 2 ) )",
            "RETURN DIVIDE ( cov, sx * sy )"]), "0.000"),
        measure("f", "Corrélation Durée-Prix", "\n".join([
            "VAR mx = AVERAGE ( flight_prices[duration] )",
            "VAR my = AVERAGE ( flight_prices[price] )",
            "VAR cov = SUMX ( flight_prices, ( flight_prices[duration] - mx ) * ( flight_prices[price] - my ) )",
            "VAR sx = SQRT ( SUMX ( flight_prices, ( flight_prices[duration] - mx ) ^ 2 ) )",
            "VAR sy = SQRT ( SUMX ( flight_prices, ( flight_prices[price] - my ) ^ 2 ) )",
            "RETURN DIVIDE ( cov, sx * sy )"]), "0.000"),
    ],
}

airlines = {
    "name": "Airlines_info",
    "lineageTag": gid("Airlines_info"),
    "columns": [col("IATA_CODE", "string"), col("Airline", "string"),
                col("Headquarters_Location", "string"), col("Country", "string")],
    "partitions": [{"name": "Airlines_info", "mode": "import", "source": m([
        "let",
        "    Source = " + CSV.format(f="Airlines_info.csv") + ",",
        "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),",
        '    Types = Table.TransformColumnTypes(EnTetes, {{"IATA_CODE", type text}, {"Airline", type text}, {"Headquarters_Location", type text}, {"Country", type text}})',
        "in",
        "    Types"])}],
}

weather = {
    "name": "Weather_conditions",
    "lineageTag": gid("Weather_conditions"),
    "columns": [col("Weather_ID", "int64", "0"), col("Weather_condition", "string")],
    "partitions": [{"name": "Weather_conditions", "mode": "import", "source": m([
        "let",
        "    Source = " + CSV.format(f="Weather_conditions.csv") + ",",
        "    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),",
        '    Types = Table.TransformColumnTypes(EnTetes, {{"Weather_ID", Int64.Type}, {"Weather_condition", type text}})',
        "in",
        "    Types"])}],
}

dim_date = {
    "name": "Dim_Date",
    "lineageTag": gid("Dim_Date"),
    "dataCategory": "Time",
    "columns": [
        col("Date", "dateTime", "Short Date", isKey=True),
        col("Année", "int64", "0"),
        col("Trimestre", "string"),
        col("Mois_Num", "int64", "0"),
        col("Mois", "string", sortByColumn="Mois_Num"),
        col("Année_Mois", "string"),
    ],
    "partitions": [{"name": "Dim_Date", "mode": "import", "source": m([
        "let",
        "    Debut = #date(2020, 1, 1),",
        "    Fin = #date(2025, 12, 31),",
        "    Dates = List.Dates(Debut, Duration.Days(Fin - Debut) + 1, #duration(1, 0, 0, 0)),",
        '    Table0 = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),',
        '    Typee = Table.TransformColumnTypes(Table0, {{"Date", type date}}),',
        '    Annee = Table.AddColumn(Typee, "Année", each Date.Year([Date]), Int64.Type),',
        '    Trim = Table.AddColumn(Annee, "Trimestre", each "T" & Text.From(Date.QuarterOfYear([Date])), type text),',
        '    MoisNum = Table.AddColumn(Trim, "Mois_Num", each Date.Month([Date]), Int64.Type),',
        '    Mois = Table.AddColumn(MoisNum, "Mois", each Date.ToText([Date], "MMM", "fr-FR"), type text),',
        '    AnneeMois = Table.AddColumn(Mois, "Année_Mois", each Date.ToText([Date], "yyyy-MM"), type text)',
        "in",
        "    AnneeMois"])}],
}

model = {
    "compatibilityLevel": 1567,
    "model": {
        "culture": "fr-FR",
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "sourceQueryCulture": "fr-FR",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "tables": [flight_prices, airlines, weather, dim_date],
        "relationships": [
            {"name": gid("rel1"), "fromTable": "flight_prices", "fromColumn": "OP_CARRIER",
             "toTable": "Airlines_info", "toColumn": "IATA_CODE"},
            {"name": gid("rel2"), "fromTable": "flight_prices", "fromColumn": "Weather_ID",
             "toTable": "Weather_conditions", "toColumn": "Weather_ID"},
            {"name": gid("rel3"), "fromTable": "flight_prices", "fromColumn": "Arrival_date",
             "toTable": "Dim_Date", "toColumn": "Date"},
        ],
        "expressions": [{
            "name": "DossierTP3", "kind": "m", "lineageTag": gid("DossierTP3"),
            "expression": '"https://raw.githubusercontent.com/mostafa-amine-zakani/Tp3/claude/project-thread-tws6zc" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
            "annotations": [{"name": "PBI_ResultType", "value": "Text"}],
        }],
        "annotations": [
            {"name": "PBI_QueryOrder",
             "value": json.dumps(["DossierTP3", "flight_prices", "Airlines_info", "Weather_conditions", "Dim_Date"])},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
        ],
    },
}

# =====================================================================================
# Rapport (format PBIR : un dossier par page et un fichier visual.json par visuel)
# =====================================================================================
VISUAL_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.3.0/schema.json"


def lit(value):
    return {"expr": {"Literal": {"Value": value}}}


def visual(vid, vtype, x, y, w, h, roles, title=None):
    """roles : {role: [(kind, table, field), ...]} avec kind = 'C' (colonne) ou 'M' (mesure)."""
    query = {}
    for role, fields in roles.items():
        projections = []
        for kind, table, field in fields:
            key = "Column" if kind == "C" else "Measure"
            p = {"field": {key: {"Expression": {"SourceRef": {"Entity": table}}, "Property": field}},
                 "queryRef": f"{table}.{field}", "nativeQueryRef": field}
            if role in ("Category", "Rows", "Values") and kind == "C" and vtype != "tableEx":
                p["active"] = True
            projections.append(p)
        query[role] = {"projections": projections}
    container = {"background": [{"properties": {"show": lit("false")}}]}
    if title:
        container["title"] = [{"properties": {"show": lit("true"),
                                              "text": lit("'" + title.replace("'", "''") + "'")}}]
    return {
        "$schema": VISUAL_SCHEMA,
        "name": vid,
        "position": {"x": x, "y": y, "z": 1000, "height": h, "width": w, "tabOrder": 1000},
        "visual": {"visualType": vtype, "query": {"queryState": query},
                   "visualContainerObjects": container, "drillFilterOtherVisuals": True},
    }


F, A, D = "flight_prices", "Airlines_info", "Dim_Date"
cards = lambda prefix, ms: [visual(f"{prefix}card{i}", "card", 20 + i * 310, 20, 290, 110, {"Values": [("M", F, mm)]})
                            for i, mm in enumerate(ms)]

pages = [
    ("Vue générale", cards("g", ["Prix Moyen", "Durée Moyenne", "Taux Vols Directs", "Nombre de Vols"]) + [
        visual("g1", "clusteredColumnChart", 20, 150, 610, 270,
               {"Category": [("C", A, "Airline")], "Y": [("M", F, "Prix Moyen")]}, "Prix moyen par compagnie"),
        visual("g2", "clusteredColumnChart", 650, 150, 610, 270,
               {"Category": [("C", F, "class")], "Y": [("M", F, "Prix Moyen")]}, "Prix moyen par classe"),
        visual("g3", "donutChart", 20, 440, 610, 260,
               {"Category": [("C", F, "Prix_Categorie")], "Y": [("M", F, "Nombre de Vols")]}, "Vols par catégorie de prix"),
        visual("g4", "clusteredBarChart", 650, 440, 610, 260,
               {"Category": [("C", A, "Airline")], "Y": [("M", F, "Durée Moyenne")]}, "Durée moyenne par compagnie"),
    ]),
    ("Vue temporelle", [
        visual("t1", "clusteredColumnChart", 20, 20, 610, 330,
               {"Category": [("C", F, "departure_time")], "Y": [("M", F, "Nombre de Vols")]}, "Vols par heure de départ"),
        visual("t2", "lineChart", 650, 20, 610, 330,
               {"Category": [("C", D, "Année_Mois")], "Y": [("M", F, "Nombre de Vols")]}, "Vols par date d'arrivée (mois)"),
        visual("t3", "clusteredColumnChart", 20, 370, 920, 330,
               {"Category": [("C", F, "departure_time")], "Series": [("C", F, "class")], "Y": [("M", F, "Prix Moyen")]},
               "Prix moyen par heure de départ et classe"),
        visual("t4", "card", 960, 370, 300, 150, {"Values": [("M", F, "Corrélation Heure-Prix")]}),
        visual("t5", "card", 960, 540, 300, 150, {"Values": [("M", F, "Corrélation Durée-Prix")]}),
    ]),
    ("Vue géographique", [
        visual("geo1", "pivotTable", 20, 20, 760, 330,
               {"Rows": [("C", F, "source_city")], "Columns": [("C", F, "destination_city")],
                "Values": [("M", F, "Nombre de Vols")]}, "Vols : ville de départ → ville d'arrivée"),
        visual("geo2", "clusteredBarChart", 800, 20, 460, 680,
               {"Category": [("C", F, "Trajet")], "Y": [("M", F, "Nombre de Vols")]}, "Vols par trajet"),
        visual("geo3", "tableEx", 20, 370, 760, 330,
               {"Values": [("C", F, "source_city"), ("M", F, "Durée Moyenne"), ("M", F, "Escales Moyennes"),
                           ("M", F, "Prix Moyen")]}, "Durée et escales par ville de départ"),
    ]),
    ("Vue analytique", [
        visual("an1", "tableEx", 20, 20, 760, 300,
               {"Values": [("C", A, "Airline"), ("M", F, "Prix Moyen Economy"), ("M", F, "Durée Moyenne Economy"),
                           ("M", F, "Taux Vols Directs Economy"), ("M", F, "Nombre de Vols Economy")]},
               "Compagnies sur le segment Economy"),
        visual("an2", "clusteredColumnChart", 800, 20, 460, 300,
               {"Category": [("C", A, "Airline")], "Y": [("M", F, "Prix Moyen Economy")]}, "Prix moyen Economy"),
        visual("an3", "card", 20, 340, 300, 150, {"Values": [("M", F, "Nb Anomalies")]}),
        visual("an4", "clusteredColumnChart", 340, 340, 920, 360,
               {"Category": [("C", A, "Airline")], "Y": [("M", F, "Nb Anomalies")]}, "Vols au prix anormal par compagnie"),
        visual("an5", "slicer", 20, 510, 300, 190, {"Values": [("C", F, "class")]}),
    ]),
]

# =====================================================================================
# Fichiers du projet
# =====================================================================================
platform = lambda typ: {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "metadata": {"type": typ, "displayName": "TP3"},
    "config": {"version": "2.0", "logicalId": gid("logical" + typ)},
}

write(ROOT / "TP3.pbip", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
                          "version": "1.0", "artifacts": [{"report": {"path": "TP3.Report"}}],
                          "settings": {"enableAutoRecovery": True}})
write(ROOT / "TP3.SemanticModel" / "definition.pbism", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
    "version": "1.0", "settings": {}})
write(ROOT / "TP3.SemanticModel" / "model.bim", model)
write(ROOT / "TP3.SemanticModel" / ".platform", platform("SemanticModel"))
import shutil
shutil.rmtree(ROOT / "TP3.Report", ignore_errors=True)
REP = ROOT / "TP3.Report"
write(REP / "definition.pbir", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
    "version": "4.0", "datasetReference": {"byPath": {"path": "../TP3.SemanticModel"}}})
write(REP / "definition" / "version.json", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
    "version": "2.0.0"})
write(REP / "definition" / "report.json", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.2.0/schema.json",
    "themeCollection": {"baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.59", "type": "SharedResources"}},
    "layoutOptimization": "None",
    "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                          "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]}],
    "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                 "allowChangeFilterTypes": True, "useDefaultAggregateDisplayName": True},
})
shutil.copy(Path(__file__).parent / "ressources" / "CY24SU10.json",
            (REP / "StaticResources" / "SharedResources" / "BaseThemes").mkdir(parents=True, exist_ok=True)
            or REP / "StaticResources" / "SharedResources" / "BaseThemes" / "CY24SU10.json")
page_ids = [f"page{i}" for i in range(len(pages))]
write(REP / "definition" / "pages" / "pages.json", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
    "pageOrder": page_ids, "activePageName": page_ids[0]})
for pid, (name, visuals) in zip(page_ids, pages):
    write(REP / "definition" / "pages" / pid / "page.json", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.2.0/schema.json",
        "name": pid, "displayName": name, "displayOption": "FitToPage", "height": 720, "width": 1280})
    for v in visuals:
        write(REP / "definition" / "pages" / pid / "visuals" / v["name"] / "visual.json", v)
write(ROOT / "TP3.Report" / ".platform", platform("Report"))
print("Projet généré : TP3.pbip")
