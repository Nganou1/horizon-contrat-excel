// Fait_RH - Heures mensuelles par salarié
// Chargée dans le modèle de données
let
    Source = Csv.Document(File.Contents(pChemin & "06_RH\Heures_Mensuelles.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), "")),
    Mois = Table.TransformColumns(Noms, {{"Mois", (m) => fxDateSouple(m & "-01"), type date}}),
    Nombres = Table.TransformColumns(Mois, {
        {"Heures_Travaillees", fxNombreSouple, type number},
        {"Heures_Supplementaires", fxNombreSouple, type number},
        {"Jours_Absence", fxNombreSouple, type number}}),
    Types = Table.TransformColumnTypes(Nombres, {
        {"Matricule", type text}, {"Code_Magasin", type text}, {"Poste", type text}})
in
    Types
