// Fait_Stocks - Stocks mensuels par magasin et produit
// Chargée dans le modèle de données
let
    Source = Csv.Document(File.Contents(pChemin & "04_Stocks\Stocks_Mensuels.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), "")),
    Mois = Table.TransformColumns(Noms, {{"Mois", (m) => fxDateSouple(m & "-01"), type date}}),
    Nombres = Table.TransformColumns(Mois, {
        {"Stock_Initial", fxNombreSouple, Int64.Type}, {"Entrees", fxNombreSouple, Int64.Type},
        {"Sorties", fxNombreSouple, Int64.Type}, {"Pertes", fxNombreSouple, Int64.Type},
        {"Stock_Final", fxNombreSouple, Int64.Type}}),
    Taux = Table.AddColumn(Nombres, "Taux_Perte", each
        if [Sorties] = null or [Sorties] = 0 then null else [Pertes] / [Sorties], type number),
    Types = Table.TransformColumnTypes(Taux, {{"Code_Magasin", type text}, {"Code_Produit", type text}})
in
    Types
