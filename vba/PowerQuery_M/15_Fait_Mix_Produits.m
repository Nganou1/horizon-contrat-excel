// Fait_Mix_Produits - Mix produits : magasin x mois x produit
// Chargée dans le modèle de données
let
    Source = Csv.Document(File.Contents(pChemin & "03_Mix_Produits\Mix_Produits_Mensuel.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), "")),
    Mois = Table.TransformColumns(Noms, {{"Mois", (m) => fxDateSouple(m & "-01"), type date}}),
    Nombres = Table.TransformColumns(Mois, {
        {"Quantite_Vendue", fxNombreSouple, Int64.Type},
        {"CA_HT", fxNombreSouple, type number},
        {"Cout_Matiere_HT", fxNombreSouple, type number}}),
    Types = Table.TransformColumnTypes(Nombres, {{"Code_Magasin", type text}, {"Code_Produit", type text}})
in
    Types
