// Fait_Achats - Lignes de commande fournisseurs, avec retard de livraison
// Chargée dans le modèle de données
let
    Source = Csv.Document(File.Contents(pChemin & "05_Achats\Achats_Fournisseurs.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    EnTetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Noms = Table.TransformColumnNames(EnTetes, (n) => Text.Replace(n, Character.FromNumber(65279), "")),
    Dates = Table.TransformColumns(Noms, {
        {"Date_Commande", fxDateSouple, type date},
        {"Date_Livraison_Reelle", fxDateSouple, type date}}),
    Nombres = Table.TransformColumns(Dates, {
        {"Quantite_Commandee", fxNombreSouple, Int64.Type},
        {"Prix_Achat_Unitaire_HT", fxNombreSouple, type number},
        {"Montant_HT", fxNombreSouple, type number}}),
    Jointure = Table.NestedJoin(Nombres, {"Code_Fournisseur"}, Dim_Fournisseurs, {"Code_Fournisseur"}, "F", JoinKind.LeftOuter),
    Delai = Table.ExpandTableColumn(Jointure, "F", {"Delai_Livraison_Contractuel_Jours"}, {"Delai_Contractuel"}),
    Retard = Table.AddColumn(Delai, "Retard_Jours", each
        fxEcartJours([Date_Commande], [Date_Livraison_Reelle], [Delai_Contractuel]), type number),
    MoisCommande = Table.AddColumn(Retard, "Mois_Commande", each Date.StartOfMonth([Date_Commande]), type date),
    Types = Table.TransformColumnTypes(MoisCommande, {
        {"Num_Commande", type text}, {"Code_Magasin", type text},
        {"Code_Fournisseur", type text}, {"Code_Produit", type text}})
in
    Types
