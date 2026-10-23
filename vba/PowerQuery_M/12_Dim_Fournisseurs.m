// Dim_Fournisseurs - Référentiel fournisseurs
// Chargée dans le modèle de données
let
    Classeur = Excel.Workbook(File.Contents(pChemin & "01_Referentiels\Referentiels_Horizon_Contrat.xlsx"), null, true),
    Feuille = Classeur{[Item = "Fournisseurs", Kind = "Sheet"]}[Data],
    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Types = Table.TransformColumnTypes(EnTetes, {
        {"Code_Fournisseur", type text}, {"Nom_Fournisseur", type text}, {"Categorie_Approvisionnee", type text},
        {"Delai_Livraison_Contractuel_Jours", Int64.Type}, {"Conditions_Paiement", type text}, {"Note_Qualite", type number}})
in
    Types
