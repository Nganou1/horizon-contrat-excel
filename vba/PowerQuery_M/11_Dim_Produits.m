// Dim_Produits - Référentiel produits
// Chargée dans le modèle de données
let
    Classeur = Excel.Workbook(File.Contents(pChemin & "01_Referentiels\Referentiels_Horizon_Contrat.xlsx"), null, true),
    Feuille = Classeur{[Item = "Produits", Kind = "Sheet"]}[Data],
    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Types = Table.TransformColumnTypes(EnTetes, {
        {"Code_Produit", type text}, {"Libelle_Produit", type text}, {"Categorie", type text},
        {"Sous_Categorie", type text}, {"Taux_TVA", type number}, {"Taux_Cout_Matiere", type number},
        {"Prix_Vente_TTC_Actuel", type number}, {"Actif", type text}})
in
    Types
