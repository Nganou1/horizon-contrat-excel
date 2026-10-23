// Dim_Magasins - Référentiel magasins dédoublonné, villes normalisées
// Chargée dans le modèle de données
let
    Classeur = Excel.Workbook(File.Contents(pChemin & "01_Referentiels\Referentiels_Horizon_Contrat.xlsx"), null, true),
    Feuille = Classeur{[Item = "Magasins", Kind = "Sheet"]}[Data],
    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    SansDoublons = Table.Distinct(EnTetes),
    Villes = Table.TransformColumns(SansDoublons, {{"Ville", fxNettoyerTexte, type text}}),
    Dates = Table.TransformColumns(Villes, {{"Date_Ouverture", fxDateSouple, type date}, {"Date_Fermeture", fxDateSouple, type date}}),
    Types = Table.TransformColumnTypes(Dates, {
        {"Code_Magasin", type text}, {"Nom_Magasin", type text}, {"Region", type text},
        {"Type_Emplacement", type text}, {"Statut", type text}, {"Surface_m2", Int64.Type},
        {"Nb_Places_Assises", Int64.Type}, {"Drive", type text}, {"Franchise_Responsable", type text}})
in
    Types
