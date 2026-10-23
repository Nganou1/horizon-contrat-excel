// Dim_Employes - Référentiel salariés, taux horaires texte convertis
// Chargée dans le modèle de données
let
    Classeur = Excel.Workbook(File.Contents(pChemin & "01_Referentiels\Referentiels_Horizon_Contrat.xlsx"), null, true),
    Feuille = Classeur{[Item = "Employes", Kind = "Sheet"]}[Data],
    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars = true]),
    Taux = Table.TransformColumns(EnTetes, {{"Taux_Horaire_Brut", (v) =>
        if v is number then v
        else Number.FromText(Text.Remove(Text.From(v), {"€", " ", Character.FromNumber(160), Character.FromNumber(8239)}), "fr-FR"),
        type number}}),
    Dates = Table.TransformColumns(Taux, {{"Date_Embauche", fxDateSouple, type date}}),
    Types = Table.TransformColumnTypes(Dates, {
        {"Matricule", type text}, {"Code_Magasin", type text}, {"Nom_Prenom", type text},
        {"Poste", type text}, {"Type_Contrat", type text}})
in
    Types
