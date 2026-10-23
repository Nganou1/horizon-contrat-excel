// qry_Rejets_Ventes - Lignes sans code magasin, comptées par fichier
// Connexion uniquement (étape intermédiaire ou fonction)
let
    Source = qry_Ventes_Brutes,
    Vides = Table.SelectRows(Source, each [Code_Magasin] = null or Text.Trim([Code_Magasin]) = ""),
    ParFichier = Table.Group(Vides, {"Mois_Fichier"}, {{"Nb_Rejets", each Table.RowCount(_), Int64.Type}})
in
    ParFichier
