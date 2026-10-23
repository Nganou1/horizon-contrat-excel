// pChemin - Chemin du dossier Donnees, lu dans la cellule nommée Chemin_Donnees
// Connexion uniquement (étape intermédiaire ou fonction)
let
    Source = Excel.CurrentWorkbook(){[Name = "Chemin_Donnees"]}[Content],
    Valeur = Text.Trim(Text.From(Source{0}[Column1])),
    AvecSeparateur = if Text.EndsWith(Valeur, "\") then Valeur else Valeur & "\"
in
    AvecSeparateur
