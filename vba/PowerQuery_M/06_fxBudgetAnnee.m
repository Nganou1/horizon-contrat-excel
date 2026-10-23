// fxBudgetAnnee - Lit un fichier Budget_AAAA.xlsx et le dépivote
// Connexion uniquement (étape intermédiaire ou fonction)
(annee as number) as table =>
let
    Classeur = Excel.Workbook(File.Contents(pChemin & "07_Budget\Budget_" & Number.ToText(annee) & ".xlsx"), null, true),
    Feuille = Classeur{0}[Data],
    SansTitres = Table.Skip(Feuille, 4),
    EnTetes = Table.PromoteHeaders(SansTitres, [PromoteAllScalars = true]),
    SansVides = Table.SelectRows(EnTetes, each [Code_Magasin] <> null and [Code_Magasin] <> ""),
    Depivote = Table.UnpivotOtherColumns(SansVides, {"Code_Magasin", "Nom_Magasin", "Region"}, "Mois_Texte", "Budget_CA_HT"),
    Mois = Table.AddColumn(Depivote, "Mois", each
        #date(Number.FromText(Text.AfterDelimiter([Mois_Texte], "/")), Number.FromText(Text.BeforeDelimiter([Mois_Texte], "/")), 1), type date),
    Colonnes = Table.SelectColumns(Mois, {"Code_Magasin", "Mois", "Budget_CA_HT"}),
    Types = Table.TransformColumnTypes(Colonnes, {{"Code_Magasin", type text}, {"Budget_CA_HT", type number}})
in
    Types
