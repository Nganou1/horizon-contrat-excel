// qry_Ventes_Brutes - Les 24 fichiers de ventes lus et convertis, sans filtre
// Connexion uniquement (étape intermédiaire ou fonction)
let
    Source = Folder.Files(pChemin & "02_Ventes"),
    Fichiers = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv" and Text.StartsWith([Name], "Ventes_")),
    AvecMois = Table.AddColumn(Fichiers, "Mois_Fichier", each Text.BetweenDelimiters([Name], "Ventes_", ".csv"), type text),
    Lues = Table.AddColumn(AvecMois, "Contenu", each
        Table.TransformColumnNames(
            Table.PromoteHeaders(
                Csv.Document([Content], [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                [PromoteAllScalars = true]),
            (n) => Text.Replace(n, Character.FromNumber(65279), ""))),
    Reduite = Table.SelectColumns(Lues, {"Mois_Fichier", "Contenu"}),
    Etendue = Table.ExpandTableColumn(Reduite, "Contenu",
        {"Date", "Code_Magasin", "Canal", "Nb_Tickets", "Nb_Articles", "CA_Brut_HT", "Remises_HT", "Cout_Matiere_HT"}),
    Dates = Table.TransformColumns(Etendue, {{"Date", fxDateSouple, type date}}),
    Nombres = Table.TransformColumns(Dates, {
        {"CA_Brut_HT", fxNombreSouple, type number},
        {"Remises_HT", fxNombreSouple, type number},
        {"Cout_Matiere_HT", fxNombreSouple, type number},
        {"Nb_Tickets", fxNombreSouple, Int64.Type},
        {"Nb_Articles", fxNombreSouple, Int64.Type}})
in
    Nombres
