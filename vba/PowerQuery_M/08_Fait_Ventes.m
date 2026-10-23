// Fait_Ventes - Ventes nettoyées : canal normalisé, doublons et codes vides retirés
// Chargée dans le modèle de données
let
    Source = qry_Ventes_Brutes,
    CanalPropre = Table.TransformColumns(Source, {{"Canal", (c) =>
        let k = Text.Lower(Text.Trim(c)) in
        if k = "sur place" then "Sur place"
        else if k = "à emporter" then "À emporter"
        else if k = "livraison" then "Livraison"
        else if k = "drive" then "Drive"
        else Text.Trim(c), type text}}),
    SansDoublons = Table.Distinct(CanalPropre,
        {"Date", "Code_Magasin", "Canal", "Nb_Tickets", "Nb_Articles", "CA_Brut_HT", "Remises_HT", "Cout_Matiere_HT"}),
    SansVides = Table.SelectRows(SansDoublons, each [Code_Magasin] <> null and Text.Trim([Code_Magasin]) <> ""),
    CANet = Table.AddColumn(SansVides, "CA_Net_HT", each [CA_Brut_HT] - [Remises_HT], type number),
    Types = Table.TransformColumnTypes(CANet, {{"Code_Magasin", type text}, {"Mois_Fichier", type text}})
in
    Types
