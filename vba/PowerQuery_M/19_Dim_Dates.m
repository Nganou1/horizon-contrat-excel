// Dim_Dates - Calendrier du 01/09/2024 au 31/08/2026, exercice septembre-août
// Chargée dans le modèle de données
let
    Debut = #date(2024, 9, 1),
    Fin = #date(2026, 8, 31),
    Jours = List.Dates(Debut, Duration.Days(Fin - Debut) + 1, #duration(1, 0, 0, 0)),
    Base = Table.FromList(Jours, Splitter.SplitByNothing(), {"Date"}),
    Typee = Table.TransformColumnTypes(Base, {{"Date", type date}}),
    C1 = Table.AddColumn(Typee, "Annee", each Date.Year([Date]), Int64.Type),
    C2 = Table.AddColumn(C1, "Mois_Num", each Date.Month([Date]), Int64.Type),
    C3 = Table.AddColumn(C2, "Mois_Debut", each Date.StartOfMonth([Date]), type date),
    C4 = Table.AddColumn(C3, "Mois_Nom", each Date.ToText([Date], "MMMM", "fr-FR"), type text),
    C5 = Table.AddColumn(C4, "Mois_Abrege", each Date.ToText([Date], "MMM yy", "fr-FR"), type text),
    C6 = Table.AddColumn(C5, "Trimestre", each "T" & Number.ToText(Date.QuarterOfYear([Date])), type text),
    C7 = Table.AddColumn(C6, "Semaine", each Date.WeekOfYear([Date], Day.Monday), Int64.Type),
    C8 = Table.AddColumn(C7, "Jour_Semaine", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type),
    C9 = Table.AddColumn(C8, "Jour_Semaine_Nom", each Date.ToText([Date], "dddd", "fr-FR"), type text),
    C10 = Table.AddColumn(C9, "Est_Weekend", each [Jour_Semaine] >= 6, type logical),
    C11 = Table.AddColumn(C10, "Annee_Exercice", each if [Mois_Num] >= 9 then [Annee] else [Annee] - 1, Int64.Type),
    C12 = Table.AddColumn(C11, "Exercice", each
        Number.ToText([Annee_Exercice]) & "/" & Text.End(Number.ToText([Annee_Exercice] + 1), 2), type text),
    C13 = Table.AddColumn(C12, "Mois_Exercice", each if [Mois_Num] >= 9 then [Mois_Num] - 8 else [Mois_Num] + 4, Int64.Type),
    C14 = Table.AddColumn(C13, "Mois_Index", each [Annee] * 100 + [Mois_Num], Int64.Type)
in
    C14
