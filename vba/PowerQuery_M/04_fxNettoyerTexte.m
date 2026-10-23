// fxNettoyerTexte - Casse de titre, particules en minuscule après un tiret (Aix-en-Provence)
// Connexion uniquement (étape intermédiaire ou fonction)
(t as nullable text) as nullable text =>
let
    Particules = {"de", "du", "des", "la", "le", "les", "sur", "en"},
    Majuscule = (w as text) as text =>
        if w = "" then w else Text.Upper(Text.Start(w, 1)) & Text.Lower(Text.Range(w, 1)),
    TraiterMot = (m as text) as text =>
        let
            Morceaux = Text.Split(m, "-"),
            Traites = List.Transform(
                List.Positions(Morceaux),
                (i) => if i > 0 and List.Contains(Particules, Text.Lower(Morceaux{i}))
                       then Text.Lower(Morceaux{i})
                       else Majuscule(Morceaux{i}))
        in
            Text.Combine(Traites, "-"),
    Resultat =
        if t = null then null
        else Text.Combine(List.Transform(List.Select(Text.Split(Text.Trim(t), " "), each _ <> ""), TraiterMot), " ")
in
    Resultat
