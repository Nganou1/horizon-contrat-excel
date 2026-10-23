// fxNombreSouple - Convertit un nombre écrit avec virgule ou point décimal
// Connexion uniquement (étape intermédiaire ou fonction)
(t as any) as nullable number =>
let
    Resultat =
        if t = null then null
        else if t is number then t
        else
            let
                s = Text.Trim(Text.From(t))
            in
                if s = "" then null
                else if Text.Contains(s, ",") then Number.FromText(s, "fr-FR")
                else Number.FromText(s, "en-US")
in
    Resultat
