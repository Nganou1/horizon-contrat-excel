// fxDateSouple - Convertit JJ/MM/AAAA, AAAA-MM-JJ ou une vraie date en date
// Connexion uniquement (étape intermédiaire ou fonction)
(t as any) as nullable date =>
let
    Resultat =
        if t = null then null
        else if t is date then t
        else if t is datetime then Date.From(t)
        else
            let
                s = Text.Trim(Text.From(t))
            in
                if s = "" then null
                else if Text.Contains(s, "/") then
                    let p = Text.Split(s, "/") in
                    #date(Number.FromText(p{2}), Number.FromText(p{1}), Number.FromText(p{0}))
                else
                    #date(Number.FromText(Text.Middle(s, 0, 4)), Number.FromText(Text.Middle(s, 5, 2)), Number.FromText(Text.Middle(s, 8, 2)))
in
    Resultat
