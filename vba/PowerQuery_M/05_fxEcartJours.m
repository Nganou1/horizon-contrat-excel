// fxEcartJours - Retard de livraison en jours au-delà du délai contractuel (null si non livré)
// Connexion uniquement (étape intermédiaire ou fonction)
(dCommande as nullable date, dLivraison as nullable date, delai as nullable number) as nullable number =>
    if dCommande = null or dLivraison = null or delai = null then null
    else Duration.Days(dLivraison - dCommande) - delai
