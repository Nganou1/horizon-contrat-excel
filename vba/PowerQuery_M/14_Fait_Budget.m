// Fait_Budget - Budgets 2025 et 2026 dépivotés : 120 magasins x 12 mois x 2 ans
// Chargée dans le modèle de données
let
    Combine = Table.Combine({fxBudgetAnnee(2025), fxBudgetAnnee(2026)})
in
    Combine
