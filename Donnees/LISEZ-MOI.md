# Jeu de données Horizon_Contrat

Réseau fictif de restauration rapide premium : 120 points de vente sur 4 régions, du 01/09/2024 au 31/08/2026 (24 mois complets). Toutes les données sont simulées.

## Arborescence

| Dossier | Contenu | Volume |
|---|---|---|
| `01_Referentiels` | Magasins, Produits, Historique_Prix, Fournisseurs, Employes (un classeur, 5 feuilles) | — |
| `02_Ventes` | `Ventes_AAAA-MM.csv`, grain magasin × jour × canal | 24 fichiers, 289 266 lignes |
| `03_Mix_Produits` | Grain magasin × mois × produit | 112 596 lignes |
| `04_Stocks` | Stocks mensuels | 112 596 lignes |
| `05_Achats` | Commandes fournisseurs | 112 596 lignes |
| `06_RH` | Heures mensuelles | 42 759 lignes |
| `07_Budget` | Budgets 2025 et 2026, mois en colonnes | 2 classeurs |

CSV : séparateur point-virgule, UTF-8 avec BOM.

## Règles de gestion

- `CA_Brut_HT` est le chiffre d'affaires avant remises ; CA net = `CA_Brut_HT − Remises_HT`.
- `Cout_Matiere_HT` est exprimé sur le CA net ; marge brute = CA brut − remises − coût matière.
- TVA : 10 % en restauration, 20 % sur les alcools.
- Les prix de la table Produits sont les prix actuels ; l'historique des prix porte la hausse de +4,5 % du 01/09/2025.
- Ventes est la source de référence pour le CA et la marge ; Mix_Produits sert au mix produit (écart inférieur à 0,2 %, arrondis).
- Stock_Final = Stock_Initial + Entrées − Sorties − Pertes, vérifié sur toutes les lignes.

## Défauts de qualité traités dans le classeur

- Dates au format JJ/MM/AAAA dans trois fichiers de ventes, AAAA-MM-JJ ailleurs.
- Virgule décimale dans deux fichiers de ventes.
- 240 lignes strictement dupliquées dans `Ventes_2025-07`.
- 422 lignes sans code magasin, réparties sur les 24 fichiers.
- Libellés de canal incohérents (« Livraison  », « livraison », « LIVRAISON », « Sur place » avec espaces).
- Feuille Magasins : 2 doublons parfaits, 12 villes en casse incohérente.
- Feuille Employes : 327 taux horaires stockés en texte (« 12,43 € »).
- Achats : 2 251 dates de livraison réelle vides (commandes en cours).
- Budgets : trois lignes de titre avant l'en-tête et les mois en colonnes, à dépivoter.

Le détail des corrections et les contrôles recalculés figurent dans la feuille `QUALITE` du classeur.
