# Application Excel/VBA – Tarification Assurance Santé

## Contexte
Ce projet a été réalisé dans le cadre de la formation en actuariat à l'ISFA. L'objectif était de créer une application interactive sous Microsoft Excel, utilisant le langage VBA, permettant de simuler la tarification d'un produit d'assurance santé.

L'utilisateur saisit les caractéristiques du client via un formulaire (UserForm), et la prime d'assurance est calculée automatiquement selon une formule personnalisable basée sur les données entrées.

## Objectif de l'application
- Permettre la saisie des informations du client (âge, IMC, sexe, couverture, options, fumeur).
- Calculer automatiquement la prime d'assurance selon une formule ajustée par coefficients.
- Fournir une interface claire et ergonomique pour l'utilisateur.
- Lancer l'application facilement via un bouton ou une image dans Excel.

## Fonctionnalités principales
1. **UserForm principal (frmTarif)**
   - Champs de saisie : Âge, IMC, Type de couverture, Options (Optique, Dentaire, Hospitalisation, Chirurgie, Soins à domicile, Fumeur)
   - Bouton pour sélectionner le sexe (frmSexe)
   - Bouton de calcul de la prime
   - Affichage dynamique du résultat (Label)

2. **UserForm secondaire (frmSexe)**
   - Sélection du sexe (Homme/Femme)
   - Validation et transfert vers le formulaire principal

3. **Macro de lancement**
   - `LancerTarification` permet d’exécuter l’application depuis un bouton ou un logo inséré dans la feuille Excel.
## Formule de tarification
Prime = Base x CoeffAge x CoeffIMC x CoeffSexe x CoeffCouverture x CoeffOptions x CoeffFumeur
- Base fixe (ex : 10 000 FCFA)
- Coefficients ajustés selon les caractéristiques du client :
  - Age : surcote à partir de 30 et 50 ans
  - IMC : surcharge pour IMC élevé
  - Sexe : légère décote pour les femmes
  - Type de couverture : coefficient selon niveau (ex : 1.5 pour Étendue)
  - Options : ajout cumulatif
  - Fumeur : surcote de 20%

## Compétences développées
- **VBA / Excel** : développement d’interfaces et macros interactives
- **Modélisation actuarielle** : calcul de primes selon règles personnalisées
- **Structuration de projet technique** : séparation des formulaires et modules
- **Interface utilisateur (UserForm)** : ergonomie et visualisation des résultats
- **Simulation et validation** : vérification des coefficients et calcul dynamique

## Fichiers inclus
- Fichier Excel (.xlsm) : application fonctionnelle pour calcul des primes
- Code VBA exporté : modules et UserForms
- Rapport PDF : description détaillée du projet et de la méthodologie

## Améliorations possibles
- Enregistrement automatique des simulations
- Export de devis en PDF
- Historique des calculs
- Vérifications avancées de saisie


