# Actualisation de la base des médicaments vétérinaires (ANSES) dans Grist

Le serveur de l'ANSES (`pro.anses.fr`) ne renvoie pas d'en-tête CORS : un widget Grist ne peut donc pas
télécharger le fichier directement. Ce dépôt contourne le problème :

1. **GitHub Actions** télécharge chaque nuit le fichier publié sur data.gouv.fr
   (`amm-vet-fr-v2.xls`) et le publie sur **GitHub Pages**, qui autorise les requêtes CORS.
2. Le **widget** (`site/index.html`), publié sur le même site, lit cette copie et met à jour
   les tables du document Grist.

| Ressource | URL |
|---|---|
| Widget | `https://dlafon31.github.io/anses-medvet/` |
| Copie du fichier | `https://dlafon31.github.io/anses-medvet/data/amm-vet-fr.xls` |
| Métadonnées | `https://dlafon31.github.io/anses-medvet/data/meta.json` |

## Mise en place

1. Créer le dépôt `anses-medvet` (public) et y pousser ces fichiers.
2. *Settings → Pages → Build and deployment → Source* : **GitHub Actions**.
3. *Actions → Miroir du fichier ANSES → Run workflow* pour la première publication.
4. Dans Grist : widget **Personnalisé** → URL `https://dlafon31.github.io/anses-medvet/`,
   accès **complet** au document.

## Utilisation

**1. Télécharger et analyser** → contrôler le récapitulatif → **2. Appliquer les modifications**.
`Medicaments` est mise à jour par `Num_Ident` (identifiants de ligne stables) ; les autres tables
sont remplacées uniquement si leur contenu a changé. Le bouton *Importer un fichier Excel…* reste
disponible si la copie est indisponible.

## Tables de paramétrage dans le document Grist

- **Admin** (colonnes `Table`, `Onglet`) : onglet du fichier qui alimente chaque table.
- **Admin_Colonnes** (colonnes `Table`, `Colonne`, `Entete`) : en-têtes Excel qui diffèrent du
  libellé de la colonne Grist (ex. `Medicaments.Num_Ident` ← « N° d'identification »).

Garde-fou : une table n'est jamais réécrite si sa colonne clé (utilisée par les références) ou
l'une de ses colonnes de référence n'est pas trouvée dans l'onglet.
L'inventaire des onglets et en-têtes du fichier est publié dans `data/inventaire.json`.
