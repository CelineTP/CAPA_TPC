# Capa Stock

Appli web (PWA) pour scanner un article en magasin, consulter ses infos (désignation, stock, prépa en cours) et définir sa capa par site, puis envoyer le résultat par mail en XLS (5.0/95) ou CSV.

## Déploiement sur GitHub Pages

1. Crée un nouveau dépôt GitHub (public), par exemple `capa-stock`.
2. Mets-y tous les fichiers de ce dossier (`index.html`, `app.js`, `styles.css`, `manifest.json`, `sw.js`, `icons/`) à la racine du dépôt.
3. Dans le dépôt : **Settings → Pages → Source : Deploy from a branch**, branche `main`, dossier `/ (root)`. Sauvegarde.
4. Après une à deux minutes, l'appli est disponible à une adresse du type :
   `https://<ton-compte-github>.github.io/capa-stock/`
5. Ouvre cette adresse sur l'iPhone avec Safari, puis **Partager → Sur l'écran d'accueil**. L'appli s'installe comme une vraie appli, avec son icône.

Chaque mise à jour : remplace les fichiers modifiés sur GitHub (commit), la version installée se met à jour automatiquement au prochain lancement (grâce au service worker).

## Utilisation

1. **Réglages → Fichier stock** : importe ton fichier CSV ou XLS. Il est mémorisé sur le téléphone (pas besoin de le réimporter à chaque lancement, sauf changement de fichier).
2. **Scanner** : scanne le code-barres de l'article (bouton caméra), ou tape le code interne / le libellé dans la barre de recherche. La fiche article s'affiche avec désignation, PV, stock entrepôt, puis une ligne par magasin (VILLE, PUNA, TRVO, PAO, ARUE, MOOREA) avec son stock et sa capa — modifiable en un coup d'œil pour n'importe quel site, pas seulement celui où tu te trouves.
3. Modifie la ou les capa nécessaires (les lignes changées se surlignent), puis **Enregistrer les modifications**. Seules les capa réellement différentes de la valeur du fichier stock sont ajoutées à la session.
4. Onglet **Session** : la liste ne contient que les capa modifiées. Un tap sur ✎ permet de corriger une quantité, ✕ pour la retirer.
5. **Envoyer en XLS (5.0/95)** : sur iPhone, le menu de partage iOS s'ouvre pour envoyer directement par Mail (ou AirDrop, Messages, etc.).

Le fichier exporté contient uniquement 3 colonnes numériques : **site (n°) / code article / capa**, uniquement les lignes modifiées — avec 1=VILLE, 2=PUNA, 3=TRVO, 4=PAO, 5=ARUE, 6=MOOREA.

## Point d'attention : codes-barres en notation scientifique

Dans le fichier stock fourni en exemple, la colonne `codebarre` est stockée comme un nombre et peut apparaître en notation scientifique dans Excel (ex. `8,00084E+12`). L'appli reconvertit ce nombre en chaîne de chiffres complète pour le comparer au code scanné. Cela fonctionne bien pour des codes-barres à 13 chiffres (EAN-13), mais :
- si le code-barres d'origine commence par un ou plusieurs **zéros**, ce zéro initial est perdu dès lors que la colonne a été stockée comme un nombre dans le fichier source (avant même l'import dans l'appli) — l'appli ne peut pas recréer une information déjà perdue en amont ;
- si jamais un article n'est pas retrouvé par scan alors qu'il existe dans le fichier, la recherche manuelle par **code interne** ou **libellé** reste toujours disponible en repli.

## Correspondance des colonnes du fichier stock

Par défaut, l'appli attend ces noms de colonnes (d'après l'exemple fourni) :

| Donnée | Colonne attendue |
|---|---|
| Code interne | `code` |
| Désignation | `libelle` |
| Code-barres | `codebarre` |
| Prix de vente (PV) | `prix_vente` |
| Stock entrepôt | `site0` |
| Stock par magasin | `site1` à `site6` |
| Capa existante par magasin | `capa1` à `capa6` |

Si le format du fichier change, ajuste ces correspondances dans **Réglages → Réglages avancés — correspondance des colonnes**, sans avoir besoin de modifier le code.

## Fonctionnement hors-ligne

Une fois l'appli ouverte une première fois (et le fichier stock importé), elle continue de fonctionner sans connexion internet : le scan, la recherche et la saisie des capas ne nécessitent pas de réseau. Seul l'envoi par mail nécessite une connexion (comme n'importe quel envoi de mail).

## Limites connues

- Le fichier stock est mémorisé dans le stockage local du navigateur (quelques Mo en général suffisants pour plusieurs milliers d'articles). Si le fichier est très volumineux, un message t'avertit s'il n'a pas pu être entièrement mémorisé — dans ce cas, réimporte-le à chaque lancement.
- Le scan caméra fonctionne dans Safari iOS 15+ (accès à la caméra à autoriser au premier lancement).
- Le partage direct par mail avec pièce jointe utilise le menu de partage iOS natif ; sur d'anciens iPhones/versions iOS, l'appli bascule automatiquement sur un téléchargement du fichier à joindre manuellement.
