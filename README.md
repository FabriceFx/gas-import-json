# gas-import-json

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

Bibliothèque Google Apps Script moderne pour importer et aplanir des données JSON directement dans Google Sheets. Cette version est optimisée pour le moteur V8 (ES6+).

## Fonctionnalités clés

* **Importation GET** : Récupération standard depuis des API publiques.
* **Importation POST** : Support des API nécessitant un payload (corps de requête).
* **Aplatissement Automatique** : Conversion des objets JSON imbriqués en tableau 2D lisible.
* **Filtrage par Chemin** : Extraction ciblée de sous-ensembles de données (ex: `/data/resultats`).
* **Options Flexibles** : Gestion des en-têtes (`noHeaders`, `rawHeaders`).

## Installation manuelle

1.  Ouvrez votre Google Sheet.
2.  Allez dans **Extensions** > **Apps Script**.
3.  Copiez le contenu du fichier `Code.gs` dans l'éditeur.
4.  Sauvegardez le projet (`Ctrl + S`).

## Utilisation

### Syntaxe de base (GET)

```excel
=IMPORTER_JSON(url; [chemin_requete]; [options])
