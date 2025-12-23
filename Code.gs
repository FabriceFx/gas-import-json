/**
 * @fileoverview Bibliothèque moderne pour importer des données JSON dans Google Sheets.
 * Refonte complète de ImportJSON pour le moteur V8 et ES6+.
 * @author Fabrice Faucheux (Expert Apps Script)
 */

/**
 * Fonction personnalisée pour importer un flux JSON depuis une URL.
 * S'utilise directement dans une cellule : =IMPORTER_JSON("https://api.ex.com/users", "/data", "noHeaders")
 *
 * @param {string} url - L'URL de l'API publique (GET).
 * @param {string} [cheminRequete=""] - (Optionnel) Chemin pour filtrer les données (ex: "/items").
 * @param {string} [options=""] - (Optionnel) Options séparées par des virgules (ex: "noHeaders, rawHeaders").
 * @return {Array<Array<string>>} Un tableau 2D prêt pour l'affichage dans Sheets.
 * @customfunction
 */
const IMPORTER_JSON = (url, cheminRequete = "", options = "") => {
  return executerImportation_(url, { method: "get" }, cheminRequete, options);
};

/**
 * Fonction personnalisée pour importer un flux JSON via une requête POST.
 * Utile pour les API nécessitant un corps de requête (payload).
 *
 * @param {string} url - L'URL de l'API.
 * @param {string} payload - Le contenu du corps de la requête (souvent x-www-form-urlencoded).
 * @param {string} [cheminRequete=""] - (Optionnel) Chemin pour filtrer les données.
 * @param {string} [options=""] - (Optionnel) Options d'affichage.
 * @return {Array<Array<string>>} Tableau 2D des résultats.
 * @customfunction
 */
const IMPORTER_JSON_POST = (url, payload, cheminRequete = "", options = "") => {
  const optionsRequete = {
    method: "post",
    payload: payload,
    contentType: "application/x-www-form-urlencoded" // Par défaut, modifiable si besoin
  };
  return executerImportation_(url, optionsRequete, cheminRequete, options);
};

/**
 * Cœur logique de l'importation. Gère le fetch, le parsing et la transformation.
 * * @param {string} url - L'URL cible.
 * @param {Object} optionsRequete - Options pour UrlFetchApp (method, payload, etc.).
 * @param {string} cheminRequete - Le chemin pour filtrer le JSON.
 * @param {string} optionsUtilisateur - Chaîne d'options (ex: "noHeaders").
 * @return {Array<Array<any>>} Les données formatées pour Sheets.
 * @private
 */
const executerImportation_ = (url, optionsRequete, cheminRequete, optionsUtilisateur) => {
  if (!url) return [["Erreur : URL manquante"]];

  try {
    // 1. Appel API
    const reponse = UrlFetchApp.fetch(url, {
      ...optionsRequete,
      muteHttpExceptions: true // Pour gérer les erreurs 404/500 gracieusement
    });

    if (reponse.getResponseCode() !== 200) {
      throw new Error(`Erreur HTTP ${reponse.getResponseCode()}: ${reponse.getContentText().slice(0, 100)}...`);
    }

    const jsonBrut = JSON.parse(reponse.getContentText());

    // 2. Traitement des options
    const config = parserOptions_(optionsUtilisateur);

    // 3. Filtrage et Aplatissement
    // Si un chemin est spécifié, on descend dans l'objet
    const donneesCible = naviguerVersChemin_(jsonBrut, cheminRequete);
    
    // Aplatissement en tableau 2D
    const tableauFinal = transformerEnTableau2D_(donneesCible, config);

    return tableauFinal.length > 0 ? tableauFinal : [["Aucune donnée trouvée"]];

  } catch (erreur) {
    console.error(`Erreur dans IMPORTER_JSON : ${erreur.message}`);
    return [[`Erreur : ${erreur.message}`]];
  }
};

/**
 * Convertit des données JSON complexes (Objets/Tableaux imbriqués) en un tableau 2D pour Sheets.
 * * @param {Object|Array} json - Les données JSON à transformer.
 * @param {Object} config - Configuration (headers, etc.).
 * @return {Array<Array<any>>} Tableau 2D.
 * @private
 */
const transformerEnTableau2D_ = (json, config) => {
  const listeObjetsAplanis = [];
  const enTetesSet = new Set();

  // Fonction récursive pour aplanir un objet
  const aplanirObjet = (obj, prefixe = "") => {
    const resultat = {};
    
    Object.keys(obj).forEach(cle => {
      const valeur = obj[cle];
      const nouvelleCle = prefixe ? `${prefixe}/${cle}` : cle; // Séparateur de chemin standard

      if (valeur && typeof valeur === 'object' && !Array.isArray(valeur)) {
        // C'est un objet imbriqué -> Récursion
        const sousObjet = aplanirObjet(valeur, nouvelleCle);
        Object.assign(resultat, sousObjet);
      } else if (Array.isArray(valeur)) {
        // C'est un tableau -> On joint les valeurs scalaires ou on laisse tel quel pour le traitement par ligne
        // Simplification pour Sheets : on convertit souvent les tableaux en chaîne pour tenir dans une cellule
        // Ou on pourrait étendre les lignes. Ici, choix de concaténer pour simplicité d'affichage.
        resultat[nouvelleCle] = valeur.map(v => (typeof v === 'object' ? JSON.stringify(v) : v)).join(", ");
        enTetesSet.add(nouvelleCle);
      } else {
        // Valeur scalaire (string, number, boolean)
        resultat[nouvelleCle] = valeur;
        enTetesSet.add(nouvelleCle);
      }
    });
    return resultat;
  };

  // Normalisation de l'entrée en tableau d'objets
  let tableauDonnees = [];
  if (Array.isArray(json)) {
    tableauDonnees = json;
  } else if (typeof json === 'object' && json !== null) {
    tableauDonnees = [json];
  } else {
    return [[json]]; // Cas scalaire simple
  }

  // Traitement de chaque ligne
  tableauDonnees.forEach(item => {
    if (typeof item === 'object' && item !== null) {
      listeObjetsAplanis.push(aplanirObjet(item));
    } else {
      listeObjetsAplanis.push({ "valeur": item });
      enTetesSet.add("valeur");
    }
  });

  // Construction de la matrice finale
  const enTetes = Array.from(enTetesSet);
  const matrice = [];

  // Ajout des en-têtes si demandé
  if (!config.noHeaders) {
    const enTetesFormates = config.rawHeaders ? enTetes : enTetes.map(t => formaterEnTete_(t));
    matrice.push(enTetesFormates);
  }

  // Remplissage des données
  listeObjetsAplanis.forEach(obj => {
    const ligne = enTetes.map(header => {
      return obj[header] !== undefined ? obj[header] : ""; // Vide si pas de valeur
    });
    matrice.push(ligne);
  });

  return matrice;
};

/**
 * Navigue dans l'objet JSON selon un chemin de type "/data/items".
 * @private
 */
const naviguerVersChemin_ = (json, chemin) => {
  if (!chemin || chemin === "/") return json;
  
  const segments = chemin.split("/").filter(s => s.length > 0);
  let pointeur = json;

  for (const segment of segments) {
    if (pointeur && pointeur[segment] !== undefined) {
      pointeur = pointeur[segment];
    } else {
      return null;
    }
  }
  return pointeur;
};

/**
 * Parse la chaîne d'options utilisateur en objet de configuration.
 * @private
 */
const parserOptions_ = (chaineOptions) => {
  const options = {
    noHeaders: false,
    rawHeaders: false
  };
  
  if (!chaineOptions) return options;
  
  const tableauOptions = chaineOptions.split(",").map(s => s.trim());
  if (tableauOptions.includes("noHeaders")) options.noHeaders = true;
  if (tableauOptions.includes("rawHeaders")) options.rawHeaders = true;
  
  return options;
};

/**
 * Formate un en-tête (ex: "user/id" -> "User Id") pour l'affichage.
 * @private
 */
const formaterEnTete_ = (texte) => {
  return texte
    .replace(/[\/\_]/g, " ") // Remplace / et _ par espace
    .replace(/(\w)(\w*)/g, (g0, g1, g2) => g1.toUpperCase() + g2.toLowerCase()); // Title Case
};

