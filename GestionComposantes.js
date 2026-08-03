// ─── Composantes de classement configurables ──────────────────────────────────
//
// Trois tables normalisées dans l'onglet Configuration :
//   A. Définitions       — AA:AL (cols 27-38)
//   B. Tarifs annuels    — AN:AR (cols 40-44)  [AM=39 est un séparateur vide]
//   C. Répartitions      — AT:BB (cols 46-54)  [AS=45 est un séparateur vide]
//
// Ligne 4 : titre de section  |  Ligne 5 : en-têtes  |  Ligne 6+ : données
// Le libellé est la clé stable d'une composante — ne jamais renommer, désactiver.

// ─── Constantes de configuration ─────────────────────────────────────────────

const CONFIG_COMPOSANTES_ = {
  nomFeuille: 'Configuration',

  // Section A — Définitions (AA:AL)
  DEF_COL_DEBUT:        27,  // AA
  DEF_LIGNE_TITRE:       4,
  DEF_LIGNE_ENTETES:     5,
  DEF_PREMIERE_LIGNE:    6,
  DEF_NOMBRE_COLONNES:  12,
  DEF_IDX_ID:            0,  // AA
  DEF_IDX_ACTIF:         1,  // AB
  DEF_IDX_ORDRE:         2,  // AC
  DEF_IDX_LIBELLE:       3,  // AD
  DEF_IDX_NATURE:        4,  // AE
  DEF_IDX_COMPTE:        5,  // AF
  DEF_IDX_SOURCE:        6,  // AG
  DEF_IDX_UGS:           7,  // AH
  DEF_IDX_FORFAIT:       8,  // AI
  DEF_IDX_CPTE_COUT:     9,  // AJ
  DEF_IDX_CPTE_INV:     10,  // AK
  DEF_IDX_NOTES:        11,  // AL

  // Section B — Tarifs annuels (AN:AR)
  TAR_COL_DEBUT:        40,  // AN
  TAR_LIGNE_TITRE:       4,
  TAR_LIGNE_ENTETES:     5,
  TAR_PREMIERE_LIGNE:    6,
  TAR_NOMBRE_COLONNES:   5,
  TAR_IDX_SAISON:        0,  // AN
  TAR_IDX_ID:            1,  // AO
  TAR_IDX_PRIX:          2,  // AP
  TAR_IDX_ACTIF:         3,  // AQ
  TAR_IDX_NOTES:         4,  // AR

  // Section C — Répartitions comptables (AT:BB)
  REP_COL_DEBUT:        46,  // AT
  REP_LIGNE_TITRE:       4,
  REP_LIGNE_ENTETES:     5,
  REP_PREMIERE_LIGNE:    6,
  REP_NOMBRE_COLONNES:   9,
  REP_IDX_SAISON:        0,  // AT
  REP_IDX_ID:            1,  // AU
  REP_IDX_ACTIF:         2,  // AV
  REP_IDX_ORDRE:         3,  // AW
  REP_IDX_PROGRAMME:     4,  // AX
  REP_IDX_PROJET:        5,  // AY
  REP_IDX_PART:          6,  // AZ
  REP_IDX_COMPTE_SUB:    7,  // BA
  REP_IDX_NOTES:         8   // BB
};

const ENTETES_DEF_ = [
  'ID composante', 'Actif', 'Ordre', 'Libellé', 'Nature',
  'Code compte revenu', 'Source du prix', 'UGS inventaire',
  'Crée un forfait', 'Compte coût des ventes', 'Compte inventaire', 'Notes'
];
const ENTETES_TAR_ = ['Saison', 'ID composante', 'Prix', 'Actif', 'Notes'];
const ENTETES_REP_ = [
  'Saison', 'ID composante', 'Actif', 'Ordre',
  'Programme', 'Projet', 'Part', 'Code compte substitut', 'Notes'
];

const NATURES_COMP_VALIDES_   = ['Entrée', 'Forfait', 'Marchandise'];
const SOURCES_PRIX_VALIDES_   = ['Tarif annuel', 'Inventaire'];

// G:H sont des colonnes de PRÉSENTATION HISTORIQUE dans Forfaits — elles ne participent
// ni aux écritures comptables, ni au calcul du compte, ni à la création d'un forfait,
// ni à la répartition du Journal. Ces constantes servent uniquement à calculer les
// colonnes G:H (formules dynamiques) et M:N (migration des lignes L:O existantes).
const PROGRAMMES_PION_    = ['Pion joues-tu?'];
const PROGRAMMES_CARTIER_ = ['Pion joues-tu? – Cartier'];

// ─── Données de migration ─────────────────────────────────────────────────────
// Ordre : ID, Actif, Ordre, Libellé, Nature, Compte, Source, UGS,
//         CréeForfait, CpteCoût, CpteInv, Notes

const DEFS_MIGRATION_ = [
  ['COMP-0001','Oui', 10,'Entrée – Pion joues-tu?','Entrée','4000','Tarif annuel','','Non','','',''],
  ['COMP-0002','Oui', 20,'Entrée – Cartier','Entrée','4000','Tarif annuel','','Non','','',''],
  ['COMP-0003','Oui', 30,'Forfait Pion joues-tu?','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0004','Oui', 40,'Forfait Cartier','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0005','Oui', 50,'Forfait combiné','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0006','Oui', 60,'T-shirt','Marchandise','4020','Inventaire','MERCH-TS-NOIR','Non','5000','1300',''],
  ['COMP-0007','Oui', 70,'Chandail à manches longues','Marchandise','4020','Inventaire','MERCH-LS-NOIR','Non','5000','1300',''],
  ['COMP-0008','Oui', 80,'Hoodie','Marchandise','4020','Inventaire','MERCH-HD-NOIR','Non','5000','1300',''],
  ['COMP-0009','Non', 90,'Forfait Pion des bois – 1 jour','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0010','Non',100,'Forfait Pion des bois – 1 nuit','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0011','Non',110,'Forfait Pion des bois – 2 nuits','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0012','Non',120,'Forfait Pion des bois – 3 nuits','Forfait','4010','Tarif annuel','','Oui','','',''],
  ['COMP-0013','Non',130,'Entrée Pit à pion','Entrée','4000','Tarif annuel','','Non','','','']
];

// Ordre : Saison, ID composante, Prix, Actif, Notes
const TARIFS_MIGRATION_ = [
  [2026,'COMP-0001', 10,'Oui',''],
  [2026,'COMP-0002',  2,'Oui',''],
  [2026,'COMP-0003', 30,'Oui',''],
  [2026,'COMP-0004', 20,'Oui',''],
  [2026,'COMP-0005', 45,'Oui',''],
  [2027,'COMP-0003', 30,'Oui',''],
  [2027,'COMP-0004', 20,'Oui',''],
  [2027,'COMP-0005', 50,'Oui',''],
  [2026,'COMP-0013', 15,'Oui','']
];

// Ordre : Saison, ID composante, Actif, Ordre, Programme, Projet, Part, CpteSubstitut, Notes
const REPS_MIGRATION_ = [
  [2026,'COMP-0001','Oui',1,'Pion joues-tu?','PJT-2026',100,'',''],
  [2026,'COMP-0002','Oui',1,'Pion joues-tu? – Cartier','PJC-2026',100,'',''],
  [2026,'COMP-0003','Oui',1,'Pion joues-tu?','PJT-2026',100,'',''],
  [2026,'COMP-0004','Oui',1,'Pion joues-tu? – Cartier','PJC-2026',100,'',''],
  [2026,'COMP-0005','Oui',1,'Pion joues-tu?','PJT-2026',60,'',''],
  [2026,'COMP-0005','Oui',2,'Pion joues-tu? – Cartier','PJC-2026',40,'',''],
  ['Toutes','COMP-0006','Oui',1,'Marchandise','',100,'',''],
  ['Toutes','COMP-0007','Oui',1,'Marchandise','',100,'',''],
  ['Toutes','COMP-0008','Oui',1,'Marchandise','',100,'',''],
  [2026,'COMP-0009','Non',1,'Pion des bois','',100,'',''],
  [2026,'COMP-0010','Non',1,'Pion des bois','',100,'',''],
  [2026,'COMP-0011','Non',1,'Pion des bois','',100,'',''],
  [2026,'COMP-0012','Non',1,'Pion des bois','',100,'',''],
  [2026,'COMP-0013','Non',1,'Pit à pion','',100,'','']
];

// ─── Installateur public ──────────────────────────────────────────────────────

function installerComposantesConfigurables() {
  var verrou = LockService.getDocumentLock();
  verrou.waitLock(30000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Préparer les trois sections dans Configuration
    preparerSectionsComposantes_(ss);

    // 2. Migrer les données initiales
    var nbDefs = migrerDefinitionsInitiales_(ss);
    var nbTarifs = migrerTarifsInitiaux_(ss);
    var nbReps = migrerRepartitionsInitiales_(ss);

    // 3. Préparer Forfaits : migration L:O des lignes historiques, puis formules G:H
    preparerForfaitsExtendus_(ss);  // throw si ligne historique non associable

    // 3b. Valider les formules Forfaits après migration
    var erreurs = validerForfaitsExtendus_(ss);

    // 4. Validations dropdown dans Configuration
    appliquerValidationsComposantes_(ss);

    // 5. Valider l'intégrité des composantes
    erreurs = erreurs.concat(validerConfigurationComposantes_(ss));

    if (erreurs.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Composantes — ' + erreurs.length + ' erreur(s)',
        'Installation incomplète — corrigez les erreurs avant d\'utiliser le système.\n\n' +
        erreurs.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return { erreurs: erreurs };
    }

    // 6. Mettre à jour les validations dépendantes (Répartition, Règles bancaires, Import bancaire)
    var definitions = chargerDefinitionsComposantes_(ss, { inclureInactives: false });
    var libellesActifs = definitions.map(function(d) { return d.libelle; });
    mettreAJourValidationsComposante_(ss, libellesActifs);

    SpreadsheetApp.flush();

    var actives  = definitions.length;
    var defsAll  = chargerDefinitionsComposantes_(ss, { inclureInactives: true });
    var inactives = defsAll.length - actives;

    SpreadsheetApp.getUi().alert(
      'Composantes de classement installées',
      'Définitions migrées : ' + nbDefs + '\n' +
      'Tarifs migrés : ' + nbTarifs + '\n' +
      'Répartitions migrées : ' + nbReps + '\n' +
      'Définitions actives : ' + actives + '\n' +
      'Définitions inactives : ' + inactives,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return { creees: nbDefs, actives: actives, inactives: inactives, erreurs: [] };

  } finally {
    verrou.releaseLock();
  }
}

// ─── Préparation des trois sections ──────────────────────────────────────────

function preparerSectionsComposantes_(ss) {
  var feuille = ss.getSheetByName(CONFIG_COMPOSANTES_.nomFeuille);
  if (!feuille) throw new Error('Onglet Configuration introuvable.');

  // Assurer assez de colonnes (BB = col 54)
  if (feuille.getMaxColumns() < 54) {
    feuille.insertColumnsAfter(feuille.getMaxColumns(), 54 - feuille.getMaxColumns());
  }

  preparerUneSectionConfig_(feuille,
    'Composantes de classement',
    CONFIG_COMPOSANTES_.DEF_COL_DEBUT,
    CONFIG_COMPOSANTES_.DEF_LIGNE_TITRE,
    CONFIG_COMPOSANTES_.DEF_LIGNE_ENTETES,
    CONFIG_COMPOSANTES_.DEF_NOMBRE_COLONNES,
    ENTETES_DEF_
  );

  preparerUneSectionConfig_(feuille,
    'Tarifs annuels',
    CONFIG_COMPOSANTES_.TAR_COL_DEBUT,
    CONFIG_COMPOSANTES_.TAR_LIGNE_TITRE,
    CONFIG_COMPOSANTES_.TAR_LIGNE_ENTETES,
    CONFIG_COMPOSANTES_.TAR_NOMBRE_COLONNES,
    ENTETES_TAR_
  );

  preparerUneSectionConfig_(feuille,
    'Répartitions comptables',
    CONFIG_COMPOSANTES_.REP_COL_DEBUT,
    CONFIG_COMPOSANTES_.REP_LIGNE_TITRE,
    CONFIG_COMPOSANTES_.REP_LIGNE_ENTETES,
    CONFIG_COMPOSANTES_.REP_NOMBRE_COLONNES,
    ENTETES_REP_
  );
}

function preparerUneSectionConfig_(feuille, titre, colDebut, ligneTitre, ligneEntetes, nbCols, entetes) {
  // Titre
  var cellTitre = feuille.getRange(ligneTitre, colDebut);
  var valTitre  = String(cellTitre.getValue() || '').trim();
  if (!valTitre) {
    cellTitre.setValue(titre)
      .setFontWeight('bold')
      .setBackground('#c9daf8')
      .setFontSize(10);
  } else if (valTitre !== titre) {
    throw new Error(
      'Configuration ligne ' + ligneTitre + ' col ' + colDebut +
      ' : titre « ' + valTitre + ' » incompatible avec « ' + titre + ' ».'
    );
  }

  // En-têtes — idempotent
  var valeursEntetes = feuille
    .getRange(ligneEntetes, colDebut, 1, nbCols)
    .getDisplayValues()[0];

  var premierEntete = String(valeursEntetes[0] || '').trim();

  if (!premierEntete) {
    feuille
      .getRange(ligneEntetes, colDebut, 1, nbCols)
      .setValues([entetes])
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true);
  } else {
    var incompatibles = [];
    entetes.forEach(function(attendu, i) {
      var actuel = String(valeursEntetes[i] || '').trim();
      if (actuel !== attendu) {
        incompatibles.push(
          'col ' + (colDebut + i) + ' : attendu « ' + attendu + ' », trouvé « ' + actuel + ' »'
        );
      }
    });
    if (incompatibles.length > 0) {
      throw new Error(
        'Section « ' + titre + ' » — en-têtes incompatibles :\n' + incompatibles.join('\n') +
        '\nCorrigez manuellement avant de continuer.'
      );
    }
  }
}

// ─── Migration des données initiales ─────────────────────────────────────────

function migrerDefinitionsInitiales_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  var derniereLigne = feuille.getLastRow();

  var idsExistants = new Set();
  var libellesExistants = new Set();

  if (derniereLigne >= C.DEF_PREMIERE_LIGNE) {
    feuille
      .getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT, derniereLigne - C.DEF_PREMIERE_LIGNE + 1, C.DEF_NOMBRE_COLONNES)
      .getDisplayValues()
      .forEach(function(l) {
        var id  = String(l[C.DEF_IDX_ID]      || '').trim();
        var lib = String(l[C.DEF_IDX_LIBELLE] || '').trim();
        if (id)  idsExistants.add(id);
        if (lib) libellesExistants.add(lib);
      });
  }

  var nouvelles = DEFS_MIGRATION_.filter(function(d) {
    return !idsExistants.has(String(d[C.DEF_IDX_ID])) &&
           !libellesExistants.has(String(d[C.DEF_IDX_LIBELLE]));
  });

  if (!nouvelles.length) return 0;

  var ligneInsertion = Math.max(derniereLigne + 1, C.DEF_PREMIERE_LIGNE);
  feuille
    .getRange(ligneInsertion, C.DEF_COL_DEBUT, nouvelles.length, C.DEF_NOMBRE_COLONNES)
    .setValues(nouvelles)
    .setNumberFormat('@');

  // Ordre en numérique
  feuille
    .getRange(ligneInsertion, C.DEF_COL_DEBUT + C.DEF_IDX_ORDRE, nouvelles.length, 1)
    .setNumberFormat('0');

  return nouvelles.length;
}

function migrerTarifsInitiaux_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  var derniereLigne = feuille.getLastRow();

  var clefsExistantes = new Set();

  if (derniereLigne >= C.TAR_PREMIERE_LIGNE) {
    feuille
      .getRange(C.TAR_PREMIERE_LIGNE, C.TAR_COL_DEBUT, derniereLigne - C.TAR_PREMIERE_LIGNE + 1, C.TAR_NOMBRE_COLONNES)
      .getValues()
      .forEach(function(l) {
        var saison = normaliserSaison_(l[C.TAR_IDX_SAISON]);
        var id     = String(l[C.TAR_IDX_ID] || '').trim();
        if (id) clefsExistantes.add(String(saison) + '_' + id);
      });
  }

  var nouveaux = TARIFS_MIGRATION_.filter(function(t) {
    var clef = String(t[C.TAR_IDX_SAISON]) + '_' + String(t[C.TAR_IDX_ID]);
    return !clefsExistantes.has(clef);
  });

  if (!nouveaux.length) return 0;

  var ligneInsertion = Math.max(derniereLigne + 1, C.TAR_PREMIERE_LIGNE);
  feuille
    .getRange(ligneInsertion, C.TAR_COL_DEBUT, nouveaux.length, C.TAR_NOMBRE_COLONNES)
    .setValues(nouveaux);

  // Saison et Prix en numérique
  feuille
    .getRange(ligneInsertion, C.TAR_COL_DEBUT + C.TAR_IDX_SAISON, nouveaux.length, 1)
    .setNumberFormat('0');
  feuille
    .getRange(ligneInsertion, C.TAR_COL_DEBUT + C.TAR_IDX_PRIX, nouveaux.length, 1)
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  feuille
    .getRange(ligneInsertion, C.TAR_COL_DEBUT + C.TAR_IDX_ID, nouveaux.length, 1)
    .setNumberFormat('@');

  return nouveaux.length;
}

function migrerRepartitionsInitiales_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  var derniereLigne = feuille.getLastRow();

  var clefsExistantes = new Set();

  if (derniereLigne >= C.REP_PREMIERE_LIGNE) {
    feuille
      .getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT, derniereLigne - C.REP_PREMIERE_LIGNE + 1, C.REP_NOMBRE_COLONNES)
      .getValues()
      .forEach(function(l) {
        var saison = normaliserSaison_(l[C.REP_IDX_SAISON]);
        var id     = String(l[C.REP_IDX_ID]    || '').trim();
        var ordre  = Number(l[C.REP_IDX_ORDRE] || 0);
        if (id) clefsExistantes.add(String(saison) + '_' + id + '_' + ordre);
      });
  }

  var nouvelles = REPS_MIGRATION_.filter(function(r) {
    var clef = String(r[C.REP_IDX_SAISON]) + '_' + String(r[C.REP_IDX_ID]) + '_' + r[C.REP_IDX_ORDRE];
    return !clefsExistantes.has(clef);
  });

  if (!nouvelles.length) return 0;

  var ligneInsertion = Math.max(derniereLigne + 1, C.REP_PREMIERE_LIGNE);
  feuille
    .getRange(ligneInsertion, C.REP_COL_DEBUT, nouvelles.length, C.REP_NOMBRE_COLONNES)
    .setValues(nouvelles);

  feuille
    .getRange(ligneInsertion, C.REP_COL_DEBUT + C.REP_IDX_ORDRE, nouvelles.length, 1)
    .setNumberFormat('0');
  feuille
    .getRange(ligneInsertion, C.REP_COL_DEBUT + C.REP_IDX_PART, nouvelles.length, 1)
    .setNumberFormat('0');
  feuille
    .getRange(ligneInsertion, C.REP_COL_DEBUT + C.REP_IDX_ID, nouvelles.length, 1)
    .setNumberFormat('@');

  return nouvelles.length;
}

// ─── Chargement des définitions ───────────────────────────────────────────────

// options.inclureInactives = true → retourne aussi les composantes inactives.
function chargerDefinitionsComposantes_(ss, options) {
  var opts = options || {};
  var inclureInactives = Boolean(opts.inclureInactives);
  var C = CONFIG_COMPOSANTES_;

  var feuille = ss.getSheetByName(C.nomFeuille);
  if (!feuille || feuille.getLastRow() < C.DEF_PREMIERE_LIGNE) return [];

  var nbLignes = feuille.getLastRow() - C.DEF_PREMIERE_LIGNE + 1;
  var valeurs  = feuille
    .getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT, nbLignes, C.DEF_NOMBRE_COLONNES)
    .getValues();

  var definitions  = [];
  var idsVus       = new Set();
  var libellesVus  = new Set();

  valeurs.forEach(function(ligne, index) {
    var id = String(ligne[C.DEF_IDX_ID] || '').trim();
    if (!id) return;

    var actif = String(ligne[C.DEF_IDX_ACTIF] || '').trim() === 'Oui';
    if (!actif && !inclureInactives) return;

    var libelle = String(ligne[C.DEF_IDX_LIBELLE] || '').trim();

    if (idsVus.has(id)) {
      throw new Error(
        'Composantes : ID en doublon « ' + id +
        ' » (ligne ' + (C.DEF_PREMIERE_LIGNE + index) + ').'
      );
    }
    if (libelle && libellesVus.has(libelle)) {
      throw new Error(
        'Composantes : libellé en doublon « ' + libelle +
        ' » (ligne ' + (C.DEF_PREMIERE_LIGNE + index) + ').'
      );
    }
    idsVus.add(id);
    if (libelle) libellesVus.add(libelle);

    definitions.push({
      id:               id,
      actif:            actif,
      ordre:            Number(ligne[C.DEF_IDX_ORDRE])    || 0,
      libelle:          libelle,
      nature:           String(ligne[C.DEF_IDX_NATURE]    || '').trim(),
      codeCompte:       String(ligne[C.DEF_IDX_COMPTE]    || '').trim(),
      sourcePrix:       String(ligne[C.DEF_IDX_SOURCE]    || '').trim(),
      ugsInventaire:    String(ligne[C.DEF_IDX_UGS]       || '').trim(),
      creeForfait:      String(ligne[C.DEF_IDX_FORFAIT]   || '').trim() === 'Oui',
      compteVentes:     String(ligne[C.DEF_IDX_CPTE_COUT] || '').trim(),
      compteInventaire: String(ligne[C.DEF_IDX_CPTE_INV]  || '').trim(),
      notes:            String(ligne[C.DEF_IDX_NOTES]     || '').trim()
    });
  });

  definitions.sort(function(a, b) {
    if (a.ordre !== b.ordre) return a.ordre - b.ordre;
    return a.libelle.localeCompare(b.libelle, 'fr');
  });

  return definitions;
}

// ─── Chargement des tarifs ────────────────────────────────────────────────────

function chargerTarifsComposantes_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  if (!feuille || feuille.getLastRow() < C.TAR_PREMIERE_LIGNE) return [];

  var nbLignes = feuille.getLastRow() - C.TAR_PREMIERE_LIGNE + 1;
  var valeurs  = feuille
    .getRange(C.TAR_PREMIERE_LIGNE, C.TAR_COL_DEBUT, nbLignes, C.TAR_NOMBRE_COLONNES)
    .getValues();

  var tarifs = [];

  valeurs.forEach(function(ligne) {
    var id = String(ligne[C.TAR_IDX_ID] || '').trim();
    if (!id) return;

    var saison = normaliserSaison_(ligne[C.TAR_IDX_SAISON]);
    if (!saison) return;

    tarifs.push({
      saison:       saison,
      idComposante: id,
      prix:         Number(ligne[C.TAR_IDX_PRIX])  || 0,
      actif:        String(ligne[C.TAR_IDX_ACTIF] || '').trim() === 'Oui',
      notes:        String(ligne[C.TAR_IDX_NOTES] || '').trim()
    });
  });

  return tarifs;
}

// ─── Chargement des répartitions ──────────────────────────────────────────────

function chargerRepartitionsComposantes_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  if (!feuille || feuille.getLastRow() < C.REP_PREMIERE_LIGNE) return [];

  var nbLignes = feuille.getLastRow() - C.REP_PREMIERE_LIGNE + 1;
  var valeurs  = feuille
    .getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT, nbLignes, C.REP_NOMBRE_COLONNES)
    .getValues();

  var repartitions = [];

  valeurs.forEach(function(ligne) {
    var id = String(ligne[C.REP_IDX_ID] || '').trim();
    if (!id) return;

    var saison = normaliserSaison_(ligne[C.REP_IDX_SAISON]);
    if (saison === null || saison === '') return;

    repartitions.push({
      saison:          saison,
      idComposante:    id,
      actif:           String(ligne[C.REP_IDX_ACTIF]      || '').trim() === 'Oui',
      ordre:           Number(ligne[C.REP_IDX_ORDRE])     || 0,
      programme:       String(ligne[C.REP_IDX_PROGRAMME]  || '').trim(),
      projet:          String(ligne[C.REP_IDX_PROJET]     || '').trim(),
      part:            Number(ligne[C.REP_IDX_PART])      || 0,
      compteSubstitut: String(ligne[C.REP_IDX_COMPTE_SUB] || '').trim(),
      notes:           String(ligne[C.REP_IDX_NOTES]      || '').trim()
    });
  });

  return repartitions;
}

// ─── Compilation pour une saison ─────────────────────────────────────────────

// Retourne { byLibelle: {libelle: compiledComposante}, byId: {id: compiled}, ordered: [] }
//
// compiledComposante = { definition, prix, repartitions, disponible }
//
// options.inclureInactives = true → charge les inactives (pour annulations historiques).
// Active + incomplète → erreur explicite.
// Inactive + incomplète → disponible=false, sans erreur.
function compilerComposantesPourSaison_(ss, dateTransaction, options) {
  var opts = options || {};
  var inclureInactives = Boolean(opts.inclureInactives);
  var date = (dateTransaction instanceof Date) ? dateTransaction : new Date(dateTransaction);
  var saison = date.getFullYear();

  if (isNaN(saison)) throw new Error('Date de transaction invalide pour la compilation des composantes.');

  var definitions   = chargerDefinitionsComposantes_(ss, { inclureInactives: inclureInactives });
  var tarifs        = chargerTarifsComposantes_(ss);
  var allReps       = chargerRepartitionsComposantes_(ss);

  // Charger l'inventaire une seule fois si nécessaire
  var inventaireParUgs = null;
  var besoinInventaire = definitions.some(function(d) { return d.sourcePrix === 'Inventaire'; });
  if (besoinInventaire) {
    inventaireParUgs = chargerInventaireParUgs_(ss);
  }

  var byLibelle = {};
  var byId      = {};
  var ordered   = [];

  definitions.forEach(function(def) {
    var prix = null;
    var repartitions = [];
    var erreur = null;

    if (def.actif) {
      // Composante active : tout doit être complet, sinon erreur bloquante.
      try {
        prix = obtenirPrixComposante_(def, saison, tarifs, inventaireParUgs);
      } catch (e) {
        throw new Error('[' + def.id + ' – ' + def.libelle + '] ' + e.message);
      }
      repartitions = obtenirRepartitionsComposante_(def, saison, allReps);
      if (repartitions.length === 0) {
        throw new Error(
          '[' + def.id + ' – ' + def.libelle + '] ' +
          'Aucune répartition active pour la saison ' + saison + '.'
        );
      }
      var somme = repartitions.reduce(function(s, r) { return s + r.part; }, 0);
      if (Math.abs(somme - 100) >= 0.005) {
        throw new Error(
          '[' + def.id + ' – ' + def.libelle + '] ' +
          'La somme des parts est ' + somme.toFixed(2) + ' % (100 % attendu).'
        );
      }
      // Les Entrées et Forfaits actifs doivent avoir un projet dans chaque répartition.
      // Les Marchandises peuvent avoir un projet vide.
      if (def.nature !== 'Marchandise') {
        repartitions.forEach(function(r) {
          if (!r.projet) {
            throw new Error(
              '[' + def.id + ' – ' + def.libelle + '] ' +
              'La répartition #' + r.ordre + ' (programme « ' + r.programme + ' ») ' +
              'n\'a pas de projet. Configurez un projet ou désactivez la composante.'
            );
          }
        });
      }
    } else {
      // Composante inactive : collecter ce qu'on peut, sans bloquer.
      try { prix = obtenirPrixComposante_(def, saison, tarifs, inventaireParUgs); } catch (e) { prix = null; }
      repartitions = obtenirRepartitionsComposante_(def, saison, allReps);
    }

    var compiled = {
      definition:   def,
      prix:         prix !== null ? prix : 0,
      repartitions: repartitions,
      disponible:   def.actif && erreur === null
    };

    if (def.libelle) byLibelle[def.libelle] = compiled;
    byId[def.id] = compiled;
    ordered.push(compiled);
  });

  return { byLibelle: byLibelle, byId: byId, ordered: ordered };
}

// ─── Résolution du prix ───────────────────────────────────────────────────────

function obtenirPrixComposante_(def, saison, tarifs, inventaireParUgs) {
  if (def.sourcePrix === 'Tarif annuel') {
    var tarif = null;
    tarifs.forEach(function(t) {
      if (t.idComposante === def.id && t.saison === saison && t.actif) {
        tarif = t;
      }
    });
    if (!tarif) {
      throw new Error(
        'Aucun tarif actif trouvé pour la saison ' + saison + '. ' +
        'Ajoutez une ligne dans la table Tarifs annuels.'
      );
    }
    if (tarif.prix < 0) {
      throw new Error('Le tarif ' + saison + ' est négatif (' + tarif.prix + ').');
    }
    return tarif.prix;
  }

  if (def.sourcePrix === 'Inventaire') {
    if (!inventaireParUgs) {
      throw new Error('Inventaire non disponible.');
    }
    if (!def.ugsInventaire) {
      throw new Error('UGS inventaire non configurée.');
    }
    if (!(def.ugsInventaire in inventaireParUgs)) {
      throw new Error(
        'UGS « ' + def.ugsInventaire + ' » introuvable dans Inventaire.'
      );
    }
    return inventaireParUgs[def.ugsInventaire];
  }

  throw new Error(
    'Source du prix inconnue : « ' + def.sourcePrix + ' ». ' +
    'Valeurs permises : ' + SOURCES_PRIX_VALIDES_.join(', ') + '.'
  );
}

// ─── Résolution des répartitions pour une saison ──────────────────────────────

// Règle de priorité : les lignes de l'année exacte remplacent « Toutes ».
// Retourne les répartitions actives triées par ordre.
function obtenirRepartitionsComposante_(def, saison, allRepartitions) {
  var pourCette = allRepartitions.filter(function(r) {
    return r.idComposante === def.id;
  });

  var exactes = pourCette.filter(function(r) {
    return r.saison === saison && r.actif;
  });

  if (exactes.length > 0) {
    return exactes.slice().sort(function(a, b) { return a.ordre - b.ordre; });
  }

  return pourCette
    .filter(function(r) { return r.saison === 'Toutes' && r.actif; })
    .sort(function(a, b) { return a.ordre - b.ordre; });
}

// ─── Chargement de l'inventaire ───────────────────────────────────────────────

function chargerInventaireParUgs_(ss) {
  var inv = ss.getSheetByName('Inventaire');
  var resultat = {};
  if (inv && inv.getLastRow() >= 6) {
    inv.getRange(6, 1, inv.getLastRow() - 5, 11).getValues()
      .forEach(function(ligne) {
        var u = String(ligne[0] || '').trim();
        if (u) resultat[u] = Number(ligne[10]) || 0;  // col K = index 10
      });
  }
  return resultat;
}

// ─── Recherche par libellé ou par ID ─────────────────────────────────────────

function obtenirDefinitionComposante_(definitions, libelle) {
  for (var i = 0; i < definitions.length; i++) {
    if (definitions[i].libelle === libelle) return definitions[i];
  }
  return null;
}

function obtenirDefinitionComposanteParId_(definitions, id) {
  for (var i = 0; i < definitions.length; i++) {
    if (definitions[i].id === id) return definitions[i];
  }
  return null;
}

// ─── Construction des options pour l'interface ────────────────────────────────

// Retourne [{valeur, prix}, ...] pour les composantes actives et disponibles.
function construireOptionsComposantesConfigurables_(ss, dateTransaction) {
  var date = (dateTransaction instanceof Date) ? dateTransaction : new Date(dateTransaction);
  var compiled = compilerComposantesPourSaison_(ss, date, { inclureInactives: false });

  var options = compiled.ordered
    .filter(function(c) { return c.disponible; })
    .map(function(c) { return { valeur: c.definition.libelle, prix: c.prix }; });

  if (!options.length) {
    throw new Error(
      'Aucune composante active et disponible pour la saison ' +
      date.getFullYear() + '. ' +
      'Exécutez « Installer / mettre à jour les composantes ».'
    );
  }

  return options;
}

// ─── Validation de la configuration ──────────────────────────────────────────

// Retourne un tableau de messages d'erreur (vide si tout est valide).
function validerConfigurationComposantes_(ss) {
  var erreurs = [];

  var definitions = chargerDefinitionsComposantes_(ss, { inclureInactives: true });
  var tarifs      = chargerTarifsComposantes_(ss);
  var repartitions = chargerRepartitionsComposantes_(ss);

  // Références externes (lectures groupées)
  var feuille = ss.getSheetByName('Configuration');
  var planComptable = {};
  var programmes    = new Set();
  if (feuille && feuille.getLastRow() >= 6) {
    feuille.getRange(6, 1, feuille.getLastRow() - 5, 3).getDisplayValues()
      .forEach(function(l) {
        var code = String(l[0] || '').trim();
        if (code) planComptable[code] = String(l[1] || '').trim();
      });
    feuille.getRange(6, 8, Math.max(1, feuille.getLastRow() - 5), 1).getDisplayValues()
      .forEach(function(l) { var p = String(l[0] || '').trim(); if (p) programmes.add(p); });
  }

  var ugsSet = new Set();
  var inv = ss.getSheetByName('Inventaire');
  if (inv && inv.getLastRow() >= 6) {
    inv.getRange(6, 1, inv.getLastRow() - 5, 1).getDisplayValues()
      .forEach(function(l) { var u = String(l[0] || '').trim(); if (u) ugsSet.add(u); });
  }

  var projets = new Set();
  var projetsSheet = ss.getSheetByName('Projets');
  if (projetsSheet && projetsSheet.getLastRow() >= 6) {
    projetsSheet.getRange(6, 1, projetsSheet.getLastRow() - 5, 4).getDisplayValues()
      .forEach(function(l) {
        var id = String(l[0] || '').trim();
        if (id && String(l[3] || '').trim() !== 'Exemple') projets.add(id);
      });
  }

  function err(id, libelle, msg) {
    erreurs.push('[' + id + ' – ' + libelle + '] ' + msg);
  }

  // Vérifier les définitions
  definitions.forEach(function(def) {
    if (!def.libelle) { err(def.id, def.id, 'Libellé obligatoire.'); return; }
    if (!def.nature) {
      err(def.id, def.libelle, 'Nature obligatoire.');
    } else if (NATURES_COMP_VALIDES_.indexOf(def.nature) === -1) {
      err(def.id, def.libelle, 'Nature « ' + def.nature + ' » invalide (' + NATURES_COMP_VALIDES_.join('/') + ').');
    }
    if (!def.sourcePrix) {
      err(def.id, def.libelle, 'Source du prix obligatoire.');
    } else if (SOURCES_PRIX_VALIDES_.indexOf(def.sourcePrix) === -1) {
      err(def.id, def.libelle, 'Source du prix « ' + def.sourcePrix + ' » invalide (' + SOURCES_PRIX_VALIDES_.join('/') + ').');
    }
    if (def.codeCompte && Object.keys(planComptable).length > 0 && !planComptable[def.codeCompte]) {
      err(def.id, def.libelle, 'Compte revenu « ' + def.codeCompte + ' » absent du plan comptable.');
    }
    if (def.sourcePrix === 'Inventaire' && !def.ugsInventaire) {
      err(def.id, def.libelle, 'UGS obligatoire pour Source du prix = Inventaire.');
    }
    if (def.ugsInventaire && ugsSet.size > 0 && !ugsSet.has(def.ugsInventaire)) {
      err(def.id, def.libelle, 'UGS « ' + def.ugsInventaire + ' » introuvable dans Inventaire.');
    }
    if (def.nature === 'Marchandise') {
      if (!def.compteVentes)     err(def.id, def.libelle, 'Compte coût des ventes obligatoire pour Marchandise.');
      if (!def.compteInventaire) err(def.id, def.libelle, 'Compte inventaire obligatoire pour Marchandise.');
      if (def.compteVentes && Object.keys(planComptable).length > 0 && !planComptable[def.compteVentes]) {
        err(def.id, def.libelle, 'Compte coût des ventes « ' + def.compteVentes + ' » absent du plan comptable.');
      }
      if (def.compteInventaire && Object.keys(planComptable).length > 0 && !planComptable[def.compteInventaire]) {
        err(def.id, def.libelle, 'Compte inventaire « ' + def.compteInventaire + ' » absent du plan comptable.');
      }
    }
  });

  // Vérifier les comptes des huit composantes historiques
  // Détecte une migration incorrecte (mauvais account copié depuis l'ancienne structure).
  var COMPTES_HISTORIQUES_ = {
    'COMP-0001': '4000', 'COMP-0002': '4000',
    'COMP-0003': '4010', 'COMP-0004': '4010', 'COMP-0005': '4010',
    'COMP-0006': '4020', 'COMP-0007': '4020', 'COMP-0008': '4020'
  };
  definitions.forEach(function(def) {
    var attendu = COMPTES_HISTORIQUES_[def.id];
    if (!attendu) return;
    if (def.codeCompte !== attendu) {
      err(def.id, def.libelle,
        'Compte attendu « ' + attendu + ' », valeur lue « ' + def.codeCompte + ' ». ' +
        'Migration incorrecte — supprimez la ligne et relancez l\'installation.'
      );
    }
  });

  // Vérifier les tarifs (doublons actif par saison+id)
  var clefsActives = {};
  tarifs.forEach(function(t) {
    if (!t.actif) return;
    var clef = String(t.saison) + '_' + t.idComposante;
    if (clefsActives[clef]) {
      erreurs.push(
        'Tarifs : doublon actif pour saison ' + t.saison +
        ' + ID ' + t.idComposante + '.'
      );
    }
    clefsActives[clef] = true;
    if (t.prix < 0) {
      erreurs.push(
        'Tarifs : prix négatif (' + t.prix + ') pour ' + t.idComposante + ' saison ' + t.saison + '.'
      );
    }
  });

  // Index nature des définitions pour la validation des projets
  var natureParId = {};
  definitions.forEach(function(def) { natureParId[def.id] = def.nature; });

  // Vérifier les répartitions (somme par composante+saison)
  var parCompSaison = {};
  repartitions.forEach(function(r) {
    if (!r.actif) return;
    var clef = r.idComposante + '_' + String(r.saison);
    if (!parCompSaison[clef]) parCompSaison[clef] = 0;
    parCompSaison[clef] += r.part;

    if (r.programme && programmes.size > 0 && !programmes.has(r.programme)) {
      erreurs.push(
        'Répartitions [' + r.idComposante + ' saison ' + r.saison + '] : ' +
        'Programme « ' + r.programme + ' » inconnu.'
      );
    }
    if (r.projet && projets.size > 0 && !projets.has(r.projet)) {
      erreurs.push(
        'Répartitions [' + r.idComposante + ' saison ' + r.saison + '] : ' +
        'Projet « ' + r.projet + ' » inconnu.'
      );
    }
    // Les Entrées et Forfaits actifs doivent avoir un projet (pas les Marchandises)
    var nature = natureParId[r.idComposante];
    if (nature && nature !== 'Marchandise' && !r.projet) {
      erreurs.push(
        'Répartitions [' + r.idComposante + ' saison ' + r.saison + '] : ' +
        'Projet obligatoire pour une ' + nature + ' (ordre #' + r.ordre + ').'
      );
    }
  });

  Object.keys(parCompSaison).forEach(function(clef) {
    var somme = parCompSaison[clef];
    if (Math.abs(somme - 100) >= 0.005) {
      erreurs.push(
        'Répartitions [' + clef + '] : somme des parts = ' + somme.toFixed(2) + ' % (100 % attendu).'
      );
    }
  });

  return erreurs;
}

// ─── Validations dropdown dans Configuration ──────────────────────────────────

function appliquerValidationsComposantes_(ss) {
  var C = CONFIG_COMPOSANTES_;
  var feuille = ss.getSheetByName(C.nomFeuille);
  if (!feuille || feuille.getLastRow() < C.DEF_PREMIERE_LIGNE) return;

  var maxLigne = feuille.getMaxRows();
  var nbDef    = maxLigne - C.DEF_PREMIERE_LIGNE + 1;
  var nbTar    = maxLigne - C.TAR_PREMIERE_LIGNE + 1;
  var nbRep    = maxLigne - C.REP_PREMIERE_LIGNE + 1;

  function vlListe(liste, permissif) {
    return SpreadsheetApp.newDataValidation()
      .requireValueInList(liste, true)
      .setAllowInvalid(Boolean(permissif))
      .build();
  }

  if (nbDef > 0) {
    feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_ACTIF,  nbDef, 1).setDataValidation(vlListe(['Oui','Non'], false));
    feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_NATURE, nbDef, 1).setDataValidation(vlListe(NATURES_COMP_VALIDES_, false));
    feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_SOURCE, nbDef, 1).setDataValidation(vlListe(SOURCES_PRIX_VALIDES_, false));
    feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_FORFAIT, nbDef, 1).setDataValidation(vlListe(['Oui','Non'], false));

    var planCodes = [];
    if (feuille.getLastRow() >= 6) {
      feuille.getRange(6, 1, feuille.getLastRow() - 5, 1).getDisplayValues()
        .forEach(function(l) { var c = String(l[0]||'').trim(); if (c) planCodes.push(c); });
    }
    if (planCodes.length > 0) {
      var vCpte = vlListe(planCodes, true);
      feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_COMPTE,   nbDef, 1).setDataValidation(vCpte);
      feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_CPTE_COUT, nbDef, 1).setDataValidation(vCpte);
      feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_CPTE_INV,  nbDef, 1).setDataValidation(vCpte);
    }

    var ugsList = [];
    var invSheet = ss.getSheetByName('Inventaire');
    if (invSheet && invSheet.getLastRow() >= 6) {
      invSheet.getRange(6, 1, invSheet.getLastRow() - 5, 1).getDisplayValues()
        .forEach(function(l) { var u = String(l[0]||'').trim(); if (u) ugsList.push(u); });
    }
    if (ugsList.length > 0) {
      feuille.getRange(C.DEF_PREMIERE_LIGNE, C.DEF_COL_DEBUT + C.DEF_IDX_UGS, nbDef, 1).setDataValidation(vlListe(ugsList, true));
    }
  }

  if (nbTar > 0) {
    feuille.getRange(C.TAR_PREMIERE_LIGNE, C.TAR_COL_DEBUT + C.TAR_IDX_ACTIF, nbTar, 1).setDataValidation(vlListe(['Oui','Non'], false));

    var idsComp = chargerDefinitionsComposantes_(ss, { inclureInactives: true }).map(function(d) { return d.id; });
    if (idsComp.length > 0) {
      feuille.getRange(C.TAR_PREMIERE_LIGNE, C.TAR_COL_DEBUT + C.TAR_IDX_ID, nbTar, 1).setDataValidation(vlListe(idsComp, true));
    }
  }

  if (nbRep > 0) {
    feuille.getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT + C.REP_IDX_ACTIF, nbRep, 1).setDataValidation(vlListe(['Oui','Non'], false));
    var idsComp2 = chargerDefinitionsComposantes_(ss, { inclureInactives: true }).map(function(d) { return d.id; });
    if (idsComp2.length > 0) {
      feuille.getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT + C.REP_IDX_ID, nbRep, 1).setDataValidation(vlListe(idsComp2, true));
    }

    var progList = [];
    if (feuille.getLastRow() >= 6) {
      feuille.getRange(6, 8, Math.max(1, feuille.getLastRow() - 5), 1).getDisplayValues()
        .forEach(function(l) { var p = String(l[0]||'').trim(); if (p) progList.push(p); });
    }
    if (progList.length > 0) {
      feuille.getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT + C.REP_IDX_PROGRAMME, nbRep, 1).setDataValidation(vlListe(progList, true));
    }

    var projSheet = ss.getSheetByName('Projets');
    if (projSheet && projSheet.getLastRow() >= 6) {
      var idsProjets = projSheet.getRange(6, 1, projSheet.getLastRow() - 5, 4)
        .getDisplayValues()
        .filter(function(l) { return l[0] && l[3] !== 'Exemple'; })
        .map(function(l) { return String(l[0]).trim(); });
      if (idsProjets.length > 0) {
        feuille.getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT + C.REP_IDX_PROJET, nbRep, 1).setDataValidation(vlListe(idsProjets, true));
      }
    }

    if (planCodes && planCodes.length > 0) {
      feuille.getRange(C.REP_PREMIERE_LIGNE, C.REP_COL_DEBUT + C.REP_IDX_COMPTE_SUB, nbRep, 1).setDataValidation(vlListe(planCodes, true));
    }
  }
}

// ─── Mise à jour des validations dépendantes ─────────────────────────────────

// Met à jour les dropdowns composantes dans Répartition (col F),
// Règles bancaires (col H) et Import bancaire (col U).
function mettreAJourValidationsComposante_(ss, libellesActifs) {
  if (!libellesActifs.length) return;

  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(libellesActifs, true)
    .setAllowInvalid(true)
    .build();

  var repartition = ss.getSheetByName('Répartition');
  if (repartition && repartition.getMaxRows() >= 6) {
    repartition.getRange(6, 6, repartition.getMaxRows() - 5, 1).setDataValidation(validation);
  }

  var regles = ss.getSheetByName('Règles bancaires');
  if (regles && regles.getMaxRows() >= 6) {
    regles.getRange(6, 8, regles.getMaxRows() - 5, 1).setDataValidation(validation);
  }

  var importBancaire = ss.getSheetByName('Import bancaire');
  if (importBancaire && importBancaire.getMaxRows() >= 6) {
    importBancaire.getRange(6, 21, importBancaire.getMaxRows() - 5, 1).setDataValidation(validation);
  }
}

// ─── Préparation des colonnes étendues de Forfaits ───────────────────────────

// Prépare les colonnes L:O de la feuille Forfaits et installe/répare les formules G:H.
// Migration sécurisée de Forfaits :
// 1. Vérifie séparément L5:O5 (incompatible = erreur).
// 2. Pour chaque ligne historique (A non vide, L vide) : calcule L:O en mémoire depuis
//    les composantes. Si une ligne ne peut être associée → erreur avant toute écriture.
// 3. Écrit L:O manquantes, flush, relit pour confirmer.
// 4. Seulement ensuite installe/répare G:H sur toutes les lignes.
// Préserve toute valeur L:O existante non vide.
function preparerForfaitsExtendus_(ss) {
  var feuille = ss.getSheetByName('Forfaits');
  if (!feuille) return;

  // 1. Assurer au moins 15 colonnes (O = col 15)
  if (feuille.getMaxColumns() < 15) {
    feuille.insertColumnsAfter(feuille.getMaxColumns(), 15 - feuille.getMaxColumns());
  }

  if (feuille.getLastRow() < 5) return;

  // 2. Vérifier séparément chaque en-tête L5:O5
  var ENTETES_LO_ = ['ID composante', 'Part Pion joues-tu? (%)', 'Part Cartier (%)', 'Projet(s)'];
  var COLS_LO_    = ['L', 'M', 'N', 'O'];
  var row5LO      = feuille.getRange(5, 12, 1, 4).getValues()[0];
  var vals5       = row5LO.map(function(v) { return String(v || '').trim(); });
  var tousVides5  = vals5.every(function(v) { return !v; });
  var tousOK5     = vals5.every(function(v, i) { return v === ENTETES_LO_[i]; });

  if (!tousVides5 && !tousOK5) {
    var pbHeaders = [];
    vals5.forEach(function(v, i) {
      if (v && v !== ENTETES_LO_[i]) {
        pbHeaders.push(COLS_LO_[i] + '5 = «' + v + '» (attendu «' + ENTETES_LO_[i] + '»)');
      } else if (!v) {
        pbHeaders.push(COLS_LO_[i] + '5 vide (attendu «' + ENTETES_LO_[i] + '»)');
      }
    });
    throw new Error(
      'En-têtes Forfaits L:O incompatibles — ' + pbHeaders.join('; ') +
      '. Corrigez manuellement avant de relancer l\'installation.'
    );
  }

  if (tousVides5) {
    feuille.getRange(5, 12, 1, 4)
      .setValues([ENTETES_LO_])
      .setFontWeight('bold')
      .setBackground('#f1f3f4');
  }

  // 3. Format numérique M/N (entiers 0-100, pas de format pourcentage Sheets)
  var maxLigne = feuille.getMaxRows();
  if (maxLigne >= 6) {
    feuille.getRange(6, 13, maxLigne - 5, 2).setNumberFormat('0.##');
  }

  // 4. Migration L:O pour les lignes historiques (A non vide, L vide)
  //    Tout calculer en mémoire — aucune écriture avant la validation complète.
  var derniereLigneData = feuille.getLastRow();
  if (derniereLigneData < 6) {
    poserFormulasGH_(feuille, maxLigne);
    return;
  }

  var nbDataRows   = derniereLigneData - 5;
  var dataForfaits = feuille.getRange(6, 1, nbDataRows, 15).getValues();

  // Lectures groupées : composantes (incl. inactives) et répartitions
  var allDefs = chargerDefinitionsComposantes_(ss, { inclureInactives: true });
  var allReps = chargerRepartitionsComposantes_(ss);

  var defParLibelle = {};
  allDefs.forEach(function(d) { if (d.libelle) defParLibelle[d.libelle] = d; });

  // Passe 1 : validation complète en mémoire
  var fileEcritures  = [];
  var incompatibles  = [];

  dataForfaits.forEach(function(row, i) {
    var idForfait = String(row[0] || '').trim();  // col A
    if (!idForfait) return;

    var lVal = String(row[11] || '').trim();  // col L
    if (lVal) return;  // déjà migré, conserver

    var libelle = String(row[3] || '').trim();   // col D
    var def     = defParLibelle[libelle];
    if (!def) {
      incompatibles.push({ numeroLigne: i + 6, libelle: libelle || '(vide)' });
      return;
    }

    // Saison : col K, puis année de la date en col B
    var saisonBrute = row[10];  // col K
    var saison = normaliserSaison_(saisonBrute);
    if (saison === null || saison === 'Toutes') {
      var dateB = row[1];  // col B
      if (dateB instanceof Date && !isNaN(dateB)) saison = dateB.getFullYear();
      else saison = normaliserSaison_(String(dateB || '').trim());
    }
    if (typeof saison !== 'number' || saison < 2000) {
      incompatibles.push({ numeroLigne: i + 6, libelle: libelle, raison: 'Saison indéterminée' });
      return;
    }

    fileEcritures.push({ rowIndex: i, numeroLigne: i + 6, def: def, saison: saison });
  });

  if (incompatibles.length > 0) {
    throw new Error(
      'Forfaits — ' + incompatibles.length + ' ligne(s) historique(s) ne peuvent pas être migrées : ' +
      incompatibles.map(function(x) {
        return 'ligne ' + x.numeroLigne + ' (« ' + x.libelle + '»' +
               (x.raison ? ' — ' + x.raison : '') + ')';
      }).join(', ') +
      '. Vérifiez les libellés (col D) et la saison (col K), puis relancez l\'installation.'
    );
  }

  // Passe 2 : calcul L:O (PROGRAMMES_PION_/PROGRAMMES_CARTIER_ = colonnes présentation)
  var valeurs = fileEcritures.map(function(item) {
    var reps = obtenirRepartitionsComposante_(item.def, item.saison, allReps);
    var partPion = 0, partCartier = 0, projets = [];
    reps.forEach(function(r) {
      if (PROGRAMMES_PION_.indexOf(r.programme) !== -1)    partPion    += r.part;
      if (PROGRAMMES_CARTIER_.indexOf(r.programme) !== -1) partCartier += r.part;
      if (r.projet && projets.indexOf(r.projet) === -1)   projets.push(r.projet);
    });
    return {
      numeroLigne: item.numeroLigne,
      idComp:      item.def.id,
      partPion:    partPion,
      partCartier: partCartier,
      projetsStr:  projets.join(', ')
    };
  });

  // Passe 3 : écriture, flush, confirmation groupée
  if (valeurs.length > 0) {
    valeurs.forEach(function(e) {
      feuille.getRange(e.numeroLigne, 12).setValue(e.idComp);
      feuille.getRange(e.numeroLigne, 13).setValue(e.partPion);
      feuille.getRange(e.numeroLigne, 14).setValue(e.partCartier);
      feuille.getRange(e.numeroLigne, 15).setValue(e.projetsStr);
    });
    SpreadsheetApp.flush();

    // Relecture groupée de la colonne L pour confirmer
    var confirmL    = feuille.getRange(6, 12, nbDataRows, 1).getValues();
    var confirmErrs = [];
    valeurs.forEach(function(e) {
      var lConfirm = String(confirmL[e.numeroLigne - 6][0] || '').trim();
      if (lConfirm !== e.idComp) confirmErrs.push('ligne ' + e.numeroLigne);
    });
    if (confirmErrs.length > 0) {
      throw new Error(
        'Forfaits — migration L:O non confirmée pour : ' + confirmErrs.join(', ') + '.'
      );
    }
  }

  // 5. Installer/réparer G:H après migration entièrement validée
  poserFormulasGH_(feuille, maxLigne);
}

// Pose les formules G:H sur toutes les lignes données (installe et répare).
// G:H = colonnes de présentation historiques, basées sur M:N (entiers 0-100).
function poserFormulasGH_(feuille, maxLigne) {
  if (maxLigne < 6) return;
  var fmtG = '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IFERROR(RC5*RC13/100,0)))';
  var fmtH = '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IFERROR(RC5*RC14/100,0)))';
  feuille.getRange(6, 7, maxLigne - 5, 1).setFormulaR1C1(fmtG);
  feuille.getRange(6, 8, maxLigne - 5, 1).setFormulaR1C1(fmtH);
}

// ─── Validation des formules Forfaits après installation ─────────────────────

// Vérifie pour chaque ligne Forfaits existante (A non vide) :
// - L non vide;
// - M et N numériques dans [0, 100];
// - G ≈ E × M / 100 (sauf Annulé/Remboursé).
// M + N ≠ 100 est autorisé (forfait Pion des bois : M=0, N=0).
// Retourne un tableau de messages d'erreur (vide si tout est valide).
function validerForfaitsExtendus_(ss) {
  var feuille = ss.getSheetByName('Forfaits');
  if (!feuille || feuille.getLastRow() < 6) return [];

  SpreadsheetApp.flush();
  var nbRows = feuille.getLastRow() - 5;
  // Colonnes : A(1), E(5), G(7), H(8), I(9), L(12), M(13), N(14)
  var data = feuille.getRange(6, 1, nbRows, 14).getValues();
  var msgs = [];

  data.forEach(function(row, i) {
    var idForfait = String(row[0] || '').trim();
    if (!idForfait) return;

    var ligne   = i + 6;
    var montant = Number(row[4] || 0);   // col E
    var g       = Number(row[6] || 0);   // col G
    var h       = Number(row[7] || 0);   // col H
    var statut  = String(row[8] || '').trim();  // col I
    var lVal    = String(row[11] || '').trim(); // col L
    var mVal    = row[12];                // col M
    var nVal    = row[13];                // col N

    if (!lVal) {
      msgs.push('Forfaits ligne ' + ligne + ' : L vide alors que A = «' + idForfait + '».');
    }

    var mNum = Number(mVal);
    var nNum = Number(nVal);
    if (mVal === '' || mVal === null || isNaN(mNum)) {
      msgs.push('Forfaits ligne ' + ligne + ' : M non numérique (valeur : «' + mVal + '»).');
    } else if (mNum < 0 || mNum > 100) {
      msgs.push('Forfaits ligne ' + ligne + ' : M = ' + mNum + ' hors plage [0, 100].');
    }
    if (nVal === '' || nVal === null || isNaN(nNum)) {
      msgs.push('Forfaits ligne ' + ligne + ' : N non numérique (valeur : «' + nVal + '»).');
    } else if (nNum < 0 || nNum > 100) {
      msgs.push('Forfaits ligne ' + ligne + ' : N = ' + nNum + ' hors plage [0, 100].');
    }

    if (statut !== 'Annulé' && statut !== 'Remboursé') {
      var attenduG = Math.round(montant * mNum * 100) / 10000;
      var attenduH = Math.round(montant * nNum * 100) / 10000;
      if (Math.abs(g - attenduG) > 0.005) {
        msgs.push('Forfaits ligne ' + ligne + ' : G = ' + g + ' ≠ E×M/100 = ' + attenduG + '.');
      }
      if (Math.abs(h - attenduH) > 0.005) {
        msgs.push('Forfaits ligne ' + ligne + ' : H = ' + h + ' ≠ E×N/100 = ' + attenduH + '.');
      }
    }
  });

  return msgs;
}

// ─── Helpers internes ────────────────────────────────────────────────────────

// Normalise une valeur saison : retourne un nombre (année) ou la chaîne 'Toutes'.
function normaliserSaison_(valeur) {
  if (typeof valeur === 'number' && !isNaN(valeur) && valeur > 1000) {
    return Math.round(valeur);
  }
  var s = String(valeur || '').trim();
  if (/^\d{4}$/.test(s)) return Number(s);
  if (s === 'Toutes') return 'Toutes';
  return null;
}
