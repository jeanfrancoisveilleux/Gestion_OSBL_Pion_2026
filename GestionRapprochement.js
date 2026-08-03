// ─── Rapprochement bancaire mensuel — Phase 3B ───────────────────────────────

const CONFIG_RAPPROCHEMENT = {
  nomFeuille: 'Rapprochement bancaire',
  ligneEntetes: 5,
  premiereLigne: 6,
  COL_MOIS: 1,
  COL_SOL_OUV: 2,
  COL_ENTREES: 3,
  COL_SORTIES: 4,
  COL_SOL_FIN: 5,
  COL_SOL_COMPT: 6,
  COL_EQUILIBRE: 7,
  COL_SANS_LIEN: 8,
  COL_A_CLASSER: 9,
  COL_ECART: 10,
  COL_STATUT: 11,
  COL_DATE_RAPPR: 12,
  COL_NOTES: 13,
  NOMBRE_COLONNES: 13,
  COMPTE_BANCAIRE: '1000',
  TOLERANCE: 0.005,
  ANNEE_DEBUT: 2026
};

// ─── Fonctions publiques ──────────────────────────────────────────────────────

function installerRapprochementBancaire() {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. ' +
      'Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    preparerOngletRapprochement_(ss);

    SpreadsheetApp.getUi().alert(
      'L’onglet « Rapprochement bancaire » est prêt.\n\n' +
      'Utilisez « Ouvrir le rapprochement bancaire » pour calculer ' +
      'et consulter les données mensuelles.'
    );
  } finally {
    verrou.releaseLock();
  }
}

function ouvrirRapprochementBancaire() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  preparerOngletRapprochement_(ss);

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('RapprochementBancaire')
      .setWidth(1100).setHeight(780),
    'Rapprochement bancaire'
  );
}

function obtenirDonneesRapprochementBancaire() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Années tirées directement de Import bancaire — source primaire indépendante
  // de l'état de l'onglet Rapprochement bancaire. Permet de peupler le sélecteur
  // dès la première ouverture, avant toute actualisation.
  const annees = obtenirAnneesDepuisImportBancaire_(ss);
  const anneeDefaut = annees.length > 0
    ? annees[0]
    : new Date().getFullYear();

  const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);

  if (!feuille || feuille.getLastRow() < CONFIG_RAPPROCHEMENT.premiereLigne) {
    return {
      lignes: [],
      annees: annees,
      anneeDefaut: anneeDefaut
    };
  }

  const nbLignes =
    feuille.getLastRow() - CONFIG_RAPPROCHEMENT.premiereLigne + 1;

  const valeurs = feuille
    .getRange(
      CONFIG_RAPPROCHEMENT.premiereLigne,
      1,
      nbLignes,
      CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
    )
    .getValues();

  const lignes = [];

  valeurs.forEach(function(ligne) {
    const cle = normaliserCleMoisRappr_(ligne[0]);

    if (!cle) {
      return;
    }

    const parties = cle.split('-');
    const annee = Number(parties[0]);
    const mois = Number(parties[1]);

    lignes.push({
      cle: cle,
      annee: annee,
      mois: mois,
      libelle: libelleMoisRappr_(annee, mois),
      soldeBancaireOuverture: Number(ligne[1] || 0),
      entreesBancaires: Number(ligne[2] || 0),
      sortiesBancaires: Number(ligne[3] || 0),
      soldeBancaireFin: Number(ligne[4] || 0),
      soldeComptable: Number(ligne[5] || 0),
      journalEquilibre: String(ligne[6] || '').trim() === 'Oui',
      ecritures1000SansLien: Number(ligne[7] || 0),
      transactionsAClasser: Number(ligne[8] || 0),
      ecart: Number(ligne[9] || 0),
      statut: String(ligne[10] || '').trim(),
      dateRapprochement: formaterDateRappr_(ligne[11]),
      notes: String(ligne[12] || ''),
      estCloture: ligne[11] instanceof Date && !isNaN(ligne[11].getTime())
    });
  });

  return { lignes: lignes, annees: annees, anneeDefaut: anneeDefaut };
}

function actualiserRapprochementBancaire() {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. ' +
      'Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    preparerOngletRapprochement_(ss);
    calculerEtEcrireRapprochement_(ss);
    return obtenirDonneesRapprochementBancaire();
  } finally {
    verrou.releaseLock();
  }
}

function auditerClotureMoisComptable(cleMois) {
  const cleStr = String(cleMois || '').trim();

  if (!cleStr.match(/^\d{4}-\d{2}$/)) {
    throw new Error(
      'La clé de mois « ' + cleStr +
      ' » est invalide. Format attendu : YYYY-MM.'
    );
  }

  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. ' +
      'Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return auditerClotureMoisComptableSansVerrou_(ss, cleStr);
  } finally {
    verrou.releaseLock();
  }
}

// Effectue le recalcul et retourne le résultat de l'audit sans acquérir de verrou.
// Appelé par auditerClotureMoisComptable, fermerMoisComptable et reouvrirMoisComptable.
function auditerClotureMoisComptableSansVerrou_(ss, cleStr) {
  const parties = cleStr.split('-');
  const annee = Number(parties[0]);
  const mois = Number(parties[1]);
  const libelleMois = libelleMoisRappr_(annee, mois);

  // 1. Préparer l'onglet et recalculer
  preparerOngletRapprochement_(ss);
  calculerEtEcrireRapprochement_(ss);

  // 2. Lire les données recalculées
  const donnees = obtenirDonneesRapprochementBancaire();
  const ligneMois = donnees.lignes.filter(function(l) {
    return l.cle === cleStr;
  })[0] || null;

  // 3. Construire les contrôles
  const controles = [];

  // PERIODE_EXISTANTE
  const periodeExiste = ligneMois !== null;
  controles.push({
    code: 'PERIODE_EXISTANTE',
    libelle: 'Période présente dans le rapprochement',
    ok: periodeExiste,
    valeur: periodeExiste ? 'Oui' : 'Non',
    detail: periodeExiste
      ? null
      : 'Le mois ' + cleStr +
        ' ne figure pas dans les données recalculées. ' +
        'Vérifiez que des transactions existent dans Import bancaire pour ce mois.'
  });

  if (!periodeExiste) {
    return {
      cleMois: cleStr,
      libelleMois: libelleMois,
      admissible: false,
      dateRapprochement: '',
      notes: '',
      controles: controles
    };
  }

  // IMPORT_CLASSE
  const aClasser = ligneMois.transactionsAClasser;
  controles.push({
    code: 'IMPORT_CLASSE',
    libelle: 'Transactions bancaires classées',
    ok: aClasser === 0,
    valeur: aClasser === 0
      ? 'Toutes classées'
      : aClasser + ' transaction(s) encore à classer',
    detail: null
  });

  // JOURNAL_EQUILIBRE
  const journalEq = ligneMois.journalEquilibre;
  controles.push({
    code: 'JOURNAL_EQUILIBRE',
    libelle: 'Journal comptable équilibré',
    ok: journalEq,
    valeur: journalEq ? 'Équilibré' : 'Non équilibré',
    detail: journalEq
      ? null
      : 'La somme des débits ne correspond pas à la somme des crédits pour ce mois.'
  });

  // ECRITURES_LIEES
  const sansLien = ligneMois.ecritures1000SansLien;
  controles.push({
    code: 'ECRITURES_LIEES',
    libelle: 'Écritures compte 1000 liées à une transaction bancaire',
    ok: sansLien === 0,
    valeur: sansLien === 0
      ? 'Toutes liées'
      : sansLien + ' écriture(s) sans lien bancaire',
    detail: null
  });

  // SOLDE_CONCORDE
  const ecart = ligneMois.ecart;
  const soldeConcorde = Math.abs(ecart) < CONFIG_RAPPROCHEMENT.TOLERANCE;
  controles.push({
    code: 'SOLDE_CONCORDE',
    libelle: 'Solde bancaire concordant avec le solde comptable 1000',
    ok: soldeConcorde,
    valeur: 'Écart : ' +
      (ecart >= 0 ? '+' : '') +
      ecart.toFixed(2) + ' $',
    detail: soldeConcorde
      ? null
      : 'L\'écart dépasse la tolérance de ' +
        CONFIG_RAPPROCHEMENT.TOLERANCE + ' $. ' +
        'Solde bancaire fin : ' + ligneMois.soldeBancaireFin.toFixed(2) +
        ' $ — Solde comptable : ' + ligneMois.soldeComptable.toFixed(2) + ' $.'
  });

  // PIECES_DEPENSES
  const depenses = chargerDepensesMois_Rappr_(ss, cleStr);
  const depensesSansPiece = depenses.filter(function(d) {
    return d.pieceCtrl !== 'OK';
  });
  const LIMITE_DETAIL = 20;
  var detailPieces = null;

  if (depensesSansPiece.length > 0) {
    var ids = depensesSansPiece
      .slice(0, LIMITE_DETAIL)
      .map(function(d) { return d.idTransaction; })
      .join(', ');

    if (depensesSansPiece.length > LIMITE_DETAIL) {
      ids += ' … (+' + (depensesSansPiece.length - LIMITE_DETAIL) + ' autres)';
    }

    detailPieces = ids;
  }

  controles.push({
    code: 'PIECES_DEPENSES',
    libelle: 'Pièces justificatives des dépenses',
    ok: depensesSansPiece.length === 0,
    valeur: depenses.length === 0
      ? 'Aucune dépense ce mois'
      : depensesSansPiece.length === 0
        ? 'Toutes les dépenses (' + depenses.length + ') ont une pièce'
        : depensesSansPiece.length + ' dépense(s) sur ' + depenses.length + ' sans pièce',
    detail: detailPieces
  });

  // MOIS_PRECEDENTS_CLOTURES
  const moisPrecedentsOuverts = [];
  const estPremierMois =
    (annee === CONFIG_RAPPROCHEMENT.ANNEE_DEBUT && mois === 1);

  if (!estPremierMois) {
    donnees.lignes.forEach(function(l) {
      if (l.cle < cleStr && !l.estCloture) {
        moisPrecedentsOuverts.push(l.libelle);
      }
    });
  }

  controles.push({
    code: 'MOIS_PRECEDENTS_CLOTURES',
    libelle: 'Mois précédents clôturés',
    ok: moisPrecedentsOuverts.length === 0,
    valeur: moisPrecedentsOuverts.length === 0
      ? (estPremierMois ? 'Aucun mois précédent' : 'Tous clôturés')
      : moisPrecedentsOuverts.length + ' mois précédent(s) non clôturé(s)',
    detail: moisPrecedentsOuverts.length > 0
      ? moisPrecedentsOuverts.join(', ')
      : null
  });

  const admissible = controles.every(function(c) { return c.ok; });

  return {
    cleMois: cleStr,
    libelleMois: libelleMois,
    admissible: admissible,
    dateRapprochement: ligneMois.dateRapprochement,
    notes: ligneMois.notes,
    controles: controles
  };
}

function fermerMoisComptable(cleMois, note) {
  const cleStr = String(cleMois || '').trim();

  if (!cleStr.match(/^\d{4}-\d{2}$/)) {
    throw new Error(
      'La clé de mois « ' + cleStr +
      ' » est invalide. Format attendu : YYYY-MM.'
    );
  }

  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. ' +
      'Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Audit complet sous verrou — inclut recalcul
    const audit = auditerClotureMoisComptableSansVerrou_(ss, cleStr);

    // Refus si période inexistante
    const ctrlPeriode = audit.controles.filter(function(c) {
      return c.code === 'PERIODE_EXISTANTE';
    })[0];
    if (!ctrlPeriode || !ctrlPeriode.ok) {
      throw new Error(
        'Le mois ' + cleStr +
        ' n\'existe pas dans le rapprochement recalculé.'
      );
    }

    // Refus si déjà fermé
    if (audit.dateRapprochement) {
      throw new Error(
        'Le mois ' + audit.libelleMois +
        ' est déjà clôturé (le ' + audit.dateRapprochement + ').'
      );
    }

    // Refus si un contrôle échoue
    const echoues = audit.controles.filter(function(c) { return !c.ok; });
    if (echoues.length > 0) {
      throw new Error(
        'Impossible de clôturer ' + audit.libelleMois + '. ' +
        echoues.length + ' contrôle(s) échoué(s) :\n' +
        echoues.map(function(c) {
          return '• ' + c.libelle + ' — ' + c.valeur;
        }).join('\n')
      );
    }

    const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);
    const numLigne = trouverLigneCleMoisRappr_(feuille, cleStr);

    if (!numLigne) {
      throw new Error('Ligne introuvable pour le mois ' + cleStr + '.');
    }

    const tz = ss.getSpreadsheetTimeZone();
    const maintenant = new Date();
    const horodatage = Utilities.formatDate(
      maintenant, tz, 'yyyy-MM-dd HH:mm'
    );

    const noteExistante = String(
      feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_NOTES).getValue() || ''
    );
    const noteUtil = String(note || '').trim();
    const entree = '[' + horodatage + '] Clôture du mois' +
      (noteUtil ? ' — ' + noteUtil : '');
    const nouvelleNote = noteExistante
      ? noteExistante + '\n' + entree
      : entree;

    feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR)
      .setValue(maintenant);
    feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_NOTES)
      .setValue(nouvelleNote);

    SpreadsheetApp.flush();
    calculerEtEcrireRapprochement_(ss);

    const parties = cleStr.split('-');
    return {
      succes: true,
      libelleMois: libelleMoisRappr_(Number(parties[0]), Number(parties[1])),
      horodatage: horodatage,
      donnees: obtenirDonneesRapprochementBancaire()
    };

  } finally {
    verrou.releaseLock();
  }
}

function reouvrirMoisComptable(cleMois, motif) {
  const cleStr = String(cleMois || '').trim();

  if (!cleStr.match(/^\d{4}-\d{2}$/)) {
    throw new Error(
      'La clé de mois « ' + cleStr +
      ' » est invalide. Format attendu : YYYY-MM.'
    );
  }

  const motifStr = String(motif || '').trim();
  if (!motifStr) {
    throw new Error('Un motif de réouverture est obligatoire.');
  }

  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. ' +
      'Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);

    if (!feuille || feuille.getLastRow() < CONFIG_RAPPROCHEMENT.premiereLigne) {
      throw new Error(
        'L\'onglet « Rapprochement bancaire » est vide ou introuvable.'
      );
    }

    const numLigne = trouverLigneCleMoisRappr_(feuille, cleStr);

    if (!numLigne) {
      throw new Error(
        'Le mois ' + cleStr + ' est introuvable dans le rapprochement.'
      );
    }

    // Vérifier que le mois est effectivement fermé
    const dateL = feuille
      .getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR)
      .getValue();
    const estFerme = dateL instanceof Date && !isNaN(dateL.getTime());

    if (!estFerme) {
      const p = cleStr.split('-');
      throw new Error(
        'Le mois ' +
        libelleMoisRappr_(Number(p[0]), Number(p[1])) +
        ' n\'est pas clôturé.'
      );
    }

    // Vérifier qu'aucun mois ultérieur n'est clôturé
    const nbLignes =
      feuille.getLastRow() - CONFIG_RAPPROCHEMENT.premiereLigne + 1;
    const toutesValeurs = feuille
      .getRange(
        CONFIG_RAPPROCHEMENT.premiereLigne,
        1,
        nbLignes,
        CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
      )
      .getValues();

    const moisFermesUlterieurs = [];
    toutesValeurs.forEach(function(ligneData) {
      const cleRow = normaliserCleMoisRappr_(ligneData[0]);
      if (!cleRow || cleRow <= cleStr) return;
      const dateLRow = ligneData[CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR - 1];
      if (dateLRow instanceof Date && !isNaN(dateLRow.getTime())) {
        const p2 = cleRow.split('-');
        moisFermesUlterieurs.push(
          libelleMoisRappr_(Number(p2[0]), Number(p2[1]))
        );
      }
    });

    if (moisFermesUlterieurs.length > 0) {
      const p = cleStr.split('-');
      throw new Error(
        'Impossible de réouvrir ' +
        libelleMoisRappr_(Number(p[0]), Number(p[1])) +
        '. Des mois ultérieurs sont clôturés : ' +
        moisFermesUlterieurs.join(', ') +
        '. Réouvrez-les en ordre inverse (du plus récent au plus ancien).'
      );
    }

    const tz = ss.getSpreadsheetTimeZone();
    const horodatage = Utilities.formatDate(
      new Date(), tz, 'yyyy-MM-dd HH:mm'
    );

    const noteExistante = String(
      feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_NOTES).getValue() || ''
    );
    const entree = '[' + horodatage + '] Réouverture du mois — ' + motifStr;
    const nouvelleNote = noteExistante
      ? noteExistante + '\n' + entree
      : entree;

    feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_NOTES)
      .setValue(nouvelleNote);
    feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR)
      .clearContent();

    SpreadsheetApp.flush();

    // Vérification explicite : col M doit contenir l'entrée de réouverture
    // ET préserver les notes de clôture préexistantes (jamais clearContent sur M).
    const noteApres = String(
      feuille.getRange(numLigne, CONFIG_RAPPROCHEMENT.COL_NOTES).getValue() || ''
    );
    if (noteApres.indexOf(horodatage) === -1) {
      throw new Error(
        'Anomalie interne : l\'entrée de réouverture (' + horodatage +
        ') est absente de la col M du mois ' + cleStr +
        '. L\'opération a été interrompue.'
      );
    }

    calculerEtEcrireRapprochement_(ss);

    const parties = cleStr.split('-');
    return {
      succes: true,
      libelleMois: libelleMoisRappr_(Number(parties[0]), Number(parties[1])),
      donnees: obtenirDonneesRapprochementBancaire()
    };

  } finally {
    verrou.releaseLock();
  }
}

function trouverLigneCleMoisRappr_(feuille, cleMois) {
  if (!feuille || feuille.getLastRow() < CONFIG_RAPPROCHEMENT.premiereLigne) {
    return null;
  }

  const nbLignes = feuille.getLastRow() - CONFIG_RAPPROCHEMENT.premiereLigne + 1;
  const valeurs = feuille
    .getRange(CONFIG_RAPPROCHEMENT.premiereLigne, 1, nbLignes, 1)
    .getValues();

  for (var i = 0; i < valeurs.length; i++) {
    if (normaliserCleMoisRappr_(valeurs[i][0]) === cleMois) {
      return CONFIG_RAPPROCHEMENT.premiereLigne + i;
    }
  }

  return null;
}

// ─── Préparation de l'onglet ──────────────────────────────────────────────────

function preparerOngletRapprochement_(ss) {
  let feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);

  if (!feuille) {
    feuille = ss.insertSheet(CONFIG_RAPPROCHEMENT.nomFeuille);
  }

  feuille.getRange('A1')
    .setValue('Rapprochement bancaire mensuel')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#e8f0fe');

  feuille.getRange('A3')
    .setValue('Phase 3B — Rapprochement et clôture mensuelle')
    .setFontColor('#5f6368');

  const entetes = [
    'Mois',
    'Solde bancaire ouverture',
    'Entrées bancaires',
    'Sorties bancaires',
    'Solde bancaire fin',
    'Solde comptable 1000',
    'Journal équilibré',
    'Écritures 1000 sans lien',
    'Transactions à classer',
    'Écart bancaire/comptable',
    'Statut',
    'Date de rapprochement',
    'Notes'
  ];

  feuille
    .getRange(CONFIG_RAPPROCHEMENT.ligneEntetes, 1, 1, CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES)
    .setValues([entetes])
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle');

  feuille.setFrozenRows(5);

  const largeurs = [115, 140, 130, 130, 130, 150, 110, 150, 140, 160, 150, 150, 210];

  largeurs.forEach(function(largeur, index) {
    feuille.setColumnWidth(index + 1, largeur);
  });

  const nbFormatLignes = Math.max(
    1,
    feuille.getMaxRows() - CONFIG_RAPPROCHEMENT.ligneEntetes
  );

  // Format monétaire pour les colonnes de soldes et montants
  [2, 3, 4, 5, 6, 10].forEach(function(col) {
    feuille
      .getRange(CONFIG_RAPPROCHEMENT.premiereLigne, col, nbFormatLignes, 1)
      .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
  });

  return feuille;
}

// ─── Calcul principal ─────────────────────────────────────────────────────────

function calculerEtEcrireRapprochement_(ss) {
  // 1. Charger toutes les données sources
  const importData = chargerImportBancaire_Rappr_(ss);
  const journalData = chargerJournal_Rappr_(ss);
  const transData = chargerTransactions_Rappr_(ss);
  const soldesOuv = chargerSoldesOuverture_Rappr_(ss);
  const depensesParMois = chargerToutesDepensesParMois_Rappr_(ss);

  // 2. Valider la présence du compte 1000 dans Soldes ouverture
  if (!soldesOuv[CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE]) {
    throw new Error(
      'Le compte ' + CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE +
      ' est introuvable dans l’onglet ' +
      '« Soldes ouverture ».'
    );
  }

  const entree1000 = soldesOuv[CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE];
  const soldeOuverture1000 = arrondirRappr_(
    entree1000.debit - entree1000.credit
  );

  if (importData.length === 0) {
    return;
  }

  // 3. Déterminer le dernier mois présent dans Import bancaire
  const derniereMoisImport = importData.reduce(function(max, r) {
    const cle = cleAnneesMoisRappr_(r.date);
    return cle > max ? cle : max;
  }, '');

  if (!derniereMoisImport) {
    return;
  }

  const partiesFin = derniereMoisImport.split('-');
  const anneeFin = Number(partiesFin[0]);
  const moisFin = Number(partiesFin[1]);

  // 4. Générer la liste des mois de ANNEE_DEBUT-01 à derniereMoisImport
  const periodes = [];
  let a = CONFIG_RAPPROCHEMENT.ANNEE_DEBUT;
  let m = 1;

  while (a < anneeFin || (a === anneeFin && m <= moisFin)) {
    const cle = a + '-' + String(m).padStart(2, '0');
    periodes.push({ annee: a, mois: m, cle: cle });
    m++;

    if (m > 12) {
      m = 1;
      a++;
    }
  }

  // 5. Indexer Import bancaire par mois (trié par date puis ordre original) et par ID
  const importParMois = {};
  const importParId = {};

  importData.forEach(function(r) {
    const cle = cleAnneesMoisRappr_(r.date);

    if (!importParMois[cle]) {
      importParMois[cle] = [];
    }

    importParMois[cle].push(r);
    importParId[r.idImport] = r;
  });

  // 6. Calculer le solde bancaire d'ouverture depuis la première transaction
  //    de janvier ANNEE_DEBUT et valider contre Soldes ouverture
  //    AVANT toute écriture dans l'onglet Rapprochement bancaire.
  //    Formule : premier solde bancaire (col E) − premier montant (col D)
  //    de la première transaction triée par date puis ordre original.
  const cleDebut = CONFIG_RAPPROCHEMENT.ANNEE_DEBUT + '-01';
  const lignesJanDebut = (importParMois[cleDebut] || []).slice().sort(
    function(a, b) {
      return a.date - b.date || a.ordreOriginal - b.ordreOriginal;
    }
  );

  const soldeOuvertureBancaire0 = lignesJanDebut.length > 0
    ? arrondirRappr_(lignesJanDebut[0].solde - lignesJanDebut[0].montant)
    : 0;

  if (
    Math.abs(soldeOuvertureBancaire0 - soldeOuverture1000) >=
    CONFIG_RAPPROCHEMENT.TOLERANCE
  ) {
    throw new Error(
      'Validation des soldes d’ouverture échouée.\n' +
      'Solde bancaire calculé (première transaction ' +
      '« Import bancaire ») : ' +
      soldeOuvertureBancaire0.toFixed(2) + ' $\n' +
      'Solde comptable compte ' +
      CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE +
      ' (« Soldes ouverture ») : ' +
      soldeOuverture1000.toFixed(2) + ' $\n' +
      'Vérifiez l’onglet « Soldes ouverture » ' +
      'ou les premières transactions de l’onglet ' +
      '« Import bancaire ».'
    );
  }

  // 7. Indexer Journal par mois
  const journalParMois = {};

  journalData.forEach(function(r) {
    const cle = cleAnneesMoisRappr_(r.date);

    if (!journalParMois[cle]) {
      journalParMois[cle] = [];
    }

    journalParMois[cle].push(r);
  });

  // 8. Lire les colonnes L (Date de rapprochement) et M (Notes) pour les préserver
  const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);
  const existant = {};
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne >= CONFIG_RAPPROCHEMENT.premiereLigne) {
    const nbEx = derniereLigne - CONFIG_RAPPROCHEMENT.premiereLigne + 1;
    const vals = feuille
      .getRange(
        CONFIG_RAPPROCHEMENT.premiereLigne,
        1,
        nbEx,
        CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
      )
      .getValues();

    vals.forEach(function(ligne) {
      const cle = normaliserCleMoisRappr_(ligne[0]);

      if (cle) {
        existant[cle] = {
          dateRappr: ligne[CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR - 1],
          notes: ligne[CONFIG_RAPPROCHEMENT.COL_NOTES - 1]
        };
      }
    });
  }

  // 9. Calculer tous les mois entièrement en mémoire.
  //    Aucune écriture dans l'onglet tant que tous les calculs ne sont pas terminés.
  let soldeOuvertureBancaire = soldeOuvertureBancaire0;
  let soldeCumul1000 = soldeOuverture1000;

  const resultats = periodes.map(function(periode) {
    // Tri : premièrement par date, deuxièmement par numéro de ligne original.
    // Le deuxième critère est indispensable quand plusieurs transactions
    // partagent la même date dans Import bancaire.
    const lignesImport = (importParMois[periode.cle] || []).slice().sort(
      function(a, b) {
        return a.date - b.date || a.ordreOriginal - b.ordreOriginal;
      }
    );

    const lignesJournal = journalParMois[periode.cle] || [];

    // Solde bancaire d'ouverture :
    // — premier mois (ANNEE_DEBUT-01) : soldeOuvertureBancaire0 déjà validé
    // — mois suivants : solde bancaire de fin du mois précédent
    const ouvertureBancaire = soldeOuvertureBancaire;

    // Entrées et sorties bancaires
    let entrees = 0;
    let sorties = 0;

    lignesImport.forEach(function(r) {
      if (r.montant > 0) {
        entrees += r.montant;
      } else {
        sorties += Math.abs(r.montant);
      }
    });

    entrees = arrondirRappr_(entrees);
    sorties = arrondirRappr_(sorties);

    // Solde bancaire de fin : dernier solde de la colonne E du mois
    // (dernière ligne après tri par date puis ordre original)
    const soldeBancaireFin = lignesImport.length > 0
      ? arrondirRappr_(lignesImport[lignesImport.length - 1].solde)
      : ouvertureBancaire;

    // Solde comptable compte 1000 : cumulatif depuis Soldes ouverture
    // jusqu'à la fin du mois analysé.
    // Les écritures d'annulation (col A commence par ANN-) sont incluses
    // normalement car elles compensent les écritures originales.
    lignesJournal.forEach(function(r) {
      if (r.codeCompte === CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE) {
        soldeCumul1000 = arrondirRappr_(
          soldeCumul1000 + r.debit - r.credit
        );
      }
    });

    // Journal équilibré pour ce mois :
    // — Oui si |total débits − total crédits du mois| < 0,005 $
    // — Oui si le mois ne contient aucune écriture Journal
    //   (mois sans écriture ≠ déséquilibre ; mais restera « À compléter »
    //    si des transactions bancaires ne sont pas encore classées)
    // — Non si la différence absolue est ≥ 0,005 $
    let totalDebits = 0;
    let totalCredits = 0;

    lignesJournal.forEach(function(r) {
      totalDebits += r.debit;
      totalCredits += r.credit;
    });

    const journalEquilibre =
      lignesJournal.length === 0 ||
      Math.abs(totalDebits - totalCredits) < CONFIG_RAPPROCHEMENT.TOLERANCE;

    // Écritures 1000 sans lien bancaire (dans ce mois uniquement) :
    // Pour chaque ligne Journal dont col D = compte 1000, suivre la chaîne :
    //   col B (idTransaction) → Transactions col A → col N (refBancaire) → Import bancaire col A
    // Note : on utilise col B (idTransaction), jamais col A (idEcriture).
    // Les écritures dont col A commence par ANN- restent valides si leur col B
    // mène à une transaction possédant une référence bancaire valide.
    let sansLien = 0;

    lignesJournal.forEach(function(r) {
      if (r.codeCompte !== CONFIG_RAPPROCHEMENT.COMPTE_BANCAIRE) {
        return;
      }

      const trans = transData[r.idTransaction];

      if (!trans || !trans.refBancaire || !importParId[trans.refBancaire]) {
        sansLien++;
      }
    });

    // Transactions bancaires à classer dans ce mois
    const aClasser = lignesImport.filter(function(r) {
      return r.statut !== 'Classée';
    }).length;

    // Écart : solde bancaire fin − solde comptable 1000
    const ecart = arrondirRappr_(soldeBancaireFin - soldeCumul1000);

    // Préserver les colonnes L et M (date clôture + notes)
    const preserve = existant[periode.cle] || {};
    const moisFerme =
      preserve.dateRappr instanceof Date &&
      !isNaN(preserve.dateRappr.getTime());

    // Statut — priorité stricte :
    // Mois fermé → 'Rapproché' ou 'Rapproché — anomalie' selon les contrôles
    // Sinon : priorité habituelle À compléter → À vérifier → Prêt à rapprocher
    let statut;

    if (moisFerme) {
      const depMois = depensesParMois[periode.cle] || { total: 0, sansPiece: 0 };
      const tousControles =
        aClasser === 0 &&
        journalEquilibre &&
        sansLien === 0 &&
        Math.abs(ecart) < CONFIG_RAPPROCHEMENT.TOLERANCE &&
        depMois.sansPiece === 0;
      statut = tousControles ? 'Rapproché' : 'Rapproché — anomalie';
    } else if (aClasser > 0) {
      statut = 'À compléter';
    } else if (
      !journalEquilibre ||
      sansLien > 0 ||
      Math.abs(ecart) >= CONFIG_RAPPROCHEMENT.TOLERANCE
    ) {
      statut = 'À vérifier';
    } else {
      statut = 'Prêt à rapprocher';
    }

    // Prochain solde bancaire d'ouverture = solde fin du mois courant
    soldeOuvertureBancaire = soldeBancaireFin;

    return {
      cle: periode.cle,
      ouvertureBancaire: ouvertureBancaire,
      entrees: entrees,
      sorties: sorties,
      soldeBancaireFin: soldeBancaireFin,
      soldeComptable: soldeCumul1000,
      journalEquilibre: journalEquilibre,
      sansLien: sansLien,
      aClasser: aClasser,
      ecart: ecart,
      statut: statut,
      dateRappr: preserve.dateRappr !== undefined ? preserve.dateRappr : '',
      notes: preserve.notes !== undefined ? String(preserve.notes) : ''
    };
  });

  // 10. Écriture atomique :
  //     Toutes les données calculées et validées en mémoire avant toute modification.
  //     Si une erreur survient avant ce point, l'ancien rapport est intégralement conservé
  //     (colonnes L et M comprises).
  if (derniereLigne >= CONFIG_RAPPROCHEMENT.premiereLigne) {
    feuille
      .getRange(
        CONFIG_RAPPROCHEMENT.premiereLigne,
        1,
        derniereLigne - CONFIG_RAPPROCHEMENT.premiereLigne + 1,
        CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
      )
      .clearContent();
  }

  if (resultats.length > 0) {
    const donnees = resultats.map(function(r) {
      return [
        r.cle,
        r.ouvertureBancaire,
        r.entrees,
        r.sorties,
        r.soldeBancaireFin,
        r.soldeComptable,
        r.journalEquilibre ? 'Oui' : 'Non',
        r.sansLien,
        r.aClasser,
        r.ecart,
        r.statut,
        r.dateRappr,
        r.notes
      ];
    });

    feuille
      .getRange(
        CONFIG_RAPPROCHEMENT.premiereLigne,
        1,
        donnees.length,
        CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
      )
      .setValues(donnees);

    // Couleur de fond col K selon le statut (appliquée ligne par ligne)
    resultats.forEach(function(r, index) {
      const ligne = CONFIG_RAPPROCHEMENT.premiereLigne + index;
      const couleur =
        r.statut === 'Rapproché' ? '#c8e6c9' :
        r.statut === 'Rapproché — anomalie' ? '#fce8e6' :
        r.statut === 'Prêt à rapprocher' ? '#e6f4ea' :
        r.statut === 'À vérifier' ? '#fce8e6' :
        '#fff7df';

      feuille
        .getRange(ligne, CONFIG_RAPPROCHEMENT.COL_STATUT)
        .setBackground(couleur);
    });
  }

  SpreadsheetApp.flush();
}

// ─── Chargement des données sources ──────────────────────────────────────────

function obtenirAnneesDepuisImportBancaire_(ss) {
  const feuille = ss.getSheetByName('Import bancaire');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 2).getValues();
  const anneesIndex = {};

  valeurs.forEach(function(ligne) {
    const idImport = String(ligne[0] || '').trim();
    const date = ligne[1];

    if (!idImport || !(date instanceof Date) || isNaN(date.getTime())) {
      return;
    }

    anneesIndex[date.getFullYear()] = true;
  });

  // Tri décroissant : l'année la plus récente est affichée en premier dans le sélecteur.
  return Object.keys(anneesIndex)
    .map(Number)
    .sort(function(a, b) { return b - a; });
}

function chargerImportBancaire_Rappr_(ss) {
  const feuille = ss.getSheetByName('Import bancaire');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 11).getValues();
  const lignes = [];

  valeurs.forEach(function(ligne, index) {
    const idImport = String(ligne[0] || '').trim();
    const date = ligne[1];
    const montant = Number(ligne[3] || 0);
    const solde = Number(ligne[4] || 0);
    const idTransaction = String(ligne[9] || '').trim();
    const statut = String(ligne[10] || '').trim();

    if (!idImport || !(date instanceof Date) || isNaN(date.getTime())) {
      return;
    }

    lignes.push({
      idImport: idImport,
      date: date,
      montant: montant,
      solde: solde,
      idTransaction: idTransaction,
      statut: statut,
      ordreOriginal: index
    });
  });

  return lignes;
}

function chargerJournal_Rappr_(ss) {
  const feuille = ss.getSheetByName('Journal');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 8).getValues();
  const lignes = [];

  valeurs.forEach(function(ligne) {
    const idEcriture = String(ligne[0] || '').trim();
    const idTransaction = String(ligne[1] || '').trim();
    const date = ligne[2];
    const codeCompte = String(ligne[3] || '').trim();
    const debit = Number(ligne[6] || 0);
    const credit = Number(ligne[7] || 0);

    if (!idEcriture || !(date instanceof Date) || isNaN(date.getTime())) {
      return;
    }

    lignes.push({
      idEcriture: idEcriture,
      idTransaction: idTransaction,
      date: date,
      codeCompte: codeCompte,
      debit: debit,
      credit: credit
    });
  });

  return lignes;
}

function chargerTransactions_Rappr_(ss) {
  const feuille = ss.getSheetByName('Transactions');

  if (!feuille || feuille.getLastRow() < 6) {
    return {};
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 15).getValues();
  const index = {};

  valeurs.forEach(function(ligne) {
    const idTransaction = String(ligne[0] || '').trim();
    const refBancaire = String(ligne[13] || '').trim();
    const statut = String(ligne[14] || '').trim();

    if (!idTransaction) {
      return;
    }

    index[idTransaction] = { refBancaire: refBancaire, statut: statut };
  });

  return index;
}

function chargerSoldesOuverture_Rappr_(ss) {
  const feuille = ss.getSheetByName('Soldes ouverture');

  if (!feuille) {
    throw new Error(
      'L’onglet « Soldes ouverture » est introuvable.'
    );
  }

  if (feuille.getLastRow() < 6) {
    throw new Error(
      'L’onglet « Soldes ouverture » ne contient aucune donnée.'
    );
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 5).getValues();
  const index = {};

  valeurs.forEach(function(ligne) {
    const code = String(ligne[0] || '').trim();
    const debit = Number(ligne[2] || 0);
    const credit = Number(ligne[3] || 0);

    if (!code) {
      return;
    }

    index[code] = { debit: debit, credit: credit };
  });

  return index;
}

// Charge les dépenses actives d'un mois depuis Transactions (colonnes A:Q).
// Exclut les transactions annulées.
// Retourne un tableau {idTransaction, pieceCtrl} pour le contrôle PIECES_DEPENSES.
function chargerDepensesMois_Rappr_(ss, cleMois) {
  const feuille = ss.getSheetByName('Transactions');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 17).getValues();
  const depenses = [];

  valeurs.forEach(function(ligne) {
    const idTransaction = String(ligne[0] || '').trim();
    const date = ligne[1];
    const type = String(ligne[2] || '').trim();
    const statut = String(ligne[14] || '').trim();
    const pieceCtrl = String(ligne[16] || '').trim();

    if (!idTransaction) return;
    if (!(date instanceof Date) || isNaN(date.getTime())) return;
    if (cleAnneesMoisRappr_(date) !== cleMois) return;
    if (type !== 'Dépense') return;
    if (statut === 'Annulée') return;

    depenses.push({
      idTransaction: idTransaction,
      pieceCtrl: pieceCtrl
    });
  });

  return depenses;
}

// Charge toutes les dépenses actives regroupées par mois depuis Transactions (A:Q).
// Retourne { cleMois: { total, sansPiece } } — utilisé par calculerEtEcrireRapprochement_.
function chargerToutesDepensesParMois_Rappr_(ss) {
  const feuille = ss.getSheetByName('Transactions');

  if (!feuille || feuille.getLastRow() < 6) {
    return {};
  }

  const nbLignes = feuille.getLastRow() - 5;
  const valeurs = feuille.getRange(6, 1, nbLignes, 17).getValues();
  const parMois = {};

  valeurs.forEach(function(ligne) {
    const idTransaction = String(ligne[0] || '').trim();
    const date = ligne[1];
    const type = String(ligne[2] || '').trim();
    const statut = String(ligne[14] || '').trim();
    const pieceCtrl = String(ligne[16] || '').trim();

    if (!idTransaction) return;
    if (!(date instanceof Date) || isNaN(date.getTime())) return;
    if (type !== 'Dépense') return;
    if (statut === 'Annulée') return;

    const cle = cleAnneesMoisRappr_(date);
    if (!parMois[cle]) parMois[cle] = { total: 0, sansPiece: 0 };
    parMois[cle].total++;
    if (pieceCtrl !== 'OK') parMois[cle].sansPiece++;
  });

  return parMois;
}

// ─── Garde de période comptable ───────────────────────────────────────────────

// Dérive la clé YYYY-MM depuis une Date comptable.
// Utilise le fuseau du classeur (America/Toronto).
function obtenirCleMoisDepuisDateComptable_(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error(
      'Date comptable invalide : ' + String(date)
    );
  }
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'yyyy-MM');
}

// Lit l'onglet Rapprochement bancaire et retourne l'état de la période.
// Ne lance aucun verrou — doit être appelée sous verrou existant.
// Retourne { cle, estFermee, libelle }.
function obtenirEtatPeriodeComptable_(ss, date) {
  const cle = obtenirCleMoisDepuisDateComptable_(date);
  const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);

  if (!feuille || feuille.getLastRow() < CONFIG_RAPPROCHEMENT.premiereLigne) {
    return { cle: cle, estFermee: false };
  }

  const nbLignes =
    feuille.getLastRow() - CONFIG_RAPPROCHEMENT.premiereLigne + 1;
  const valeurs = feuille
    .getRange(
      CONFIG_RAPPROCHEMENT.premiereLigne,
      1,
      nbLignes,
      CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
    )
    .getValues();

  for (var i = 0; i < valeurs.length; i++) {
    const cleRow = normaliserCleMoisRappr_(valeurs[i][0]);
    if (cleRow !== cle) continue;

    const dateL = valeurs[i][CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR - 1];
    if (dateL instanceof Date && !isNaN(dateL.getTime())) {
      const p = cle.split('-');
      return {
        cle: cle,
        estFermee: true,
        libelle: libelleMoisRappr_(Number(p[0]), Number(p[1]))
      };
    }
    // L non vide mais invalide (texte, nombre, Date NaN) = anomalie bloquante.
    // Ne jamais traiter cette situation comme une période ouverte.
    if (dateL !== '' && dateL !== null && dateL !== undefined) {
      throw new Error(
        'La colonne L du mois ' + cle +
        ' contient une valeur non vide invalide : « ' + String(dateL) + ' ». ' +
        'Corrigez cette cellule dans l\'onglet « Rapprochement bancaire » avant de continuer.'
      );
    }
    return { cle: cle, estFermee: false };
  }

  return { cle: cle, estFermee: false };
}

// Lance une erreur bloquante si la période comptable est fermée.
// action : courte description à l'infinitif, ex. « créer une transaction ».
function verifierPeriodeComptableOuverte_(ss, date, action) {
  const etat = obtenirEtatPeriodeComptable_(ss, date);
  if (etat.estFermee) {
    const p = etat.cle.split('-');
    const libelle = libelleMoisRappr_(Number(p[0]), Number(p[1]));
    throw new Error(
      'La période ' + libelle +
      ' est fermée. Réouvrez-la avant de ' + action + '.'
    );
  }
}

// Retourne la liste des libellés de tous les mois actuellement clôturés.
// Utilisé par ReinitialisationTests.js pour bloquer la remise à zéro.
function obtenirMoisFermesRapprochement_(ss) {
  const feuille = ss.getSheetByName(CONFIG_RAPPROCHEMENT.nomFeuille);

  if (!feuille || feuille.getLastRow() < CONFIG_RAPPROCHEMENT.premiereLigne) {
    return [];
  }

  const nbLignes =
    feuille.getLastRow() - CONFIG_RAPPROCHEMENT.premiereLigne + 1;
  const valeurs = feuille
    .getRange(
      CONFIG_RAPPROCHEMENT.premiereLigne,
      1,
      nbLignes,
      CONFIG_RAPPROCHEMENT.NOMBRE_COLONNES
    )
    .getValues();

  return valeurs.reduce(function(acc, ligne) {
    const cle = normaliserCleMoisRappr_(ligne[0]);
    const dateL = ligne[CONFIG_RAPPROCHEMENT.COL_DATE_RAPPR - 1];
    if (cle && dateL instanceof Date && !isNaN(dateL.getTime())) {
      const p = cle.split('-');
      acc.push(libelleMoisRappr_(Number(p[0]), Number(p[1])));
    }
    return acc;
  }, []);
}

// ─── Utilitaires ─────────────────────────────────────────────────────────────

function arrondirRappr_(valeur) {
  const n = Number(valeur || 0);
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function cleAnneesMoisRappr_(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }

  return (
    date.getFullYear() +
    '-' +
    String(date.getMonth() + 1).padStart(2, '0')
  );
}

function normaliserCleMoisRappr_(valeur) {
  if (valeur instanceof Date && !isNaN(valeur.getTime())) {
    return cleAnneesMoisRappr_(valeur);
  }
  const s = String(valeur || '').trim();
  return s.match(/^\d{4}-\d{2}$/) ? s : '';
}

function libelleMoisRappr_(annee, mois) {
  const noms = [
    'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
    'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'
  ];

  return noms[mois - 1] + ' ' + annee;
}

function formaterDateRappr_(valeur) {
  if (!valeur) {
    return '';
  }

  const d = valeur instanceof Date ? valeur : new Date(valeur);

  if (isNaN(d.getTime())) {
    return '';
  }

  return (
    d.getFullYear() +
    '-' +
    String(d.getMonth() + 1).padStart(2, '0') +
    '-' +
    String(d.getDate()).padStart(2, '0')
  );
}
