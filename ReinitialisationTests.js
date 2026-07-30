// ReinitialisationTests.js
// Module de remise à zéro contrôlée des données de test
// Ne modifie jamais les onglets maîtres. N'appelle jamais les fonctions comptables.
// Utilise clearContent() uniquement — aucun deleteRow(), clear() ou clearDataValidations().

// ─── Constantes ──────────────────────────────────────────────────────────────

const ONGLETS_MAITRES_REINIT_ = [
  'Inventaire', 'Soldes ouverture', 'Règles bancaires',
  'Fournisseurs', 'Contacts', 'Configuration', 'Projets', 'Accueil'
];

const ONGLETS_DERIVES_REINIT_ = ['Rapports', 'Rapport fournisseurs'];

const MOT_CONFIRMATION_REINIT_ = 'REINITIALISER LES TESTS';

// Spécification de chaque onglet à réinitialiser.
// plagesEffacer     : [{col, nbCols, label}]  — jamais superposées avec formulesProtegees.
// formulesProtegees : [{col, nbCols, label}]  — lues avant et vérifiées après.
// headersAttendus   : [{col, valeur}] — valeur exacte attendue en ligne 5 (vérification structurelle).
// nbColsMin         : nombre minimum de colonnes requis (bloquant si inférieur).
// colIdVerif        : colonne dont on vérifie l'absence de données après remise à zéro.
// plagesEffacer null = détection dynamique du nombre de colonnes (onglet Documents).
const SPECS_ONGLETS_REINIT_ = [
  {
    nom: 'Import bancaire',
    ligneDepart: 6,
    colId: 1,
    // A:W = toutes les données d'import et colonnes de suggestions (H:I pré-classement,
    //       O action, P:S fournisseur/contact suggéré, T:W règle/composante/projet/ID règle)
    plagesEffacer: [{ col: 1, nbCols: 23, label: 'A:W' }],
    formulesProtegees: [],
    headersAttendus: [
      { col: 15, valeur: 'Action' },
      { col: 16, valeur: 'ID fournisseur suggéré' },
      { col: 20, valeur: 'Type de classement suggéré' },
      { col: 23, valeur: 'ID règle appliquée' }
    ],
    nbColsMin: 23,
    colIdVerif: 1,
    labelVerif: 'IDs bancaires (colonne A)'
  },
  {
    nom: 'Transactions',
    ligneDepart: 6,
    colId: 1,
    // A:O = données saisies + col H (formule par ligne, réécrite à chaque transaction)
    // R:T = fournisseur/contact (données importées)
    // P:Q = formules permanentes TEXT et contrôle pièce — préservées
    plagesEffacer: [
      { col: 1, nbCols: 15, label: 'A:O' },
      { col: 18, nbCols: 3, label: 'R:T' }
    ],
    formulesProtegees: [{ col: 16, nbCols: 2, label: 'P:Q' }],
    headersAttendus: [
      { col: 18, valeur: 'ID fournisseur' },
      { col: 19, valeur: 'Fournisseur' },
      { col: 20, valeur: 'ID contact' }
    ],
    nbColsMin: 20,
    colIdVerif: 1,
    labelVerif: 'IDs transactions (colonne A)'
  },
  {
    nom: 'Journal',
    ligneDepart: 6,
    colId: 1,
    // A:M = données comptables (E:F = formules VLOOKUP par ligne, réécrites à chaque écriture)
    // O:P = fournisseur (données ajoutées par l'installeur fournisseurs)
    // N   = formule TEXT(date,"yyyy-mm") — préservée car utilisée par les rapports
    plagesEffacer: [
      { col: 1, nbCols: 13, label: 'A:M' },
      { col: 15, nbCols: 2, label: 'O:P' }
    ],
    formulesProtegees: [{ col: 14, nbCols: 1, label: 'N' }],
    headersAttendus: [
      { col: 15, valeur: 'ID fournisseur' },
      { col: 16, valeur: 'Fournisseur' }
    ],
    nbColsMin: 16,
    colIdVerif: 1,
    labelVerif: 'IDs écritures (colonne A)'
  },
  {
    nom: 'Répartition',
    ligneDepart: 6,
    colId: 1,
    // A:H = données de répartition (saisies ou importées)
    // J:L = compte/programme/projet — valeurs pour dépenses, formules pour composantes :
    //        les deux cas sont des données de test à supprimer
    // P:Q = notes et synchronisation (données de test)
    // I   = formule montant attribué (qty × prix) — préservée
    // M:O = formules total réparti / reste / statut — préservées
    plagesEffacer: [
      { col: 1, nbCols: 8, label: 'A:H' },
      { col: 10, nbCols: 3, label: 'J:L' },
      { col: 16, nbCols: 2, label: 'P:Q' }
    ],
    formulesProtegees: [
      { col: 9, nbCols: 1, label: 'I' },
      { col: 13, nbCols: 3, label: 'M:O' }
    ],
    headersAttendus: [
      { col: 1, valeur: 'ID import bancaire' },
      { col: 9, valeur: 'Montant attribué' },
      { col: 16, valeur: 'Notes' }
    ],
    nbColsMin: 16,
    colIdVerif: 1,
    labelVerif: 'IDs import bancaire (colonne A)'
  },
  {
    nom: 'Forfaits',
    ligneDepart: 6,
    colId: 1,
    // A:F = données du forfait (ID, date, acheteur, type, montant, ID transaction)
    // I:J = statut et notes (données)
    // G:H = formules de calcul des parts Pion/Cartier — préservées
    plagesEffacer: [
      { col: 1, nbCols: 6, label: 'A:F' },
      { col: 9, nbCols: 2, label: 'I:J' }
    ],
    formulesProtegees: [{ col: 7, nbCols: 2, label: 'G:H' }],
    headersAttendus: [
      { col: 1, valeur: 'ID forfait' },
      { col: 7, valeur: 'Part Pion joues-tu?' },
      { col: 10, valeur: 'Notes' }
    ],
    nbColsMin: 10,
    colIdVerif: null,
    labelVerif: null
  },
  {
    nom: 'Budget 2026',
    ligneDepart: 6,
    colId: 1,
    // A:D = année, période, programme, code compte (saisie)
    // G   = montant budgété (saisie)
    // K:L = statut et notes (saisie)
    // E:F = formules VLOOKUP nom/type du compte — préservées
    // H:J = formules SUMIFS réel, écart, % utilisé — préservées
    plagesEffacer: [
      { col: 1, nbCols: 4, label: 'A:D' },
      { col: 7, nbCols: 1, label: 'G' },
      { col: 11, nbCols: 2, label: 'K:L' }
    ],
    formulesProtegees: [
      { col: 5, nbCols: 2, label: 'E:F' },
      { col: 8, nbCols: 3, label: 'H:J' }
    ],
    headersAttendus: [
      { col: 1, valeur: 'Année' },
      { col: 8, valeur: 'Réel' },
      { col: 12, valeur: 'Notes' }
    ],
    nbColsMin: 12,
    colIdVerif: null,
    labelVerif: null
  },
  {
    nom: 'Documents',
    ligneDepart: 6,
    colId: 1,
    // A:K = 11 colonnes de données (ID document à Notes)
    plagesEffacer: [{ col: 1, nbCols: 11, label: 'A:K' }],
    formulesProtegees: [],
    headersAttendus: [
      { col: 1, valeur: 'ID document' },
      { col: 11, valeur: 'Notes' }
    ],
    nbColsMin: 11,
    colIdVerif: null,
    labelVerif: null
  },
  {
    nom: 'Rapprochement bancaire',
    ligneDepart: 6,
    colId: 1,
    // A:M = données calculées (toutes des valeurs, aucune formule)
    plagesEffacer: [{ col: 1, nbCols: 13, label: 'A:M' }],
    formulesProtegees: [],
    headersAttendus: [
      { col: 1, valeur: 'Mois' },
      { col: 11, valeur: 'Statut' },
      { col: 13, valeur: 'Notes' }
    ],
    nbColsMin: 13,
    colIdVerif: null,
    labelVerif: null
  }
];

// ─── Fonctions publiques ──────────────────────────────────────────────────────

function auditerReinitialisationDonneesTests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var rapport = executerAuditReinit_(ss);

  var lignes = [];
  lignes.push('=== AUDIT DE REMISE A ZERO ===\n');

  lignes.push('-- Onglets maitres (conserves) --');
  rapport.ongletsMaitresTrouves.forEach(function(nom) {
    lignes.push('  OK : ' + nom);
  });
  if (rapport.ongletsAbsentsMAITRES.length) {
    rapport.ongletsAbsentsMAITRES.forEach(function(nom) {
      lignes.push('  ABSENT (bloquant) : ' + nom);
    });
  }

  lignes.push('\n-- Onglets derives (empreintes capturees avant effacement) --');
  rapport.ongletsDerivesPresents.forEach(function(nom) {
    lignes.push('  OK : ' + nom);
  });
  if (rapport.ongletsAbsentsDERIVES.length) {
    rapport.ongletsAbsentsDERIVES.forEach(function(nom) {
      lignes.push('  ABSENT (bloquant) : ' + nom);
    });
  }

  lignes.push('\n-- Onglets a reinitialiser --');
  rapport.ongletsReinitInfo.forEach(function(info) {
    var ligneInfo = '  ' + info.nom + ' : ' + info.nbLignesDonnees + ' ligne(s)';
    if (info.colonnesEffacees) {
      ligneInfo += ' | Effacer : ' + info.colonnesEffacees;
    }
    if (info.formulesProtegees.length) {
      ligneInfo += ' | Formules protegees : ' + info.formulesProtegees.join(', ');
    }
    lignes.push(ligneInfo);
  });
  if (rapport.ongletsAbsentsREINIT.length) {
    rapport.ongletsAbsentsREINIT.forEach(function(nom) {
      lignes.push('  ABSENT (bloquant) : ' + nom);
    });
  }

  if (rapport.problemes.length) {
    lignes.push('\nPROBLEMES DETECTES :');
    rapport.problemes.forEach(function(p) {
      lignes.push('  * ' + p);
    });
    if (rapport.bloquant) {
      lignes.push('\nREMISE A ZERO BLOQUEE — corrigez les problemes avant de relancer.');
    }
  } else {
    lignes.push('\nAucun probleme — remise a zero possible.');
  }

  ui.alert('Audit de remise a zero\n\n' + lignes.join('\n'));

  return rapport;
}

function reinitialiserDonneesTests() {
  var verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    SpreadsheetApp.getUi().alert(
      'Une autre operation est en cours. Reessayez dans quelques secondes.'
    );
    return;
  }

  try {
    reinitialiserDonneesTestsAvecVerrou_();
  } finally {
    verrou.releaseLock();
  }
}

// ─── Execution principale (protegee par verrou) ───────────────────────────────

function reinitialiserDonneesTestsAvecVerrou_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. Audit complet avant toute écriture
  var rapport = executerAuditReinit_(ss);

  if (rapport.bloquant) {
    ui.alert(
      'Remise a zero impossible\n\n' +
      'Problemes bloquants :\n\n' +
      rapport.problemes.map(function(p) { return '* ' + p; }).join('\n') +
      '\n\nCorrigez ces problemes avant de relancer.'
    );
    return;
  }

  // 2. Afficher le bilan de l'audit
  var lignesBilan = ['BILAN DE L\'AUDIT\n'];
  lignesBilan.push(
    'Onglets maitres verifies : ' +
    rapport.ongletsMaitresTrouves.length + '/' + ONGLETS_MAITRES_REINIT_.length
  );
  lignesBilan.push(
    'Onglets derives verifies : ' +
    rapport.ongletsDerivesPresents.length + '/' + ONGLETS_DERIVES_REINIT_.length
  );
  lignesBilan.push('\nDonnees qui seront effacees :');
  rapport.ongletsReinitInfo.forEach(function(info) {
    lignesBilan.push(
      '  ' + info.nom + ' : ' + info.nbLignesDonnees +
      ' ligne(s) — colonnes ' + info.colonnesEffacees
    );
  });
  lignesBilan.push('\nCETTE ACTION EST IRREVERSIBLE.');
  lignesBilan.push('La copie de sauvegarde Google Sheet est votre mecanisme de recuperation.');

  ui.alert('Remise a zero des donnees de test\n\n' + lignesBilan.join('\n'));

  // 3. Confirmation explicite
  var reponse = ui.prompt(
    'Confirmation requise',
    'Pour confirmer la remise a zero, saisissez exactement :\n\n' +
    MOT_CONFIRMATION_REINIT_ + '\n\n' +
    '(Cliquez Annuler pour abandonner sans modification)',
    ui.ButtonSet.OK_CANCEL
  );

  if (reponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Remise a zero annulee — aucune modification effectuee.');
    return;
  }

  if (reponse.getResponseText().trim() !== MOT_CONFIRMATION_REINIT_) {
    ui.alert(
      'Texte de confirmation incorrect — aucune modification effectuee.\n\n' +
      'Attendu : ' + MOT_CONFIRMATION_REINIT_
    );
    return;
  }

  // 4. Capturer les empreintes des onglets maitres avant toute écriture
  var empreintesAvant = {};
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    var feuille = ss.getSheetByName(nom);
    if (feuille) {
      empreintesAvant[nom] = empreindreOnglet_(feuille);
    }
  });

  // 5. Capturer les empreintes des onglets dérivés avant tout effacement.
  // L'empreinte utilise le texte des formules (pas les résultats calculés),
  // donc elle reste stable après l'effacement des données source.
  // Un changement d'empreinte signale une modification inattendue de structure.
  var empreintesDerivesAvant = {};
  ONGLETS_DERIVES_REINIT_.forEach(function(nom) {
    var feuille = ss.getSheetByName(nom);
    if (feuille) {
      empreintesDerivesAvant[nom] = empreindreFormulesOngletDerive_(feuille);
    }
  });

  // 6. Effacements onglet par onglet — arrêt immédiat sur erreur
  var bilanFinal = [];
  var erreurs = [];
  var ongletsTraites = [];
  var arreter = false;

  for (var i = 0; i < SPECS_ONGLETS_REINIT_.length; i++) {
    if (arreter) break;

    var spec = SPECS_ONGLETS_REINIT_[i];
    var feuille = ss.getSheetByName(spec.nom);

    if (!feuille) {
      // L'audit aurait bloqué si l'onglet était absent — défense en profondeur
      erreurs.push(
        spec.nom + ' : onglet introuvable apres audit — arret immediat.\n' +
        'Onglets deja traites : ' +
        (ongletsTraites.length ? ongletsTraites.join(', ') : 'aucun') +
        '\nVerifiez et restaurez depuis la copie de sauvegarde si necessaire.'
      );
      bilanFinal.push('  ' + spec.nom + ' : ERREUR — onglet introuvable');
      arreter = true;
      break;
    }

    var lastRow = feuille.getLastRow();
    var nbLignes = Math.max(0, lastRow - spec.ligneDepart + 1);

    if (nbLignes === 0) {
      bilanFinal.push('  ' + spec.nom + ' : aucune donnee a effacer');
      ongletsTraites.push(spec.nom);
      continue;
    }

    try {
      // Capturer les formules protégées avant l'effacement
      var formulasAvant = capturerFormulasProtegees_(
        feuille, spec.ligneDepart, nbLignes, spec.formulesProtegees
      );

      var plages = spec.plagesEffacer;

      // Effacer les plages (clearContent préserve validations, formats, largeurs de colonne)
      for (var p = 0; p < plages.length; p++) {
        feuille.getRange(spec.ligneDepart, plages[p].col, nbLignes, plages[p].nbCols)
          .clearContent();
      }

      SpreadsheetApp.flush();

      // Vérifier les résultats de l'effacement — arrêt immédiat sur toute anomalie
      var erreursOnglet = [];

      verifierPlagesEffacees_(feuille, spec.ligneDepart, nbLignes, plages)
        .forEach(function(e) { erreursOnglet.push(e); });

      verifierFormulasProtegees_(
        feuille, spec.ligneDepart, nbLignes, formulasAvant, spec.formulesProtegees
      ).forEach(function(e) { erreursOnglet.push(e); });

      if (spec.colIdVerif) {
        var nbReste = compterLignesAvecDonnees_(
          feuille, spec.colIdVerif, spec.ligneDepart, nbLignes
        );
        if (nbReste > 0) {
          erreursOnglet.push(
            spec.nom + ' : ' + nbReste +
            ' ligne(s) encore non videe(s) (' + spec.labelVerif + ')'
          );
        }
      }

      if (erreursOnglet.length > 0) {
        throw new Error(
          'Anomalies de verification (' + erreursOnglet.length + ') :\n' +
          erreursOnglet.join('\n')
        );
      }

      var labelsEffaces = plages.map(function(pp) { return pp.label; }).join(', ');
      bilanFinal.push(
        '  ' + spec.nom + ' : ' + nbLignes +
        ' ligne(s) effacee(s) — colonnes ' + labelsEffaces
      );
      ongletsTraites.push(spec.nom);

    } catch (e) {
      erreurs.push(
        spec.nom + ' (effacement) : ' + e.message + '\n' +
        'Onglets deja traites : ' +
        (ongletsTraites.length ? ongletsTraites.join(', ') : 'aucun') +
        '\nVerifiez et restaurez depuis la copie de sauvegarde si necessaire.'
      );
      bilanFinal.push('  ' + spec.nom + ' : ERREUR — ' + e.message);
      arreter = true;
    }
  }

  SpreadsheetApp.flush();

  // 7. Vérifier l'intégrité des onglets maitres
  var erreursMaitres = [];
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    var feuilleM = ss.getSheetByName(nom);
    if (!feuilleM || !empreintesAvant[nom]) return;

    var empreinteApres = empreindreOnglet_(feuilleM);
    if (empreinteApres !== empreintesAvant[nom]) {
      erreursMaitres.push('MODIFICATION sur l\'onglet maitre : ' + nom);
      erreurs.push(
        'Integrite compromise : ' + nom +
        ' — empreinte modifiee. Verifiez et restaurez depuis la copie de sauvegarde.'
      );
    }
  });

  // 8. Vérifier l'intégrité des onglets dérivés.
  // Les formules sont stables (empreinte basée sur le texte, pas les résultats).
  // Un changement signale une modification inattendue de structure.
  var erreursDerives = [];
  ONGLETS_DERIVES_REINIT_.forEach(function(nom) {
    var feuilleD = ss.getSheetByName(nom);
    if (!feuilleD || !empreintesDerivesAvant[nom]) return;

    var empreinteApres = empreindreFormulesOngletDerive_(feuilleD);
    if (empreinteApres !== empreintesDerivesAvant[nom]) {
      erreursDerives.push('Structure modifiee sur l\'onglet derive : ' + nom);
      erreurs.push(
        'Integrite derive compromise : ' + nom +
        ' — formules ou valeurs fixes modifiees. Verifiez depuis la copie de sauvegarde.'
      );
    }
  });

  // 9. Afficher le bilan final
  var titreMessage = arreter ? 'REMISE A ZERO INTERROMPUE' : 'REMISE A ZERO TERMINEE';

  var messageFinal = titreMessage + '\n\nDetail par onglet :\n' + bilanFinal.join('\n');

  if (erreursMaitres.length) {
    messageFinal +=
      '\n\n!!! INTEGRITE DES ONGLETS MAITRES COMPROMISE !!!\n' +
      erreursMaitres.join('\n') +
      '\n\nVerifiez immediatement et restaurez depuis la copie de sauvegarde.';
  } else {
    messageFinal += '\n\nOnglets maitres integres — empreintes verifiees.';
  }

  if (erreursDerives.length) {
    messageFinal +=
      '\n\n!!! INTEGRITE DES ONGLETS DERIVES COMPROMISE !!!\n' +
      erreursDerives.join('\n') +
      '\n\nVerifiez et restaurez depuis la copie de sauvegarde.';
  } else {
    messageFinal += '\n\nOnglets derives integres — empreintes verifiees.';
  }

  if (erreurs.length) {
    messageFinal +=
      '\n\nANOMALIES (' + erreurs.length + ') :\n' +
      erreurs.map(function(e) { return '* ' + e; }).join('\n');
  } else {
    messageFinal +=
      '\n\nAucune anomalie detectee.' +
      '\nVous pouvez maintenant importer le CSV bancaire pour tester le systeme.';
  }

  ui.alert('Remise a zero des donnees de test\n\n' + messageFinal);
}

// ─── Audit interne ────────────────────────────────────────────────────────────

function executerAuditReinit_(ss) {
  var rapport = {
    ongletsMaitresTrouves: [],
    ongletsAbsentsMAITRES: [],
    ongletsDerivesPresents: [],
    ongletsAbsentsDERIVES: [],
    ongletsReinitInfo: [],
    ongletsAbsentsREINIT: [],
    problemes: [],
    bloquant: false
  };

  // Vérifier les onglets maîtres (leur absence est bloquante)
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    if (ss.getSheetByName(nom)) {
      rapport.ongletsMaitresTrouves.push(nom);
    } else {
      rapport.ongletsAbsentsMAITRES.push(nom);
      rapport.problemes.push('Onglet maitre absent : « ' + nom + ' »');
      rapport.bloquant = true;
    }
  });

  // Vérifier les onglets dérivés (leur absence est bloquante)
  ONGLETS_DERIVES_REINIT_.forEach(function(nom) {
    if (ss.getSheetByName(nom)) {
      rapport.ongletsDerivesPresents.push(nom);
    } else {
      rapport.ongletsAbsentsDERIVES.push(nom);
      rapport.problemes.push('Onglet derive absent : « ' + nom + ' »');
      rapport.bloquant = true;
    }
  });

  // Inspecter chaque onglet à réinitialiser
  SPECS_ONGLETS_REINIT_.forEach(function(spec) {
    var feuille = ss.getSheetByName(spec.nom);

    if (!feuille) {
      rapport.ongletsAbsentsREINIT.push(spec.nom);
      rapport.problemes.push('Onglet a reinitialiser absent : « ' + spec.nom + ' »');
      rapport.bloquant = true;
      return;
    }

    // Vérifier le nombre minimum de colonnes
    var lastCol = feuille.getLastColumn();
    if (lastCol < spec.nbColsMin) {
      rapport.problemes.push(
        spec.nom + ' : ' + lastCol + ' colonne(s) trouvee(s), minimum ' +
        spec.nbColsMin + ' attendu — structure incompatible'
      );
      rapport.bloquant = true;
    }

    // Vérifier les en-têtes attendus (ligne 5, cellule par cellule)
    if (spec.headersAttendus.length) {
      if (feuille.getLastRow() < 5) {
        rapport.problemes.push(
          spec.nom + ' : ligne 5 absente (getLastRow = ' + feuille.getLastRow() +
          ') — en-tetes non verifiables, structure incompatible'
        );
        rapport.bloquant = true;
      } else {
        spec.headersAttendus.forEach(function(ha) {
          var valeurReelle = String(feuille.getRange(5, ha.col).getValue() || '').trim();
          if (valeurReelle !== ha.valeur) {
            rapport.problemes.push(
              spec.nom + ' col ' + colLetter_(ha.col) + '5 : en-tete « ' +
              valeurReelle + ' » (attendu « ' + ha.valeur + ' ») — structure incompatible'
            );
            rapport.bloquant = true;
          }
        });
      }
    }

    // Vérifier l'absence de superposition entre plages à effacer et formules protégées
    var superpositions = detecterSuperpositions_(spec.plagesEffacer, spec.formulesProtegees);
    superpositions.forEach(function(msg) {
      rapport.problemes.push(spec.nom + ' : ' + msg);
      rapport.bloquant = true;
    });

    // Compter les lignes avec données dans la colonne identifiant
    var lastRow = feuille.getLastRow();
    var nbLignesDonnees = 0;
    if (lastRow >= spec.ligneDepart) {
      nbLignesDonnees = compterLignesAvecDonnees_(
        feuille, spec.colId, spec.ligneDepart,
        lastRow - spec.ligneDepart + 1
      );
    }

    var labelsEffaces = spec.plagesEffacer
      ? spec.plagesEffacer.map(function(p) { return p.label; }).join(', ')
      : 'aucune';

    rapport.ongletsReinitInfo.push({
      nom: spec.nom,
      nbLignesDonnees: nbLignesDonnees,
      formulesProtegees: spec.formulesProtegees.map(function(fp) { return fp.label; }),
      colonnesEffacees: labelsEffaces
    });
  });

  return rapport;
}

// ─── Helpers privés ───────────────────────────────────────────────────────────

// Produit une empreinte SHA-256 de l'onglet, cellule par cellule.
// Pour une cellule formule : utilise le texte de la formule (résistant au recalcul).
// Pour une cellule valeur  : utilise la valeur affichée.
// Ajoute la note de chaque cellule pour détecter les modifications de commentaires.
function empreindreOnglet_(feuille) {
  var lastRow = feuille.getLastRow();
  var lastCol = feuille.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return 'VIDE-0x0';
  }

  var plage         = feuille.getRange(1, 1, lastRow, lastCol);
  var formulas      = plage.getFormulas();
  var displayValues = plage.getDisplayValues();
  var notes         = plage.getNotes();

  var parties = [String(lastRow), String(lastCol)];
  for (var i = 0; i < lastRow; i++) {
    for (var j = 0; j < lastCol; j++) {
      var formule = formulas[i][j];
      var valeur  = formule ? formule : displayValues[i][j];
      var note    = notes[i][j] || '';
      parties.push(valeur + '\x01' + note);
    }
  }

  var contenu = parties.join('\x02');

  var octets = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    contenu,
    Utilities.Charset.UTF_8
  );

  return octets.map(function(o) {
    return ('0' + (o < 0 ? o + 256 : o).toString(16)).slice(-2);
  }).join('');
}

// Détecte les superpositions de colonnes entre plages à effacer et formules protégées.
// Retourne un tableau de messages descriptifs (vide si aucune superposition).
function detecterSuperpositions_(plagesEffacer, formulesProtegees) {
  var superpositions = [];
  if (!plagesEffacer || !formulesProtegees || !formulesProtegees.length) {
    return superpositions;
  }

  plagesEffacer.forEach(function(pe) {
    formulesProtegees.forEach(function(fp) {
      var peDebut = pe.col;
      var peFin   = pe.col + pe.nbCols - 1;
      var fpDebut = fp.col;
      var fpFin   = fp.col + fp.nbCols - 1;
      if (Math.max(peDebut, fpDebut) <= Math.min(peFin, fpFin)) {
        superpositions.push(
          'Superposition entre plage a effacer « ' + pe.label +
          ' » et formules protegees « ' + fp.label + ' »'
        );
      }
    });
  });

  return superpositions;
}

// Compte les lignes non vides dans une colonne à partir d'une ligne de départ.
function compterLignesAvecDonnees_(feuille, col, ligneDepart, nbLignes) {
  if (nbLignes <= 0) return 0;

  return feuille
    .getRange(ligneDepart, col, nbLignes, 1)
    .getValues()
    .filter(function(r) { return String(r[0] || '').trim() !== ''; })
    .length;
}

// Capture les formules des colonnes protégées avant effacement.
function capturerFormulasProtegees_(feuille, ligneDepart, nbLignes, formulesProtegees) {
  var capture = {};
  if (nbLignes <= 0 || !formulesProtegees.length) return capture;

  formulesProtegees.forEach(function(fp) {
    var plage = feuille.getRange(ligneDepart, fp.col, nbLignes, fp.nbCols);
    capture[fp.label] = {
      formulas: plage.getFormulas()
    };
  });

  return capture;
}

// Vérifie que les colonnes protégées sont inchangées après l'effacement.
// Contrôle toutes les lignes — s'arrête à la première erreur par colonne protégée.
// Retourne un tableau de messages d'erreur (vide si tout est intact).
function verifierFormulasProtegees_(feuille, ligneDepart, nbLignes, formulasAvant, formulesProtegees) {
  var erreurs = [];
  if (nbLignes <= 0 || !formulesProtegees.length) return erreurs;

  for (var k = 0; k < formulesProtegees.length; k++) {
    var fp = formulesProtegees[k];
    var avant = formulasAvant[fp.label];
    if (!avant) continue;

    var formulesApres = feuille
      .getRange(ligneDepart, fp.col, nbLignes, fp.nbCols)
      .getFormulas();

    var trouveErreur = false;
    for (var i = 0; i < formulesApres.length && !trouveErreur; i++) {
      for (var j = 0; j < formulesApres[i].length && !trouveErreur; j++) {
        if (formulesApres[i][j] !== avant.formulas[i][j]) {
          erreurs.push(
            feuille.getName() + ' col ' + fp.label +
            ' ligne ' + (ligneDepart + i) +
            ' : formule modifiee (attendu « ' + avant.formulas[i][j] +
            ' », obtenu « ' + formulesApres[i][j] + ' »)'
          );
          trouveErreur = true;
        }
      }
    }
  }

  return erreurs;
}

// Vérifie que toutes les cellules des plages effacées sont réellement vides.
// S'arrête à la première cellule non vide par plage pour limiter le bruit.
// Retourne un tableau de messages d'erreur (vide si tout est vide).
function verifierPlagesEffacees_(feuille, ligneDepart, nbLignes, plages) {
  var erreurs = [];
  if (nbLignes <= 0 || !plages || !plages.length) return erreurs;

  plages.forEach(function(plage) {
    var valeurs = feuille
      .getRange(ligneDepart, plage.col, nbLignes, plage.nbCols)
      .getValues();

    var premierErreur = null;
    outer: for (var i = 0; i < valeurs.length; i++) {
      for (var j = 0; j < valeurs[i].length; j++) {
        var v = valeurs[i][j];
        if (v !== '' && v !== null && v !== undefined) {
          premierErreur = { ligne: ligneDepart + i, valeur: v };
          break outer;
        }
      }
    }

    if (premierErreur) {
      erreurs.push(
        feuille.getName() + ' col ' + plage.label +
        ' ligne ' + premierErreur.ligne +
        ' : cellule non videe apres clearContent (valeur : « ' +
        premierErreur.valeur + ' »)'
      );
    }
  });

  return erreurs;
}

// Produit une empreinte SHA-256 basée uniquement sur les formules d'un onglet dérivé.
// Conçue pour les rapports à formules QUERY/déversées (Rapports, Rapport fournisseurs) :
// les résultats déversés peuvent apparaître ou disparaître après effacement des données
// source, donc les valeurs affichées sont exclues de l'empreinte.
// Inclut : getMaxRows(), getMaxColumns(), coordonnées et texte de chaque cellule formule.
function empreindreFormulesOngletDerive_(feuille) {
  var maxRow = feuille.getMaxRows();
  var maxCol = feuille.getMaxColumns();

  if (maxRow < 1 || maxCol < 1) {
    return 'VIDE-0x0';
  }

  var formulas = feuille.getRange(1, 1, maxRow, maxCol).getFormulas();

  var parties = [String(maxRow), String(maxCol)];
  for (var i = 0; i < maxRow; i++) {
    for (var j = 0; j < maxCol; j++) {
      var formule = formulas[i][j];
      if (formule) {
        // Coordonnées + texte exact : détecte tout déplacement ou modification de formule
        parties.push((i + 1) + '\x01' + (j + 1) + '\x01' + formule);
      }
    }
  }

  var contenu = parties.join('\x02');

  var octets = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    contenu,
    Utilities.Charset.UTF_8
  );

  return octets.map(function(o) {
    return ('0' + (o < 0 ? o + 256 : o).toString(16)).slice(-2);
  }).join('');
}

// Convertit un numéro de colonne en lettre(s) Excel (1=A, 26=Z, 27=AA ...).
function colLetter_(col) {
  var lettre = '';
  var c = col;
  while (c > 0) {
    c--;
    lettre = String.fromCharCode(65 + (c % 26)) + lettre;
    c = Math.floor(c / 26);
  }
  return lettre;
}
