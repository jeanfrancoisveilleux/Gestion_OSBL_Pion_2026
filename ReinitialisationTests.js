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
// plagesEffacer  : [{col, nbCols, label}]  — jamais superposées avec formulesProtegees.
// formulesProtegees : [{col, nbCols, label}]  — lues avant et vérifiées après.
// colIdVerif : colonne dont on vérifie l'absence de données après remise à zéro.
// plagesEffacer null = détection dynamique du nombre de colonnes (onglet Documents).
const SPECS_ONGLETS_REINIT_ = [
  {
    nom: 'Import bancaire',
    ligneDepart: 6,
    colId: 1,
    plagesEffacer: [{ col: 1, nbCols: 23, label: 'A:W' }],
    formulesProtegees: [],
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
    colIdVerif: 1,
    labelVerif: 'IDs transactions (colonne A)'
  },
  {
    nom: 'Journal',
    ligneDepart: 6,
    colId: 1,
    // A:M = données comptables (E:F = formules VLOOKUP par ligne, réécrites à chaque écriture)
    // O:P = fournisseur (données ajoutées par l'installeur fournisseurs)
    // N = formule TEXT(date,"yyyy-mm") — préservée car utilisée par les rapports
    plagesEffacer: [
      { col: 1, nbCols: 13, label: 'A:M' },
      { col: 15, nbCols: 2, label: 'O:P' }
    ],
    formulesProtegees: [{ col: 14, nbCols: 1, label: 'N' }],
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
    colIdVerif: null,
    labelVerif: null
  },
  {
    nom: 'Documents',
    ligneDepart: 6,
    colId: 1,
    // Structure inconnue du code — toutes les colonnes sont effacées dynamiquement.
    plagesEffacer: null,
    formulesProtegees: [],
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
    lignes.push('  ABSENTS (bloquant) :');
    rapport.ongletsAbsentsMAITRES.forEach(function(nom) {
      lignes.push('    ABSENT : ' + nom);
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
    lignes.push('  Absents (sautes) : ' + rapport.ongletsAbsentsREINIT.join(', '));
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

  // 1. Audit avant toute ecriture
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
  lignesBilan.push('\nDonnees qui seront effacees :');
  rapport.ongletsReinitInfo.forEach(function(info) {
    lignesBilan.push(
      '  ' + info.nom + ' : ' + info.nbLignesDonnees +
      ' ligne(s) — colonnes ' + info.colonnesEffacees
    );
  });
  if (rapport.ongletsAbsentsREINIT.length) {
    lignesBilan.push('\nOnglets absents (sautes) : ' + rapport.ongletsAbsentsREINIT.join(', '));
  }
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

  // 4. Capturer les empreintes des onglets maitres avant toute ecriture
  var empreintesAvant = {};
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    var feuille = ss.getSheetByName(nom);
    if (feuille) {
      empreintesAvant[nom] = empreindreOnglet_(feuille);
    }
  });

  // 5. Effacements onglet par onglet
  var bilanFinal = [];
  var erreurs = [];

  SPECS_ONGLETS_REINIT_.forEach(function(spec) {
    var feuille = ss.getSheetByName(spec.nom);

    if (!feuille) {
      bilanFinal.push('  ' + spec.nom + ' : absent (saute)');
      return;
    }

    var lastRow = feuille.getLastRow();
    var nbLignes = Math.max(0, lastRow - spec.ligneDepart + 1);

    if (nbLignes === 0) {
      bilanFinal.push('  ' + spec.nom + ' : aucune donnee a effacer');
      return;
    }

    try {
      // Capturer les formules protegees avant l'effacement
      var formulasAvant = capturerFormulasProtegees_(
        feuille, spec.ligneDepart, nbLignes, spec.formulesProtegees
      );

      // Determiner les plages a effacer
      var plages = spec.plagesEffacer;
      if (!plages) {
        // Documents : detecter le nombre de colonnes dynamiquement
        var lastCol = Math.max(1, feuille.getLastColumn());
        plages = [{ col: 1, nbCols: lastCol, label: 'A:' + colLetter_(lastCol) }];
      }

      // Effacer les plages (clearContent preserve validations, formats, largeurs)
      plages.forEach(function(plage) {
        feuille.getRange(spec.ligneDepart, plage.col, nbLignes, plage.nbCols).clearContent();
      });

      SpreadsheetApp.flush();

      // 10. Verifier les formules protegees apres effacement
      var erreursFormulas = verifierFormulasProtegees_(
        feuille, spec.ligneDepart, nbLignes, formulasAvant, spec.formulesProtegees
      );
      erreursFormulas.forEach(function(e) { erreurs.push(e); });

      // 10. Verifier que la colonne identifiant est vide
      if (spec.colIdVerif) {
        var nbReste = compterLignesAvecDonnees_(
          feuille, spec.colIdVerif, spec.ligneDepart, nbLignes
        );
        if (nbReste > 0) {
          erreurs.push(
            spec.nom + ' : ' + nbReste +
            ' ligne(s) encore non videe(s) (' + spec.labelVerif + ')'
          );
        }
      }

      var labelsEffaces = plages.map(function(p) { return p.label; }).join(', ');
      bilanFinal.push(
        '  ' + spec.nom + ' : ' + nbLignes +
        ' ligne(s) effacee(s) — colonnes ' + labelsEffaces
      );

    } catch (e) {
      erreurs.push(spec.nom + ' (etape d\'effacement) : ' + e.message);
      bilanFinal.push('  ' + spec.nom + ' : ERREUR — ' + e.message);
    }
  });

  SpreadsheetApp.flush();

  // 11. Verifier l'integrite des onglets maitres
  var erreursMaitres = [];
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    var feuille = ss.getSheetByName(nom);
    if (!feuille || !empreintesAvant[nom]) return;

    var empreinteApres = empreindreOnglet_(feuille);
    if (empreinteApres !== empreintesAvant[nom]) {
      erreursMaitres.push('MODIFICATION sur l\'onglet maitre : ' + nom);
      erreurs.push('Integrite compromise : ' + nom +
        ' — empreinte modifiee. Verifiez et restaurez depuis la copie de sauvegarde.');
    }
  });

  // 12. Afficher le bilan final
  var messageFinal = 'REMISE A ZERO TERMINEE\n\n';
  messageFinal += 'Detail par onglet :\n' + bilanFinal.join('\n');

  if (erreursMaitres.length) {
    messageFinal +=
      '\n\n!!! INTEGRITE DES ONGLETS MAITRES COMPROMISE !!!\n' +
      erreursMaitres.join('\n') +
      '\n\nVerifiez immediatement et restaurez depuis la copie de sauvegarde.';
  } else {
    messageFinal += '\n\nOnglets maitres integres — empreintes verifiees.';
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
    ongletsReinitInfo: [],
    ongletsAbsentsREINIT: [],
    problemes: [],
    bloquant: false
  };

  // Verifier les onglets maitres (leur absence est bloquante)
  ONGLETS_MAITRES_REINIT_.forEach(function(nom) {
    if (ss.getSheetByName(nom)) {
      rapport.ongletsMaitresTrouves.push(nom);
    } else {
      rapport.ongletsAbsentsMAITRES.push(nom);
      rapport.problemes.push('Onglet maitre absent : « ' + nom + ' »');
      rapport.bloquant = true;
    }
  });

  // Inspecter chaque onglet a reinitialiser
  SPECS_ONGLETS_REINIT_.forEach(function(spec) {
    var feuille = ss.getSheetByName(spec.nom);

    if (!feuille) {
      rapport.ongletsAbsentsREINIT.push(spec.nom);
      return;
    }

    // Verifier la presence de l'en-tete (ligne 5, colonne A) — structure minimale
    var headerA = '';
    if (feuille.getLastRow() >= 5) {
      headerA = String(feuille.getRange(5, 1).getValue() || '').trim();
    }
    if (!headerA && spec.nom !== 'Documents') {
      rapport.problemes.push(
        spec.nom + ' : en-tete colonne A (ligne 5) absent — structure incompatible'
      );
      rapport.bloquant = true;
    }

    // Compter les lignes avec donnees dans la colonne identifiant
    var lastRow = feuille.getLastRow();
    var nbLignesDonnees = 0;
    if (lastRow >= spec.ligneDepart) {
      nbLignesDonnees = compterLignesAvecDonnees_(
        feuille, spec.colId, spec.ligneDepart,
        lastRow - spec.ligneDepart + 1
      );
    }

    // Determiner le label des colonnes effacees
    var labelsEffaces = 'aucune';
    if (spec.plagesEffacer) {
      labelsEffaces = spec.plagesEffacer.map(function(p) { return p.label; }).join(', ');
    } else {
      var lastCol = Math.max(1, feuille.getLastColumn());
      labelsEffaces = 'A:' + colLetter_(lastCol) + ' (detecte)';
    }

    rapport.ongletsReinitInfo.push({
      nom: spec.nom,
      nbLignesDonnees: nbLignesDonnees,
      formulesProtegees: spec.formulesProtegees.map(function(fp) { return fp.label; }),
      colonnesEffacees: labelsEffaces
    });
  });

  return rapport;
}

// ─── Helpers prives ───────────────────────────────────────────────────────────

// Produit une empreinte SHA-256 du contenu d'un onglet (valeurs, formules, notes).
function empreindreOnglet_(feuille) {
  var lastRow = feuille.getLastRow();
  var lastCol = feuille.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return 'VIDE-0x0';
  }

  var plage = feuille.getRange(1, 1, lastRow, lastCol);

  var serArr = function(arr2d) {
    return arr2d.map(function(r) { return r.join('\x01'); }).join('\x02');
  };

  var contenu = [
    lastRow,
    lastCol,
    serArr(plage.getDisplayValues()),
    serArr(plage.getFormulas()),
    serArr(plage.getNotes())
  ].join('\x03');

  var octets = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    contenu,
    Utilities.Charset.UTF_8
  );

  return octets.map(function(o) {
    return ('0' + (o < 0 ? o + 256 : o).toString(16)).slice(-2);
  }).join('');
}

// Compte les lignes non vides dans une colonne a partir d'une ligne de depart.
function compterLignesAvecDonnees_(feuille, col, ligneDepart, nbLignes) {
  if (nbLignes <= 0) return 0;

  return feuille
    .getRange(ligneDepart, col, nbLignes, 1)
    .getValues()
    .filter(function(r) { return String(r[0] || '').trim() !== ''; })
    .length;
}

// Capture les formules et valeurs affichees des colonnes protegees avant effacement.
function capturerFormulasProtegees_(feuille, ligneDepart, nbLignes, formulesProtegees) {
  var capture = {};
  if (nbLignes <= 0 || !formulesProtegees.length) return capture;

  formulesProtegees.forEach(function(fp) {
    var plage = feuille.getRange(ligneDepart, fp.col, nbLignes, fp.nbCols);
    capture[fp.label] = {
      formulas: plage.getFormulas(),
      displayValues: plage.getDisplayValues()
    };
  });

  return capture;
}

// Verifie que les colonnes protegees sont inchangees apres l'effacement.
// Retourne un tableau de messages d'erreur (vide si tout est intact).
function verifierFormulasProtegees_(feuille, ligneDepart, nbLignes, formulasAvant, formulesProtegees) {
  var erreurs = [];
  if (nbLignes <= 0 || !formulesProtegees.length) return erreurs;

  formulesProtegees.forEach(function(fp) {
    var avant = formulasAvant[fp.label];
    if (!avant) return;

    var formulesApres = feuille
      .getRange(ligneDepart, fp.col, nbLignes, fp.nbCols)
      .getFormulas();

    // Verifier les premieres lignes ayant une formule avant (echantillon representatif)
    var nCheck = Math.min(10, formulesApres.length);
    for (var i = 0; i < nCheck; i++) {
      for (var j = 0; j < formulesApres[i].length; j++) {
        if (formulesApres[i][j] !== avant.formulas[i][j]) {
          erreurs.push(
            feuille.getName() + ' col ' + fp.label +
            ' ligne ' + (ligneDepart + i) +
            ' : formule modifiee (attendu « ' + avant.formulas[i][j] +
            ' », obtenu « ' + formulesApres[i][j] + ' »)'
          );
          return;
        }
      }
    }
  });

  return erreurs;
}

// Convertit un numero de colonne en lettre(s) Excel (1=A, 26=Z, 27=AA ...).
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
