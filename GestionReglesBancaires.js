const CONFIG_REGLES_BANCAIRES_ = {
  nomFeuille: 'Règles bancaires',
  ligneEntetes: 5,
  premiereLigne: 6,
  COL_ID: 1,
  COL_ACTIF: 2,
  COL_PRIORITE: 3,
  COL_SENS: 4,
  COL_CORRESPONDANCE: 5,
  COL_TEXTE: 6,
  COL_TYPE: 7,
  COL_COMPOSANTE: 8,
  COL_FOURNISSEUR: 9,
  COL_CONTACT: 10,
  COL_COMPTE: 11,
  COL_PROGRAMME: 12,
  COL_PROJET: 13,
  COL_NOTES: 14,
  NOMBRE_COLONNES: 14
};

const COMPOSANTES_BANCAIRES_ = [
  'Entrée – Pion joues-tu?',
  'Entrée – Cartier',
  'Forfait Pion joues-tu?',
  'Forfait Cartier',
  'Forfait combiné',
  'T-shirt',
  'Chandail à manches longues',
  'Hoodie'
];

const TYPES_CLASSEMENT_REGLES_ = [
  'Dépense',
  'Revenu comptable direct',
  'Entrées, forfaits ou marchandises'
];

// ─── Installer public ──────────────────────────────────────────────────────────

function installerReglesBancairesConfigurables() {
  const verrou = LockService.getDocumentLock();
  verrou.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = preparerOngletReglesBancaires_(ss);
    const nombreMigrees = migrerReglesBancairesDepuisConfig_(ss, feuille);
    const nombreIdentifies = attribuerIdsReglesBancaires_(feuille);
    verifierIntegriteIdsRegles_(feuille);
    preparerColonnesExtenduesImportBancaire_(ss);

    SpreadsheetApp.getUi().alert(
      'Module de règles bancaires installé.\n\n' +
      'Règles migrées depuis Configuration!T:Z : ' + nombreMigrees + '\n' +
      'Identifiants attribués : ' + nombreIdentifies
    );
  } finally {
    verrou.releaseLock();
  }
}

// ─── Commandes publiques ───────────────────────────────────────────────────────

function validerReglesBancairesConfigurables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const regles = chargerReglesBancairesActives_(ss);
    SpreadsheetApp.getUi().alert(
      'Validation réussie\n\n' +
      regles.length + ' règle(s) active(s) valide(s) chargée(s).'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Erreur de validation des règles bancaires\n\n' + e.message
    );
  }
}

function reappliquerReglesBancairesAuxTransactionsAClasser() {
  const ui = SpreadsheetApp.getUi();

  const reponse = ui.alert(
    'Réappliquer les règles bancaires',
    'Cette opération va vider et recalculer les suggestions (colonnes H, I et P:W) ' +
    'pour toutes les transactions « À classer » dans Import bancaire.\n\n' +
    'Les transactions « Classée » et leurs données comptables ne seront pas modifiées.\n\n' +
    'Confirmer ?',
    ui.ButtonSet.YES_NO
  );

  if (reponse !== ui.Button.YES) {
    return;
  }

  const verrou = LockService.getDocumentLock();
  verrou.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = ss.getSheetByName('Import bancaire');

    if (!feuille || feuille.getLastRow() < 6) {
      ui.alert('Import bancaire est introuvable ou ne contient aucune transaction.');
      return;
    }

    preparerColonnesExtenduesImportBancaire_(ss);

    let regles;
    try {
      regles = chargerReglesBancairesActives_(ss);
    } catch (e) {
      ui.alert('Impossible de charger les règles bancaires.\n\n' + e.message);
      return;
    }

    const dictos = chargerDictionnairesNoms_(ss);
    const premiereTransaction = 6;
    const derniereLigne = feuille.getLastRow();
    const nbTotal = derniereLigne - premiereTransaction + 1;

    if (nbTotal <= 0) {
      ui.alert('Aucune transaction dans Import bancaire.');
      return;
    }

    // Lire A:W (23 colonnes) pour avoir toutes les données nécessaires
    const donnees = feuille
      .getRange(premiereTransaction, 1, nbTotal, 23)
      .getValues();

    let lignesAnalysees = 0;
    let reglesTrouvees = 0;
    let sansCorrespondance = 0;
    let suggestionsModifiees = 0;
    const conflits = [];

    for (let i = 0; i < nbTotal; i++) {
      const statut = String(donnees[i][10] || '').trim();  // col K (index 10)
      if (statut !== 'À classer') continue;

      lignesAnalysees += 1;

      const description = String(donnees[i][2] || '').trim();  // col C (index 2)
      const montant = Number(donnees[i][3]) || 0;               // col D (index 3)

      // Comptabiliser les suggestions précédentes (H, P, ou W non vides)
      const hActuel = String(donnees[i][7] || '').trim();   // col H (index 7)
      const pActuel = String(donnees[i][15] || '').trim();  // col P (index 15)
      const wActuel = String(donnees[i][22] || '').trim();  // col W (index 22)
      if (hActuel || pActuel || wActuel) {
        suggestionsModifiees += 1;
      }

      let suggestion = null;
      let estConflit = false;

      try {
        suggestion = rechercherRegleBancaireDansListe_(regles, description, montant);
      } catch (e) {
        estConflit = true;
        conflits.push({
          ligne: premiereTransaction + i,
          description: description,
          message: e.message
        });
      }

      if (suggestion) {
        reglesTrouvees += 1;
      } else if (!estConflit) {
        sansCorrespondance += 1;
      }

      const ligneSheet = premiereTransaction + i;

      if (suggestion) {
        const nomFournisseur = suggestion.idFournisseur
          ? (dictos.fournisseurs[suggestion.idFournisseur] || '')
          : '';
        const nomContact = suggestion.idContact
          ? (dictos.contacts[suggestion.idContact] || '')
          : '';

        // H:I — setNumberFormat('@') avant setValues
        feuille.getRange(ligneSheet, 8, 1, 2)
          .setNumberFormat('@')
          .setValues([[suggestion.codeCompte || '', suggestion.programme || '']]);

        // P:W — setNumberFormat('@') avant setValues
        feuille.getRange(ligneSheet, 16, 1, 8)
          .setNumberFormat('@')
          .setValues([[
            suggestion.idFournisseur  || '',
            nomFournisseur,
            suggestion.idContact      || '',
            nomContact,
            suggestion.typeClassement || '',
            suggestion.composante     || '',
            suggestion.projet         || '',
            suggestion.idRegle        || ''
          ]]);
      } else {
        // Vider les suggestions (conflit ou aucune correspondance)
        feuille.getRange(ligneSheet, 8, 1, 2).clearContent();
        feuille.getRange(ligneSheet, 16, 1, 8).clearContent();
      }
    }

    SpreadsheetApp.flush();

    const conflitsMsg = conflits.length > 0
      ? '\n\nConflits de règles (' + conflits.length + ') :\n' +
        conflits.map(function(c) {
          return '• Ligne ' + c.ligne + ' – ' +
            c.description.slice(0, 40) + '\n  ' + c.message;
        }).join('\n')
      : '';

    ui.alert(
      'Réapplication terminée\n\n' +
      'Lignes À classer analysées         : ' + lignesAnalysees + '\n' +
      'Règles correspondantes             : ' + reglesTrouvees + '\n' +
      'Sans correspondance                : ' + sansCorrespondance + '\n' +
      'Conflits de règles                 : ' + conflits.length + '\n' +
      'Suggestions effacées / remplacées  : ' + suggestionsModifiees +
      conflitsMsg
    );
  } finally {
    verrou.releaseLock();
  }
}

// ─── Loader public ─────────────────────────────────────────────────────────────

function chargerReglesBancairesActives_(ss) {
  const feuille = ss.getSheetByName(CONFIG_REGLES_BANCAIRES_.nomFeuille);
  const C = CONFIG_REGLES_BANCAIRES_;

  if (!feuille || feuille.getLastRow() < C.premiereLigne) {
    return [];
  }

  const nbLignes = feuille.getLastRow() - C.ligneEntetes;

  if (nbLignes <= 0) {
    return [];
  }

  const valeurs = feuille
    .getRange(C.premiereLigne, 1, nbLignes, C.NOMBRE_COLONNES)
    .getDisplayValues();

  const refs = chargerReferencesValidationRegles_(ss);
  const idsVus = new Set();
  const regles = [];
  const erreurs = [];

  valeurs.forEach(function(ligne, index) {
    const numLigne = C.premiereLigne + index;
    const actif = String(ligne[C.COL_ACTIF - 1] || '').trim();
    const texte = String(ligne[C.COL_TEXTE - 1] || '').trim();

    if (actif !== 'Oui' || !texte) {
      return;
    }

    const id = String(ligne[C.COL_ID - 1] || '').trim();
    const prioriteRaw = String(ligne[C.COL_PRIORITE - 1] || '').trim();
    const prioriteNum = Number(prioriteRaw);
    const prioriteValide = prioriteRaw !== '' && Number.isInteger(prioriteNum) && prioriteNum > 0;
    const sens = String(ligne[C.COL_SENS - 1] || '').trim();
    const correspondance = String(ligne[C.COL_CORRESPONDANCE - 1] || '').trim();
    const typeClassement = String(ligne[C.COL_TYPE - 1] || '').trim();
    const composante = String(ligne[C.COL_COMPOSANTE - 1] || '').trim();
    const idFournisseur = String(ligne[C.COL_FOURNISSEUR - 1] || '').trim();
    const idContact = String(ligne[C.COL_CONTACT - 1] || '').trim();
    const codeCompte = String(ligne[C.COL_COMPTE - 1] || '').trim();
    const programme = String(ligne[C.COL_PROGRAMME - 1] || '').trim();
    const projet = String(ligne[C.COL_PROJET - 1] || '').trim();

    let regleValide = true;

    function ajouterErreur(cause) {
      const msg = 'Ligne ' + numLigne +
        (id ? ' [' + id + ']' : '') +
        ' : ' + cause;
      erreurs.push(msg);
      Logger.log('Règles bancaires — ' + msg);
      regleValide = false;
    }

    if (!id) {
      ajouterErreur('ID manquant.');
    } else if (!/^REG-\d{4,}$/.test(id)) {
      ajouterErreur('ID « ' + id + ' » non conforme (format REG-0001 attendu).');
    }

    if (id && idsVus.has(id)) {
      ajouterErreur('ID « ' + id + ' » en doublon.');
    } else if (id) {
      idsVus.add(id);
    }

    if (!prioriteValide) {
      ajouterErreur('Priorité « ' + prioriteRaw + ' » invalide (entier positif attendu).');
    }

    if (sens !== 'Entrée' && sens !== 'Sortie') {
      ajouterErreur('Sens « ' + sens + ' » invalide (Entrée ou Sortie attendu).');
    }

    if (correspondance !== 'Contient' && correspondance !== 'Exacte') {
      ajouterErreur('Correspondance « ' + correspondance + ' » invalide (Contient ou Exacte attendu).');
    }

    if (typeClassement && TYPES_CLASSEMENT_REGLES_.indexOf(typeClassement) === -1) {
      ajouterErreur(
        'Type de classement « ' + typeClassement + ' » invalide (' +
        TYPES_CLASSEMENT_REGLES_.join(', ') + ' attendu).'
      );
    }

    if (idFournisseur) {
      if (refs.fournisseurs === null) {
        ajouterErreur(
          'Fournisseur « ' + idFournisseur + ' » renseigné mais la feuille Fournisseurs est absente.'
        );
      } else if (!refs.fournisseurs.has(idFournisseur)) {
        ajouterErreur('Fournisseur « ' + idFournisseur + ' » introuvable.');
      }
    }

    if (idContact) {
      if (refs.contacts === null) {
        ajouterErreur(
          'Contact « ' + idContact + ' » renseigné mais la feuille Contacts est absente.'
        );
      } else if (!refs.contacts.has(idContact)) {
        ajouterErreur('Contact « ' + idContact + ' » introuvable.');
      }
    }

    if (codeCompte) {
      if (refs.comptes === null) {
        ajouterErreur(
          'Compte « ' + codeCompte + ' » renseigné mais la feuille Configuration est absente.'
        );
      } else if (!refs.comptes.has(codeCompte)) {
        ajouterErreur('Compte « ' + codeCompte + ' » introuvable ou inactif.');
      }
    }

    if (programme) {
      if (refs.programmes === null) {
        ajouterErreur(
          'Programme « ' + programme + ' » renseigné mais Configuration!I6:I14 est absent.'
        );
      } else if (!refs.programmes.has(programme)) {
        ajouterErreur('Programme « ' + programme + ' » introuvable.');
      }
    }

    if (projet) {
      if (refs.projets === null) {
        ajouterErreur(
          'Projet « ' + projet + ' » renseigné mais la feuille Projets est absente.'
        );
      } else if (refs.projetsExemple.has(projet)) {
        ajouterErreur('Projet « ' + projet + ' » est un exemple.');
      } else if (!refs.projets.has(projet)) {
        ajouterErreur('Projet « ' + projet + ' » introuvable.');
      }
    }

    if (composante && !refs.composantes.has(composante)) {
      ajouterErreur('Composante « ' + composante + ' » inconnue.');
    }

    if (regleValide) {
      regles.push({
        idRegle: id,
        priorite: prioriteValide ? prioriteNum : 9999,
        sens: sens,
        correspondance: correspondance,
        texte: texte,
        typeClassement: typeClassement,
        composante: composante,
        idFournisseur: idFournisseur,
        idContact: idContact,
        codeCompte: codeCompte,
        programme: programme,
        projet: projet,
        numeroLigne: index
      });
    }
  });

  if (erreurs.length > 0) {
    throw new Error(
      'Règles bancaires invalides (' + erreurs.length + ') :\n' +
      erreurs.join('\n')
    );
  }

  return regles.sort(function(a, b) {
    if (a.priorite !== b.priorite) {
      return a.priorite - b.priorite;
    }
    return a.numeroLigne - b.numeroLigne;
  });
}

function chargerReferencesValidationRegles_(ss) {
  var fournisseurs = null;
  var contacts = null;
  var comptes = null;
  var programmes = null;
  var projets = null;
  var projetsExemple = null;

  var feuilleFournisseurs = ss.getSheetByName('Fournisseurs');
  if (feuilleFournisseurs && feuilleFournisseurs.getLastRow() >= 6) {
    fournisseurs = new Set();
    feuilleFournisseurs
      .getRange(6, 1, feuilleFournisseurs.getLastRow() - 5, 1)
      .getDisplayValues()
      .forEach(function(l) { if (l[0]) fournisseurs.add(String(l[0]).trim()); });
  }

  var feuilleContacts = ss.getSheetByName('Contacts');
  if (feuilleContacts && feuilleContacts.getLastRow() >= 6) {
    contacts = new Set();
    feuilleContacts
      .getRange(6, 1, feuilleContacts.getLastRow() - 5, 1)
      .getDisplayValues()
      .forEach(function(l) { if (l[0]) contacts.add(String(l[0]).trim()); });
  }

  var feuilleConfig = ss.getSheetByName('Configuration');
  if (feuilleConfig && feuilleConfig.getLastRow() >= 6) {
    comptes = new Set();
    programmes = new Set();
    feuilleConfig
      .getRange(6, 1, feuilleConfig.getLastRow() - 5, 5)
      .getDisplayValues()
      .forEach(function(l) {
        if (l[0] && l[4] === 'Oui') {
          comptes.add(String(l[0]).trim());
        }
      });
    feuilleConfig
      .getRange('I6:I14')
      .getDisplayValues()
      .forEach(function(l) { if (l[0]) programmes.add(String(l[0]).trim()); });
  }

  var feuilleProjets = ss.getSheetByName('Projets');
  if (feuilleProjets && feuilleProjets.getLastRow() >= 6) {
    projets = new Set();
    projetsExemple = new Set();
    feuilleProjets
      .getRange(6, 1, feuilleProjets.getLastRow() - 5, 4)
      .getDisplayValues()
      .forEach(function(l) {
        if (!l[0]) return;
        var id = String(l[0]).trim();
        if (String(l[3] || '').trim() === 'Exemple') {
          projetsExemple.add(id);
        } else {
          projets.add(id);
        }
      });
  }

  return {
    fournisseurs: fournisseurs,
    contacts: contacts,
    comptes: comptes,
    programmes: programmes,
    projets: projets,
    projetsExemple: projetsExemple,
    composantes: new Set(COMPOSANTES_BANCAIRES_)
  };
}

// ─── Moteur de recherche public ────────────────────────────────────────────────

function rechercherRegleBancaireDansListe_(regles, description, montant) {
  if (!montant) {
    return null;
  }

  const sensRecherche = montant > 0 ? 'Entrée' : 'Sortie';
  const descNormalisee = normaliserTexteRechercheRegle_(description);

  const correspondantes = regles.filter(function(r) {
    if (r.sens !== sensRecherche) {
      return false;
    }

    if (!r.texte) {
      return false;
    }

    const texteNorm = normaliserTexteRechercheRegle_(r.texte);

    if (r.correspondance === 'Exacte') {
      return descNormalisee === texteNorm;
    }

    return descNormalisee.indexOf(texteNorm) !== -1;
  });

  if (!correspondantes.length) {
    return null;
  }

  const meilleeurePriorite = Math.min.apply(
    null,
    correspondantes.map(function(r) { return r.priorite; })
  );

  const gagnantes = correspondantes.filter(function(r) {
    return r.priorite === meilleeurePriorite;
  });

  if (gagnantes.length > 1) {
    throw new Error(
      'Conflit de règles bancaires : plusieurs règles de même priorité (' +
      meilleeurePriorite + ') correspondent à « ' + description + ' » : ' +
      gagnantes.map(function(r) { return r.idRegle || '(sans ID)'; }).join(', ')
    );
  }

  const r = gagnantes[0];

  return {
    idRegle: r.idRegle,
    typeClassement: r.typeClassement,
    composante: r.composante,
    idFournisseur: r.idFournisseur,
    idContact: r.idContact,
    codeCompte: r.codeCompte,
    programme: r.programme,
    projet: r.projet
  };
}

// ─── Préparation de l'onglet ───────────────────────────────────────────────────

function preparerOngletReglesBancaires_(ss) {
  let feuille = ss.getSheetByName(CONFIG_REGLES_BANCAIRES_.nomFeuille);
  const C = CONFIG_REGLES_BANCAIRES_;

  if (!feuille) {
    feuille = ss.insertSheet(C.nomFeuille);
    feuille.setHiddenGridlines(true);
  }

  if (!String(feuille.getRange(1, 1).getValue() || '').trim()) {
    feuille.getRange(1, 1, 1, C.NOMBRE_COLONNES)
      .merge()
      .setValue('Règles bancaires')
      .setBackground('#e8f0fe')
      .setFontWeight('bold');
  }

  if (!String(feuille.getRange(3, 1).getValue() || '').trim()) {
    feuille.getRange(3, 1, 1, C.NOMBRE_COLONNES)
      .merge()
      .setValue(
        'Ces règles servent de suggestions uniquement et ne déclenchent aucune comptabilisation automatique.'
      )
      .setFontStyle('italic')
      .setWrap(true);
  }

  const entetes = [
    'ID règle', 'Actif', 'Priorité', 'Sens', 'Correspondance',
    'Texte à reconnaître', 'Type de classement', 'Composante',
    'ID fournisseur', 'ID contact', 'Code compte', 'Programme', 'Projet', 'Notes'
  ];

  verifierEtEcrireEntetesRegles_(feuille, entetes);

  if (feuille.getFrozenRows() < C.ligneEntetes) {
    feuille.setFrozenRows(C.ligneEntetes);
  }

  if (!feuille.getFilter()) {
    feuille
      .getRange(C.ligneEntetes, 1, feuille.getMaxRows() - C.ligneEntetes + 1, C.NOMBRE_COLONNES)
      .createFilter();
  }

  appliquerValidationsReglesBancaires_(ss, feuille);

  return feuille;
}

function verifierEtEcrireEntetesRegles_(feuille, entetes) {
  const C = CONFIG_REGLES_BANCAIRES_;
  const cellule = feuille.getRange(C.ligneEntetes, 1, 1, C.NOMBRE_COLONNES);
  const valeurs = cellule.getDisplayValues()[0];
  const premierEntete = String(valeurs[0] || '').trim();

  if (!premierEntete) {
    cellule
      .setValues([entetes])
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
    return;
  }

  const incompatibles = entetes.filter(function(titre, i) {
    return String(valeurs[i] || '').trim() !== titre;
  });

  if (incompatibles.length > 0) {
    throw new Error(
      "L'onglet « " + C.nomFeuille +
      " » existe avec des en-têtes incompatibles dans les colonnes : " +
      incompatibles.join(', ') +
      '. Colonnes attendues : ' + entetes.join(', ')
    );
  }
}

function appliquerValidationsReglesBancaires_(ss, feuille) {
  const C = CONFIG_REGLES_BANCAIRES_;
  const plageData = feuille.getMaxRows() - C.premiereLigne + 1;

  if (plageData <= 0) {
    return;
  }

  feuille.getRange(C.premiereLigne, C.COL_ID, plageData, 1)
    .setNumberFormat('@');

  feuille.getRange(C.premiereLigne, C.COL_ACTIF, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Oui', 'Non'], true)
        .setAllowInvalid(false)
        .build()
    );

  feuille.getRange(C.premiereLigne, C.COL_PRIORITE, plageData, 1)
    .setNumberFormat('0');

  feuille.getRange(C.premiereLigne, C.COL_SENS, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Entrée', 'Sortie'], true)
        .setAllowInvalid(false)
        .build()
    );

  feuille.getRange(C.premiereLigne, C.COL_CORRESPONDANCE, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Contient', 'Exacte'], true)
        .setAllowInvalid(false)
        .build()
    );

  feuille.getRange(C.premiereLigne, C.COL_TYPE, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(TYPES_CLASSEMENT_REGLES_, true)
        .setAllowInvalid(false)
        .build()
    );

  feuille.getRange(C.premiereLigne, C.COL_COMPOSANTE, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(COMPOSANTES_BANCAIRES_, true)
        .setAllowInvalid(true)
        .build()
    );

  feuille.getRange(C.premiereLigne, C.COL_FOURNISSEUR, plageData, 1)
    .setNumberFormat('@');

  const feuilleFournisseurs = ss.getSheetByName('Fournisseurs');
  if (feuilleFournisseurs && feuilleFournisseurs.getLastRow() >= 6) {
    feuille.getRange(C.premiereLigne, C.COL_FOURNISSEUR, plageData, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(
            feuilleFournisseurs.getRange(6, 1, feuilleFournisseurs.getLastRow() - 5, 1),
            true
          )
          .setAllowInvalid(true)
          .build()
      );
  }

  feuille.getRange(C.premiereLigne, C.COL_CONTACT, plageData, 1)
    .setNumberFormat('@');

  const feuilleContacts = ss.getSheetByName('Contacts');
  if (feuilleContacts && feuilleContacts.getLastRow() >= 6) {
    feuille.getRange(C.premiereLigne, C.COL_CONTACT, plageData, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(
            feuilleContacts.getRange(6, 1, feuilleContacts.getLastRow() - 5, 1),
            true
          )
          .setAllowInvalid(true)
          .build()
      );
  }

  feuille.getRange(C.premiereLigne, C.COL_COMPTE, plageData, 1)
    .setNumberFormat('@');

  const feuilleConfig = ss.getSheetByName('Configuration');
  if (feuilleConfig && feuilleConfig.getLastRow() >= 6) {
    const codesComptes = feuilleConfig
      .getRange(6, 1, feuilleConfig.getLastRow() - 5, 5)
      .getDisplayValues()
      .filter(function(l) { return l[0] && l[4] === 'Oui'; })
      .map(function(l) { return String(l[0]).trim(); });

    if (codesComptes.length > 0) {
      feuille.getRange(C.premiereLigne, C.COL_COMPTE, plageData, 1)
        .setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(codesComptes, true)
            .setAllowInvalid(true)
            .build()
        );
    }

    const programmes = feuilleConfig
      .getRange('I6:I14')
      .getDisplayValues()
      .flat()
      .filter(String);

    if (programmes.length > 0) {
      feuille.getRange(C.premiereLigne, C.COL_PROGRAMME, plageData, 1)
        .setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(programmes, true)
            .setAllowInvalid(true)
            .build()
        );
    }
  }

  const feuilleProjets = ss.getSheetByName('Projets');
  if (feuilleProjets && feuilleProjets.getLastRow() >= 6) {
    const idsProjets = feuilleProjets
      .getRange(6, 1, feuilleProjets.getLastRow() - 5, 4)
      .getDisplayValues()
      .filter(function(l) { return l[0] && l[3] !== 'Exemple'; })
      .map(function(l) { return String(l[0]).trim(); });

    if (idsProjets.length > 0) {
      feuille.getRange(C.premiereLigne, C.COL_PROJET, plageData, 1)
        .setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(idsProjets, true)
            .setAllowInvalid(true)
            .build()
        );
    }
  }
}

function migrerReglesBancairesDepuisConfig_(ss, feuille) {
  const config = ss.getSheetByName('Configuration');
  const C = CONFIG_REGLES_BANCAIRES_;

  if (!config || config.getLastRow() < 6) {
    return 0;
  }

  const reglesSources = config
    .getRange(6, 20, config.getLastRow() - 5, 7)
    .getValues();

  const clesExistantes = new Set();
  const derniereLigneFeuille = feuille.getLastRow();

  if (derniereLigneFeuille >= C.premiereLigne) {
    feuille
      .getRange(C.premiereLigne, 1, derniereLigneFeuille - C.ligneEntetes, C.NOMBRE_COLONNES)
      .getDisplayValues()
      .forEach(function(ligne) {
        const sens = String(ligne[C.COL_SENS - 1] || '').trim();
        const texte = String(ligne[C.COL_TEXTE - 1] || '').trim();
        if (sens && texte) {
          clesExistantes.add(sens + '|' + normaliserTexteRechercheRegle_(texte));
        }
      });
  }

  let nombreMigrees = 0;
  let ligneEcriture = Math.max(feuille.getLastRow() + 1, C.premiereLigne);

  reglesSources.forEach(function(ligne) {
    const sens = String(ligne[0] || '').trim();
    const texte = String(ligne[1] || '').trim();
    const idFournisseur = String(ligne[2] || '').trim();
    const idContact = String(ligne[3] || '').trim();
    const codeCompte = String(ligne[4] || '').trim();
    const programme = String(ligne[5] || '').trim();
    const actif = String(ligne[6] || '').trim();

    if (!sens || !texte) {
      return;
    }

    const cle = sens + '|' + normaliserTexteRechercheRegle_(texte);
    if (clesExistantes.has(cle)) {
      return;
    }

    let type = 'Dépense';
    if (sens === 'Entrée') {
      type = ['4000', '4010', '4020'].indexOf(codeCompte) !== -1
        ? 'Entrées, forfaits ou marchandises'
        : 'Revenu comptable direct';
    }

    feuille.getRange(ligneEcriture, 1, 1, C.NOMBRE_COLONNES).setValues([[
      '',
      actif || 'Oui',
      100,
      sens,
      'Contient',
      texte,
      type,
      '',
      idFournisseur,
      idContact,
      codeCompte,
      programme,
      '',
      'Migrée depuis Configuration!T:Z'
    ]]);

    feuille.getRange(ligneEcriture, C.COL_ID).setNumberFormat('@');
    feuille.getRange(ligneEcriture, C.COL_FOURNISSEUR).setNumberFormat('@');
    feuille.getRange(ligneEcriture, C.COL_CONTACT).setNumberFormat('@');
    feuille.getRange(ligneEcriture, C.COL_COMPTE).setNumberFormat('@');

    clesExistantes.add(cle);
    ligneEcriture += 1;
    nombreMigrees += 1;
  });

  return nombreMigrees;
}

function attribuerIdsReglesBancaires_(feuille) {
  const C = CONFIG_REGLES_BANCAIRES_;
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < C.premiereLigne) {
    return 0;
  }

  const nbLignes = derniereLigne - C.ligneEntetes;

  if (nbLignes <= 0) {
    return 0;
  }

  const colonnesIDTexte = feuille
    .getRange(C.premiereLigne, 1, nbLignes, C.COL_TEXTE)
    .getDisplayValues();

  // Passe 1 : déterminer le numéro maximum déjà attribué
  let numeroMax = 0;
  colonnesIDTexte.forEach(function(ligne) {
    const id = String(ligne[C.COL_ID - 1] || '').trim();
    const match = id.match(/^REG-(\d+)$/);
    if (match) {
      numeroMax = Math.max(numeroMax, Number(match[1]));
    }
  });

  // Passe 2 : identifier les lignes sans ID ayant un texte à reconnaître
  const aEcrire = [];
  colonnesIDTexte.forEach(function(ligne, i) {
    const id = String(ligne[C.COL_ID - 1] || '').trim();
    const texte = String(ligne[C.COL_TEXTE - 1] || '').trim();
    if (!id && texte) {
      numeroMax += 1;
      aEcrire.push({
        ligneSheet: C.premiereLigne + i,
        id: 'REG-' + String(numeroMax).padStart(4, '0')
      });
    }
  });

  if (aEcrire.length === 0) {
    return 0;
  }

  // Écriture : setNumberFormat('@') AVANT setValue pour garantir le stockage en texte
  aEcrire.forEach(function(item) {
    feuille.getRange(item.ligneSheet, C.COL_ID)
      .setNumberFormat('@')
      .setValue(item.id);
  });

  SpreadsheetApp.flush();

  // Relecture groupée pour confirmer la persistance de chaque identifiant
  const minLigne = aEcrire[0].ligneSheet;
  const maxLigne = aEcrire[aEcrire.length - 1].ligneSheet;
  const relues = feuille
    .getRange(minLigne, C.COL_ID, maxLigne - minLigne + 1, 1)
    .getDisplayValues();

  let nombreAttribues = 0;
  const echecEcritures = [];

  aEcrire.forEach(function(item) {
    const valeur = String(relues[item.ligneSheet - minLigne][0] || '').trim();
    if (valeur === item.id) {
      nombreAttribues += 1;
    } else {
      echecEcritures.push(
        'Ligne ' + item.ligneSheet + ' : attendu « ' + item.id + ' », lu « ' + valeur + ' ».'
      );
    }
  });

  if (echecEcritures.length > 0) {
    throw new Error(
      'Échec d\'écriture des identifiants dans la colonne A (' +
      echecEcritures.length + ' cellule(s)) :\n' +
      echecEcritures.join('\n')
    );
  }

  return nombreAttribues;
}

function verifierIntegriteIdsRegles_(feuille) {
  const C = CONFIG_REGLES_BANCAIRES_;
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < C.premiereLigne) {
    return;
  }

  const nbLignes = derniereLigne - C.ligneEntetes;

  if (nbLignes <= 0) {
    return;
  }

  const colonnesIDTexte = feuille
    .getRange(C.premiereLigne, 1, nbLignes, C.COL_TEXTE)
    .getDisplayValues();

  const erreurs = [];
  const idsVus = new Set();

  colonnesIDTexte.forEach(function(ligne, i) {
    const id = String(ligne[C.COL_ID - 1] || '').trim();
    const texte = String(ligne[C.COL_TEXTE - 1] || '').trim();

    if (!texte) {
      return;
    }

    if (!id) {
      erreurs.push(
        'Ligne ' + (C.premiereLigne + i) + ' : texte présent mais identifiant manquant en colonne A.'
      );
    } else if (idsVus.has(id)) {
      erreurs.push(
        'Ligne ' + (C.premiereLigne + i) + ' : identifiant « ' + id + ' » en doublon.'
      );
    } else {
      idsVus.add(id);
    }
  });

  if (erreurs.length > 0) {
    throw new Error(
      'Intégrité des identifiants compromise (' + erreurs.length + ' problème(s)) :\n' +
      erreurs.join('\n')
    );
  }
}

// ─── Helpers d'application (partagés avec ImportCsv.js) ──────────────────────

function chargerDictionnairesNoms_(ss) {
  const fournisseurs = {};
  const contacts = {};

  const feuilleFournisseurs = ss.getSheetByName('Fournisseurs');
  if (feuilleFournisseurs && feuilleFournisseurs.getLastRow() >= 6) {
    feuilleFournisseurs
      .getRange(6, 1, feuilleFournisseurs.getLastRow() - 5, 2)
      .getDisplayValues()
      .forEach(function(l) {
        const id = String(l[0] || '').trim();
        if (id) fournisseurs[id] = String(l[1] || '').trim();
      });
  }

  const feuilleContacts = ss.getSheetByName('Contacts');
  if (feuilleContacts && feuilleContacts.getLastRow() >= 6) {
    feuilleContacts
      .getRange(6, 1, feuilleContacts.getLastRow() - 5, 4)
      .getDisplayValues()
      .forEach(function(l) {
        const id = String(l[0] || '').trim();
        if (id) contacts[id] = String(l[3] || '').trim();
      });
  }

  return { fournisseurs: fournisseurs, contacts: contacts };
}

function preparerColonnesExtenduesImportBancaire_(ss) {
  const feuille = ss.getSheetByName('Import bancaire');
  if (!feuille) return;

  // ── 1. Agrandir la feuille avant tout getRange sur T:W ────────────────────
  const colMax = feuille.getMaxColumns();
  if (colMax < 23) {
    feuille.insertColumnsAfter(colMax, 23 - colMax);
  }

  // ── 2. En-têtes T:W (idempotent, erreur si incompatible) ──────────────────
  const lettreColonne = { 20: 'T', 21: 'U', 22: 'V', 23: 'W' };

  const attendus = [
    { col: 20, valeur: 'Type de classement suggéré', largeur: 190 },
    { col: 21, valeur: 'Composante suggérée',         largeur: 170 },
    { col: 22, valeur: 'Projet suggéré',              largeur: 130 },
    { col: 23, valeur: 'ID règle appliquée',          largeur: 130 }
  ];

  attendus.forEach(function(entete) {
    const cellule = feuille.getRange(5, entete.col);
    const valeurActuelle = String(cellule.getValue() || '').trim();

    if (!valeurActuelle) {
      cellule
        .setValue(entete.valeur)
        .setBackground('#f1f3f4')
        .setFontWeight('bold')
        .setWrap(true);
      feuille.setColumnWidth(entete.col, entete.largeur);
    } else if (valeurActuelle !== entete.valeur) {
      throw new Error(
        'Import bancaire : la cellule ' +
        lettreColonne[entete.col] + '5 contient « ' + valeurActuelle +
        ' » au lieu de « ' + entete.valeur +
        ' ». Corrigez manuellement avant de continuer.'
      );
    }
  });

  // ── 3. Formats et validations sur les données (ligne 6 et plus) ────────────
  const derniereLigne = feuille.getMaxRows();
  if (derniereLigne < 6) return;

  const plageData = derniereLigne - 5;

  // Format texte sur l'ensemble de T:W
  feuille.getRange(6, 20, plageData, 4).setNumberFormat('@');

  // T (col 20) : validation Type de classement
  feuille.getRange(6, 20, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(TYPES_CLASSEMENT_REGLES_, true)
        .setAllowInvalid(true)
        .build()
    );

  // U (col 21) : validation Composante
  feuille.getRange(6, 21, plageData, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(COMPOSANTES_BANCAIRES_, true)
        .setAllowInvalid(true)
        .build()
    );

  // V (col 22) : validation Projets réels (excluant Exemple)
  const feuilleProjets = ss.getSheetByName('Projets');
  if (feuilleProjets && feuilleProjets.getLastRow() >= 6) {
    const idsProjets = feuilleProjets
      .getRange(6, 1, feuilleProjets.getLastRow() - 5, 4)
      .getDisplayValues()
      .filter(function(l) { return l[0] && l[3] !== 'Exemple'; })
      .map(function(l) { return String(l[0]).trim(); });

    if (idsProjets.length > 0) {
      feuille.getRange(6, 22, plageData, 1)
        .setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(idsProjets, true)
            .setAllowInvalid(true)
            .build()
        );
    }
  }

  // W (col 23) : format texte uniquement (déjà inclus dans le setNumberFormat groupé)
}
