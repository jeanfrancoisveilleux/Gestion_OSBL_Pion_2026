const CONFIG_BUDGET_ANNUEL = {
  nomFeuille: 'Budget 2026',
  ligneEntetes: 5,
  premiereLigne: 6,
  nombreColonnes: 12,
  periodeAnnuelle: 'Annuel',
  statutActif: 'Actif',
  statutInactif: 'Inactif'
};

function installerGestionBudgetAnnuel() {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultat = preparerStructureBudgetAnnuel_(ss);

    installerDeclencheurMenuBudgetAnnuel_(ss);
    ajouterMenuBudgetAnnuelAuDemarrage();

    SpreadsheetApp.getUi().alert(
      'Gestion du budget annuel installée.\n\n' +
      'Les budgets sont saisis par année, programme et compte comptable.\n' +
      'Les montants réels proviennent automatiquement du Journal.\n' +
      'Lignes budgétaires actives : ' +
      resultat.nombreLignesActives +
      '.'
    );
  } finally {
    verrou.releaseLock();
  }
}

function ajouterMenuBudgetAnnuelAuDemarrage() {
  SpreadsheetApp.getUi()
    .createMenu('Budget annuel')
    .addItem(
      'Gérer le budget annuel',
      'afficherGestionBudgetAnnuel'
    )
    .addSeparator()
    .addItem(
      'Installer ou réparer le budget annuel',
      'installerGestionBudgetAnnuel'
    )
    .addToUi();
}

function afficherGestionBudgetAnnuel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  preparerStructureBudgetAnnuel_(ss);

  const modele = HtmlService.createTemplateFromFile('BudgetAnnuel');
  modele.DONNEES_INITIALES = obtenirDonneesGestionBudgetAnnuel();

  const html = modele
    .evaluate()
    .setWidth(1080)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Gestion du budget annuel'
  );
}

function obtenirDonneesGestionBudgetAnnuel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = obtenirFeuilleBudgetAnnuel_(ss);
  const comptes = chargerComptesBudgetAnnuel_(ss);
  const programmes = chargerProgrammesBudgetAnnuel_(ss);
  const anneeDefaut = obtenirAnneeDefautBudgetAnnuel_(feuille);

  SpreadsheetApp.flush();

  const derniereLigne = feuille.getLastRow();
  const lignes = [];

  if (derniereLigne >= CONFIG_BUDGET_ANNUEL.premiereLigne) {
    const nombreLignes =
      derniereLigne -
      CONFIG_BUDGET_ANNUEL.premiereLigne +
      1;

    const valeurs = feuille
      .getRange(
        CONFIG_BUDGET_ANNUEL.premiereLigne,
        1,
        nombreLignes,
        CONFIG_BUDGET_ANNUEL.nombreColonnes
      )
      .getValues();

    valeurs.forEach(function(ligne, index) {
      const annee = Number(ligne[0] || 0);
      const periode = String(ligne[1] || '').trim();
      const programme = String(ligne[2] || '').trim();
      const codeCompte = String(ligne[3] || '').trim();
      const statut = String(ligne[10] || '').trim();

      if (
        !annee ||
        periode !== CONFIG_BUDGET_ANNUEL.periodeAnnuelle ||
        !programme ||
        !codeCompte ||
        statut === CONFIG_BUDGET_ANNUEL.statutInactif
      ) {
        return;
      }

      lignes.push({
        numeroLigne:
          CONFIG_BUDGET_ANNUEL.premiereLigne + index,
        annee: annee,
        programme: programme,
        codeCompte: codeCompte,
        nomCompte: String(ligne[4] || '').trim(),
        type: String(ligne[5] || '').trim(),
        budget: arrondirBudgetAnnuel_(ligne[6]),
        reel: arrondirBudgetAnnuel_(ligne[7]),
        ecart: arrondirBudgetAnnuel_(ligne[8]),
        pourcentage: Number(ligne[9] || 0),
        statut:
          statut || CONFIG_BUDGET_ANNUEL.statutActif,
        notes: String(ligne[11] || '')
      });
    });
  }

  lignes.sort(function(a, b) {
    return (
      a.annee - b.annee ||
      a.programme.localeCompare(b.programme, 'fr') ||
      a.codeCompte.localeCompare(b.codeCompte, 'fr')
    );
  });

  return {
    anneeDefaut: anneeDefaut,
    periode: CONFIG_BUDGET_ANNUEL.periodeAnnuelle,
    comptes: comptes,
    programmes: programmes,
    lignes: lignes
  };
}

function enregistrerBudgetAnnuel(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = obtenirFeuilleBudgetAnnuel_(ss);
    const validation = validerDonneesBudgetAnnuel_(
      ss,
      donnees || {}
    );
    const numeroLigneDemande = Number(
      donnees.numeroLigne || 0
    );
    const derniereLigne = Math.max(
      feuille.getLastRow(),
      CONFIG_BUDGET_ANNUEL.premiereLigne
    );
    const nombreLignes =
      derniereLigne -
      CONFIG_BUDGET_ANNUEL.premiereLigne +
      1;
    const valeursExistantes = feuille
      .getRange(
        CONFIG_BUDGET_ANNUEL.premiereLigne,
        1,
        nombreLignes,
        CONFIG_BUDGET_ANNUEL.nombreColonnes
      )
      .getValues();

    let numeroLigne = 0;

    if (numeroLigneDemande) {
      if (
        numeroLigneDemande <
          CONFIG_BUDGET_ANNUEL.premiereLigne ||
        numeroLigneDemande > feuille.getMaxRows()
      ) {
        throw new Error(
          'La ligne budgétaire à modifier est invalide.'
        );
      }

      const ligneActuelle = feuille
        .getRange(
          numeroLigneDemande,
          1,
          1,
          CONFIG_BUDGET_ANNUEL.nombreColonnes
        )
        .getValues()[0];

      if (
        String(ligneActuelle[1] || '').trim() !==
          CONFIG_BUDGET_ANNUEL.periodeAnnuelle ||
        String(ligneActuelle[10] || '').trim() ===
          CONFIG_BUDGET_ANNUEL.statutInactif
      ) {
        throw new Error(
          'Cette ligne budgétaire ne peut plus être modifiée.'
        );
      }

      numeroLigne = numeroLigneDemande;
    }

    valeursExistantes.forEach(function(ligne, index) {
      const ligneFeuille =
        CONFIG_BUDGET_ANNUEL.premiereLigne + index;

      if (ligneFeuille === numeroLigne) {
        return;
      }

      const memeCle =
        Number(ligne[0] || 0) === validation.annee &&
        String(ligne[1] || '').trim() ===
          CONFIG_BUDGET_ANNUEL.periodeAnnuelle &&
        String(ligne[2] || '').trim() ===
          validation.programme &&
        String(ligne[3] || '').trim() ===
          validation.codeCompte &&
        String(ligne[10] || '').trim() !==
          CONFIG_BUDGET_ANNUEL.statutInactif;

      if (memeCle) {
        throw new Error(
          'Un budget actif existe déjà pour cette année, ' +
          'ce programme et ce compte.'
        );
      }
    });

    if (!numeroLigne) {
      numeroLigne = trouverLigneVideBudgetAnnuel_(
        feuille,
        valeursExistantes
      );
    }

    feuille
      .getRange(numeroLigne, 1)
      .setNumberFormat('0')
      .setValue(validation.annee);

    feuille
      .getRange(numeroLigne, 2)
      .setValue(CONFIG_BUDGET_ANNUEL.periodeAnnuelle);

    feuille
      .getRange(numeroLigne, 3)
      .setValue(validation.programme);

    feuille
      .getRange(numeroLigne, 4)
      .setNumberFormat('@')
      .setValue(validation.codeCompte);

    feuille
      .getRange(numeroLigne, 7)
      .setNumberFormat('$#,##0.00')
      .setValue(validation.montant);

    feuille
      .getRange(numeroLigne, 11)
      .setValue(CONFIG_BUDGET_ANNUEL.statutActif);

    feuille
      .getRange(numeroLigne, 12)
      .setValue(validation.notes);

    appliquerFormulesLigneBudgetAnnuel_(
      feuille,
      numeroLigne
    );
    appliquerValidationsLigneBudgetAnnuel_(
      ss,
      feuille,
      numeroLigne
    );

    SpreadsheetApp.flush();

    return {
      message: numeroLigneDemande
        ? 'Budget annuel modifié.'
        : 'Budget annuel ajouté.',
      donnees: obtenirDonneesGestionBudgetAnnuel()
    };
  } finally {
    verrou.releaseLock();
  }
}

function desactiverBudgetAnnuel(numeroLigne) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = obtenirFeuilleBudgetAnnuel_(ss);
    const ligne = Number(numeroLigne || 0);

    if (
      ligne < CONFIG_BUDGET_ANNUEL.premiereLigne ||
      ligne > feuille.getLastRow()
    ) {
      throw new Error(
        'La ligne budgétaire à désactiver est invalide.'
      );
    }

    const valeurs = feuille
      .getRange(
        ligne,
        1,
        1,
        CONFIG_BUDGET_ANNUEL.nombreColonnes
      )
      .getValues()[0];

    if (
      String(valeurs[1] || '').trim() !==
      CONFIG_BUDGET_ANNUEL.periodeAnnuelle
    ) {
      throw new Error(
        'Cette ligne ne correspond pas à un budget annuel.'
      );
    }

    feuille
      .getRange(ligne, 11)
      .setValue(CONFIG_BUDGET_ANNUEL.statutInactif);

    SpreadsheetApp.flush();

    return {
      message:
        'Budget annuel désactivé. Son historique est conservé.',
      donnees: obtenirDonneesGestionBudgetAnnuel()
    };
  } finally {
    verrou.releaseLock();
  }
}

function preparerStructureBudgetAnnuel_(ss) {
  const feuille = obtenirFeuilleBudgetAnnuel_(ss);
  const maxLignesSouhaite = 500;

  if (
    feuille.getMaxColumns() <
    CONFIG_BUDGET_ANNUEL.nombreColonnes
  ) {
    feuille.insertColumnsAfter(
      feuille.getMaxColumns(),
      CONFIG_BUDGET_ANNUEL.nombreColonnes -
        feuille.getMaxColumns()
    );
  }

  if (feuille.getMaxRows() < maxLignesSouhaite) {
    feuille.insertRowsAfter(
      feuille.getMaxRows(),
      maxLignesSouhaite - feuille.getMaxRows()
    );
  }

  feuille
    .getRange('A1:L1')
    .merge()
    .setValue(
      'Budget annuel par programme et compte — ' +
      'les résultats réels proviennent du Journal.'
    )
    .setBackground('#d9e8fb')
    .setFontWeight('bold')
    .setFontSize(13);

  feuille
    .getRange('A2:L2')
    .merge()
    .setValue(
      'Utilisez le menu Budget annuel. Une ligne active ' +
      'par année, programme et compte.'
    )
    .setFontStyle('italic')
    .setFontColor('#5f6368');

  feuille
    .getRange(
      CONFIG_BUDGET_ANNUEL.ligneEntetes,
      1,
      1,
      12
    )
    .setValues([[
      'Année',
      'Mois',
      'Programme',
      'Code compte',
      'Nom du compte',
      'Type',
      'Budget',
      'Réel',
      'Écart favorable',
      '% utilisé',
      'Statut',
      'Notes'
    ]])
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setWrap(true);

  feuille.setFrozenRows(
    CONFIG_BUDGET_ANNUEL.ligneEntetes
  );

  const nombreLignes =
    feuille.getMaxRows() -
    CONFIG_BUDGET_ANNUEL.premiereLigne +
    1;

  appliquerFormulesPlageBudgetAnnuel_(
    feuille,
    CONFIG_BUDGET_ANNUEL.premiereLigne,
    nombreLignes
  );

  appliquerValidationsBudgetAnnuel_(
    ss,
    feuille,
    CONFIG_BUDGET_ANNUEL.premiereLigne,
    nombreLignes
  );

  feuille
    .getRange(
      CONFIG_BUDGET_ANNUEL.premiereLigne,
      7,
      nombreLignes,
      3
    )
    .setNumberFormat('$#,##0.00');

  feuille
    .getRange(
      CONFIG_BUDGET_ANNUEL.premiereLigne,
      10,
      nombreLignes,
      1
    )
    .setNumberFormat('0.0%');

  feuille.setColumnWidth(1, 75);
  feuille.setColumnWidth(2, 85);
  feuille.setColumnWidth(3, 190);
  feuille.setColumnWidth(4, 105);
  feuille.setColumnWidth(5, 250);
  feuille.setColumnWidth(6, 95);
  feuille.setColumnWidths(7, 3, 115);
  feuille.setColumnWidth(10, 95);
  feuille.setColumnWidth(11, 90);
  feuille.setColumnWidth(12, 260);

  if (!feuille.getFilter()) {
    feuille
      .getRange(
        CONFIG_BUDGET_ANNUEL.ligneEntetes,
        1,
        feuille.getMaxRows() -
          CONFIG_BUDGET_ANNUEL.ligneEntetes +
          1,
        CONFIG_BUDGET_ANNUEL.nombreColonnes
      )
      .createFilter();
  }

  const derniereLigne = feuille.getLastRow();
  let nombreLignesActives = 0;

  if (
    derniereLigne >=
    CONFIG_BUDGET_ANNUEL.premiereLigne
  ) {
    const valeurs = feuille
      .getRange(
        CONFIG_BUDGET_ANNUEL.premiereLigne,
        1,
        derniereLigne -
          CONFIG_BUDGET_ANNUEL.premiereLigne +
          1,
        11
      )
      .getValues();

    valeurs.forEach(function(ligne, index) {
      const contientBudget =
        ligne[0] !== '' &&
        String(ligne[2] || '').trim() &&
        String(ligne[3] || '').trim();

      if (!contientBudget) {
        return;
      }

      const celluleStatut = feuille.getRange(
        CONFIG_BUDGET_ANNUEL.premiereLigne + index,
        11
      );

      if (!String(ligne[10] || '').trim()) {
        celluleStatut.setValue(
          CONFIG_BUDGET_ANNUEL.statutActif
        );
      }

      if (
        String(
          celluleStatut.getValue() || ''
        ).trim() === CONFIG_BUDGET_ANNUEL.statutActif
      ) {
        nombreLignesActives++;
      }
    });
  }

  mettreAJourPlagesBudgetRapports_(ss);

  return {
    nombreLignesActives: nombreLignesActives
  };
}

function obtenirFeuilleBudgetAnnuel_(ss) {
  let feuille = ss.getSheetByName(
    CONFIG_BUDGET_ANNUEL.nomFeuille
  );

  if (!feuille) {
    const feuillesBudget = ss
      .getSheets()
      .filter(function(candidate) {
        return /^Budget\s+\d{4}$/i.test(
          candidate.getName()
        );
      });

    if (feuillesBudget.length) {
      feuille = feuillesBudget[0];
    }
  }

  if (!feuille) {
    feuille = ss.insertSheet(
      CONFIG_BUDGET_ANNUEL.nomFeuille
    );
  }

  return feuille;
}

function chargerComptesBudgetAnnuel_(ss) {
  const configuration = ss.getSheetByName(
    'Configuration'
  );

  if (!configuration) {
    throw new Error(
      "L’onglet Configuration est introuvable."
    );
  }

  const derniereLigne = configuration.getLastRow();

  if (derniereLigne < 6) {
    return [];
  }

  const valeurs = configuration
    .getRange(6, 1, derniereLigne - 5, 5)
    .getDisplayValues();

  return valeurs
    .filter(function(ligne) {
      const type = String(ligne[2] || '').trim();
      const actif = String(ligne[4] || '')
        .trim()
        .toLowerCase();

      return (
        String(ligne[0] || '').trim() &&
        (type === 'Revenu' || type === 'Dépense') &&
        actif === 'oui'
      );
    })
    .map(function(ligne) {
      return {
        code: String(ligne[0] || '').trim(),
        nom: String(ligne[1] || '').trim(),
        type: String(ligne[2] || '').trim()
      };
    })
    .sort(function(a, b) {
      return (
        a.type.localeCompare(b.type, 'fr') ||
        a.code.localeCompare(b.code, 'fr')
      );
    });
}

function chargerProgrammesBudgetAnnuel_(ss) {
  const configuration = ss.getSheetByName(
    'Configuration'
  );

  if (!configuration) {
    throw new Error(
      "L’onglet Configuration est introuvable."
    );
  }

  const derniereLigne = configuration.getLastRow();

  if (derniereLigne < 6) {
    return [];
  }

  const valeurs = configuration
    .getRange(6, 9, derniereLigne - 5, 1)
    .getDisplayValues()
    .map(function(ligne) {
      return String(ligne[0] || '').trim();
    })
    .filter(function(valeur) {
      return valeur !== '';
    });

  return valeurs.filter(function(valeur, index) {
    return valeurs.indexOf(valeur) === index;
  });
}

function validerDonneesBudgetAnnuel_(ss, donnees) {
  const annee = Number(donnees.annee || 0);
  const programme = String(
    donnees.programme || ''
  ).trim();
  const codeCompte = String(
    donnees.codeCompte || ''
  ).trim();
  const montant = Number(donnees.montant);
  const notes = String(donnees.notes || '').trim();

  if (
    !Number.isInteger(annee) ||
    annee < 2000 ||
    annee > 2100
  ) {
    throw new Error(
      'L’année budgétaire est invalide.'
    );
  }

  const programmes = chargerProgrammesBudgetAnnuel_(ss);

  if (programmes.indexOf(programme) === -1) {
    throw new Error(
      'Le programme sélectionné est invalide ou inactif.'
    );
  }

  const comptes = chargerComptesBudgetAnnuel_(ss);
  const compte = comptes.find(function(item) {
    return item.code === codeCompte;
  });

  if (!compte) {
    throw new Error(
      'Le compte comptable sélectionné est invalide ou inactif.'
    );
  }

  if (!Number.isFinite(montant) || montant < 0) {
    throw new Error(
      'Le montant du budget doit être un nombre positif ou zéro.'
    );
  }

  return {
    annee: annee,
    programme: programme,
    codeCompte: codeCompte,
    montant: arrondirBudgetAnnuel_(montant),
    notes: notes
  };
}

function trouverLigneVideBudgetAnnuel_(
  feuille,
  valeursExistantes
) {
  for (
    let index = 0;
    index < valeursExistantes.length;
    index++
  ) {
    const ligne = valeursExistantes[index];

    if (
      String(ligne[0] || '').trim() === '' &&
      String(ligne[3] || '').trim() === ''
    ) {
      return (
        CONFIG_BUDGET_ANNUEL.premiereLigne + index
      );
    }
  }

  const ancienneTaille = feuille.getMaxRows();
  feuille.insertRowsAfter(ancienneTaille, 50);

  appliquerFormulesPlageBudgetAnnuel_(
    feuille,
    ancienneTaille + 1,
    50
  );

  appliquerValidationsBudgetAnnuel_(
    SpreadsheetApp.getActiveSpreadsheet(),
    feuille,
    ancienneTaille + 1,
    50
  );

  return ancienneTaille + 1;
}

function appliquerFormulesPlageBudgetAnnuel_(
  feuille,
  premiereLigne,
  nombreLignes
) {
  if (nombreLignes <= 0) {
    return;
  }

  const formulesNom = [];
  const formulesType = [];
  const formulesReel = [];
  const formulesEcart = [];
  const formulesPourcentage = [];

  for (
    let ligne = premiereLigne;
    ligne < premiereLigne + nombreLignes;
    ligne++
  ) {
    const formules =
      construireFormulesBudgetAnnuel_(ligne);
    formulesNom.push([formules.nom]);
    formulesType.push([formules.type]);
    formulesReel.push([formules.reel]);
    formulesEcart.push([formules.ecart]);
    formulesPourcentage.push([
      formules.pourcentage
    ]);
  }

  feuille
    .getRange(
      premiereLigne,
      5,
      nombreLignes,
      1
    )
    .setFormulas(formulesNom);

  feuille
    .getRange(
      premiereLigne,
      6,
      nombreLignes,
      1
    )
    .setFormulas(formulesType);

  feuille
    .getRange(
      premiereLigne,
      8,
      nombreLignes,
      1
    )
    .setFormulas(formulesReel);

  feuille
    .getRange(
      premiereLigne,
      9,
      nombreLignes,
      1
    )
    .setFormulas(formulesEcart);

  feuille
    .getRange(
      premiereLigne,
      10,
      nombreLignes,
      1
    )
    .setFormulas(formulesPourcentage);
}

function appliquerFormulesLigneBudgetAnnuel_(
  feuille,
  ligne
) {
  const formules =
    construireFormulesBudgetAnnuel_(ligne);

  feuille.getRange(ligne, 5).setFormula(formules.nom);
  feuille.getRange(ligne, 6).setFormula(formules.type);
  feuille.getRange(ligne, 8).setFormula(formules.reel);
  feuille.getRange(ligne, 9).setFormula(formules.ecart);
  feuille
    .getRange(ligne, 10)
    .setFormula(formules.pourcentage);
}

function construireFormulesBudgetAnnuel_(ligne) {
  const debutPeriode =
    'DATE($A' +
    ligne +
    ',IF($B' +
    ligne +
    '="' +
    CONFIG_BUDGET_ANNUEL.periodeAnnuelle +
    '",1,VALUE($B' +
    ligne +
    ')),1)';

  const finPeriode =
    'IF($B' +
    ligne +
    '="' +
    CONFIG_BUDGET_ANNUEL.periodeAnnuelle +
    '",DATE($A' +
    ligne +
    '+1,1,1),EDATE(DATE($A' +
    ligne +
    ',VALUE($B' +
    ligne +
    '),1),1))';

  const criteres =
    ',Journal!$C$6:$C,">="&' +
    debutPeriode +
    ',Journal!$C$6:$C,"<"&' +
    finPeriode +
    ',Journal!$D$6:$D,$D' +
    ligne +
    ',Journal!$I$6:$I,$C' +
    ligne;

  const revenus =
    'SUMIFS(Journal!$H$6:$H' +
    criteres +
    ')' +
    '-SUMIFS(Journal!$G$6:$G' +
    criteres +
    ')';

  const depenses =
    'SUMIFS(Journal!$G$6:$G' +
    criteres +
    ')' +
    '-SUMIFS(Journal!$H$6:$H' +
    criteres +
    ')';

  return {
    nom:
      '=IF($D' +
      ligne +
      '="","",IFERROR(VLOOKUP($D' +
      ligne +
      ',Configuration!$A$6:$C,2,FALSE),' +
      '"Compte introuvable"))',
    type:
      '=IF($D' +
      ligne +
      '="","",IFERROR(VLOOKUP($D' +
      ligne +
      ',Configuration!$A$6:$C,3,FALSE),""))',
    reel:
      '=IF(OR($A' +
      ligne +
      '="",$B' +
      ligne +
      '="",$C' +
      ligne +
      '="",$D' +
      ligne +
      '=""),"",IF($F' +
      ligne +
      '="Revenu",' +
      revenus +
      ',IF($F' +
      ligne +
      '="Dépense",' +
      depenses +
      ',"")))',
    ecart:
      '=IF(OR($G' +
      ligne +
      '="",$H' +
      ligne +
      '="",$F' +
      ligne +
      '=""),"",IF($F' +
      ligne +
      '="Revenu",$H' +
      ligne +
      '-$G' +
      ligne +
      ',$G' +
      ligne +
      '-$H' +
      ligne +
      '))',
    pourcentage:
      '=IF(OR($G' +
      ligne +
      '="",$G' +
      ligne +
      '=0),"",$H' +
      ligne +
      '/$G' +
      ligne +
      ')'
  };
}

function appliquerValidationsBudgetAnnuel_(
  ss,
  feuille,
  premiereLigne,
  nombreLignes
) {
  if (nombreLignes <= 0) {
    return;
  }

  const programmes =
    chargerProgrammesBudgetAnnuel_(ss);
  const comptes = chargerComptesBudgetAnnuel_(ss)
    .map(function(compte) {
      return compte.code;
    });

  const validationAnnee =
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(2000, 2100)
      .setAllowInvalid(false)
      .build();

  const validationPeriode =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [
          CONFIG_BUDGET_ANNUEL.periodeAnnuelle,
          '1',
          '2',
          '3',
          '4',
          '5',
          '6',
          '7',
          '8',
          '9',
          '10',
          '11',
          '12'
        ],
        true
      )
      .setAllowInvalid(false)
      .build();

  const validationMontant =
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .build();

  const validationStatut =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [
          CONFIG_BUDGET_ANNUEL.statutActif,
          CONFIG_BUDGET_ANNUEL.statutInactif
        ],
        true
      )
      .setAllowInvalid(false)
      .build();

  feuille
    .getRange(
      premiereLigne,
      1,
      nombreLignes,
      1
    )
    .setDataValidation(validationAnnee);

  feuille
    .getRange(
      premiereLigne,
      2,
      nombreLignes,
      1
    )
    .setDataValidation(validationPeriode);

  if (programmes.length) {
    const validationProgramme =
      SpreadsheetApp.newDataValidation()
        .requireValueInList(programmes, true)
        .setAllowInvalid(false)
        .build();

    feuille
      .getRange(
        premiereLigne,
        3,
        nombreLignes,
        1
      )
      .setDataValidation(validationProgramme);
  }

  if (comptes.length) {
    const validationCompte =
      SpreadsheetApp.newDataValidation()
        .requireValueInList(comptes, true)
        .setAllowInvalid(false)
        .build();

    feuille
      .getRange(
        premiereLigne,
        4,
        nombreLignes,
        1
      )
      .setDataValidation(validationCompte);
  }

  feuille
    .getRange(
      premiereLigne,
      7,
      nombreLignes,
      1
    )
    .setDataValidation(validationMontant);

  feuille
    .getRange(
      premiereLigne,
      11,
      nombreLignes,
      1
    )
    .setDataValidation(validationStatut);
}

function appliquerValidationsLigneBudgetAnnuel_(
  ss,
  feuille,
  numeroLigne
) {
  appliquerValidationsBudgetAnnuel_(
    ss,
    feuille,
    numeroLigne,
    1
  );
}

function mettreAJourPlagesBudgetRapports_(ss) {
  const rapports = ss.getSheetByName('Rapports');

  if (!rapports || rapports.getLastRow() < 1) {
    return;
  }

  const plage = rapports.getDataRange();
  const formules = plage.getFormulas();
  const misesAJour = [];

  for (
    let ligne = 0;
    ligne < formules.length;
    ligne++
  ) {
    for (
      let colonne = 0;
      colonne < formules[ligne].length;
      colonne++
    ) {
      const formule = formules[ligne][colonne];

      if (
        !formule ||
        formule.indexOf("'Budget 2026'!") === -1
      ) {
        continue;
      }

      const nouvelleFormule = formule
        .replace(
          /'Budget 2026'!\$G\$6:\$G\$204/g,
          "'Budget 2026'!$G$6:$G"
        )
        .replace(
          /'Budget 2026'!\$C\$6:\$C\$204/g,
          "'Budget 2026'!$C$6:$C"
        )
        .replace(
          /'Budget 2026'!\$F\$6:\$F\$204/g,
          "'Budget 2026'!$F$6:$F"
        )
        .replace(
          /'Budget 2026'!\$K\$6:\$K\$204/g,
          "'Budget 2026'!$K$6:$K"
        );

      if (nouvelleFormule !== formule) {
        misesAJour.push({
          ligne: ligne + 1,
          colonne: colonne + 1,
          formule: nouvelleFormule
        });
      }
    }
  }

  misesAJour.forEach(function(miseAJour) {
    rapports
      .getRange(
        miseAJour.ligne,
        miseAJour.colonne
      )
      .setFormula(miseAJour.formule);
  });
}

function installerDeclencheurMenuBudgetAnnuel_(ss) {
  const nomGestionnaire =
    'ajouterMenuBudgetAnnuelAuDemarrage';

  ScriptApp.getProjectTriggers()
    .filter(function(declencheur) {
      return (
        declencheur.getHandlerFunction() ===
        nomGestionnaire
      );
    })
    .forEach(function(declencheur) {
      ScriptApp.deleteTrigger(declencheur);
    });

  ScriptApp.newTrigger(nomGestionnaire)
    .forSpreadsheet(ss)
    .onOpen()
    .create();
}

function obtenirAnneeDefautBudgetAnnuel_(feuille) {
  const correspondance = feuille
    .getName()
    .match(/(\d{4})/);

  if (correspondance) {
    return Number(correspondance[1]);
  }

  return new Date().getFullYear();
}

function arrondirBudgetAnnuel_(valeur) {
  const nombre = Number(valeur || 0);
  return (
    Math.round(
      (nombre + Number.EPSILON) * 100
    ) / 100
  );
}
