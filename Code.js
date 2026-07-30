const CONFIG_OSBL = {
  dossierPrincipal: 'Gestion OSBL - Pion joues-tu - 2026',
  sousDossiers: [
    '01 - Reçus et factures',
    '02 - Contrats',
    '03 - Relevés bancaires',
    '04 - Projets',
    '05 - Subventions et commandites',
    '06 - Rapports'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gestion OSBL')
    .addItem(
      'Saisir une transaction',
      'afficherSaisieTransaction'
    )
    .addItem(
      'Importer un relevé Desjardins',
      'afficherImportBancaire'
    )
    .addSeparator()
    .addItem(
      'Ouvrir le rapprochement bancaire',
      'ouvrirRapprochementBancaire'
    )
    .addItem(
      'Installer le rapprochement bancaire',
      'installerRapprochementBancaire'
    )
    .addSeparator()
    .addItem(
      'Installer Fournisseurs et Contacts',
      'installerGestionFournisseursContacts'
    )
    .addItem(
      'Installer les règles bancaires',
      'installerReglesBancairesConfigurables'
    )
    .addItem(
      'Valider les règles bancaires',
      'validerReglesBancairesConfigurables'
    )
    .addItem(
      'Réappliquer les règles aux transactions à classer',
      'reappliquerReglesBancairesAuxTransactionsAClasser'
    )
    .addItem(
      'Installer le rapport fournisseurs',
      'installerRapportDepensesFournisseurs'
    )
    .addSeparator()
    .addItem(
      'Initialiser les dossiers Drive',
      'initialiserDossiersDrive'
    )
    .addItem(
      'Vérifier le système',
      'verifierSysteme'
    )
    .addToUi();
}

function initialiserDossiersDrive() {
  const classeur = SpreadsheetApp.getActive();
  const configuration = classeur.getSheetByName('Configuration');

  if (!configuration) {
    throw new Error("L’onglet Configuration est introuvable.");
  }

  const ligneConfiguration = trouverLigneConfiguration_(
    configuration,
    'Dossier Drive principal'
  );

  let dossierPrincipal = obtenirDossierConfigure_(
    configuration.getRange(ligneConfiguration, 12).getValue()
  );

  if (!dossierPrincipal) {
    dossierPrincipal = DriveApp.createFolder(
      CONFIG_OSBL.dossierPrincipal
    );
  }

  CONFIG_OSBL.sousDossiers.forEach(function(nom) {
    trouverOuCreerSousDossier_(dossierPrincipal, nom);
  });

  configuration
    .getRange(ligneConfiguration, 12)
    .setValue(dossierPrincipal.getUrl());

  configuration
    .getRange(ligneConfiguration, 14)
    .setValue('Confirmé');

  classeur.toast(
    'Les dossiers Drive ont été créés et liés au système.',
    'Gestion OSBL',
    6
  );
}

function trouverLigneConfiguration_(feuille, parametre) {
  const valeurs = feuille
    .getRange(1, 11, feuille.getLastRow(), 1)
    .getValues();

  for (let index = 0; index < valeurs.length; index++) {
    if (valeurs[index][0] === parametre) {
      return index + 1;
    }
  }

  throw new Error(
    'Le paramètre « ' + parametre + ' » est introuvable.'
  );
}

function obtenirDossierConfigure_(url) {
  if (!url) {
    return null;
  }

  const resultat = String(url).match(/\/folders\/([^/?]+)/);

  if (!resultat) {
    return null;
  }

  try {
    return DriveApp.getFolderById(resultat[1]);
  } catch (erreur) {
    return null;
  }
}

function trouverOuCreerSousDossier_(dossierParent, nom) {
  const dossiers = dossierParent.getFoldersByName(nom);

  if (dossiers.hasNext()) {
    return dossiers.next();
  }

  return dossierParent.createFolder(nom);
}

function verifierSysteme() {
  const classeur = SpreadsheetApp.getActive();
  const ongletsRequis = [
    'Accueil',
    'Transactions',
    'Budget 2026',
    'Projets',
    'Inventaire',
    'Import bancaire',
    'Documents',
    'Soldes ouverture',
    'Rapports',
    'Journal',
    'Configuration'
  ];

  const manquants = ongletsRequis.filter(function(nom) {
    return !classeur.getSheetByName(nom);
  });

  if (manquants.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Onglets manquants : ' + manquants.join(', ')
    );
    return;
  }

  const soldes = classeur.getSheetByName('Soldes ouverture');
  const debit = Number(soldes.getRange('C14').getValue()) || 0;
  const credit = Number(soldes.getRange('D14').getValue()) || 0;
  const ecart = Math.round((debit - credit) * 100) / 100;

  SpreadsheetApp.getUi().alert(
    'Vérification terminée\n\n' +
    'Les 11 onglets sont présents.\n' +
    'Écart des soldes d’ouverture : ' +
    ecart.toFixed(2) +
    ' $'
  );
}

function installerProtectionsDonneesGenerees() {
  const classeur = SpreadsheetApp.getActive();
  const ongletsCibles = [
    'Transactions',
    'Journal',
    'Répartition',
    'Forfaits',
    'Rapport fournisseurs'
  ];
  const prefixeDescription = 'Gestion OSBL – données générées – ';

  let protectionsCreees = 0;
  let protectionsActualisees = 0;
  let doublonsSupprimes = 0;
  const ongletsManquants = [];

  ongletsCibles.forEach(function(nomOnglet) {
    const feuille = classeur.getSheetByName(nomOnglet);

    if (!feuille) {
      ongletsManquants.push(nomOnglet);
      return;
    }

    const description = prefixeDescription + nomOnglet;
    const protections = feuille
      .getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .filter(function(protection) {
        return String(protection.getDescription() || '') === description;
      });

    let protectionReference = protections[0] || null;

    if (!protectionReference) {
      protectionReference = feuille.protect();
      protectionsCreees += 1;
    } else {
      protectionsActualisees += 1;
    }

    protectionReference.setDescription(description);
    protectionReference.setWarningOnly(true);

    protections.slice(1).forEach(function(protectionDoublon) {
      protectionDoublon.remove();
      doublonsSupprimes += 1;
    });
  });

  const message =
    'Installation des protections terminée.\n\n' +
    'Protections créées : ' + protectionsCreees + '\n' +
    'Protections actualisées : ' + protectionsActualisees + '\n' +
    'Doublons supprimés : ' + doublonsSupprimes + '\n' +
    'Onglets manquants : ' +
    (ongletsManquants.length
      ? ongletsManquants.join(', ')
      : 'Aucun');

  SpreadsheetApp.getUi().alert(message);
}

function afficherSaisieTransaction() {
  const interface = HtmlService
    .createHtmlOutputFromFile('Transaction')
    .setTitle('Nouvelle transaction');

  SpreadsheetApp.getUi().showSidebar(interface);
}

function obtenirOptionsTransaction() {
  const classeur = SpreadsheetApp.getActive();
  const configuration = classeur.getSheetByName('Configuration');
  const projets = classeur.getSheetByName('Projets');

  const comptes = configuration
    .getRange(6, 1, configuration.getLastRow() - 5, 5)
    .getValues()
    .filter(function(ligne) {
      return ligne[0] && ligne[4] === 'Oui';
    })
    .map(function(ligne) {
      return {
        code: String(ligne[0]),
        nom: ligne[1],
        type: ligne[2]
      };
    });

  const programmes = configuration
    .getRange('I6:I14')
    .getValues()
    .flat()
    .filter(String);

  const listeProjets = projets
    .getRange(6, 1, Math.max(projets.getLastRow() - 5, 1), 4)
    .getValues()
    .filter(function(ligne) {
      return ligne[0] && ligne[3] !== 'Exemple';
    })
    .map(function(ligne) {
      return {
        id: ligne[0],
        nom: ligne[1],
        statut: ligne[3]
      };
    });

  return {
    date: Utilities.formatDate(
      new Date(),
      classeur.getSpreadsheetTimeZone(),
      'yyyy-MM-dd'
    ),
    comptes: comptes,
    programmes: programmes,
    projets: listeProjets
  };
}

function enregistrerTransaction(donnees) {
  const verrou = LockService.getDocumentLock();
  verrou.waitLock(30000);

  try {
    const classeur = SpreadsheetApp.getActive();
    const feuille = classeur.getSheetByName('Transactions');
    const journal = classeur.getSheetByName('Journal');
    const configuration = classeur.getSheetByName('Configuration');

    const type = String(donnees.type || '').trim();
    const description = String(donnees.description || '').trim();
    const codeCompte = String(donnees.compte || '').trim();
    const comptePaiement = String(
      donnees.comptePaiement || '1000'
    ).trim();

    const montant = Number(
      String(donnees.montant || '')
        .replace(/\s/g, '')
        .replace(',', '.')
        .replace('$', '')
    );

    if (!['Revenu', 'Dépense', 'Transfert'].includes(type)) {
      throw new Error('Le type de transaction est invalide.');
    }

    if (!donnees.date || !description || !codeCompte) {
      throw new Error(
        'La date, la description et le compte sont obligatoires.'
      );
    }

    if (!montant || montant <= 0) {
      throw new Error('Le montant doit être supérieur à zéro.');
    }

    if (codeCompte === comptePaiement) {
      throw new Error(
        'Le compte comptable et le compte de paiement doivent être différents.'
      );
    }

    verifierCompteExiste_(configuration, codeCompte);
    verifierCompteExiste_(configuration, comptePaiement);

    const date = new Date(donnees.date + 'T12:00:00');
    const identifiant =
      'T-' +
      Utilities.formatDate(
        new Date(),
        classeur.getSpreadsheetTimeZone(),
        'yyyyMMdd-HHmmss'
      ) +
      '-' +
      Utilities.getUuid().slice(0, 4).toUpperCase();

    const ligne = prochaineLigneLibre_(feuille, 1, 6);

    feuille.getRange(ligne, 1, 1, 7).setValues([[
      identifiant,
      date,
      type,
      donnees.contact || '',
      description,
      montant,
      codeCompte
    ]]);

    feuille.getRange(ligne, 8).setFormula(
      '=IF($G' + ligne + '="","",IFERROR(VLOOKUP($G' +
      ligne +
      ',Configuration!$A$6:$B$47,2,FALSE),"Compte inconnu"))'
    );

    feuille.getRange(ligne, 9, 1, 7).setValues([[
      donnees.programme || '',
      donnees.projet || '',
      comptePaiement,
      donnees.document || '',
      'Manuelle',
      '',
      'Validée'
    ]]);

    feuille.getRange(ligne, 16).setFormula(
      '=IF($B' + ligne + '="","",TEXT($B' +
      ligne +
      ',"yyyy-mm"))'
    );

    feuille.getRange(ligne, 17).setFormula(
      '=IF($A' + ligne +
      '="","",IF(AND($C' + ligne +
      '="Dépense",$L' + ligne +
      '=""),"Pièce requise","OK"))'
    );

    const ligneJournal = prochaineLigneLibre_(journal, 1, 6);

    if (type === 'Revenu') {
      ecrireLigneJournal_(
        journal, ligneJournal, identifiant, date,
        comptePaiement, montant, 0, donnees
      );

      ecrireLigneJournal_(
        journal, ligneJournal + 1, identifiant, date,
        codeCompte, 0, montant, donnees
      );
    } else {
      ecrireLigneJournal_(
        journal, ligneJournal, identifiant, date,
        codeCompte, montant, 0, donnees
      );

      ecrireLigneJournal_(
        journal, ligneJournal + 1, identifiant, date,
        comptePaiement, 0, montant, donnees
      );
    }

    SpreadsheetApp.flush();

    return {
      succes: true,
      identifiant: identifiant
    };
  } finally {
    verrou.releaseLock();
  }
}

function ecrireLigneJournal_(
  feuille,
  ligne,
  identifiant,
  date,
  codeCompte,
  debit,
  credit,
  donnees
) {
  const idEcriture =
    'J-' + identifiant + '-' + (debit > 0 ? 'D' : 'C');

  feuille.getRange(ligne, 1, 1, 4).setValues([[
    idEcriture,
    identifiant,
    date,
    codeCompte
  ]]);

  feuille.getRange(ligne, 5).setFormula(
    '=IF($D' + ligne + '="","",IFERROR(VLOOKUP($D' +
    ligne +
    ',Configuration!$A$6:$C$47,2,FALSE),"Compte inconnu"))'
  );

  feuille.getRange(ligne, 6).setFormula(
    '=IF($D' + ligne + '="","",IFERROR(VLOOKUP($D' +
    ligne +
    ',Configuration!$A$6:$C$47,3,FALSE),""))'
  );

  feuille.getRange(ligne, 7, 1, 7).setValues([[
    debit,
    credit,
    donnees.programme || '',
    donnees.projet || '',
    donnees.description || '',
    donnees.document || '',
    'Oui'
  ]]);

  feuille.getRange(ligne, 14).setFormula(
    '=IF($C' + ligne + '="","",TEXT($C' +
    ligne +
    ',"yyyy-mm"))'
  );
}

function verifierCompteExiste_(configuration, code) {
  const codes = configuration
    .getRange(6, 1, configuration.getLastRow() - 5, 1)
    .getDisplayValues()
    .flat();

  if (!codes.includes(String(code))) {
    throw new Error('Le compte ' + code + ' est introuvable.');
  }
}

function prochaineLigneLibre_(feuille, colonne, premiereLigne) {
  const nombre = feuille.getMaxRows() - premiereLigne + 1;
  const valeurs = feuille
    .getRange(premiereLigne, colonne, nombre, 1)
    .getValues();

  for (let index = 0; index < valeurs.length; index++) {
    if (!valeurs[index][0]) {
      return premiereLigne + index;
    }
  }

  feuille.insertRowsAfter(feuille.getMaxRows(), 50);
  return feuille.getMaxRows() - 49;
}

function installerEvenements2026() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Projets');

  if (!feuille) {
    throw new Error("L'onglet Projets est introuvable.");
  }

  const evenements = [
    ['PJC-2026-01-14', 'Pion joues-tu? au Cartier – 14 janvier 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 0, 14), new Date(2026, 0, 14), 0],
    ['PJC-2026-01-28', 'Pion joues-tu? au Cartier – 28 janvier 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 0, 28), new Date(2026, 0, 28), 0],
    ['PJC-2026-02-11', 'Pion joues-tu? au Cartier – 11 février 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 1, 11), new Date(2026, 1, 11), 0],
    ['PJT-2026-02-21', 'Pion joues-tu? – 21 février 2026', 'Pion joues-tu?', 'Terminé', new Date(2026, 1, 21), new Date(2026, 1, 21), 0],
    ['PJC-2026-02-25', 'Pion joues-tu? au Cartier – 25 février 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 1, 25), new Date(2026, 1, 25), 0],
    ['PJC-2026-03-11', 'Pion joues-tu? au Cartier – 11 mars 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 2, 11), new Date(2026, 2, 11), 0],
    ['PJC-2026-03-25', 'Pion joues-tu? au Cartier – 25 mars 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 2, 25), new Date(2026, 2, 25), 0],
    ['PJT-2026-03-28', 'Pion joues-tu? – 28 mars 2026', 'Pion joues-tu?', 'Terminé', new Date(2026, 2, 28), new Date(2026, 2, 28), 0],
    ['PJC-2026-04-08', 'Pion joues-tu? au Cartier – 8 avril 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 3, 8), new Date(2026, 3, 8), 0],
    ['PJC-2026-04-22', 'Pion joues-tu? au Cartier – 22 avril 2026', 'Pion joues-tu? – Cartier', 'Terminé', new Date(2026, 3, 22), new Date(2026, 3, 22), 0],
    ['PJC-2026-09-23', 'Pion joues-tu? au Cartier – 23 septembre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 8, 23), new Date(2026, 8, 23), 0],
    ['PJC-2026-10-07', 'Pion joues-tu? au Cartier – 7 octobre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 9, 7), new Date(2026, 9, 7), 0],
    ['PJT-2026-10-17', 'Pion joues-tu? – 17 octobre 2026', 'Pion joues-tu?', 'Planifié', new Date(2026, 9, 17), new Date(2026, 9, 17), 0],
    ['PJC-2026-10-21', 'Pion joues-tu? au Cartier – 21 octobre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 9, 21), new Date(2026, 9, 21), 0],
    ['PJC-2026-11-04', 'Pion joues-tu? au Cartier – 4 novembre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 10, 4), new Date(2026, 10, 4), 0],
    ['PJC-2026-11-18', 'Pion joues-tu? au Cartier – 18 novembre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 10, 18), new Date(2026, 10, 18), 0],
    ['PJC-2026-12-02', 'Pion joues-tu? au Cartier – 2 décembre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 11, 2), new Date(2026, 11, 2), 0],
    ['PJT-2026-12-12', 'Pion joues-tu? – 12 décembre 2026', 'Pion joues-tu?', 'Planifié', new Date(2026, 11, 12), new Date(2026, 11, 12), 0],
    ['PJC-2026-12-16', 'Pion joues-tu? au Cartier – 16 décembre 2026', 'Pion joues-tu? – Cartier', 'Planifié', new Date(2026, 11, 16), new Date(2026, 11, 16), 0]
  ];

  const idsPermis = new Set(evenements.map(e => e[0]));
  const idsExistants = feuille.getRange(6, 1, evenements.length, 1)
    .getDisplayValues()
    .flat();

  const donneesImprevues = idsExistants.filter(
    id => id && id !== 'EXEMPLE-P001' && !idsPermis.has(id)
  );

  if (donneesImprevues.length) {
    throw new Error(
      'Des projets existants seraient remplacés : ' +
      donneesImprevues.join(', ')
    );
  }

  feuille.getRange(6, 1, evenements.length, 7).setValues(evenements);
  feuille.getRange(6, 5, evenements.length, 2).setNumberFormat('yyyy-mm-dd');
}

function installerModuleForfaits() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();

  if (classeur.getSheetByName('Forfaits')) {
    SpreadsheetApp.getUi().alert("L'onglet Forfaits existe déjà.");
    return;
  }

  const feuille = classeur.insertSheet('Forfaits', 2);
  feuille.setHiddenGridlines(true);

  feuille.getRange('A1:J1').merge()
    .setValue('Ventes de forfaits')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('A2:J2').merge()
    .setValue(
      'Une ligne par forfait vendu. Les parts attribuées à Pion joues-tu? et au Cartier sont calculées automatiquement.'
    )
    .setFontStyle('italic')
    .setWrap(true);

  const entetes = [[
    'ID forfait',
    'Date de vente',
    'Acheteur',
    'Type de forfait',
    'Montant reçu',
    'ID transaction bancaire',
    'Part Pion joues-tu?',
    'Part Cartier',
    'Statut',
    'Notes'
  ]];

  feuille.getRange(5, 1, 1, 10)
    .setValues(entetes)
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  const types = [
    'Forfait Pion joues-tu? – 30 $',
    'Forfait Cartier – 20 $',
    'Forfait combiné – 50 $'
  ];

  const validationType = SpreadsheetApp.newDataValidation()
    .requireValueInList(types, true)
    .setAllowInvalid(false)
    .build();

  feuille.getRange('D6:D1000').setDataValidation(validationType);

  const validationStatut = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Actif', 'Annulé', 'Remboursé'], true)
    .setAllowInvalid(false)
    .build();

  feuille.getRange('I6:I1000').setDataValidation(validationStatut);

  const nombreLignes = 995;
  const formulesPion = [];
  const formulesCartier = [];

  for (let i = 0; i < nombreLignes; i++) {
    formulesPion.push([`
      =IF(RC[-6]="","",
        IF(OR(RC[2]="Annulé",RC[2]="Remboursé"),0,
          IF(RC[-3]="Forfait Pion joues-tu? – 30 $",RC[-2],
            IF(RC[-3]="Forfait combiné – 50 $",RC[-2]*0.6,0)
          )
        )
      )
    `.replace(/\s+/g, '')]);

    formulesCartier.push([`
      =IF(RC[-7]="","",
        IF(OR(RC[1]="Annulé",RC[1]="Remboursé"),0,
          IF(RC[-4]="Forfait Cartier – 20 $",RC[-3],
            IF(RC[-4]="Forfait combiné – 50 $",RC[-3]*0.4,0)
          )
        )
      )
    `.replace(/\s+/g, '')]);
  }

  feuille.getRange(6, 7, nombreLignes, 1)
    .setFormulasR1C1(formulesPion);

  feuille.getRange(6, 8, nombreLignes, 1)
    .setFormulasR1C1(formulesCartier);

  feuille.getRange('B6:B1000').setNumberFormat('yyyy-mm-dd');
  feuille.getRange('E6:H1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  feuille.getRange('A6:F1000').setBackground('#fff7df');
  feuille.getRange('G6:H1000').setBackground('#f1f3f4');
  feuille.getRange('I6:J1000').setBackground('#fff7df');

  feuille.setFrozenRows(5);

  const largeurs = [120, 105, 180, 235, 110, 165, 130, 115, 105, 220];
  largeurs.forEach((largeur, index) =>
    feuille.setColumnWidth(index + 1, largeur)
  );

  feuille.getRange(5, 1, 996, 10).createFilter();

  SpreadsheetApp.getUi().alert(
    "L'onglet Forfaits a été créé avec succès."
  );
}

function corrigerFormulesForfaits() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Forfaits');

  if (!feuille) {
    throw new Error("L'onglet Forfaits est introuvable.");
  }

  const formulePion =
    '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IF(RC4="Forfait Pion joues-tu? – 30 $",RC5,IF(RC4="Forfait combiné – 50 $",RC5*0.6,0))))';

  const formuleCartier =
    '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IF(RC4="Forfait Cartier – 20 $",RC5,IF(RC4="Forfait combiné – 50 $",RC5*0.4,0))))';

  feuille.getRange('G6:G1000').setFormulaR1C1(formulePion);
  feuille.getRange('H6:H1000').setFormulaR1C1(formuleCartier);

  SpreadsheetApp.getUi().alert('Les formules des forfaits sont corrigées.');
}

function configurerTarifsForfaitsParSaison() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const forfaits = classeur.getSheetByName('Forfaits');
  const configuration = classeur.getSheetByName('Configuration');

  if (!forfaits || !configuration) {
    throw new Error("Les onglets Forfaits ou Configuration sont introuvables.");
  }

  // Grille de tarifs conservée dans Configuration
  configuration.getRange('P4:S4').breakApart().merge()
    .setValue('Tarification des forfaits')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  configuration.getRange('P5:S7').setValues([
    ['Saison', 'Forfait Pion', 'Forfait Cartier', 'Forfait combiné'],
    [2026, 30, 20, 45],
    [2027, 30, 20, 50]
  ]);

  configuration.getRange('P5:S5')
    .setBackground('#f1f3f4')
    .setFontWeight('bold');

  configuration.getRange('Q6:S7')
    .setNumberFormat('$#,##0.00');

  // Les noms des forfaits ne contiennent plus leur prix
  const typesForfaits = [
    'Forfait Pion joues-tu?',
    'Forfait Cartier',
    'Forfait combiné'
  ];

  const validationType = SpreadsheetApp.newDataValidation()
    .requireValueInList(typesForfaits, true)
    .setAllowInvalid(false)
    .build();

  forfaits.getRange('D6:D1000').setDataValidation(validationType);

  // Conversion d'anciens noms, s'il y en a
  const plageTypes = forfaits.getRange('D6:D1000');
  const valeursTypes = plageTypes.getValues().map(ligne => {
    const valeur = ligne[0];

    if (valeur && valeur.indexOf('Forfait Pion joues-tu?') === 0) {
      return ['Forfait Pion joues-tu?'];
    }

    if (valeur && valeur.indexOf('Forfait Cartier') === 0) {
      return ['Forfait Cartier'];
    }

    if (valeur && valeur.indexOf('Forfait combiné') === 0) {
      return ['Forfait combiné'];
    }

    return [valeur];
  });

  plageTypes.setValues(valeursTypes);

  // Ajout de la saison en colonne K
  forfaits.getRange('K5').setValue('Saison')
    .setBackground('#f1f3f4')
    .setFontWeight('bold');

  const validationSaison = SpreadsheetApp.newDataValidation()
    .requireValueInRange(configuration.getRange('P6:P7'), true)
    .setAllowInvalid(false)
    .build();

  forfaits.getRange('K6:K1000')
    .setDataValidation(validationSaison)
    .setBackground('#fff7df')
    .setNumberFormat('0');

  // Répartition selon le montant réellement reçu
  const formulePion =
    '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IF(RC4="Forfait Pion joues-tu?",RC5,IF(RC4="Forfait combiné",RC5*0.6,0))))';

  const formuleCartier =
    '=IF(RC1="","",IF(OR(RC9="Annulé",RC9="Remboursé"),0,IF(RC4="Forfait Cartier",RC5,IF(RC4="Forfait combiné",RC5*0.4,0))))';

  forfaits.getRange('G6:G1000').setFormulaR1C1(formulePion);
  forfaits.getRange('H6:H1000').setFormulaR1C1(formuleCartier);

  // Agrandissement des titres et du filtre jusqu'à la colonne Saison
  forfaits.getRange('A1:K2').breakApart();
  forfaits.getRange('A1:K1').merge();
  forfaits.getRange('A2:K2').merge();

  const filtre = forfaits.getFilter();
  if (filtre) filtre.remove();

  forfaits.getRange(5, 1, 996, 11).createFilter();
  forfaits.setColumnWidth(11, 90);

  SpreadsheetApp.getUi().alert(
    'Tarifs 2026 et 2027 configurés avec succès.'
  );
}

function installerRapportEntreesForfaits() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Rapports');

  if (!feuille) {
    throw new Error("L'onglet Rapports est introuvable.");
  }

  feuille.getRange('G8:K8').breakApart().merge()
    .setValue("Revenus d'entrées et de laissez-passer – 2026")
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('G9:K12').clearContent();

  feuille.getRange('G9:K9').setValues([[
    'Type de revenu',
    'Pion joues-tu?',
    'Cartier',
    'Non attribué',
    'Total'
  ]]);

  feuille.getRange('G10:G12').setValues([
    ['Entrées'],
    ['Laissez-passer'],
    ['Total']
  ]);

  // Revenus d'entrées — compte 4000
  feuille.getRange('H10').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4000",Journal!$J$6:$J$1005,"PJT-2026")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4000",Journal!$J$6:$J$1005,"PJT-2026")'
  );

  feuille.getRange('I10').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4000",Journal!$J$6:$J$1005,"PJC-2026")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4000",Journal!$J$6:$J$1005,"PJC-2026")'
  );

  feuille.getRange('J10').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4000")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4000")' +
    '-H10-I10'
  );

  feuille.getRange('K10').setFormula('=SUM(H10:J10)');

  // Revenus de laissez-passer — compte 4010
  feuille.getRange('H11').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4010",Journal!$J$6:$J$1005,"PJT-2026")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4010",Journal!$J$6:$J$1005,"PJT-2026")'
  );

  feuille.getRange('I11').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4010",Journal!$J$6:$J$1005,"PJC-2026")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4010",Journal!$J$6:$J$1005,"PJC-2026")'
  );

  feuille.getRange('J11').setFormula(
    '=SUMIFS(Journal!$H$6:$H$1005,Journal!$D$6:$D$1005,"4010")' +
    '-SUMIFS(Journal!$G$6:$G$1005,Journal!$D$6:$D$1005,"4010")' +
    '-H11-I11'
  );

  feuille.getRange('K11').setFormula('=SUM(H11:J11)');

  // Totaux
  feuille.getRange('H12:K12').setFormulas([[
    '=SUM(H10:H11)',
    '=SUM(I10:I11)',
    '=SUM(J10:J11)',
    '=SUM(K10:K11)'
  ]]);

  feuille.getRange('G9:K9')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  feuille.getRange('G12:K12')
    .setFontWeight('bold')
    .setBackground('#e8f0fe');

  feuille.getRange('J10:J11').setBackground('#fff7df');
  feuille.getRange('H10:K12')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  feuille.getRange('G9:K12')
    .setBorder(true, true, true, true, true, true);

  feuille.setColumnWidth(7, 145);
  feuille.setColumnWidths(8, 4, 115);

  SpreadsheetApp.getUi().alert(
    'Le rapport Entrées / Laissez-passer a été créé.'
  );
}

function suggererClassementRevenusInterac() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Import bancaire');

  if (!feuille) {
    throw new Error("L'onglet Import bancaire est introuvable.");
  }

  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < 6) {
    SpreadsheetApp.getUi().alert('Aucune transaction bancaire trouvée.');
    return;
  }

  const plage = feuille.getRange(6, 1, derniereLigne - 5, 13);
  const donnees = plage.getValues();
  let nombreSuggestions = 0;

  donnees.forEach(ligne => {
    const description = String(ligne[2] || '');
    const montant = Number(ligne[3] || 0);
    const statut = String(ligne[10] || '');

    const estInteracRecu =
      montant > 0 &&
      /^Virement Interac de/i.test(description) &&
      statut === 'À classer';

    if (!estInteracRecu) return;

    if (!ligne[7]) {
      ligne[7] = "Revenus d'entrées aux événements";
    }

    if (!ligne[8]) {
      ligne[8] = 'Pion joues-tu?';
    }

    nombreSuggestions++;
  });

  plage.setValues(donnees);

  SpreadsheetApp.getUi().alert(
    nombreSuggestions +
    ' virements reçus ont été préremplis. Ils ne sont pas encore comptabilisés.'
  );
}

function installerAutomatisationForfaitsDepuisBanque() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const importBanque = classeur.getSheetByName('Import bancaire');

  if (!importBanque) {
    throw new Error("L'onglet Import bancaire est introuvable.");
  }

  // Supprimer le déclencheur onEdit du compte courant (le trigger « Other user »
  // doit être supprimé manuellement depuis le panneau Apps Script).
  ScriptApp.getProjectTriggers()
    .filter(function(declencheur) {
      return declencheur.getHandlerFunction() ===
        'creerOuMettreAJourForfaitDepuisBanque';
    })
    .forEach(function(declencheur) {
      ScriptApp.deleteTrigger(declencheur);
    });

  const nombreLignes = importBanque.getMaxRows() - 5;
  const plageN = importBanque.getRange(6, 14, nombreLignes, 1);

  // Retirer la validation de données, la note et la couleur de la colonne N.
  plageN.clearDataValidations();
  plageN.clearNote();
  plageN.setBackground(null);

  // Effacer uniquement les cellules contenant une ancienne commande de forfait.
  const anciensTypes = [
    'Forfait Pion joues-tu?',
    'Forfait Cartier',
    'Forfait combiné'
  ];

  const valeurs = plageN.getValues();

  valeurs.forEach(function(ligne, index) {
    const valeur = String(ligne[0] || '').trim();

    if (anciensTypes.indexOf(valeur) !== -1) {
      importBanque.getRange(6 + index, 14).clearContent();
    }
  });

  // Renommer l'en-tête et masquer la colonne N.
  importBanque.getRange('N5')
    .setValue('Ancien type de forfait (inactif)')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true)
    .clearNote();

  importBanque.hideColumns(14);

  SpreadsheetApp.getUi().alert(
    'La colonne « Type de forfait » a été désactivée.\n\n' +
    'Les forfaits sont maintenant sélectionnés directement dans ' +
    'l’interface « Classer » lors du classement bancaire.'
  );
}


// Compatibilité inactive. L'ancien déclencheur onEdit de cette fonction peut
// appartenir à un autre utilisateur (« Other user ») et doit être supprimé via
// installerInterfaceTransactionsMixtes(). Les forfaits sont désormais créés par
// creerForfaitDepuisRepartitionMixte_ lors du classement bancaire universel.
function creerOuMettreAJourForfaitDepuisBanque(e) {
  return;
}


function extraireNomInterac_(description) {
  const resultat = description.match(/de\s*\/([^/]+)\//i);

  if (resultat && resultat[1]) {
    return resultat[1].replace(/\s+/g, ' ').trim();
  }

  return description.replace(/\s+/g, ' ').trim();
}


function genererProchainIdForfait_(feuille, saison) {
  const valeurs = feuille
    .getRange(6, 1, feuille.getMaxRows() - 5, 1)
    .getDisplayValues()
    .flat();

  let numeroMaximum = 0;
  const expression = new RegExp(
    '^FOR-' + saison + '-(\\d{4})$'
  );

  valeurs.forEach(valeur => {
    const correspondance = String(valeur).match(expression);

    if (correspondance) {
      numeroMaximum = Math.max(
        numeroMaximum,
        Number(correspondance[1])
      );
    }
  });

  return 'FOR-' +
    saison +
    '-' +
    String(numeroMaximum + 1).padStart(4, '0');
}

function installerModuleRepartition() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const importBanque = classeur.getSheetByName('Import bancaire');

  if (!importBanque) {
    throw new Error("L'onglet Import bancaire est introuvable.");
  }

  let repartition = classeur.getSheetByName('Répartition');

  if (!repartition) {
    repartition = classeur.insertSheet('Répartition', 3);
  }

  repartition.setHiddenGridlines(true);

  repartition.getRange('A1:P2').breakApart();

  repartition.getRange('A1:P1').merge()
    .setValue('Répartition des transactions mixtes')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  repartition.getRange('A2:P2').merge()
    .setValue(
      'Une transaction bancaire peut contenir plusieurs composantes. Le total réparti doit correspondre exactement au montant bancaire.'
    )
    .setFontStyle('italic')
    .setWrap(true);

  repartition.getRange('A5:P5').setValues([[
    'ID import bancaire',
    'Date',
    'Description bancaire',
    'Total bancaire',
    'Ligne',
    'Composante',
    'Quantité',
    'Prix unitaire',
    'Montant attribué',
    'Compte',
    'Programme',
    'Projet',
    'Total réparti',
    'Reste à répartir',
    'Statut',
    'Notes'
  ]]);

  repartition.getRange('A5:P5')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  const composantes = [
    'Entrée – Pion joues-tu?',
    'Entrée – Cartier',
    'Forfait Pion joues-tu?',
    'Forfait Cartier',
    'Forfait combiné',
    'T-shirt',
    'Chandail à manches longues',
    'Hoodie'
  ];

  const validationComposante = SpreadsheetApp.newDataValidation()
    .requireValueInList(composantes, true)
    .setAllowInvalid(false)
    .build();

  const premiereLigne = 6;
  const derniereLigne = repartition.getMaxRows();
  const nombreLignes = derniereLigne - premiereLigne + 1;

  repartition
    .getRange(premiereLigne, 6, nombreLignes, 1)
    .setDataValidation(validationComposante);

  // Montant attribué = quantité × prix
  repartition
    .getRange(premiereLigne, 9, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(OR(RC1="",RC6=""),"",RC7*RC8)'
    );

  // Compte de revenu
  repartition
    .getRange(premiereLigne, 10, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC1="","",IF(OR(RC6="Entrée – Pion joues-tu?",RC6="Entrée – Cartier"),"Revenus d\'entrées aux événements",IF(OR(RC6="Forfait Pion joues-tu?",RC6="Forfait Cartier",RC6="Forfait combiné"),"Revenus de laissez-passer et forfaits",IF(OR(RC6="T-shirt",RC6="Chandail à manches longues",RC6="Hoodie"),"Ventes de marchandises et de jeux",""))))'
    );

  // Programme
  repartition
    .getRange(premiereLigne, 11, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC1="","",IF(OR(RC6="Entrée – Pion joues-tu?",RC6="Forfait Pion joues-tu?"),"Pion joues-tu?",IF(OR(RC6="Entrée – Cartier",RC6="Forfait Cartier"),"Pion joues-tu? – Cartier",IF(OR(RC6="T-shirt",RC6="Chandail à manches longues",RC6="Hoodie"),"Marchandise",""))))'
    );

  // Projet annuel
  repartition
    .getRange(premiereLigne, 12, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC11="Pion joues-tu?","PJT-2026",IF(RC11="Pion joues-tu? – Cartier","PJC-2026",""))'
    );

  // Somme de toutes les composantes de la même transaction
  repartition
    .getRange(premiereLigne, 13, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC1="","",SUMIF(R6C1:R1000C1,RC1,R6C9:R1000C9))'
    );

  // Montant restant
  repartition
    .getRange(premiereLigne, 14, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC1="","",RC4-RC13)'
    );

  // État de la répartition
  repartition
    .getRange(premiereLigne, 15, nombreLignes, 1)
    .setFormulaR1C1(
      '=IF(RC1="","",IF(ABS(RC14)<0.005,"Prêt",IF(RC14<0,"Dépassement","À compléter")))'
    );

  repartition.getRange('B6:B1000')
    .setNumberFormat('yyyy-mm-dd');

  repartition.getRange('D6:D1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  repartition.getRange('H6:I1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  repartition.getRange('M6:N1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  repartition.getRange('A6:E1000').setBackground('#f1f3f4');
  repartition.getRange('F6:H1000').setBackground('#fff7df');
  repartition.getRange('I6:O1000').setBackground('#f1f3f4');
  repartition.getRange('P6:P1000').setBackground('#fff7df');

  repartition.setFrozenRows(5);

  const largeurs = [
    190, 105, 240, 115, 60, 210, 80, 105,
    115, 210, 170, 120, 115, 115, 105, 180
  ];

  largeurs.forEach((largeur, index) => {
    repartition.setColumnWidth(index + 1, largeur);
  });

  if (!repartition.getFilter()) {
    repartition.getRange('A5:P1000').createFilter();
  }

  // Colonne Action : déclencheur du classement bancaire universel
  importBanque.getRange('O5')
    .setValue('Action')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  const validationMode = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Classer'], true)
    .setAllowInvalid(false)
    .build();

  const lignesImport = importBanque.getMaxRows() - 5;

  importBanque.getRange(6, 15, lignesImport, 1)
    .setDataValidation(validationMode)
    .setBackground('#fff7df');

  importBanque.setColumnWidth(15, 100);

  // Migrer et nettoyer la colonne Action :
  // - vider l'ancienne valeur 'Transaction mixte' sur toutes les lignes (migration);
  // - ne jamais la réécrire par script, car seul un clic utilisateur declenche l'interface;
  // - vider 'Classer' uniquement sur les lignes deja classees;
  // - laisser 'Classer' intact sur les lignes encore a classer.
  const derniereLigneImport = importBanque.getLastRow();
  if (derniereLigneImport >= 6) {
    const nombreLignesData = derniereLigneImport - 5;
    const colonnesKO = importBanque
      .getRange(6, 11, nombreLignesData, 5)
      .getValues();
    colonnesKO.forEach(function(ligne, index) {
      const statut = String(ligne[0] || '').trim();
      const action = String(ligne[4] || '').trim();
      if (
        action === 'Transaction mixte' ||
        (statut === 'Classée' && action === 'Classer')
      ) {
        importBanque.getRange(6 + index, 15).clearContent();
      }
    });
  }

  installerInterfaceTransactionsMixtes();

  SpreadsheetApp.getUi().alert(
    'Le module de classement bancaire est installé.\n\n' +
    'Choisissez « Classer » dans la colonne « Action » de l’onglet ' +
    '« Import bancaire » pour classer une transaction.'
  );
}


// Compatibilité inactive. L'ancien flux de préparation de cinq lignes dans
// Répartition a été remplacé par l'interface universelle « Classer » (col O).
// Le déclencheur onEdit associé a été supprimé par installerInterfaceTransactionsMixtes().
function preparerTransactionMixteDepuisBanque(e) {
  return;
}
