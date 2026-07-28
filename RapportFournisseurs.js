// ─── Rapport automatisé des dépenses par fournisseur ────────────────────────

function installerRapportDepensesFournisseurs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Journal');

  if (!journal) {
    throw new Error("L'onglet Journal est introuvable.");
  }

  verifierEntetesJournalRapportFournisseurs_(journal);

  let feuille = ss.getSheetByName('Rapport fournisseurs');

  if (!feuille) {
    feuille = ss.insertSheet('Rapport fournisseurs');
  }

  preparerOngletRapportFournisseurs_(feuille);
  ecrireFormulesRapportFournisseurs_(feuille);
  protegerRapportFournisseurs_(feuille);

  feuille.activate();

  SpreadsheetApp.getUi().alert(
    'Le rapport des dépenses par fournisseur est prêt.\n\n' +
    'Il se met à jour automatiquement à partir du Journal.\n' +
    'Les annulations sont compensées par les écritures inverses.'
  );
}

function verifierEntetesJournalRapportFournisseurs_(journal) {
  const entetesAttendues = {
    C: 'Date',
    D: 'Code compte',
    E: 'Nom du compte',
    F: 'Type compte',
    G: 'Débit',
    H: 'Crédit',
    I: 'Programme',
    O: 'ID fournisseur',
    P: 'Fournisseur'
  };

  Object.keys(entetesAttendues).forEach(function(colonne) {
    const valeur = String(
      journal.getRange(colonne + '5').getDisplayValue() || ''
    ).trim();

    if (valeur !== entetesAttendues[colonne]) {
      throw new Error(
        'La colonne ' + colonne + ' du Journal doit porter le titre « ' +
        entetesAttendues[colonne] + ' ». Valeur trouvée : « ' +
        (valeur || 'vide') + ' ».'
      );
    }
  });
}

function preparerOngletRapportFournisseurs_(feuille) {
  feuille.setHiddenGridlines(true);

  const colonnesNecessaires = 10;

  if (feuille.getMaxColumns() < colonnesNecessaires) {
    feuille.insertColumnsAfter(
      feuille.getMaxColumns(),
      colonnesNecessaires - feuille.getMaxColumns()
    );
  }

  const lignesNecessaires = 1000;

  if (feuille.getMaxRows() < lignesNecessaires) {
    feuille.insertRowsAfter(
      feuille.getMaxRows(),
      lignesNecessaires - feuille.getMaxRows()
    );
  }

  const zone = feuille.getRange(
    1,
    1,
    feuille.getMaxRows(),
    colonnesNecessaires
  );

  const filtre = feuille.getFilter();

  if (filtre) {
    filtre.remove();
  }

  zone.breakApart();
  zone.clearContent();
  zone.clearFormat();
  zone.clearDataValidations();

  feuille.getRange('A1:J1')
    .merge()
    .setValue('Dépenses par fournisseur')
    .setBackground('#e8f0fe')
    .setFontWeight('bold')
    .setFontSize(15)
    .setHorizontalAlignment('left');

  feuille.getRange('A2:J2')
    .merge()
    .setValue(
      'Comparaison automatique entre les sorties bancaires et les ' +
      'dépenses comptabilisées du Journal.'
    )
    .setFontColor('#5f6368')
    .setWrap(true);

  feuille.getRange('A3')
    .setValue('Année')
    .setFontWeight('bold');

  feuille.getRange('B3')
    .setValue(2026)
    .setNumberFormat('0')
    .setBackground('#fff7df')
    .setHorizontalAlignment('center')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberBetween(2020, 2100)
        .setAllowInvalid(false)
        .build()
    );

  feuille.getRange('A4')
    .setValue('Total achats')
    .setFontWeight('bold');

  feuille.getRange('B4')
    .setFormula('=IFERROR(SUM(C7:C),0)')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('D4')
    .setValue('Dépenses comptabilisées')
    .setFontWeight('bold');

  feuille.getRange('E4')
    .setFormula('=IFERROR(SUM(D7:D),0)')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('G4')
    .setValue('Différence totale')
    .setFontWeight('bold');

  feuille.getRange('H4')
    .setFormula('=IFERROR(SUM(E7:E),0)')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('A5:E5')
    .merge()
    .setValue('Sommaire par fournisseur')
    .setBackground('#f1f3f4')
    .setFontWeight('bold');

  feuille.setFrozenRows(6);

  const largeurs = [
    125, 210, 135, 175, 135, 25, 150, 135, 25, 25
  ];

  largeurs.forEach(function(largeur, index) {
    feuille.setColumnWidth(index + 1, largeur);
  });

  feuille.getRange('A6:E6')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  feuille.getRange('C7:E1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function ecrireFormulesRapportFournisseurs_(feuille) {
  feuille.getRange('A6').setFormula(
    formuleSommaireRapportFournisseurs_()
  );
}

function formuleSommaireRapportFournisseurs_() {
  return '=IFERROR(QUERY(QUERY(FILTER({' +
    'Journal!$O$6:$O,' +
    'Journal!$P$6:$P,' +
    'ARRAYFORMULA((' +
    'TO_TEXT(Journal!$D$6:$D)="1000")*(' +
    'Journal!$H$6:$H-Journal!$G$6:$G)),' +
    'ARRAYFORMULA((' +
    'TO_TEXT(Journal!$D$6:$D)<>"1000")*((' +
    'Journal!$F$6:$F="Dépense")+(' +
    'Journal!$F$6:$F="Actif"))*(' +
    'Journal!$G$6:$G-Journal!$H$6:$H)),' +
    'ARRAYFORMULA(((' +
    'TO_TEXT(Journal!$D$6:$D)="1000")*(' +
    'Journal!$H$6:$H-Journal!$G$6:$G))-((' +
    'TO_TEXT(Journal!$D$6:$D)<>"1000")*((' +
    'Journal!$F$6:$F="Dépense")+(' +
    'Journal!$F$6:$F="Actif"))*(' +
    'Journal!$G$6:$G-Journal!$H$6:$H)))' +
    '},' +
    'Journal!$O$6:$O<>"",' +
    'ARRAYFORMULA(YEAR(Journal!$C$6:$C))=$B$3' +
    '),' +
    '"select Col1,Col2,sum(Col3),sum(Col4),sum(Col5) ' +
    'group by Col1,Col2 ' +
    'label Col1 \'ID fournisseur\',Col2 \'Fournisseur\',' +
    'sum(Col3) \'Total achats\',' +
    'sum(Col4) \'Dépenses comptabilisées\',' +
    'sum(Col5) \'Différence\'",' +
    '0),' +
    '"select Col1,Col2,Col3,Col4,Col5 ' +
    'where Col3 <> 0 or Col4 <> 0 ' +
    'order by Col2",' +
    '1),' +
    '{"ID fournisseur","Fournisseur","Total achats",' +
    '"Dépenses comptabilisées","Différence"})';
}

function protegerRapportFournisseurs_(feuille) {
  const description =
    'Gestion OSBL – données générées – Rapport fournisseurs';

  const protections = feuille
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(function(protection) {
      return protection.getDescription() === description;
    });

  let protection;

  if (protections.length > 0) {
    protection = protections.shift();

    protections.forEach(function(doublon) {
      doublon.remove();
    });
  } else {
    protection = feuille.protect();
  }

  protection
    .setDescription(description)
    .setWarningOnly(true);
}
