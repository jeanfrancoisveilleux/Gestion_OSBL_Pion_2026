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
      'Vue automatique basée sur le Journal. Les annulations sont ' +
      'compensées par les écritures inverses.'
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
    .setValue('Dépenses nettes fournisseurs')
    .setFontWeight('bold');

  feuille.getRange('B4')
    .setFormula('=IFERROR(SUM(C7:C),0)')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('D4')
    .setValue('Fournisseur principal')
    .setFontWeight('bold');

  feuille.getRange('E4')
    .setFormula(
      '=IF(B4=0,"—",IFERROR(INDEX(B7:B,MATCH(MAX(C7:C),C7:C,0)),"—"))'
    )
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('G4')
    .setValue('Montant principal')
    .setFontWeight('bold');

  feuille.getRange('H4')
    .setFormula('=IFERROR(MAX(C7:C),0)')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00')
    .setBackground('#e8f0fe')
    .setFontWeight('bold');

  feuille.getRange('A5:C5')
    .merge()
    .setValue('Sommaire par fournisseur')
    .setBackground('#f1f3f4')
    .setFontWeight('bold');

  feuille.getRange('E5:J5')
    .merge()
    .setValue('Détail par fournisseur, programme et compte')
    .setBackground('#f1f3f4')
    .setFontWeight('bold');

  feuille.setFrozenRows(6);

  const largeurs = [
    125, 210, 135, 25, 125, 190, 170, 90, 210, 135
  ];

  largeurs.forEach(function(largeur, index) {
    feuille.setColumnWidth(index + 1, largeur);
  });

  feuille.getRange('A6:J6')
    .setBackground('#f1f3f4')
    .setFontWeight('bold')
    .setWrap(true);

  feuille.getRange('C7:C1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  feuille.getRange('J7:J1000')
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');
}

function ecrireFormulesRapportFournisseurs_(feuille) {
  feuille.getRange('A6').setFormula(
    formuleSommaireRapportFournisseurs_()
  );

  feuille.getRange('E6').setFormula(
    formuleDetailRapportFournisseurs_()
  );
}

function formuleSommaireRapportFournisseurs_() {
  return '=IFERROR(QUERY(QUERY({' +
    'Journal!$O$6:$O,' +
    'Journal!$P$6:$P,' +
    'Journal!$C$6:$C,' +
    'Journal!$D$6:$D,' +
    'Journal!$F$6:$F,' +
    'Journal!$G$6:$G-Journal!$H$6:$H' +
    '},' +
    '"select Col1,Col2,sum(Col6) ' +
    'where Col1 is not null ' +
    'and year(Col3) = "&$B$3&" ' +
    'and Col4 <> \'1000\' ' +
    'and (Col5 = \'Dépense\' or Col5 = \'Actif\') ' +
    'group by Col1,Col2 ' +
    'label Col1 \'ID fournisseur\',Col2 \'Fournisseur\',' +
    'sum(Col6) \'Dépenses nettes\'",' +
    '0),' +
    '"select Col1,Col2,Col3 ' +
    'where Col3 <> 0 ' +
    'order by Col3 desc",' +
    '1),' +
    '{"ID fournisseur","Fournisseur","Dépenses nettes"})';
}

function formuleDetailRapportFournisseurs_() {
  return '=IFERROR(QUERY(QUERY({' +
    'Journal!$O$6:$O,' +
    'Journal!$P$6:$P,' +
    'Journal!$C$6:$C,' +
    'Journal!$I$6:$I,' +
    'Journal!$D$6:$D,' +
    'Journal!$E$6:$E,' +
    'Journal!$F$6:$F,' +
    'Journal!$G$6:$G-Journal!$H$6:$H' +
    '},' +
    '"select Col1,Col2,Col4,Col5,Col6,sum(Col8) ' +
    'where Col1 is not null ' +
    'and year(Col3) = "&$B$3&" ' +
    'and Col5 <> \'1000\' ' +
    'and (Col7 = \'Dépense\' or Col7 = \'Actif\') ' +
    'group by Col1,Col2,Col4,Col5,Col6 ' +
    'label Col1 \'ID fournisseur\',Col2 \'Fournisseur\',' +
    'Col4 \'Programme\',Col5 \'Code compte\',' +
    'Col6 \'Compte\',sum(Col8) \'Dépenses nettes\'",' +
    '0),' +
    '"select Col1,Col2,Col3,Col4,Col5,Col6 ' +
    'where Col6 <> 0 ' +
    'order by Col2,Col6 desc",' +
    '1),' +
    '{"ID fournisseur","Fournisseur","Programme",' +
    '"Code compte","Compte","Dépenses nettes"})';
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
