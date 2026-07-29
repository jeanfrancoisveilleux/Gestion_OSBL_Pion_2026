function installerInterfaceTransactionsMixtes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gestionnairesDeclencheurs = [
    'ajouterMenuTransactionsMixtesAuDemarrage',
    'preparerTransactionMixteDepuisBanque',
    'ouvrirInterfaceMixteDepuisImport'
  ];

  ScriptApp.getProjectTriggers()
    .filter(function(declencheur) {
      return gestionnairesDeclencheurs.indexOf(
        declencheur.getHandlerFunction()
      ) !== -1;
    })
    .forEach(function(declencheur) {
      ScriptApp.deleteTrigger(declencheur);
    });

  ScriptApp.newTrigger('ajouterMenuTransactionsMixtesAuDemarrage')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  ScriptApp.newTrigger('ouvrirInterfaceMixteDepuisImport')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ajouterMenuTransactionsMixtesAuDemarrage();

  SpreadsheetApp.getUi().alert(
    'L’interface de classement bancaire est installée.\n\n' +
    'Dès que vous choisissez « Classer » dans la colonne ' +
    '« Action » de l’onglet « Import bancaire », la fenêtre ' +
    'de classement s’ouvre automatiquement.'
  );
}

function ajouterMenuTransactionsMixtesAuDemarrage() {
  SpreadsheetApp.getUi()
    .createMenu('Classement bancaire')
    .addItem(
      'Classer la transaction sélectionnée',
      'ouvrirRepartitionTransactionSelectionnee'
    )
    .addItem(
      'Modifier la transaction sélectionnée',
      'modifierTransactionSelectionnee'
    )
    .addSeparator()
    .addItem(
      'Annuler la transaction sélectionnée',
      'annulerTransactionSelectionnee'
    )
    .addToUi();
}

function annulerTransactionSelectionnee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getActiveSheet();
  const cellule = feuille.getActiveCell();
  const ui = SpreadsheetApp.getUi();

  if (feuille.getName() !== 'Transactions' || cellule.getRow() < 6) {
    ui.alert(
      'Sélectionnez d’abord une cellule sur la transaction à annuler ' +
      'dans l’onglet « Transactions ».'
    );
    return;
  }

  const valeurs = feuille
    .getRange(cellule.getRow(), 1, 1, 17)
    .getValues()[0];
  const idTransaction = String(valeurs[0] || '').trim();
  const statut = String(valeurs[14] || '').trim();

  if (!idTransaction) {
    ui.alert('La ligne sélectionnée ne contient aucune transaction.');
    return;
  }

  if (statut === 'Annulée') {
    ui.alert('Cette transaction est déjà annulée.');
    return;
  }

  if (statut === 'Exemple') {
    ui.alert('La ligne d’exemple ne peut pas être annulée.');
    return;
  }

  const referenceBancaire = String(valeurs[13] || '').trim();
  const transactionsCibles = obtenirTransactionsCiblesAnnulation_(
    feuille,
    idTransaction,
    referenceBancaire
  );
  const total = arrondirMontantMixte_(
    transactionsCibles.reduce(function(somme, transaction) {
      return somme + Number(transaction.valeurs[5] || 0);
    }, 0)
  );
  const confirmation = ui.alert(
    'Confirmer l’annulation',
    'Cette opération annulera ' +
    transactionsCibles.length +
    ' ligne(s) de transaction pour un total de ' +
    total.toFixed(2) +
    ' $.\n\n' +
    'Le journal sera inversé, les données liées seront synchronisées ' +
    'et la transaction bancaire redeviendra « À classer ».\n\n' +
    'Voulez-vous continuer?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  const resultat = annulerGroupeTransactionsOSBL_(
    idTransaction,
    referenceBancaire
  );

  ui.alert(
    'Annulation terminée',
    resultat.transactions +
    ' transaction(s) annulée(s), ' +
    resultat.ecritures +
    ' écriture(s) d’annulation créée(s), ' +
    resultat.forfaits +
    ' forfait(s) annulé(s) et ' +
    resultat.mouvementsInventaire +
    ' mouvement(s) de stock restauré(s).',
    ui.ButtonSet.OK
  );
}

function modifierTransactionSelectionnee() {
  const ui = SpreadsheetApp.getUi();

  try {
    const selection = obtenirSelectionTransactionActivePourRevision_();

    ouvrirRepartitionPourRevisionTransaction_(
      selection.referenceBancaire,
      selection.idTransaction
    );
  } catch (erreur) {
    ui.alert(erreur && erreur.message ? erreur.message : String(erreur));
  }
}

function obtenirSelectionTransactionActivePourRevision_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getActiveSheet();
  const cellule = feuille.getActiveCell();

  if (feuille.getName() !== 'Transactions' || cellule.getRow() < 6) {
    throw new Error(
      'Sélectionnez d’abord une cellule sur la transaction à modifier dans l’onglet « Transactions ».'
    );
  }

  const valeurs = feuille
    .getRange(cellule.getRow(), 1, 1, 17)
    .getValues()[0];
  const idTransaction = String(valeurs[0] || '').trim();
  const source = String(valeurs[12] || '').trim();
  const referenceBancaire = String(valeurs[13] || '').trim();
  const statut = String(valeurs[14] || '').trim();

  if (!idTransaction) {
    throw new Error('La ligne sélectionnée ne contient aucune transaction.');
  }

  if (statut === 'Annulée' || statut === 'Exemple') {
    throw new Error('Seules les transactions actives peuvent être modifiées.');
  }

  if (!referenceBancaire) {
    throw new Error(
      'Cette transaction ne contient aucune référence bancaire. La révision contrôlée est impossible.'
    );
  }

  if (source !== 'Import') {
    throw new Error(
      'Cette transaction ne provient pas de l’import bancaire. Utilisez la correction manuelle appropriée.'
    );
  }

  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const lignesActives = lireLignesActivesRepartitionMixte_(
    repartition,
    referenceBancaire
  );

  if (lignesActives.length === 0) {
    throw new Error(
      'Aucune répartition active n’a été trouvée pour cette transaction.'
    );
  }

  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire,
    1,
    referenceBancaire,
    6
  );

  if (!ligneImport) {
    throw new Error(
      'La référence bancaire de cette transaction est introuvable dans « Import bancaire ».'
    );
  }

  return {
    idTransaction: idTransaction,
    referenceBancaire: referenceBancaire,
    ligneImport: ligneImport
  };
}

function obtenirTransactionsCiblesAnnulation_(
  transactions,
  idTransaction,
  referenceBancaire
) {
  const derniereLigne = transactions.getLastRow();

  if (derniereLigne < 6) {
    throw new Error('L’onglet « Transactions » est vide.');
  }

  const valeurs = transactions
    .getRange(6, 1, derniereLigne - 5, 17)
    .getValues();
  const cibles = [];

  valeurs.forEach(function(ligne, index) {
    const id = String(ligne[0] || '').trim();
    const reference = String(ligne[13] || '').trim();
    const statut = String(ligne[14] || '').trim();
    const correspond =
      referenceBancaire
        ? reference === referenceBancaire
        : id === idTransaction;

    if (
      correspond &&
      id &&
      statut !== 'Annulée' &&
      statut !== 'Exemple'
    ) {
      cibles.push({
        numero: index + 6,
        valeurs: ligne
      });
    }
  });

  if (cibles.length === 0) {
    throw new Error(
      'Aucune transaction active n’a été trouvée pour cette sélection.'
    );
  }

  return cibles;
}

function annulerGroupeTransactionsOSBL_(
  idTransaction,
  referenceBancaire
) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  try {
    return annulerGroupeTransactionsOSBLSansVerrou_(
      idTransaction,
      referenceBancaire,
      {}
    );
  } finally {
    verrou.releaseLock();
  }
}

function annulerGroupeTransactionsOSBLSansVerrou_(
  idTransaction,
  referenceBancaire,
  options
) {
  const optionsAnnulation = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const forfaits = obtenirFeuilleMixte_(ss, 'Forfaits');
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const inventaire = obtenirFeuilleMixte_(ss, 'Inventaire');
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const cibles = obtenirTransactionsCiblesAnnulation_(
    transactions,
    idTransaction,
    referenceBancaire
  );
  const ids = {};

  cibles.forEach(function(transaction) {
    ids[String(transaction.valeurs[0] || '').trim()] = true;
  });

  const restaurationInventaire =
    preparerRestaurationInventaireAnnulation_(
      repartition,
      inventaire,
      referenceBancaire
    );
  const ecritures = creerEcrituresAnnulationJournal_(
    journal,
    ids
  );
  const horodatage = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd HH:mm'
  );

  cibles.forEach(function(transaction) {
    const noteRevision = optionsAnnulation.idNouvelleRevision
      ? ' | Remplacée par ' + optionsAnnulation.idNouvelleRevision
      : '';

    transactions
      .getRange(transaction.numero, 15)
      .setValue('Annulée')
      .setNote('Annulée le ' + horodatage + noteRevision);
  });

  const forfaitsAnnules = annulerForfaitsParReference_(
    forfaits,
    referenceBancaire,
    horodatage
  );

  appliquerRestaurationInventaireAnnulation_(
    inventaire,
    restaurationInventaire
  );

  marquerRepartitionAnnulee_(
    repartition,
    referenceBancaire,
    horodatage
  );

  remettreImportBancaireAClasser_(
    importBancaire,
    referenceBancaire,
    horodatage
  );

  SpreadsheetApp.flush();

  return {
    transactions: cibles.length,
    ecritures: ecritures,
    forfaits: forfaitsAnnules,
    mouvementsInventaire: restaurationInventaire.length,
    horodatage: horodatage,
    idsTransactionsAnnulees: Object.keys(ids)
  };
}

function creerEcrituresAnnulationJournal_(journal, idsTransactions) {
  const derniereLigne = journal.getLastRow();

  if (derniereLigne < 6) {
    throw new Error(
      'Aucune écriture de journal n’a été trouvée pour cette transaction.'
    );
  }

  const valeurs = journal
    .getRange(6, 1, derniereLigne - 5, 16)
    .getValues();
  const idsEcrituresExistantes = {};
  const ecrituresAInverser = [];

  valeurs.forEach(function(ligne) {
    const idEcriture = String(ligne[0] || '').trim();
    const idTransaction = String(ligne[1] || '').trim();

    if (idEcriture) {
      idsEcrituresExistantes[idEcriture] = true;
    }

    if (
      idsTransactions[idTransaction] &&
      idEcriture &&
      idEcriture.indexOf('ANN-') !== 0
    ) {
      ecrituresAInverser.push(ligne);
    }
  });

  if (ecrituresAInverser.length === 0) {
    throw new Error(
      'Aucune écriture active n’a été trouvée dans « Journal ».'
    );
  }

  ecrituresAInverser.forEach(function(ligneSource) {
    const idEcritureSource = String(ligneSource[0] || '').trim();
    const idAnnulation = 'ANN-' + idEcritureSource;

    if (idsEcrituresExistantes[idAnnulation]) {
      throw new Error(
        'Une écriture d’annulation existe déjà pour ' +
        idEcritureSource +
        '.'
      );
    }
  });

  ecrituresAInverser.forEach(function(ligneSource) {
    const ligneCible = prochaineLigneVideMixte_(journal, 6);
    const cible = journal.getRange(ligneCible, 1, 1, 16);

    journal
      .getRange(6, 1, 1, 16)
      .copyTo(
        cible,
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );

    cible.setValues([[
      'ANN-' + String(ligneSource[0] || '').trim(),
      String(ligneSource[1] || '').trim(),
      ligneSource[2],
      String(ligneSource[3] || '').trim(),
      ligneSource[4],
      ligneSource[5],
      arrondirMontantMixte_(Number(ligneSource[7] || 0)),
      arrondirMontantMixte_(Number(ligneSource[6] || 0)),
      ligneSource[8],
      ligneSource[9],
      'Annulation – ' + String(ligneSource[10] || '').trim(),
      ligneSource[11],
      'Oui',
      '',
      String(ligneSource[14] || '').trim(),
      String(ligneSource[15] || '').trim()
    ]]);

    cible
      .getCell(1, 4)
      .setNumberFormat('@')
      .setValue(String(ligneSource[3] || '').trim());

    if (String(ligneSource[14] || '').trim()) {
      cible.getCell(1, 15).setNumberFormat('@')
        .setValue(String(ligneSource[14] || '').trim());
    }
  });

  return ecrituresAInverser.length;
}

function annulerForfaitsParReference_(
  forfaits,
  referenceBancaire,
  horodatage
) {
  if (!referenceBancaire || forfaits.getLastRow() < 6) {
    return 0;
  }

  const valeurs = forfaits
    .getRange(6, 1, forfaits.getLastRow() - 5, 10)
    .getValues();
  let compteur = 0;

  valeurs.forEach(function(ligne, index) {
    const reference = String(ligne[5] || '').trim();
    const statut = String(ligne[8] || '').trim();

    if (
      reference === referenceBancaire &&
      statut !== 'Annulé' &&
      statut !== 'Remboursé'
    ) {
      const numero = index + 6;
      const notes = String(ligne[9] || '').trim();

      forfaits.getRange(numero, 9).setValue('Annulé');
      forfaits
        .getRange(numero, 10)
        .setValue(
          notes +
          (notes ? ' | ' : '') +
          'Transaction annulée le ' +
          horodatage
        );
      compteur += 1;
    }
  });

  return compteur;
}

function preparerRestaurationInventaireAnnulation_(
  repartition,
  inventaire,
  referenceBancaire
) {
  if (!referenceBancaire) {
    return [];
  }

  const correspondances = {
    'T-shirt': 'MERCH-TS-NOIR',
    'Chandail à manches longues': 'MERCH-LS-NOIR',
    'Hoodie': 'MERCH-HD-NOIR'
  };
  const lignesRepartition = lireLignesActivesRepartitionMixte_(
    repartition,
    referenceBancaire
  );
  const quantites = {};

  lignesRepartition.forEach(function(ligne) {
    const composante = String(ligne.valeurs[5] || '').trim();
    const marqueur = String(ligne.valeurs[16] || '').trim();
    const ugs = correspondances[composante];

    if (ugs && marqueur.indexOf('Annulée') !== 0) {
      quantites[ugs] =
        Number(quantites[ugs] || 0) +
        Number(ligne.valeurs[6] || 0);
    }
  });

  if (Object.keys(quantites).length === 0) {
    return [];
  }

  const derniereLigneInventaire = inventaire.getLastRow();

  if (derniereLigneInventaire < 6) {
    throw new Error(
      'Impossible de restaurer les ventes de marchandise : aucun article exploitable n’existe dans « Inventaire » à partir de la ligne 6.'
    );
  }

  const valeursInventaire = inventaire
    .getRange(6, 1, derniereLigneInventaire - 5, 6)
    .getValues();
  const mouvements = [];

  Object.keys(quantites).forEach(function(ugs) {
    let trouve = false;

    valeursInventaire.forEach(function(ligne, index) {
      if (String(ligne[0] || '').trim() === ugs) {
        const ventesActuelles = Number(ligne[5] || 0);
        const quantite = Number(quantites[ugs] || 0);

        if (ventesActuelles < quantite) {
          throw new Error(
            'Impossible de restaurer ' +
            quantite +
            ' unité(s) de ' +
            ugs +
            ' : seulement ' +
            ventesActuelles +
            ' vente(s) sont inscrites dans « Inventaire ».'
          );
        }

        mouvements.push({
          ligne: index + 6,
          ventesApresAnnulation: ventesActuelles - quantite
        });
        trouve = true;
      }
    });

    if (!trouve) {
      throw new Error(
        'L’article ' + ugs + ' est introuvable dans « Inventaire ».'
      );
    }
  });

  return mouvements;
}

function appliquerRestaurationInventaireAnnulation_(
  inventaire,
  mouvements
) {
  mouvements.forEach(function(mouvement) {
    inventaire
      .getRange(mouvement.ligne, 6)
      .setValue(mouvement.ventesApresAnnulation);
  });
}

function marquerRepartitionAnnulee_(
  repartition,
  referenceBancaire,
  horodatage
) {
  if (!referenceBancaire) {
    return;
  }

  if (!String(repartition.getRange(5, 17).getValue() || '').trim()) {
    repartition.getRange(5, 17).setValue('Synchronisation');
  }

  const lignesActives = lireLignesActivesRepartitionMixte_(
    repartition,
    referenceBancaire
  );

  if (!lignesActives.length) {
    return;
  }

  const totalRevision = arrondirMontantMixte_(
    lignesActives.reduce(function(total, ligne) {
      return total + Number(ligne.valeurs[8] || 0);
    }, 0)
  );

  lignesActives.forEach(function(ligne) {
    const totalBancaire = arrondirMontantMixte_(
      Number(ligne.valeurs[3] || 0)
    );

    repartition.getRange(ligne.numero, 13).setValue(totalRevision);
    repartition.getRange(ligne.numero, 14).setValue(
      arrondirMontantMixte_(totalBancaire - totalRevision)
    );
    repartition.getRange(ligne.numero, 15).setValue('Annulée');
    repartition
      .getRange(ligne.numero, 17)
      .setValue('Annulée – ' + horodatage);
  });
}

function remettreImportBancaireAClasser_(
  importBancaire,
  referenceBancaire,
  horodatage
) {
  if (!referenceBancaire) {
    return;
  }

  const ligne = trouverLigneParValeurMixte_(
    importBancaire,
    1,
    referenceBancaire,
    6
  );

  if (!ligne) {
    throw new Error(
      'La référence bancaire ' +
      referenceBancaire +
      ' est introuvable dans « Import bancaire ».'
    );
  }

  const noteActuelle = String(
    importBancaire.getRange(ligne, 13).getValue() || ''
  ).trim();
  const parties = noteActuelle
    .split('|')
    .map(function(partie) {
      return partie.trim();
    })
    .filter(function(partie) {
      return (
        partie &&
        partie.indexOf('Transaction mixte enregistrée :') !== 0
      );
    });

  parties.push('Transaction annulée le ' + horodatage);

  importBancaire.getRange(ligne, 8, 1, 3).clearContent();
  importBancaire.getRange(ligne, 11).setValue('À classer');
  importBancaire.getRange(ligne, 13).setValue(parties.join(' | '));
  importBancaire.getRange(ligne, 15).clearContent();
}

function ouvrirRepartitionTransactionSelectionnee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getActiveSheet();
  const cellule = feuille.getActiveCell();

  if (feuille.getName() !== 'Import bancaire' || cellule.getRow() < 6) {
    SpreadsheetApp.getUi().alert(
      'Sélectionnez d’abord une cellule sur la transaction voulue dans l’onglet « Import bancaire ».'
    );
    return;
  }

  ouvrirRepartitionPourLigneMixte_(cellule.getRow());
}

function ouvrirInterfaceMixteDepuisImport(e) {
  if (!e || !e.range) {
    return;
  }

  const plage = e.range;
  const feuille = plage.getSheet();

  if (
    feuille.getName() !== 'Import bancaire' ||
    plage.getRow() < 6 ||
    plage.getColumn() !== 15 ||
    plage.getNumRows() !== 1 ||
    plage.getNumColumns() !== 1 ||
    String(e.value || '').trim() !== 'Classer'
  ) {
    return;
  }

  try {
    ouvrirRepartitionPourLigneMixte_(plage.getRow());
    plage.clearContent();
  } catch (erreur) {
    try {
      plage.clearContent();
    } catch (ignore) {}

    const message = erreur && erreur.message
      ? erreur.message
      : String(erreur);

    (e.source || SpreadsheetApp.getActiveSpreadsheet()).toast(
      message +
      ' Utilisez le menu « Classement bancaire » si la fenêtre ne s’est pas ouverte.',
      'Classement bancaire',
      10
    );
  }
}

function ouvrirRepartitionPourLigneMixte_(ligneImport) {
  const donnees = obtenirDonneesInterfaceMixte_(ligneImport, {
    modeRevision: false,
    idTransactionSource: ''
  });
  const modele = HtmlService.createTemplateFromFile('RepartitionMixte');
  modele.donnees = JSON.stringify(donnees);

  SpreadsheetApp.getUi().showModalDialog(
    modele.evaluate()
      .setWidth(840)
      .setHeight(610),
    'Classer une transaction bancaire'
  );
}

function ouvrirRepartitionPourRevisionTransaction_(
  referenceBancaire,
  idTransactionSource
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire,
    1,
    referenceBancaire,
    6
  );

  if (!ligneImport) {
    throw new Error(
      'La référence bancaire ' +
      referenceBancaire +
      ' est introuvable dans « Import bancaire ». '
    );
  }

  const donnees = obtenirDonneesInterfaceMixte_(ligneImport, {
    modeRevision: true,
    idTransactionSource: idTransactionSource
  });
  const modele = HtmlService.createTemplateFromFile('RepartitionMixte');

  modele.donnees = JSON.stringify(donnees);

  SpreadsheetApp.getUi().showModalDialog(
    modele.evaluate()
      .setWidth(840)
      .setHeight(610),
    'Modifier la transaction'
  );
}

function obtenirDonneesInterfaceMixte_(ligneImport, options) {
  const optionsChargement = options || {};
  const modeRevision = Boolean(optionsChargement.modeRevision);
  const idTransactionSource = String(
    optionsChargement.idTransactionSource || ''
  ).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const valeurs = importBancaire
    .getRange(ligneImport, 1, 1, 19)
    .getValues()[0];

  const idImport = String(valeurs[0] || '').trim();
  const statut = String(valeurs[10] || '').trim();
  const idTransaction = String(valeurs[9] || '').trim();
  const montant = Number(valeurs[3] || 0);

  if (!idImport) {
    throw new Error('La ligne sélectionnée ne contient aucune transaction.');
  }

  if (!modeRevision && (statut === 'Classée' || idTransaction)) {
    throw new Error(
      'Cette transaction est déjà classée' +
      (idTransaction ? ' sous ' + idTransaction : '') +
      '.'
    );
  }

  if (modeRevision && !idTransactionSource) {
    throw new Error('La transaction source de la révision est invalide.');
  }

  if (montant === 0) {
    throw new Error('Le montant bancaire ne peut pas être nul.');
  }

  const montantArrondi = arrondirMontantMixte_(montant);

  if (montant < 0) {
    const donneesComptes = obtenirComptesDépensePourMixte_(ss);

    let idFournisseur = '';
    let nomFournisseur = '';
    let idContact = '';
    let nomContact = '';
    const regle = rechercherRegleBancaire_(
      ss,
      String(valeurs[2] || '').trim(),
      montant
    );

    // Révision : priorité aux valeurs confirmées de la Transaction active (R:S:T, D)
    if (modeRevision && idTransactionSource) {
      const ligneSourceTx = trouverTransactionActivePourRevision_(
        ss,
        idTransactionSource
      );
      if (ligneSourceTx) {
        idFournisseur = String(ligneSourceTx[17] || '').trim(); // col R
        nomFournisseur = String(ligneSourceTx[18] || '').trim(); // col S
        idContact = String(ligneSourceTx[19] || '').trim();     // col T
        nomContact = String(ligneSourceTx[3] || '').trim();     // col D
      }
    }

    // Fallback : suggestions Import bancaire P:S, puis règle bancaire, puis FOU-0000
    if (!idFournisseur) {
      const idFournisseurSuggere = String(valeurs[15] || '').trim();
      const nomFournisseurSuggere = String(valeurs[16] || '').trim();
      const idContactSuggere = String(valeurs[17] || '').trim();
      const nomContactSuggere = String(valeurs[18] || '').trim();

      idFournisseur = idFournisseurSuggere || (regle ? regle.idFournisseur : '');
      nomFournisseur = nomFournisseurSuggere || (regle ? regle.nomFournisseur : '');
      idContact = idContactSuggere || (regle ? regle.idContact : '');
      nomContact = nomContactSuggere || (regle ? regle.nomContact : '');

      // FOU-0000 si aucun fournisseur identifié
      if (!idFournisseur) {
        idFournisseur = 'FOU-0000';
        nomFournisseur = obtenirNomFournisseur_(ss, 'FOU-0000') || 'À déterminer';
      }
    }

    const lignesExistantesDepense = lireLignesActivesRepartitionMixte_(
      repartition,
      idImport
    )
      .filter(function(ligne) {
        return String(ligne.valeurs[5] || '').trim() !== '';
      })
      .map(function(ligne) {
        return {
          compte: String(ligne.valeurs[5] || '').trim(),
          programme: String(ligne.valeurs[10] || '').trim(),
          projet: String(ligne.valeurs[11] || '').trim(),
          montant: Number(ligne.valeurs[7] || 0)
        };
      });

    const lignesDefaut = lignesExistantesDepense.length
      ? lignesExistantesDepense
      : [{
          compte: regle ? (regle.codeCompte || '') : '',
          programme: regle ? (regle.programme || '') : '',
          projet: '',
          montant: arrondirMontantMixte_(Math.abs(montant))
        }];

    return {
      typeMouvement: 'depense',
      mode: modeRevision ? 'revision' : 'creation',
      titre: modeRevision
        ? 'Modifier la transaction'
        : 'Classer une dépense bancaire',
      boutonPrincipal: modeRevision
        ? 'Enregistrer la révision'
        : 'Enregistrer et comptabiliser',
      idTransactionSource: idTransactionSource,
      idImport: idImport,
      date: Utilities.formatDate(
        valeurs[1],
        ss.getSpreadsheetTimeZone(),
        'yyyy-MM-dd'
      ),
      description: String(valeurs[2] || '').trim(),
      montant: montantArrondi,
      montantAbsolu: arrondirMontantMixte_(Math.abs(montant)),
      comptes: donneesComptes.comptes,
      programmes: donneesComptes.programmes,
      projets: donneesComptes.projets,
      fournisseurs: chargerFournisseursActifs_(ss),
      contacts: chargerTousContacts_(ss),
      idFournisseur: idFournisseur,
      nomFournisseur: nomFournisseur,
      idContact: idContact,
      nomContact: nomContact,
      lignes: lignesDefaut
    };
  }

  const optionsComposantes = construireOptionsComposantesMixtes_(ss, valeurs[1]);
  const donneesRevenusDirects = chargerComptesRevenusDirectsMixte_(ss);

  const lignesRepartitionExistantes = lireLignesActivesRepartitionMixte_(
    repartition,
    idImport
  ).filter(function(ligne) {
    return String(ligne.valeurs[5] || '').trim() !== '';
  });

  // Détecter le mode à partir des lignes actives de Répartition (col P = index 15)
  let modeRevenuDetecte = 'revenu';
  let regleBancaireCourant = null;

  if (lignesRepartitionExistantes.length > 0) {
    const marqueur = String(lignesRepartitionExistantes[0].valeurs[15] || '').trim();
    if (marqueur === 'Revenu comptable direct') {
      modeRevenuDetecte = 'revenu_direct';
    }
  } else if (!modeRevision) {
    regleBancaireCourant = rechercherRegleBancaire_(
      ss,
      String(valeurs[2] || '').trim(),
      montant
    );
    if (regleBancaireCourant && regleBancaireCourant.codeCompte) {
      const compteTrouve = donneesRevenusDirects.comptes.find(function(c) {
        return c.code === regleBancaireCourant.codeCompte;
      });
      if (compteTrouve) {
        modeRevenuDetecte = 'revenu_direct';
      }
    }
  }

  const dateFormatee = Utilities.formatDate(
    valeurs[1],
    ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd'
  );
  const descriptionBancaireMixte = String(valeurs[2] || '').trim();
  const titreBase = {
    mode: modeRevision ? 'revision' : 'creation',
    titre: modeRevision ? 'Modifier la transaction' : 'Classer la transaction',
    boutonPrincipal: modeRevision
      ? 'Enregistrer la révision'
      : 'Enregistrer et comptabiliser',
    idTransactionSource: idTransactionSource,
    idImport: idImport,
    date: dateFormatee,
    description: descriptionBancaireMixte,
    montant: montantArrondi
  };

  if (modeRevenuDetecte === 'revenu_direct') {
    const lignesDirectesExistantes = lignesRepartitionExistantes.map(
      function(ligne) {
        return {
          compte: String(ligne.valeurs[5] || '').trim(),
          programme: String(ligne.valeurs[10] || '').trim(),
          projet: String(ligne.valeurs[11] || '').trim(),
          montant: Number(ligne.valeurs[7] || 0)
        };
      }
    );

    let lignesDirectesDefaut;
    if (lignesDirectesExistantes.length > 0) {
      lignesDirectesDefaut = lignesDirectesExistantes;
    } else {
      const comptePreRempli = (regleBancaireCourant && regleBancaireCourant.codeCompte)
        ? regleBancaireCourant.codeCompte
        : '';
      const programmePreRempli = (regleBancaireCourant && regleBancaireCourant.programme)
        ? regleBancaireCourant.programme
        : '';
      lignesDirectesDefaut = [{
        compte: comptePreRempli,
        programme: programmePreRempli,
        projet: '',
        montant: montantArrondi
      }];
    }

    return Object.assign({}, titreBase, {
      typeMouvement: 'revenu_direct',
      comptes: donneesRevenusDirects.comptes,
      comptesRevenusDirects: donneesRevenusDirects.comptes,
      programmes: donneesRevenusDirects.programmes,
      projets: donneesRevenusDirects.projets,
      options: optionsComposantes,
      lignes: lignesDirectesDefaut
    });
  }

  const lignesExistantes = lignesRepartitionExistantes.map(function(ligne) {
    return {
      composante: String(ligne.valeurs[5] || '').trim(),
      quantite: Number(ligne.valeurs[6] || 1),
      prixUnitaire: Number(ligne.valeurs[7] || 0)
    };
  });

  return Object.assign({}, titreBase, {
    typeMouvement: 'revenu',
    comptes: donneesRevenusDirects.comptes,
    comptesRevenusDirects: donneesRevenusDirects.comptes,
    programmes: donneesRevenusDirects.programmes,
    projets: donneesRevenusDirects.projets,
    options: optionsComposantes,
    lignes: lignesExistantes.length
      ? lignesExistantes
      : [{ composante: '', quantite: 1, prixUnitaire: 0 }]
  });
}

function construireOptionsComposantesMixtes_(ss, dateTransaction) {
  const inventaire = obtenirFeuilleMixte_(ss, 'Inventaire');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');
  const annee = dateTransaction instanceof Date
    ? dateTransaction.getFullYear()
    : new Date(dateTransaction).getFullYear();

  const prixInventaire = {};
  const derniereLigneInventaire = inventaire.getLastRow();

  if (derniereLigneInventaire >= 6) {
    inventaire
      .getRange(6, 1, derniereLigneInventaire - 5, 11)
      .getValues()
      .forEach(function(ligne) {
        prixInventaire[String(ligne[0] || '').trim()] =
          Number(ligne[10] || 0);
      });
  }

  let prixCombine = annee >= 2027 ? 50 : 45;
  const tarifs = configuration.getRange('P6:S20').getValues();

  tarifs.some(function(ligne) {
    if (Number(ligne[0]) === annee) {
      prixCombine = Number(ligne[3] || prixCombine);
      return true;
    }
    return false;
  });

  return [
    { valeur: 'Entrée – Pion joues-tu?', prix: 10 },
    { valeur: 'Entrée – Cartier', prix: 2 },
    { valeur: 'Forfait Pion joues-tu?', prix: 30 },
    { valeur: 'Forfait Cartier', prix: 20 },
    { valeur: 'Forfait combiné', prix: prixCombine },
    {
      valeur: 'T-shirt',
      prix: Number(prixInventaire['MERCH-TS-NOIR'] || 25)
    },
    {
      valeur: 'Chandail à manches longues',
      prix: Number(prixInventaire['MERCH-LS-NOIR'] || 30)
    },
    {
      valeur: 'Hoodie',
      prix: Number(prixInventaire['MERCH-HD-NOIR'] || 50)
    }
  ];
}

function enregistrerEtComptabiliserTransactionMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  try {
    const contexte = preparerContexteTraitementMixte_(donnees, {
      modeRevision: false
    });

    enregistrerRepartitionTechniqueMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      { conserverHistorique: false }
    );

    const resultat = finaliserTransactionMixteParId_(
      contexte.idImport,
      {}
    );

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message:
        'Transaction enregistrée et comptabilisée sous ' +
        resultat.idGroupe +
        '.'
    };
  } finally {
    verrou.releaseLock();
  }
}

function enregistrerRevisionTransactionMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  let annulationEffectuee = false;

  try {
    const contexte = preparerContexteTraitementMixte_(donnees, {
      modeRevision: true
    });
    const idRevisionSource = extraireIdGroupeDepuisTransactionMixte_(
      contexte.idTransactionSource
    );

    annulerGroupeTransactionsOSBLSansVerrou_(
      contexte.idTransactionSource,
      contexte.idImport,
      {
        idNouvelleRevision: contexte.idGroupeCible
      }
    );
    annulationEffectuee = true;

    enregistrerRepartitionTechniqueMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      {
        conserverHistorique: true,
        noteRevision:
          'Révision active – de ' +
          idRevisionSource +
          ' vers ' +
          contexte.idGroupeCible
      }
    );

    const resultat = finaliserTransactionMixteParId_(
      contexte.idImport,
      {
        idGroupeForce: contexte.idGroupeCible,
        idRevisionSource: idRevisionSource
      }
    );

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message:
        'Révision enregistrée et comptabilisée sous ' +
        resultat.idGroupe +
        '.'
    };
  } catch (erreur) {
    if (annulationEffectuee) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
        const reference = String(donnees && donnees.idImport || '').trim();

        if (reference) {
          remettreImportBancaireAClasser_(
            importBancaire,
            reference,
            Utilities.formatDate(
              new Date(),
              ss.getSpreadsheetTimeZone(),
              'yyyy-MM-dd HH:mm'
            )
          );
        }
      } catch (erreurSecondaire) {
      }

      throw new Error(
        'L’ancienne version a été annulée, mais la nouvelle révision n’a pas pu être créée. ' +
        'La transaction bancaire a été remise à « À classer ». Détail : ' +
        (erreur && erreur.message ? erreur.message : String(erreur))
      );
    }

    throw erreur;
  } finally {
    verrou.releaseLock();
  }
}

function preparerContexteTraitementMixte_(donnees, options) {
  const optionsTraitement = options || {};
  const modeRevision = Boolean(optionsTraitement.modeRevision);

  if (!donnees || !donnees.idImport || !Array.isArray(donnees.lignes)) {
    throw new Error('Les données reçues sont incomplètes.');
  }

  const lignes = donnees.lignes
    .map(function(ligne) {
      return {
        composante: String(ligne.composante || '').trim(),
        quantite: Number(ligne.quantite || 0),
        prixUnitaire: Number(ligne.prixUnitaire || 0)
      };
    })
    .filter(function(ligne) {
      return ligne.composante !== '';
    });

  if (lignes.length === 0 || lignes.length > 5) {
    throw new Error(
      'La transaction doit contenir entre une et cinq composantes.'
    );
  }

  lignes.forEach(function(ligne, index) {
    if (!ligne.composante) {
      throw new Error(
        'La composante de la ligne ' + (index + 1) + ' est obligatoire.'
      );
    }

    if (Number(ligne.quantite || 0) <= 0) {
      throw new Error(
        'La quantité de la ligne ' + (index + 1) + ' doit être supérieure à zéro.'
      );
    }

    if (Number(ligne.prixUnitaire || 0) < 0) {
      throw new Error(
        'Le prix unitaire de la ligne ' + (index + 1) + ' est invalide.'
      );
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const inventaire = obtenirFeuilleMixte_(ss, 'Inventaire');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');
  const idImport = String(donnees.idImport || '').trim();
  const idTransactionSource = String(
    donnees.idTransactionSource || ''
  ).trim();

  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire,
    1,
    idImport,
    6
  );

  if (!ligneImport) {
    throw new Error('La transaction bancaire est introuvable.');
  }

  const valeursImport = importBancaire
    .getRange(ligneImport, 1, 1, 15)
    .getValues()[0];
  const montantBancaire = arrondirMontantMixte_(
    Number(valeursImport[3] || 0)
  );
  const totalSaisi = arrondirMontantMixte_(
    lignes.reduce(function(total, ligne) {
      return total + ligne.quantite * ligne.prixUnitaire;
    }, 0)
  );

  if (Math.abs(montantBancaire - totalSaisi) >= 0.005) {
    throw new Error(
      'Le total saisi (' +
      totalSaisi.toFixed(2) +
      ' $) doit correspondre au montant bancaire (' +
      montantBancaire.toFixed(2) +
      ' $).'
    );
  }

  if (montantBancaire <= 0) {
    throw new Error(
      'La finalisation automatique gère actuellement les revenus positifs seulement.'
    );
  }

  const planComptable = chargerPlanComptableMixte_(configuration);
  verifierComptesRequisMixte_(planComptable);

  const lignesActivesSaisies = convertirLignesSaisiesVersActivesMixte_(
    lignes,
    idImport
  );

  developperEcrituresComptablesMixtes_(lignesActivesSaisies);
  preparerMouvementsInventaireMixte_(inventaire, lignesActivesSaisies);

  if (modeRevision) {
    if (!idTransactionSource) {
      throw new Error('La transaction source de la révision est introuvable.');
    }

    const cibles = obtenirTransactionsCiblesAnnulation_(
      transactions,
      idTransactionSource,
      idImport
    );
    const ids = cibles.map(function(cible) {
      return String(cible.valeurs[0] || '').trim();
    });

    if (ids.indexOf(idTransactionSource) === -1) {
      throw new Error(
        'La transaction active sélectionnée n’a pas été retrouvée.'
      );
    }

    validerEcrituresJournalActivesMixte_(journal, ids);

    const repartitionActive = lireLignesActivesRepartitionMixte_(
      repartition,
      idImport
    );

    if (repartitionActive.length === 0) {
      throw new Error(
        'Aucune répartition active n’a été trouvée pour cette transaction.'
      );
    }
  }

  const idGroupeCible = creerIdGroupeDisponibleMixte_(transactions, idImport);

  validerAbsenceConflitIdGroupeMixte_(
    transactions,
    idGroupeCible
  );

  return {
    idImport: idImport,
    idTransactionSource: idTransactionSource,
    idGroupeCible: idGroupeCible,
    valeursImport: valeursImport,
    lignes: lignes,
    repartition: repartition
  };
}

function extraireIdGroupeDepuisTransactionMixte_(idTransaction) {
  return String(idTransaction || '').trim().replace(/-\d+$/, '');
}

function convertirLignesSaisiesVersActivesMixte_(lignes, idImport) {
  return lignes.map(function(ligne, index) {
    return {
      numero: index + 6,
      valeurs: [
        idImport,
        '',
        '',
        0,
        index + 1,
        ligne.composante,
        ligne.quantite,
        ligne.prixUnitaire,
        arrondirMontantMixte_(ligne.quantite * ligne.prixUnitaire)
      ]
    };
  });
}

function validerEcrituresJournalActivesMixte_(journal, idsTransactions) {
  if (!idsTransactions || idsTransactions.length === 0) {
    throw new Error('Aucune transaction active à valider dans le Journal.');
  }

  const derniereLigne = journal.getLastRow();

  if (derniereLigne < 6) {
    throw new Error('Aucune écriture de journal n’a été trouvée.');
  }

  const ensembleIds = {};

  idsTransactions.forEach(function(id) {
    ensembleIds[String(id || '').trim()] = false;
  });

  journal
    .getRange(6, 1, derniereLigne - 5, 2)
    .getDisplayValues()
    .forEach(function(ligne) {
      const idEcriture = String(ligne[0] || '').trim();
      const idTransaction = String(ligne[1] || '').trim();

      if (
        ensembleIds.hasOwnProperty(idTransaction) &&
        idEcriture.indexOf('ANN-') !== 0
      ) {
        ensembleIds[idTransaction] = true;
      }
    });

  const manquantes = Object.keys(ensembleIds).filter(function(id) {
    return !ensembleIds[id];
  });

  if (manquantes.length) {
    throw new Error(
      'Des écritures actives du Journal sont introuvables pour : ' +
      manquantes.join(', ')
    );
  }
}

function validerAbsenceConflitIdGroupeMixte_(transactions, idGroupe) {
  const derniereLigne = transactions.getLastRow();

  if (derniereLigne < 6) {
    return;
  }

  const valeurs = transactions
    .getRange(6, 1, derniereLigne - 5, 15)
    .getDisplayValues();
  const prefixe = idGroupe + '-';

  const conflit = valeurs.some(function(ligne) {
    const idTransaction = String(ligne[0] || '').trim();
    const statut = String(ligne[14] || '').trim();

    return (
      idTransaction.indexOf(prefixe) === 0 &&
      statut !== 'Annulée' &&
      statut !== 'Exemple'
    );
  });

  if (conflit) {
    throw new Error(
      'Conflit d’identifiant détecté pour la révision ' +
      idGroupe +
      '.'
    );
  }
}

function enregistrerRepartitionTechniqueMixte_(
  feuille,
  valeursImport,
  lignes,
  options
) {
  const optionsRepartition = options || {};
  const conserverHistorique = Boolean(
    optionsRepartition.conserverHistorique
  );
  const noteRevision = String(
    optionsRepartition.noteRevision || ''
  ).trim();
  const idImport = String(valeursImport[0] || '').trim();
  const lignesActivesExistantes = lireLignesActivesRepartitionMixte_(
    feuille,
    idImport
  );

  if (!conserverHistorique) {
    lignesActivesExistantes.forEach(function(ligneExistante) {
      feuille
        .getRange(ligneExistante.numero, 1, 1, 17)
        .clearContent();
    });
  }

  if (!conserverHistorique && lignesActivesExistantes.length) {
    SpreadsheetApp.flush();
  }

  const debut = trouverBlocVideRepartitionMixte_(
    feuille,
    lignes.length
  );

  const optionsComposantes = [
    'Entrée – Pion joues-tu?',
    'Entrée – Cartier',
    'Forfait Pion joues-tu?',
    'Forfait Cartier',
    'Forfait combiné',
    'T-shirt',
    'Chandail à manches longues',
    'Hoodie'
  ];

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(optionsComposantes, true)
    .setAllowInvalid(false)
    .build();
  const derniereLigneRepartition = feuille.getMaxRows();

  for (let index = 0; index < lignes.length; index += 1) {
    const numero = debut + index;
    const ligne = lignes[index];

    feuille.getRange(numero, 1, 1, 8).setValues([[
      idImport,
      valeursImport[1],
      valeursImport[2],
      Number(valeursImport[3] || 0),
      index + 1,
      ligne.composante,
      ligne.quantite,
      ligne.prixUnitaire
    ]]);

    feuille.getRange(numero, 6).setDataValidation(validation);
    feuille.getRange(numero, 9).setFormula(
      '=IF(OR($A' + numero + '="",$F' + numero + '=""),"",$G' +
      numero + '*$H' + numero + ')'
    );
    feuille.getRange(numero, 10).setFormula(
      '=IF($A' + numero + '="","",IF(OR($F' + numero +
      '="Entrée – Pion joues-tu?",$F' + numero +
      '="Entrée – Cartier"),"Revenus d\'entrées aux événements",IF(OR($F' +
      numero + '="Forfait Pion joues-tu?",$F' + numero +
      '="Forfait Cartier",$F' + numero +
      '="Forfait combiné"),"Revenus de laissez-passer et forfaits",IF(OR($F' +
      numero + '="T-shirt",$F' + numero +
      '="Chandail à manches longues",$F' + numero +
      '="Hoodie"),"Ventes de marchandises et de jeux",""))))'
    );
    feuille.getRange(numero, 11).setFormula(
      '=IF($A' + numero + '="","",IF(OR($F' + numero +
      '="Entrée – Pion joues-tu?",$F' + numero +
      '="Forfait Pion joues-tu?"),"Pion joues-tu?",IF(OR($F' +
      numero + '="Entrée – Cartier",$F' + numero +
      '="Forfait Cartier"),"Pion joues-tu? – Cartier",IF(OR($F' +
      numero + '="T-shirt",$F' + numero +
      '="Chandail à manches longues",$F' + numero +
      '="Hoodie"),"Marchandise",""))))'
    );
    feuille.getRange(numero, 12).setFormula(
      '=IF($K' + numero + '="Pion joues-tu?","PJT-2026",IF($K' +
      numero + '="Pion joues-tu? – Cartier","PJC-2026",""))'
    );
    feuille.getRange(numero, 13).setFormula(
      '=IF($A' + numero + '="","",SUMIFS($I$6:$I$' +
      derniereLigneRepartition + ',$A$6:$A$' + derniereLigneRepartition +
      ',$A' + numero + ',$Q$6:$Q$' + derniereLigneRepartition +
      ',"<>Annulée*"))'
    );
    feuille.getRange(numero, 14).setFormula(
      '=IF($A' + numero + '="","",$D' + numero + '-$M' + numero + ')'
    );
    feuille.getRange(numero, 15).setFormula(
      '=IF($A' + numero + '="","",IF(ABS($N' + numero +
      ')<0.005,"Prêt",IF($N' + numero +
      '<0,"Dépassement","À compléter")))'
    );

    feuille.getRange(numero, 17).setValue(
      conserverHistorique ? noteRevision : ''
    );
  }

  importBancaireModeMixte_(valeursImport[0]);
  SpreadsheetApp.flush();
}

function trouverBlocVideRepartitionMixte_(feuille, taille) {
  const maximum = feuille.getMaxRows();
  const ids = feuille
    .getRange(6, 1, maximum - 5, 1)
    .getDisplayValues();
  let consecutives = 0;

  for (let index = 0; index < ids.length; index += 1) {
    if (String(ids[index][0] || '').trim() === '') {
      consecutives += 1;
      if (consecutives === taille) {
        return index + 7 - taille;
      }
    } else {
      consecutives = 0;
    }
  }

  feuille.insertRowsAfter(maximum, taille);
  return maximum + 1;
}

function importBancaireModeMixte_(idImport) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const ligne = trouverLigneParValeurMixte_(feuille, 1, idImport, 6);

  if (ligne) {
    feuille.getRange(ligne, 15).setValue('Classer');
  }
}

function finaliserTransactionMixteParId_(idImport, options) {
  const optionsFinalisation = options || {};
  const idGroupeForce = String(
    optionsFinalisation.idGroupeForce || ''
  ).trim();
  const idRevisionSource = String(
    optionsFinalisation.idRevisionSource || ''
  ).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const forfaits = obtenirFeuilleMixte_(ss, 'Forfaits');
  const inventaire = obtenirFeuilleMixte_(ss, 'Inventaire');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');

  SpreadsheetApp.flush();

  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire,
    1,
    idImport,
    6
  );

  if (!ligneImport) {
    throw new Error(
      'La transaction ' + idImport + ' est introuvable dans « Import bancaire ».'
    );
  }

  const valeursImport = importBancaire
    .getRange(ligneImport, 1, 1, 15)
    .getValues()[0];

  const idDejaLie = String(valeursImport[9] || '').trim();
  const statutImport = String(valeursImport[10] || '').trim();

  if (idDejaLie || statutImport === 'Classée') {
    throw new Error(
      'Cette transaction bancaire est déjà classée' +
      (idDejaLie ? ' sous ' + idDejaLie : '') +
      '.'
    );
  }

  const idGroupeBase = creerIdGroupeMixte_(idImport);

  nettoyerTraitementPartielMixte_(
    transactions,
    journal,
    forfaits,
    idImport,
    idGroupeBase
  );

  if (transactionBancaireDejaPresenteMixte_(transactions, idImport)) {
    throw new Error(
      'Une transaction comptable utilise déjà la référence bancaire ' +
      idImport +
      '.'
    );
  }

  const idGroupe = idGroupeForce || creerIdGroupeDisponibleMixte_(
    transactions,
    idImport
  );

  if (idGroupeForce) {
    validerAbsenceConflitIdGroupeMixte_(transactions, idGroupe);
  }

  const lignesRepartition = lireLignesActivesRepartitionMixte_(
    repartition,
    idImport
  );

  if (lignesRepartition.length === 0) {
    throw new Error(
      'Aucune ligne de répartition n’a été trouvée pour ' + idImport + '.'
    );
  }

  const lignesActives = lignesRepartition.filter(function(ligne) {
    return String(ligne.valeurs[5] || '').trim() !== '';
  });

  if (lignesActives.length === 0) {
    throw new Error('Aucune composante n’a été sélectionnée.');
  }

  const totalBancaire = arrondirMontantMixte_(
    Number(valeursImport[3] || 0)
  );

  if (totalBancaire <= 0) {
    throw new Error(
      'La finalisation automatique gère actuellement les revenus positifs seulement.'
    );
  }

  lignesActives.forEach(function(ligne) {
    const composante = String(ligne.valeurs[5] || '').trim();
    const quantite = Number(ligne.valeurs[6] || 0);
    const prixUnitaire = Number(ligne.valeurs[7] || 0);
    const montant = Number(ligne.valeurs[8] || 0);

    if (!composante || quantite <= 0 || prixUnitaire < 0 || montant <= 0) {
      throw new Error(
        'La ligne ' +
        ligne.numero +
        ' contient une quantité, un prix ou un montant invalide.'
      );
    }
  });

  const totalReparti = arrondirMontantMixte_(
    lignesActives.reduce(function(total, ligne) {
      return total + Number(ligne.valeurs[8] || 0);
    }, 0)
  );

  if (Math.abs(totalBancaire - totalReparti) >= 0.005) {
    throw new Error(
      'Le total réparti (' +
      totalReparti.toFixed(2) +
      ' $) ne correspond pas au montant bancaire (' +
      totalBancaire.toFixed(2) +
      ' $).'
    );
  }

  const statutRepartition = String(
    lignesRepartition[0].valeurs[14] || ''
  ).trim();

  if (statutRepartition !== 'Prêt') {
    throw new Error(
      'Le statut de la répartition doit être « Prêt » avant l’enregistrement.'
    );
  }

  const planComptable = chargerPlanComptableMixte_(configuration);
  verifierComptesRequisMixte_(planComptable);

  const mouvementsInventaire = preparerMouvementsInventaireMixte_(
    inventaire,
    lignesActives
  );

  const dateTransaction = valeursImport[1];
  const descriptionBancaire = String(valeursImport[2] || '').trim();
  const acheteur = extraireNomInteracMixte_(descriptionBancaire);
  const ecrituresComptables = developperEcrituresComptablesMixtes_(
    lignesActives
  );
  const transactionParLigne = {};
  const transactionsCreees = [];

  ecrituresComptables.forEach(function(ecriture, index) {
    const idTransaction =
      idGroupe + '-' + String(index + 1).padStart(2, '0');

    const ligneTransaction = ecrireTransactionMixte_(
      transactions,
      idTransaction,
      dateTransaction,
      acheteur,
      ecriture.description,
      ecriture.montant,
      ecriture.codeCompte,
      ecriture.programme,
      ecriture.projet,
      idImport
    );

    transactionsCreees.push({
      idTransaction: idTransaction,
      ligne: ligneTransaction
    });

    ecrirePaireJournalRevenuMixte_(
      journal,
      planComptable,
      idTransaction,
      dateTransaction,
      ecriture
    );

    transactionParLigne[ecriture.numeroLigneSource] = idTransaction;
  });

  lignesActives.forEach(function(ligne) {
    const composante = String(ligne.valeurs[5] || '').trim();

    if (estForfaitMixte_(composante)) {
      creerForfaitDepuisRepartitionMixte_(
        forfaits,
        dateTransaction,
        acheteur,
        composante,
        Number(ligne.valeurs[8] || 0),
        idImport
      );
    }
  });

  mouvementsInventaire.forEach(function(mouvement) {
    const idTransaction =
      transactionParLigne[mouvement.numeroLigneSource];

    ecrirePaireJournalCoutMixte_(
      journal,
      planComptable,
      idTransaction,
      dateTransaction,
      mouvement
    );
  });

  appliquerMouvementsInventaireMixte_(
    inventaire,
    mouvementsInventaire
  );

  const ancienneNote = String(valeursImport[12] || '').trim();
  const noteRevision = idRevisionSource
    ? ' | Révision de ' + idRevisionSource + ' vers ' + idGroupe
    : '';
  const nouvelleNote =
    (ancienneNote ? ancienneNote + ' | ' : '') +
    'Transaction mixte enregistrée : ' +
    idGroupe +
    noteRevision;

  importBancaire.getRange(ligneImport, 8, 1, 2).clearContent();
  importBancaire.getRange(ligneImport, 10).setValue(idGroupe);
  importBancaire.getRange(ligneImport, 11).setValue('Classée');
  importBancaire.getRange(ligneImport, 13).setValue(nouvelleNote);
  importBancaire.getRange(ligneImport, 15).clearContent();

  if (idRevisionSource) {
    transactionsCreees.forEach(function(item) {
      transactions
        .getRange(item.ligne, 15)
        .setNote('Révision créée à partir de ' + idRevisionSource);
    });
  }

  SpreadsheetApp.flush();

  return {
    idGroupe: idGroupe
  };
}

function developperEcrituresComptablesMixtes_(lignesActives) {
  const definitions = {
    'Entrée – Pion joues-tu?': {
      codeCompte: '4000',
      programme: 'Pion joues-tu?',
      projet: 'PJT-2026'
    },
    'Entrée – Cartier': {
      codeCompte: '4000',
      programme: 'Pion joues-tu? – Cartier',
      projet: 'PJC-2026'
    },
    'Forfait Pion joues-tu?': {
      codeCompte: '4010',
      programme: 'Pion joues-tu?',
      projet: 'PJT-2026'
    },
    'Forfait Cartier': {
      codeCompte: '4010',
      programme: 'Pion joues-tu? – Cartier',
      projet: 'PJC-2026'
    },
    'T-shirt': {
      codeCompte: '4020',
      programme: 'Marchandise',
      projet: ''
    },
    'Chandail à manches longues': {
      codeCompte: '4020',
      programme: 'Marchandise',
      projet: ''
    },
    'Hoodie': {
      codeCompte: '4020',
      programme: 'Marchandise',
      projet: ''
    }
  };

  const resultat = [];

  lignesActives.forEach(function(ligne) {
    const composante = String(ligne.valeurs[5] || '').trim();
    const montant = arrondirMontantMixte_(
      Number(ligne.valeurs[8] || 0)
    );

    if (composante === 'Forfait combiné') {
      const partPion = arrondirMontantMixte_(montant * 0.6);
      const partCartier = arrondirMontantMixte_(montant - partPion);

      resultat.push({
        numeroLigneSource: ligne.numero,
        description: 'Forfait combiné – part Pion joues-tu?',
        montant: partPion,
        codeCompte: '4010',
        programme: 'Pion joues-tu?',
        projet: 'PJT-2026'
      });

      resultat.push({
        numeroLigneSource: ligne.numero,
        description: 'Forfait combiné – part Cartier',
        montant: partCartier,
        codeCompte: '4010',
        programme: 'Pion joues-tu? – Cartier',
        projet: 'PJC-2026'
      });

      return;
    }

    const definition = definitions[composante];

    if (!definition) {
      throw new Error(
        'La composante « ' + composante + ' » n’est pas reconnue.'
      );
    }

    resultat.push({
      numeroLigneSource: ligne.numero,
      description: composante,
      montant: montant,
      codeCompte: definition.codeCompte,
      programme: definition.programme,
      projet: definition.projet
    });
  });

  return resultat;
}

function ecrireTransactionMixte_(
  feuille,
  idTransaction,
  dateTransaction,
  contact,
  description,
  montant,
  codeCompte,
  programme,
  projet,
  idImport
) {
  const compte = obtenirCompteConfigurationMixte_(codeCompte);
  const ligne = prochaineLigneVideMixte_(feuille, 6);
  const modele = feuille.getRange(6, 1, 1, 17);
  const cible = feuille.getRange(ligne, 1, 1, 17);

  modele.copyTo(
    cible,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  modele.copyTo(
    cible,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );

  const formuleMois =
    '=IF($B' + ligne + '="","",TEXT($B' + ligne + ',"yyyy-mm"))';

  const formuleControlePiece =
    '=IF($A' +
    ligne +
    '="","",IF($O' +
    ligne +
    '="Exemple","Ignorée",IF(AND($C' +
    ligne +
    '="Dépense",$L' +
    ligne +
    '=""),"Pièce requise","OK")))';

  cible.setValues([[
    idTransaction,
    dateTransaction,
    'Revenu',
    contact,
    description,
    arrondirMontantMixte_(montant),
    compte.valeurCode,
    compte.nom,
    programme,
    projet,
    '1000',
    '',
    'Import',
    idImport,
    'Validée',
    formuleMois,
    formuleControlePiece
  ]]);

  cible
    .getCell(1, 7)
    .setNumberFormat('@')
    .setValue(String(compte.valeurCode));

  return ligne;
}

function obtenirCompteConfigurationMixte_(codeCompte) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');
  const derniereLigne = configuration.getLastRow();

  if (derniereLigne < 6) {
    throw new Error('Le plan comptable est vide.');
  }

  const codeRecherche = String(codeCompte || '').trim();
  const comptes = configuration
    .getRange(6, 1, derniereLigne - 5, 2)
    .getValues();

  for (let index = 0; index < comptes.length; index += 1) {
    const valeurCode = comptes[index][0];

    if (String(valeurCode || '').trim() === codeRecherche) {
      return {
        valeurCode: valeurCode,
        nom: String(comptes[index][1] || '').trim()
      };
    }
  }

  throw new Error(
    'Le compte ' + codeRecherche + ' est absent de « Configuration ».'
  );
}

function ecrirePaireJournalRevenuMixte_(
  journal,
  planComptable,
  idTransaction,
  dateTransaction,
  ecriture
) {
  const description = ecriture.description;
  const montant = arrondirMontantMixte_(ecriture.montant);

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-D',
    idTransaction,
    dateTransaction,
    '1000',
    montant,
    0,
    ecriture.programme,
    ecriture.projet,
    description
  );

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-C',
    idTransaction,
    dateTransaction,
    ecriture.codeCompte,
    0,
    montant,
    ecriture.programme,
    ecriture.projet,
    description
  );
}

function ecrirePaireJournalCoutMixte_(
  journal,
  planComptable,
  idTransaction,
  dateTransaction,
  mouvement
) {
  const coutTotal = arrondirMontantMixte_(
    mouvement.quantite * mouvement.coutUnitaire
  );

  if (coutTotal <= 0) {
    return;
  }

  const description = 'Coût des marchandises vendues – ' +
    mouvement.composante;

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-CMV-D',
    idTransaction,
    dateTransaction,
    '5000',
    coutTotal,
    0,
    'Marchandise',
    '',
    description
  );

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-CMV-C',
    idTransaction,
    dateTransaction,
    '1300',
    0,
    coutTotal,
    'Marchandise',
    '',
    description
  );
}

function ecrireLigneJournalMixte_(
  feuille,
  planComptable,
  idEcriture,
  idTransaction,
  dateTransaction,
  codeCompte,
  debit,
  credit,
  programme,
  projet,
  description,
  idFournisseur,
  nomFournisseur
) {
  const compte = planComptable[codeCompte];

  if (!compte) {
    throw new Error('Le compte ' + codeCompte + ' est introuvable.');
  }

  const ligne = prochaineLigneVideMixte_(feuille, 6);
  const modele = feuille.getRange(6, 1, 1, 13);
  const cible = feuille.getRange(ligne, 1, 1, 13);

  modele.copyTo(
    cible,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  cible.setValues([[
    idEcriture,
    idTransaction,
    dateTransaction,
    codeCompte,
    compte.nom,
    compte.type,
    arrondirMontantMixte_(debit),
    arrondirMontantMixte_(credit),
    programme,
    projet,
    description,
    '',
    'Oui'
  ]]);

  // Colonnes O (15) et P (16) : ID fournisseur et nom fournisseur
  if (idFournisseur) {
    feuille.getRange(ligne, 15).setNumberFormat('@').setValue(String(idFournisseur));
    feuille.getRange(ligne, 16).setValue(String(nomFournisseur || ''));
  }
}

function creerForfaitDepuisRepartitionMixte_(
  feuille,
  dateVente,
  acheteur,
  typeForfait,
  montant,
  idImport
) {
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne >= 6) {
    const existants = feuille
      .getRange(6, 1, derniereLigne - 5, 9)
      .getValues();

    const existe = existants.some(function(ligne) {
      const statut = String(ligne[8] || '').trim();

      return (
        String(ligne[5] || '').trim() === idImport &&
        String(ligne[3] || '').trim() === typeForfait &&
        Math.abs(Number(ligne[4] || 0) - Number(montant || 0)) < 0.005 &&
        statut !== 'Annulé' &&
        statut !== 'Remboursé'
      );
    });

    if (existe) {
      return;
    }
  }

  const saison = dateVente instanceof Date
    ? dateVente.getFullYear()
    : new Date(dateVente).getFullYear();

  const idForfait = genererProchainIdForfaitMixte_(feuille, saison);
  const ligne = prochaineLigneVideMixte_(feuille, 6);
  const modele = feuille.getRange(6, 1, 1, 11);
  const cible = feuille.getRange(ligne, 1, 1, 11);

  modele.copyTo(
    cible,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );
  modele.copyTo(
    cible,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );

  const formulePion =
    '=IF($A' +
    ligne +
    '="","",IF(OR($I' +
    ligne +
    '="Annulé",$I' +
    ligne +
    '="Remboursé"),0,IF($D' +
    ligne +
    '="Forfait Pion joues-tu?",$E' +
    ligne +
    ',IF($D' +
    ligne +
    '="Forfait combiné",$E' +
    ligne +
    '*0.6,0))))';

  const formuleCartier =
    '=IF($A' +
    ligne +
    '="","",IF(OR($I' +
    ligne +
    '="Annulé",$I' +
    ligne +
    '="Remboursé"),0,IF($D' +
    ligne +
    '="Forfait Cartier",$E' +
    ligne +
    ',IF($D' +
    ligne +
    '="Forfait combiné",$E' +
    ligne +
    '*0.4,0))))';

  cible.setValues([[
    idForfait,
    dateVente,
    acheteur,
    typeForfait,
    arrondirMontantMixte_(montant),
    idImport,
    formulePion,
    formuleCartier,
    'Actif',
    'Créé depuis une transaction mixte',
    saison
  ]]);
}

function preparerMouvementsInventaireMixte_(
  inventaire,
  lignesActives
) {
  const correspondances = {
    'T-shirt': 'MERCH-TS-NOIR',
    'Chandail à manches longues': 'MERCH-LS-NOIR',
    'Hoodie': 'MERCH-HD-NOIR'
  };

  const derniereLigne = inventaire.getLastRow();
  const valeurs = derniereLigne >= 6
    ? inventaire.getRange(6, 1, derniereLigne - 5, 14).getValues()
    : [];

  const articles = {};

  valeurs.forEach(function(ligne, index) {
    const ugs = String(ligne[0] || '').trim();

    if (ugs) {
      articles[ugs] = {
        ligne: index + 6,
        ventesActuelles: Number(ligne[5] || 0),
        quantiteDisponible: Number(ligne[8] || 0),
        coutUnitaire: Number(ligne[9] || ligne[3] || 0)
      };
    }
  });

  const mouvementsParUgs = {};

  lignesActives.forEach(function(ligne) {
    const composante = String(ligne.valeurs[5] || '').trim();
    const ugs = correspondances[composante];

    if (!ugs) {
      return;
    }

    const quantite = Number(ligne.valeurs[6] || 0);

    if (!mouvementsParUgs[ugs]) {
      mouvementsParUgs[ugs] = {
        ugs: ugs,
        composante: composante,
        quantite: 0,
        numerosLignesSources: []
      };
    }

    mouvementsParUgs[ugs].quantite += quantite;
    mouvementsParUgs[ugs].numerosLignesSources.push(ligne.numero);
  });

  const mouvements = [];

  Object.keys(mouvementsParUgs).forEach(function(ugs) {
    const mouvement = mouvementsParUgs[ugs];
    const article = articles[ugs];

    if (!article) {
      throw new Error(
        'L’article ' + ugs + ' est introuvable dans « Inventaire ».'
      );
    }

    if (article.quantiteDisponible < mouvement.quantite) {
      throw new Error(
        'Stock insuffisant pour « ' +
        mouvement.composante +
        ' » : ' +
        article.quantiteDisponible +
        ' disponible(s), ' +
        mouvement.quantite +
        ' demandé(s).'
      );
    }

    mouvements.push({
      ugs: ugs,
      composante: mouvement.composante,
      quantite: mouvement.quantite,
      numerosLignesSources: mouvement.numerosLignesSources,
      numeroLigneSource: mouvement.numerosLignesSources[0],
      ligneInventaire: article.ligne,
      ventesActuelles: article.ventesActuelles,
      coutUnitaire: article.coutUnitaire
    });
  });

  return mouvements;
}

function appliquerMouvementsInventaireMixte_(
  inventaire,
  mouvements
) {
  mouvements.forEach(function(mouvement) {
    inventaire
      .getRange(mouvement.ligneInventaire, 6)
      .setValue(
        Number(mouvement.ventesActuelles || 0) +
        Number(mouvement.quantite || 0)
      );
  });
}

function lireLignesRepartitionMixte_(feuille, idImport) {
  const derniereLigne = feuille.getMaxRows();

  if (derniereLigne < 6) {
    return [];
  }

  const correspondances = feuille
    .getRange(6, 1, derniereLigne - 5, 1)
    .createTextFinder(String(idImport))
    .matchEntireCell(true)
    .findAll();

  return correspondances
    .map(function(cellule) {
      const numero = cellule.getRow();
      const valeurs = feuille
        .getRange(numero, 1, 1, 17)
        .getValues()[0];

      return {
        numero: numero,
        valeurs: valeurs
      };
    })
    .filter(function(ligne) {
      return String(ligne.valeurs[0] || '').trim() === idImport;
    });
}

function estLigneRepartitionAnnuleeMixte_(ligne) {
  const synchronisation = String(ligne.valeurs[16] || '').trim();
  return synchronisation.indexOf('Annulée') === 0;
}

function lireLignesActivesRepartitionMixte_(feuille, idImport) {
  return lireLignesRepartitionMixte_(feuille, idImport).filter(
    function(ligne) {
      return !estLigneRepartitionAnnuleeMixte_(ligne);
    }
  );
}

function nettoyerTraitementPartielMixte_(
  transactions,
  journal,
  forfaits,
  idImport,
  idGroupe
) {
  const prefixeTransaction = idGroupe + '-';
  const lignesTransactions = [];
  const lignesJournal = [];
  const lignesForfaits = [];
  const idsTransactionsPartielles = {};

  if (transactions.getMaxRows() >= 6) {
    transactions
      .getRange(6, 1, transactions.getMaxRows() - 5, 15)
      .getDisplayValues()
      .forEach(function(ligne, index) {
        const idTransaction = String(ligne[0] || '').trim();
        const source = String(ligne[12] || '').trim();
        const reference = String(ligne[13] || '').trim();
        const statut = String(ligne[14] || '').trim();

        if (
          reference === idImport &&
          source === 'Import' &&
          idTransaction.indexOf(prefixeTransaction) === 0 &&
          statut !== 'Annulée' &&
          statut !== 'Exemple'
        ) {
          lignesTransactions.push(index + 6);
          idsTransactionsPartielles[idTransaction] = true;
        }
      });
  }

  if (journal.getMaxRows() >= 6) {
    journal
      .getRange(6, 1, journal.getMaxRows() - 5, 2)
      .getDisplayValues()
      .forEach(function(ligne, index) {
        const idEcriture = String(ligne[0] || '').trim();
        const idTransaction = String(ligne[1] || '').trim();

        if (
          idsTransactionsPartielles[idTransaction] &&
          idEcriture.indexOf('ANN-') !== 0
        ) {
          lignesJournal.push(index + 6);
        }
      });
  }

  if (forfaits.getMaxRows() >= 6) {
    forfaits
      .getRange(6, 1, forfaits.getMaxRows() - 5, 10)
      .getDisplayValues()
      .forEach(function(ligne, index) {
        const reference = String(ligne[5] || '').trim();
        const statut = String(ligne[8] || '').trim();
        const notes = String(ligne[9] || '').trim();

        if (
          reference === idImport &&
          notes.toLowerCase().indexOf('transaction mixte') !== -1 &&
          statut !== 'Annulé' &&
          statut !== 'Remboursé'
        ) {
          lignesForfaits.push(index + 6);
        }
      });
  }

  lignesTransactions.forEach(function(numeroLigne) {
    transactions.getRange(numeroLigne, 1, 1, 20).clearContent(); // A:T
  });

  lignesJournal.forEach(function(numeroLigne) {
    journal.getRange(numeroLigne, 1, 1, 16).clearContent();
  });

  lignesForfaits.forEach(function(numeroLigne) {
    forfaits.getRange(numeroLigne, 1, 1, 11).clearContent();
  });

  if (
    lignesTransactions.length ||
    lignesJournal.length ||
    lignesForfaits.length
  ) {
    SpreadsheetApp.flush();
  }
}

function nettoyerLignesRevenuDirectOrphelines_(repartition, idImport) {
  let modifie = false;

  lireLignesActivesRepartitionMixte_(repartition, idImport).forEach(
    function(ligne) {
      const marqueur = String(ligne.valeurs[15] || '').trim();
      const sync     = String(ligne.valeurs[16] || '').trim();

      if (
        marqueur === 'Revenu comptable direct' &&
        sync.indexOf('Annulée') !== 0
      ) {
        repartition.getRange(ligne.numero, 1, 1, 17).clearContent();
        modifie = true;
      }
    }
  );

  if (modifie) {
    SpreadsheetApp.flush();
  }
}

function transactionBancaireDejaPresenteMixte_(
  feuilleTransactions,
  idImport
) {
  const derniereLigne = feuilleTransactions.getLastRow();

  if (derniereLigne < 6) {
    return false;
  }

  return feuilleTransactions
    .getRange(6, 14, derniereLigne - 5, 2)
    .getDisplayValues()
    .some(function(ligne) {
      const reference = String(ligne[0] || '').trim();
      const statut = String(ligne[1] || '').trim();

      return reference === idImport && statut !== 'Annulée';
    });
}

function chargerPlanComptableMixte_(configuration) {
  const derniereLigne = configuration.getLastRow();

  if (derniereLigne < 6) {
    throw new Error('Le plan comptable est vide.');
  }

  const plan = {};

  configuration
    .getRange(6, 1, derniereLigne - 5, 3)
    .getDisplayValues()
    .forEach(function(ligne) {
      const code = String(ligne[0] || '').trim();

      if (code) {
        plan[code] = {
          nom: String(ligne[1] || '').trim(),
          type: String(ligne[2] || '').trim()
        };
      }
    });

  return plan;
}

function verifierComptesRequisMixte_(planComptable) {
  ['1000', '1300', '4000', '4010', '4020', '5000'].forEach(
    function(code) {
      if (!planComptable[code]) {
        throw new Error(
          'Le compte requis ' + code + ' est absent du plan comptable.'
        );
      }
    }
  );
}

function genererProchainIdForfaitMixte_(feuille, saison) {
  const derniereLigne = feuille.getLastRow();
  let maximum = 0;

  if (derniereLigne >= 6) {
    feuille
      .getRange(6, 1, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const correspondance = String(ligne[0] || '').match(
          /^FOR-\d{4}-(\d+)$/
        );

        if (correspondance) {
          maximum = Math.max(maximum, Number(correspondance[1]));
        }
      });
  }

  return (
    'FOR-' +
    saison +
    '-' +
    String(maximum + 1).padStart(4, '0')
  );
}

function creerIdGroupeMixte_(idImport) {
  const morceaux = String(idImport || '').split('-');

  if (
    morceaux.length >= 4 &&
    morceaux[0] === 'B'
  ) {
    return 'MIX-' + morceaux[1] + '-' + morceaux[2];
  }

  return 'MIX-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
}

function creerIdGroupeDisponibleMixte_(transactions, idImport) {
  const idBase = creerIdGroupeMixte_(idImport);
  const derniereLigne = transactions.getLastRow();
  let derniereRevision = 0;

  if (derniereLigne >= 6) {
    transactions
      .getRange(6, 1, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const idTransaction = String(ligne[0] || '').trim();

        if (idTransaction.indexOf(idBase + '-') !== 0) {
          return;
        }

        const suffixe = idTransaction.substring(idBase.length);
        const revision = suffixe.match(/^-R(\d+)-\d+$/);

        if (revision) {
          derniereRevision = Math.max(
            derniereRevision,
            Number(revision[1])
          );
        } else if (/^-\d+$/.test(suffixe)) {
          derniereRevision = Math.max(derniereRevision, 1);
        }
      });
  }

  return derniereRevision === 0
    ? idBase
    : idBase + '-R' + (derniereRevision + 1);
}

function extraireNomInteracMixte_(description) {
  const texte = String(description || '').trim();
  const correspondance = texte.match(/de\s*\/([^/]+)\//i);

  if (correspondance && correspondance[1]) {
    return correspondance[1].trim();
  }

  return texte || 'Client';
}

function estForfaitMixte_(composante) {
  return [
    'Forfait Pion joues-tu?',
    'Forfait Cartier',
    'Forfait combiné'
  ].indexOf(composante) !== -1;
}

function obtenirFeuilleMixte_(ss, nom) {
  const feuille = ss.getSheetByName(nom);

  if (!feuille) {
    throw new Error('L’onglet « ' + nom + ' » est introuvable.');
  }

  return feuille;
}

function trouverLigneParValeurMixte_(
  feuille,
  colonne,
  valeur,
  premiereLigne
) {
  const derniereLigne = feuille.getMaxRows();

  if (derniereLigne < premiereLigne) {
    return 0;
  }

  const resultat = feuille
    .getRange(
      premiereLigne,
      colonne,
      derniereLigne - premiereLigne + 1,
      1
    )
    .createTextFinder(String(valeur))
    .matchEntireCell(true)
    .findNext();

  return resultat ? resultat.getRow() : 0;
}

function prochaineLigneVideMixte_(feuille, premiereLigne) {
  const derniereLignePhysique = feuille.getMaxRows();

  if (derniereLignePhysique < premiereLigne) {
    feuille.insertRowsAfter(
      derniereLignePhysique,
      premiereLigne - derniereLignePhysique
    );
  }

  const nombreLignes =
    feuille.getMaxRows() - premiereLigne + 1;
  const ids = feuille
    .getRange(premiereLigne, 1, nombreLignes, 1)
    .getDisplayValues();

  for (let index = 0; index < ids.length; index += 1) {
    if (String(ids[index][0] || '').trim() === '') {
      return premiereLigne + index;
    }
  }

  const ancienneLimite = feuille.getMaxRows();
  feuille.insertRowsAfter(ancienneLimite, 100);
  return ancienneLimite + 1;
}

function arrondirMontantMixte_(montant) {
  return Math.round((Number(montant || 0) + Number.EPSILON) * 100) / 100;
}

// Retourne le tableau de valeurs (20 cols) de la Transaction active pour une révision.
// Retourne null si introuvable ou si les colonnes ne sont pas encore installées.
function trouverTransactionActivePourRevision_(ss, idTransactionSource) {
  try {
    const transactions = ss.getSheetByName('Transactions');
    if (!transactions || transactions.getLastRow() < 6) {
      return null;
    }
    const valeurs = transactions
      .getRange(6, 1, transactions.getLastRow() - 5, 20)
      .getValues();
    for (let i = 0; i < valeurs.length; i += 1) {
      const id = String(valeurs[i][0] || '').trim();
      const statut = String(valeurs[i][14] || '').trim();
      if (
        id === idTransactionSource &&
        statut !== 'Annulée' &&
        statut !== 'Exemple'
      ) {
        return valeurs[i];
      }
    }
  } catch (e) {
    // Retourner null si les colonnes T ne sont pas encore installées
  }
  return null;
}

function obtenirComptesDépensePourMixte_(ss) {
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');
  const projets = ss.getSheetByName('Projets');
  const derniereLigne = configuration.getLastRow();
  const comptes = [];

  if (derniereLigne >= 6) {
    configuration
      .getRange(6, 1, derniereLigne - 5, 5)
      .getValues()
      .forEach(function(ligne) {
        const code = String(ligne[0] || '').trim();
        const actif = String(ligne[4] || '').trim();
        const type = String(ligne[2] || '').trim();

        if (
          code &&
          actif === 'Oui' &&
          code !== '1000' &&
          (type === 'Dépense' || type === 'Actif')
        ) {
          comptes.push({
            code: code,
            nom: String(ligne[1] || '').trim(),
            type: type
          });
        }
      });
  }

  const programmes = configuration
    .getRange('I6:I14')
    .getValues()
    .flat()
    .filter(String);

  const listeProjets = [];

  if (projets) {
    const derniereLigneP = Math.max(projets.getLastRow(), 5);

    if (derniereLigneP >= 6) {
      projets
        .getRange(6, 1, derniereLigneP - 5, 4)
        .getValues()
        .forEach(function(ligne) {
          if (ligne[0] && String(ligne[3] || '') !== 'Exemple') {
            listeProjets.push({
              id: String(ligne[0]),
              nom: String(ligne[1] || ''),
              statut: String(ligne[3] || '')
            });
          }
        });
    }
  }

  return {
    comptes: comptes,
    programmes: programmes,
    projets: listeProjets
  };
}

function chargerComptesActifsDépense_(configuration) {
  const derniereLigne = configuration.getLastRow();
  const comptes = {};

  if (derniereLigne >= 6) {
    configuration
      .getRange(6, 1, derniereLigne - 5, 5)
      .getValues()
      .forEach(function(ligne) {
        const code = String(ligne[0] || '').trim();
        const actif = String(ligne[4] || '').trim();
        const type = String(ligne[2] || '').trim();

        if (
          code &&
          actif === 'Oui' &&
          code !== '1000' &&
          (type === 'Dépense' || type === 'Actif')
        ) {
          comptes[code] = true;
        }
      });
  }

  return comptes;
}

function enregistrerEtComptabiliserDepenseBancaireMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  try {
    const contexte = preparerContexteDepenseMixte_(donnees, {
      modeRevision: false
    });

    enregistrerRepartitionDepenseMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      { conserverHistorique: false }
    );

    const resultat = finaliserDepenseMixteParId_(contexte.idImport, {
      idFournisseur: contexte.idFournisseur,
      nomFournisseur: contexte.nomFournisseur,
      idContact: contexte.idContact,
      nomContact: contexte.nomContact
    });

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message: 'Dépense enregistrée et comptabilisée sous ' + resultat.idGroupe + '.'
    };
  } finally {
    verrou.releaseLock();
  }
}

function enregistrerRevisionDepenseBancaireMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  let annulationEffectuee = false;

  try {
    const contexte = preparerContexteDepenseMixte_(donnees, {
      modeRevision: true
    });
    const idRevisionSource = extraireIdGroupeDepuisTransactionMixte_(
      contexte.idTransactionSource
    );

    annulerGroupeTransactionsOSBLSansVerrou_(
      contexte.idTransactionSource,
      contexte.idImport,
      { idNouvelleRevision: contexte.idGroupeCible }
    );
    annulationEffectuee = true;

    enregistrerRepartitionDepenseMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      {
        conserverHistorique: true,
        noteRevision:
          'Révision active – de ' +
          idRevisionSource +
          ' vers ' +
          contexte.idGroupeCible
      }
    );

    const resultat = finaliserDepenseMixteParId_(contexte.idImport, {
      idGroupeForce: contexte.idGroupeCible,
      idRevisionSource: idRevisionSource,
      idFournisseur: contexte.idFournisseur,
      nomFournisseur: contexte.nomFournisseur,
      idContact: contexte.idContact,
      nomContact: contexte.nomContact
    });

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message: 'Révision de dépense enregistrée sous ' + resultat.idGroupe + '.'
    };
  } catch (erreur) {
    if (annulationEffectuee) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
        const reference = String(donnees && donnees.idImport || '').trim();

        if (reference) {
          remettreImportBancaireAClasser_(
            importBancaire,
            reference,
            Utilities.formatDate(
              new Date(),
              ss.getSpreadsheetTimeZone(),
              'yyyy-MM-dd HH:mm'
            )
          );
        }
      } catch (erreurSecondaire) {
      }

      throw new Error(
        "L'ancienne version a été annulée, mais la nouvelle révision n'a pas pu être créée. " +
        "La transaction bancaire a été remise à « À classer ». Détail : " +
        (erreur && erreur.message ? erreur.message : String(erreur))
      );
    }

    throw erreur;
  } finally {
    verrou.releaseLock();
  }
}

function preparerContexteDepenseMixte_(donnees, options) {
  const optionsTraitement = options || {};
  const modeRevision = Boolean(optionsTraitement.modeRevision);

  if (!donnees || !donnees.idImport || !Array.isArray(donnees.lignes)) {
    throw new Error('Les données reçues sont incomplètes.');
  }

  const lignes = donnees.lignes
    .map(function(ligne) {
      return {
        compte: String(ligne.compte || '').trim(),
        programme: String(ligne.programme || '').trim(),
        projet: String(ligne.projet || '').trim(),
        montant: arrondirMontantMixte_(Number(ligne.montant || 0))
      };
    })
    .filter(function(ligne) {
      return ligne.compte !== '';
    });

  if (lignes.length === 0 || lignes.length > 5) {
    throw new Error('La transaction doit contenir entre une et cinq lignes de dépense.');
  }

  lignes.forEach(function(ligne, index) {
    if (!ligne.compte) {
      throw new Error('Le compte de la ligne ' + (index + 1) + ' est obligatoire.');
    }

    if (Number(ligne.montant) <= 0) {
      throw new Error(
        'Le montant de la ligne ' + (index + 1) + ' doit être supérieur à zéro.'
      );
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');

  const idImport = String(donnees.idImport || '').trim();
  const idTransactionSource = String(donnees.idTransactionSource || '').trim();

  const ligneImport = trouverLigneParValeurMixte_(importBancaire, 1, idImport, 6);

  if (!ligneImport) {
    throw new Error('La transaction bancaire est introuvable.');
  }

  const valeursImport = importBancaire
    .getRange(ligneImport, 1, 1, 15)
    .getValues()[0];

  const montantBancaire = arrondirMontantMixte_(Number(valeursImport[3] || 0));

  if (montantBancaire >= 0) {
    throw new Error(
      'Cette fonction traite uniquement les dépenses (montant négatif dans Import bancaire).'
    );
  }

  const montantAbsolu = arrondirMontantMixte_(Math.abs(montantBancaire));
  const totalSaisi = arrondirMontantMixte_(
    lignes.reduce(function(total, ligne) {
      return total + ligne.montant;
    }, 0)
  );

  if (Math.abs(montantAbsolu - totalSaisi) >= 0.005) {
    throw new Error(
      'Le total saisi (' +
      totalSaisi.toFixed(2) +
      ' $) doit correspondre au montant bancaire (' +
      montantAbsolu.toFixed(2) +
      ' $).'
    );
  }

  const planComptable = chargerPlanComptableMixte_(configuration);
  const comptesActifs = chargerComptesActifsDépense_(configuration);

  if (!planComptable['1000']) {
    throw new Error('Le compte bancaire 1000 est introuvable dans le plan comptable.');
  }

  const programmesAutorises = configuration
    .getRange('I6:I14')
    .getValues()
    .flat()
    .filter(function(v) { return String(v || '').trim() !== ''; })
    .map(function(v) { return String(v).trim(); });

  const projetsSheet = ss.getSheetByName('Projets');
  const projetsAutorises = [];
  if (projetsSheet) {
    const derniereLigneP = Math.max(projetsSheet.getLastRow(), 5);
    if (derniereLigneP >= 6) {
      projetsSheet
        .getRange(6, 1, derniereLigneP - 5, 4)
        .getValues()
        .forEach(function(ligneProj) {
          if (ligneProj[0] && String(ligneProj[3] || '') !== 'Exemple') {
            projetsAutorises.push(String(ligneProj[0]).trim());
          }
        });
    }
  }

  lignes.forEach(function(ligne, index) {
    if (!planComptable[ligne.compte]) {
      throw new Error(
        'Le compte ' + ligne.compte + ' est absent du plan comptable (ligne ' +
        (index + 1) + ').'
      );
    }

    if (!comptesActifs[ligne.compte]) {
      throw new Error(
        'Le compte ' + ligne.compte +
        " n'est pas autorisé pour une dépense bancaire mixte (ligne " +
        (index + 1) + ').'
      );
    }

    if (ligne.programme && programmesAutorises.indexOf(ligne.programme) === -1) {
      throw new Error(
        'Le programme "' + ligne.programme + '" n\'est pas autorisé (ligne ' +
        (index + 1) + ').'
      );
    }

    if (ligne.projet && projetsAutorises.indexOf(ligne.projet) === -1) {
      throw new Error(
        'Le projet "' + ligne.projet + '" n\'est pas autorisé (ligne ' +
        (index + 1) + ').'
      );
    }
  });

  if (modeRevision) {
    if (!idTransactionSource) {
      throw new Error("La transaction source de la révision est introuvable.");
    }

    const cibles = obtenirTransactionsCiblesAnnulation_(
      transactions,
      idTransactionSource,
      idImport
    );
    const ids = cibles.map(function(cible) {
      return String(cible.valeurs[0] || '').trim();
    });

    if (ids.indexOf(idTransactionSource) === -1) {
      throw new Error("La transaction active sélectionnée n'a pas été retrouvée.");
    }

    validerEcrituresJournalActivesMixte_(journal, ids);

    const repartitionActive = lireLignesActivesRepartitionMixte_(repartition, idImport);

    if (repartitionActive.length === 0) {
      throw new Error("Aucune répartition active n'a été trouvée pour cette transaction.");
    }
  }

  const idGroupeCible = creerIdGroupeDisponibleMixte_(transactions, idImport);

  validerAbsenceConflitIdGroupeMixte_(transactions, idGroupeCible);

  const idFournisseur = String(donnees.idFournisseur || '').trim();
  const idContact = String(donnees.idContact || '').trim();
  const infosFC = validerFournisseurEtContact_(ss, idFournisseur, idContact);

  return {
    idImport: idImport,
    idTransactionSource: idTransactionSource,
    idGroupeCible: idGroupeCible,
    valeursImport: valeursImport,
    lignes: lignes,
    repartition: repartition,
    idFournisseur: infosFC.idFournisseur,
    nomFournisseur: infosFC.nomFournisseur,
    idContact: infosFC.idContact,
    nomContact: infosFC.nomContact
  };
}

function nettoyerLignesIncompletesRepartitionMixte_(feuille, idImport) {
  const lignes = lireLignesRepartitionMixte_(feuille, idImport);
  let nombreNettoyees = 0;

  lignes.forEach(function(ligne) {
    const colonneF = String(ligne.valeurs[5] || '').trim();
    const colonneJ = String(ligne.valeurs[9] || '').trim();
    const colonneO = String(ligne.valeurs[14] || '').trim();
    const colonneQ = String(ligne.valeurs[16] || '').trim();

    if (
      colonneF === '' &&
      colonneJ === '' &&
      colonneQ.indexOf('Annulée') !== 0
    ) {
      feuille.getRange(ligne.numero, 1, 1, 17).clearContent();
      nombreNettoyees += 1;
    }
  });

  return nombreNettoyees;
}

function enregistrerRepartitionDepenseMixte_(feuille, valeursImport, lignes, options) {
  const optionsRepartition = options || {};
  const conserverHistorique = Boolean(optionsRepartition.conserverHistorique);
  const noteRevision = String(optionsRepartition.noteRevision || '').trim();
  const idImport = String(valeursImport[0] || '').trim();
  const montantAbsolu = arrondirMontantMixte_(Math.abs(Number(valeursImport[3] || 0)));
  const lignesActivesExistantes = lireLignesActivesRepartitionMixte_(feuille, idImport);

  if (!conserverHistorique) {
    lignesActivesExistantes.forEach(function(ligneExistante) {
      feuille.getRange(ligneExistante.numero, 1, 1, 17).clearContent();
    });
  }

  if (!conserverHistorique && lignesActivesExistantes.length) {
    SpreadsheetApp.flush();
  }

  const nombreNettoyees = nettoyerLignesIncompletesRepartitionMixte_(feuille, idImport);
  if (nombreNettoyees > 0) {
    SpreadsheetApp.flush();
  }

  const debut = trouverBlocVideRepartitionMixte_(feuille, lignes.length);
  const derniereLigneRepartition = feuille.getMaxRows();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planComptable = chargerPlanComptableMixte_(
    obtenirFeuilleMixte_(ss, 'Configuration')
  );

  const lignesUtilisees = [];
  const validationsOriginales = {};

  try {
    for (let index = 0; index < lignes.length; index += 1) {
      const numero = debut + index;
      const ligne = lignes[index];
      const entreeCompte = planComptable[ligne.compte] || {};
      const nomCompte = entreeCompte.nom || ligne.compte;

      feuille.getRange(numero, 1, 1, 5).setValues([[
        idImport,
        valeursImport[1],
        valeursImport[2],
        montantAbsolu,
        index + 1
      ]]);

      lignesUtilisees.push(numero);

      const celluleF = feuille.getRange(numero, 6);
      validationsOriginales[numero] = celluleF.getDataValidation();
      celluleF.clearDataValidations();
      celluleF.setNumberFormat('@').setValue(String(ligne.compte));

      feuille.getRange(numero, 7, 1, 2).setValues([[1, arrondirMontantMixte_(ligne.montant)]]);

      feuille.getRange(numero, 9).setFormula(
        '=IF(OR($A' + numero + '="",$F' + numero + '=""),"",$G' + numero + '*$H' + numero + ')'
      );

      feuille.getRange(numero, 10).setValue(nomCompte);
      feuille.getRange(numero, 11).setValue(ligne.programme || '');
      feuille.getRange(numero, 12).setValue(ligne.projet || '');

      feuille.getRange(numero, 13).setFormula(
        '=IF($A' + numero + '="","",SUMIFS($I$6:$I$' +
        derniereLigneRepartition + ',$A$6:$A$' + derniereLigneRepartition +
        ',$A' + numero + ',$Q$6:$Q$' + derniereLigneRepartition + ',"<>Annulée*"))'
      );
      feuille.getRange(numero, 14).setFormula(
        '=IF($A' + numero + '="","",$D' + numero + '-$M' + numero + ')'
      );
      feuille.getRange(numero, 15).setFormula(
        '=IF($A' + numero + '="","",IF(ABS($N' + numero +
        ')<0.005,"Prêt",IF($N' + numero + '<0,"Dépassement","À compléter")))'
      );

      feuille.getRange(numero, 17).setValue(conserverHistorique ? noteRevision : '');
    }
  } catch (erreur) {
    lignesUtilisees.forEach(function(numero) {
      feuille.getRange(numero, 1, 1, 17).clearContent();
      const validationOriginale = validationsOriginales[numero];
      if (validationOriginale) {
        feuille.getRange(numero, 6).setDataValidation(validationOriginale);
      }
    });
    throw erreur;
  }

  importBancaireModeMixte_(valeursImport[0]);
  SpreadsheetApp.flush();
}

function finaliserDepenseMixteParId_(idImport, options) {
  const optionsFinalisation = options || {};
  const idGroupeForce = String(optionsFinalisation.idGroupeForce || '').trim();
  const idRevisionSource = String(optionsFinalisation.idRevisionSource || '').trim();
  const idFournisseur = String(optionsFinalisation.idFournisseur || '').trim();
  const nomFournisseur = String(optionsFinalisation.nomFournisseur || '').trim();
  const idContact = String(optionsFinalisation.idContact || '').trim();
  const nomContact = String(optionsFinalisation.nomContact || '').trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');

  SpreadsheetApp.flush();

  const ligneImport = trouverLigneParValeurMixte_(importBancaire, 1, idImport, 6);

  if (!ligneImport) {
    throw new Error(
      "La transaction " + idImport + " est introuvable dans « Import bancaire »."
    );
  }

  const valeursImport = importBancaire.getRange(ligneImport, 1, 1, 15).getValues()[0];
  const idDejaLie = String(valeursImport[9] || '').trim();
  const statutImport = String(valeursImport[10] || '').trim();

  if (idDejaLie || statutImport === 'Classée') {
    throw new Error(
      "Cette transaction bancaire est déjà classée" +
      (idDejaLie ? ' sous ' + idDejaLie : '') + '.'
    );
  }

  const montantBancaire = arrondirMontantMixte_(Number(valeursImport[3] || 0));

  if (montantBancaire >= 0) {
    throw new Error("La finalisation de dépense requiert un montant bancaire négatif.");
  }

  const montantAbsolu = arrondirMontantMixte_(Math.abs(montantBancaire));
  const idGroupeBase = creerIdGroupeMixte_(idImport);
  const forfaitsSheet = ss.getSheetByName('Forfaits');

  if (forfaitsSheet) {
    nettoyerTraitementPartielMixte_(
      transactions, journal, forfaitsSheet, idImport, idGroupeBase
    );
  }

  if (transactionBancaireDejaPresenteMixte_(transactions, idImport)) {
    throw new Error(
      "Une transaction comptable utilise déjà la référence bancaire " + idImport + '.'
    );
  }

  const idGroupe = idGroupeForce ||
    creerIdGroupeDisponibleMixte_(transactions, idImport);

  if (idGroupeForce) {
    validerAbsenceConflitIdGroupeMixte_(transactions, idGroupe);
  }

  const lignesRepartition = lireLignesActivesRepartitionMixte_(repartition, idImport);

  if (lignesRepartition.length === 0) {
    throw new Error(
      "Aucune ligne de répartition n'a été trouvée pour " + idImport + '.'
    );
  }

  const lignesActives = lignesRepartition.filter(function(ligne) {
    return String(ligne.valeurs[5] || '').trim() !== '';
  });

  if (lignesActives.length === 0) {
    throw new Error("Aucun compte n'a été sélectionné.");
  }

  lignesActives.forEach(function(ligne) {
    const codeCompte = String(ligne.valeurs[5] || '').trim();
    const montantLigne = Number(ligne.valeurs[8] || 0);

    if (!codeCompte || montantLigne <= 0) {
      throw new Error(
        'La ligne ' + ligne.numero + ' contient un compte ou un montant invalide.'
      );
    }
  });

  const totalReparti = arrondirMontantMixte_(
    lignesActives.reduce(function(total, ligne) {
      return total + Number(ligne.valeurs[8] || 0);
    }, 0)
  );

  if (Math.abs(montantAbsolu - totalReparti) >= 0.005) {
    throw new Error(
      'Le total réparti (' + totalReparti.toFixed(2) +
      ' $) ne correspond pas au montant bancaire (' + montantAbsolu.toFixed(2) + ' $).'
    );
  }

  const statutRepartition = String(lignesRepartition[0].valeurs[14] || '').trim();

  if (statutRepartition !== 'Prêt') {
    throw new Error(
      "Le statut de la répartition doit être « Prêt » avant l'enregistrement."
    );
  }

  const planComptable = chargerPlanComptableMixte_(configuration);
  const comptesActifs = chargerComptesActifsDépense_(configuration);

  if (!planComptable['1000']) {
    throw new Error('Le compte bancaire 1000 est introuvable dans le plan comptable.');
  }

  const dateTransaction = valeursImport[1];
  const descriptionBancaire = String(valeursImport[2] || '').trim();
  const transactionsCreees = [];

  lignesActives.forEach(function(ligne, index) {
    const codeCompte = String(ligne.valeurs[5] || '').trim();
    const montantLigne = arrondirMontantMixte_(Number(ligne.valeurs[8] || 0));
    const programme = String(ligne.valeurs[10] || '').trim();
    const projet = String(ligne.valeurs[11] || '').trim();

    if (!comptesActifs[codeCompte]) {
      throw new Error(
        "Le compte " + codeCompte +
        " n'est pas autorisé pour une dépense bancaire."
      );
    }

    const idTransaction = idGroupe + '-' + String(index + 1).padStart(2, '0');

    const ligneTransaction = ecrireTransactionDepenseMixte_(
      transactions,
      idTransaction,
      dateTransaction,
      descriptionBancaire,
      montantLigne,
      codeCompte,
      programme,
      projet,
      idImport,
      nomContact,
      idFournisseur,
      nomFournisseur,
      idContact
    );

    transactionsCreees.push({ idTransaction: idTransaction, ligne: ligneTransaction });

    ecrirePaireJournalDepenseMixte_(
      journal,
      planComptable,
      idTransaction,
      dateTransaction,
      {
        codeCompte: codeCompte,
        montant: montantLigne,
        programme: programme,
        projet: projet,
        description: descriptionBancaire,
        idFournisseur: idFournisseur,
        nomFournisseur: nomFournisseur
      }
    );
  });

  const ancienneNote = String(valeursImport[12] || '').trim();
  const noteRevision = idRevisionSource
    ? ' | Révision de ' + idRevisionSource + ' vers ' + idGroupe
    : '';
  const nouvelleNote =
    (ancienneNote ? ancienneNote + ' | ' : '') +
    'Transaction mixte enregistrée : ' + idGroupe + noteRevision;

  importBancaire.getRange(ligneImport, 8, 1, 2).clearContent();
  importBancaire.getRange(ligneImport, 10).setValue(idGroupe);
  importBancaire.getRange(ligneImport, 11).setValue('Classée');
  importBancaire.getRange(ligneImport, 13).setValue(nouvelleNote);
  importBancaire.getRange(ligneImport, 15).clearContent();

  // Retirer la validation de P:S sur cette ligne (protection contre la validation héritée de O)
  importBancaire.getRange(ligneImport, 16, 1, 4).clearDataValidations();

  // P:Q toujours écrits avec le fournisseur confirmé (obligatoire)
  importBancaire.getRange(ligneImport, 16).setNumberFormat('@').setValue(idFournisseur);
  importBancaire.getRange(ligneImport, 17).setValue(nomFournisseur);

  // R:S selon contact ; vider si absent pour ne pas laisser d'ancienne suggestion
  if (idContact) {
    importBancaire.getRange(ligneImport, 18).setNumberFormat('@').setValue(idContact);
    importBancaire.getRange(ligneImport, 19).setValue(nomContact);
  } else {
    importBancaire.getRange(ligneImport, 18, 1, 2).clearContent();
  }

  if (idRevisionSource) {
    transactionsCreees.forEach(function(item) {
      transactions
        .getRange(item.ligne, 15)
        .setNote('Révision créée à partir de ' + idRevisionSource);
    });
  }

  SpreadsheetApp.flush();

  return { idGroupe: idGroupe };
}

function ecrireTransactionDepenseMixte_(
  feuille,
  idTransaction,
  dateTransaction,
  description,
  montant,
  codeCompte,
  programme,
  projet,
  idImport,
  nomContact,
  idFournisseur,
  nomFournisseur,
  idContact
) {
  const compte = obtenirCompteConfigurationMixte_(codeCompte);
  const ligne = prochaineLigneVideMixte_(feuille, 6);
  const modele = feuille.getRange(6, 1, 1, 17);
  const cible = feuille.getRange(ligne, 1, 1, 17);

  modele.copyTo(cible, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  modele.copyTo(cible, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

  const formuleMois =
    '=IF($B' + ligne + '="","",TEXT($B' + ligne + ',"yyyy-mm"))';

  const formuleControlePiece =
    '=IF($A' + ligne +
    '="","",IF($O' + ligne +
    '="Exemple","Ignorée",IF(AND($C' + ligne +
    '="Dépense",$L' + ligne +
    '=""),"Pièce requise","OK")))';

  cible.setValues([[
    idTransaction,
    dateTransaction,
    'Dépense',
    String(nomContact || ''),
    description,
    arrondirMontantMixte_(montant),
    compte.valeurCode,
    compte.nom,
    programme,
    projet,
    '1000',
    '',
    'Import',
    idImport,
    'Validée',
    formuleMois,
    formuleControlePiece
  ]]);

  cible.getCell(1, 7).setNumberFormat('@').setValue(String(compte.valeurCode));

  // Colonnes R (18) et S (19) : ID fournisseur et nom fournisseur
  if (idFournisseur) {
    feuille.getRange(ligne, 18).setNumberFormat('@').setValue(String(idFournisseur));
    feuille.getRange(ligne, 19).setValue(String(nomFournisseur || ''));
  }

  // Colonne T (20) : ID contact
  if (idContact) {
    feuille.getRange(ligne, 20).setNumberFormat('@').setValue(String(idContact));
  }

  return ligne;
}

function ecrirePaireJournalDepenseMixte_(
  journal,
  planComptable,
  idTransaction,
  dateTransaction,
  ecriture
) {
  const montant = arrondirMontantMixte_(ecriture.montant);
  const programme = ecriture.programme || '';
  const projet = ecriture.projet || '';
  const description = ecriture.description || '';
  const idFournisseur = String(ecriture.idFournisseur || '').trim();
  const nomFournisseur = String(ecriture.nomFournisseur || '').trim();

  ecrireLigneJournalMixte_(
    journal, planComptable,
    'ECR-' + idTransaction + '-D',
    idTransaction, dateTransaction,
    ecriture.codeCompte, montant, 0,
    programme, projet, description,
    idFournisseur, nomFournisseur
  );

  ecrireLigneJournalMixte_(
    journal, planComptable,
    'ECR-' + idTransaction + '-C',
    idTransaction, dateTransaction,
    '1000', 0, montant,
    programme, projet, description,
    idFournisseur, nomFournisseur
  );
}

// ─── Revenus comptables directs ───────────────────────────────────────────────

function chargerComptesRevenusDirectsMixte_(ss) {
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');
  const projets = ss.getSheetByName('Projets');
  const derniereLigne = configuration.getLastRow();
  const exclus = ['4000', '4010', '4020'];
  const comptes = [];

  if (derniereLigne >= 6) {
    configuration
      .getRange(6, 1, derniereLigne - 5, 5)
      .getValues()
      .forEach(function(ligne) {
        const code = String(ligne[0] || '').trim();
        const actif = String(ligne[4] || '').trim();
        const type = String(ligne[2] || '').trim();

        if (
          code &&
          actif === 'Oui' &&
          type === 'Revenu' &&
          exclus.indexOf(code) === -1
        ) {
          comptes.push({
            code: code,
            nom: String(ligne[1] || '').trim(),
            type: type
          });
        }
      });
  }

  const programmes = configuration
    .getRange('I6:I14')
    .getValues()
    .flat()
    .filter(String);

  const listeProjets = [];

  if (projets) {
    const derniereLigneP = Math.max(projets.getLastRow(), 5);

    if (derniereLigneP >= 6) {
      projets
        .getRange(6, 1, derniereLigneP - 5, 4)
        .getValues()
        .forEach(function(ligne) {
          if (ligne[0] && String(ligne[3] || '') !== 'Exemple') {
            listeProjets.push({
              id: String(ligne[0]),
              nom: String(ligne[1] || ''),
              statut: String(ligne[3] || '')
            });
          }
        });
    }
  }

  return {
    comptes: comptes,
    programmes: programmes,
    projets: listeProjets
  };
}

function chargerComptesActifsRevenuDirect_(configuration) {
  const derniereLigne = configuration.getLastRow();
  const comptesActifs = {};
  const exclus = ['4000', '4010', '4020'];

  if (derniereLigne >= 6) {
    configuration
      .getRange(6, 1, derniereLigne - 5, 5)
      .getValues()
      .forEach(function(ligne) {
        const code = String(ligne[0] || '').trim();
        const actif = String(ligne[4] || '').trim();
        const type = String(ligne[2] || '').trim();

        if (
          code &&
          actif === 'Oui' &&
          type === 'Revenu' &&
          exclus.indexOf(code) === -1
        ) {
          comptesActifs[code] = true;
        }
      });
  }

  return comptesActifs;
}

function enregistrerEtComptabiliserRevenuDirectMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  let contexte = null;
  let repartitionEcrite = false;

  try {
    contexte = preparerContexteRevenuDirectMixte_(donnees, {
      modeRevision: false
    });

    enregistrerRepartitionRevenuDirectMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      { conserverHistorique: false }
    );
    repartitionEcrite = true;

    const resultat = finaliserRevenuDirectMixteParId_(contexte.idImport, {
      idGroupeForce: contexte.idGroupeCible
    });

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message: 'Revenu enregistré et comptabilisé sous ' + resultat.idGroupe + '.'
    };

  } catch (erreur) {
    if (repartitionEcrite && contexte) {
      const ss           = SpreadsheetApp.getActiveSpreadsheet();
      const horodatage   = Utilities.formatDate(
        new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'
      );
      const erreursSecondaires = [];

      try {
        nettoyerLignesRevenuDirectOrphelines_(
          contexte.repartition,
          contexte.idImport
        );
      } catch (e) {
        erreursSecondaires.push('Répartition : ' + e.message);
        console.error('Rollback Répartition :', e.message);
      }

      try {
        const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
        const journal      = obtenirFeuilleMixte_(ss, 'Journal');
        const forfaits     = ss.getSheetByName('Forfaits');
        if (forfaits) {
          nettoyerTraitementPartielMixte_(
            transactions, journal, forfaits,
            contexte.idImport, contexte.idGroupeCible
          );
        }
      } catch (e) {
        erreursSecondaires.push('Transactions/Journal : ' + e.message);
        console.error('Rollback Transactions/Journal :', e.message);
      }

      try {
        const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
        remettreImportBancaireAClasser_(importBancaire, contexte.idImport, horodatage);
      } catch (e) {
        erreursSecondaires.push('Import bancaire : ' + e.message);
        console.error('Rollback Import bancaire :', e.message);
      }

      if (erreursSecondaires.length > 0) {
        throw new Error(
          'La comptabilisation a échoué : ' + erreur.message +
          '. Rollback incomplet : ' + erreursSecondaires.join('; ') +
          '. Vérifiez Répartition, Transactions, Journal et Import bancaire.'
        );
      }
    }
    throw erreur;

  } finally {
    verrou.releaseLock();
  }
}

function enregistrerRevisionRevenuDirectMixte(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Attendez quelques secondes et réessayez.'
    );
  }

  let contexte = null;
  let annulationEffectuee = false;
  let repartitionEcrite = false;

  try {
    contexte = preparerContexteRevenuDirectMixte_(donnees, {
      modeRevision: true
    });
    const idRevisionSource = extraireIdGroupeDepuisTransactionMixte_(
      contexte.idTransactionSource
    );

    annulerGroupeTransactionsOSBLSansVerrou_(
      contexte.idTransactionSource,
      contexte.idImport,
      { idNouvelleRevision: contexte.idGroupeCible }
    );
    annulationEffectuee = true;

    enregistrerRepartitionRevenuDirectMixte_(
      contexte.repartition,
      contexte.valeursImport,
      contexte.lignes,
      {
        conserverHistorique: true,
        noteRevision:
          'Révision active – de ' +
          idRevisionSource +
          ' vers ' +
          contexte.idGroupeCible
      }
    );
    repartitionEcrite = true;

    const resultat = finaliserRevenuDirectMixteParId_(contexte.idImport, {
      idGroupeForce: contexte.idGroupeCible,
      idRevisionSource: idRevisionSource
    });

    return {
      succes: true,
      idGroupe: resultat.idGroupe,
      message: 'Révision enregistrée et comptabilisée sous ' + resultat.idGroupe + '.'
    };

  } catch (erreur) {
    if (annulationEffectuee && contexte) {
      const ss           = SpreadsheetApp.getActiveSpreadsheet();
      const horodatage   = Utilities.formatDate(
        new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'
      );
      const erreursSecondaires = [];

      if (repartitionEcrite) {
        try {
          nettoyerLignesRevenuDirectOrphelines_(
            contexte.repartition,
            contexte.idImport
          );
        } catch (e) {
          erreursSecondaires.push('Répartition : ' + e.message);
          console.error('Rollback Répartition :', e.message);
        }

        try {
          const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
          const journal      = obtenirFeuilleMixte_(ss, 'Journal');
          const forfaits     = ss.getSheetByName('Forfaits');
          if (forfaits) {
            nettoyerTraitementPartielMixte_(
              transactions, journal, forfaits,
              contexte.idImport, contexte.idGroupeCible
            );
          }
        } catch (e) {
          erreursSecondaires.push('Transactions/Journal : ' + e.message);
          console.error('Rollback Transactions/Journal :', e.message);
        }
      }

      try {
        const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
        remettreImportBancaireAClasser_(importBancaire, contexte.idImport, horodatage);
      } catch (e) {
        erreursSecondaires.push('Import bancaire : ' + e.message);
        console.error('Rollback Import bancaire :', e.message);
      }

      const messageOriginal =
        erreur && erreur.message ? erreur.message : String(erreur);

      const messageSuffix = erreursSecondaires.length > 0
        ? ' Rollback incomplet : ' + erreursSecondaires.join('; ') +
          '. Vérifiez Répartition, Transactions, Journal et Import bancaire.'
        : ' La transaction bancaire a été remise à « À classer ».';

      throw new Error(
        'L\'ancienne version a été annulée, mais la nouvelle révision n\'a pas pu être créée.' +
        messageSuffix + ' Détail : ' + messageOriginal
      );
    }

    throw erreur;

  } finally {
    verrou.releaseLock();
  }
}

function preparerContexteRevenuDirectMixte_(donnees, options) {
  const optionsTraitement = options || {};
  const modeRevision = Boolean(optionsTraitement.modeRevision);
  const exclus = ['4000', '4010', '4020'];

  if (!donnees || !donnees.idImport || !Array.isArray(donnees.lignes)) {
    throw new Error('Les données reçues sont incomplètes.');
  }

  const lignes = donnees.lignes
    .map(function(ligne) {
      return {
        compte: String(ligne.compte || '').trim(),
        programme: String(ligne.programme || '').trim(),
        projet: String(ligne.projet || '').trim(),
        montant: arrondirMontantMixte_(Number(ligne.montant || 0))
      };
    })
    .filter(function(ligne) {
      return ligne.compte !== '';
    });

  if (lignes.length === 0 || lignes.length > 5) {
    throw new Error(
      'La transaction doit contenir entre une et cinq lignes de revenu.'
    );
  }

  lignes.forEach(function(ligne, index) {
    if (!ligne.compte) {
      throw new Error(
        'Le compte de la ligne ' + (index + 1) + ' est obligatoire.'
      );
    }

    if (Number(ligne.montant) <= 0) {
      throw new Error(
        'Le montant de la ligne ' + (index + 1) + ' doit être supérieur à zéro.'
      );
    }

    if (exclus.indexOf(ligne.compte) !== -1) {
      throw new Error(
        'Le compte ' +
        ligne.compte +
        ' est réservé au mode spécialisé (entrées, forfaits, marchandises).'
      );
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');

  const idImport = String(donnees.idImport || '').trim();
  const idTransactionSource = String(donnees.idTransactionSource || '').trim();

  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire, 1, idImport, 6
  );

  if (!ligneImport) {
    throw new Error('La transaction bancaire est introuvable.');
  }

  const valeursImport = importBancaire
    .getRange(ligneImport, 1, 1, 15)
    .getValues()[0];
  const montantBancaire = arrondirMontantMixte_(Number(valeursImport[3] || 0));

  if (montantBancaire <= 0) {
    throw new Error(
      'Cette fonction traite uniquement les revenus ' +
      '(montant positif dans Import bancaire).'
    );
  }

  const totalSaisi = arrondirMontantMixte_(
    lignes.reduce(function(total, ligne) {
      return total + ligne.montant;
    }, 0)
  );

  if (Math.abs(montantBancaire - totalSaisi) >= 0.005) {
    throw new Error(
      'Le total saisi (' +
      totalSaisi.toFixed(2) +
      ' $) doit correspondre au montant bancaire (' +
      montantBancaire.toFixed(2) +
      ' $).'
    );
  }

  const planComptable = chargerPlanComptableMixte_(configuration);

  if (!planComptable['1000']) {
    throw new Error(
      'Le compte bancaire 1000 est introuvable dans le plan comptable.'
    );
  }

  const comptesActifsRevenu = chargerComptesActifsRevenuDirect_(configuration);

  const programmesAutorises = configuration
    .getRange('I6:I14')
    .getValues()
    .flat()
    .filter(function(v) {
      return String(v || '').trim() !== '';
    })
    .map(function(v) {
      return String(v).trim();
    });

  const projetsSheet = ss.getSheetByName('Projets');
  const projetsAutorises = [];

  if (projetsSheet) {
    const derniereLigneP = Math.max(projetsSheet.getLastRow(), 5);

    if (derniereLigneP >= 6) {
      projetsSheet
        .getRange(6, 1, derniereLigneP - 5, 4)
        .getValues()
        .forEach(function(ligneProj) {
          if (ligneProj[0] && String(ligneProj[3] || '') !== 'Exemple') {
            projetsAutorises.push(String(ligneProj[0]).trim());
          }
        });
    }
  }

  lignes.forEach(function(ligne, index) {
    if (!planComptable[ligne.compte]) {
      throw new Error(
        'Le compte ' +
        ligne.compte +
        ' est absent du plan comptable (ligne ' +
        (index + 1) +
        ').'
      );
    }

    if (!comptesActifsRevenu[ligne.compte]) {
      throw new Error(
        'Le compte ' +
        ligne.compte +
        ' n\'est pas un compte de revenu direct actif (ligne ' +
        (index + 1) +
        ').'
      );
    }

    if (ligne.programme && programmesAutorises.indexOf(ligne.programme) === -1) {
      throw new Error(
        'Le programme "' +
        ligne.programme +
        '" n\'est pas autorisé (ligne ' +
        (index + 1) +
        ').'
      );
    }

    if (ligne.projet && projetsAutorises.indexOf(ligne.projet) === -1) {
      throw new Error(
        'Le projet "' +
        ligne.projet +
        '" n\'est pas autorisé (ligne ' +
        (index + 1) +
        ').'
      );
    }
  });

  if (modeRevision) {
    if (!idTransactionSource) {
      throw new Error('La transaction source de la révision est introuvable.');
    }

    const cibles = obtenirTransactionsCiblesAnnulation_(
      transactions,
      idTransactionSource,
      idImport
    );
    const ids = cibles.map(function(cible) {
      return String(cible.valeurs[0] || '').trim();
    });

    if (ids.indexOf(idTransactionSource) === -1) {
      throw new Error(
        'La transaction active sélectionnée n\'a pas été retrouvée.'
      );
    }

    validerEcrituresJournalActivesMixte_(journal, ids);

    const repartitionActive = lireLignesActivesRepartitionMixte_(
      repartition,
      idImport
    );

    if (repartitionActive.length === 0) {
      throw new Error(
        'Aucune répartition active n\'a été trouvée pour cette transaction.'
      );
    }
  }

  const idGroupeCible = creerIdGroupeDisponibleMixte_(transactions, idImport);
  validerAbsenceConflitIdGroupeMixte_(transactions, idGroupeCible);

  return {
    idImport: idImport,
    idTransactionSource: idTransactionSource,
    idGroupeCible: idGroupeCible,
    valeursImport: valeursImport,
    lignes: lignes,
    repartition: repartition
  };
}

function enregistrerRepartitionRevenuDirectMixte_(
  feuille,
  valeursImport,
  lignes,
  options
) {
  const optionsRepartition = options || {};
  const conserverHistorique = Boolean(optionsRepartition.conserverHistorique);
  const noteRevision = String(optionsRepartition.noteRevision || '').trim();
  const idImport = String(valeursImport[0] || '').trim();
  const montant = arrondirMontantMixte_(Number(valeursImport[3] || 0));
  const lignesActivesExistantes = lireLignesActivesRepartitionMixte_(
    feuille,
    idImport
  );

  if (!conserverHistorique) {
    lignesActivesExistantes.forEach(function(ligneExistante) {
      feuille.getRange(ligneExistante.numero, 1, 1, 17).clearContent();
    });
  }

  if (!conserverHistorique && lignesActivesExistantes.length) {
    SpreadsheetApp.flush();
  }

  const debut = trouverBlocVideRepartitionMixte_(feuille, lignes.length);
  const derniereLigneRepartition = feuille.getMaxRows();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planComptable = chargerPlanComptableMixte_(
    obtenirFeuilleMixte_(ss, 'Configuration')
  );

  const lignesUtilisees = [];
  const validationsOriginales = {};

  try {
    for (let index = 0; index < lignes.length; index += 1) {
      const numero = debut + index;
      const ligne = lignes[index];
      const entreeCompte = planComptable[ligne.compte] || {};
      const nomCompte = entreeCompte.nom || ligne.compte;

      feuille.getRange(numero, 1, 1, 5).setValues([[
        idImport,
        valeursImport[1],
        valeursImport[2],
        montant,
        index + 1
      ]]);

      lignesUtilisees.push(numero);

      const celluleF = feuille.getRange(numero, 6);
      validationsOriginales[numero] = celluleF.getDataValidation();
      celluleF.clearDataValidations();
      celluleF.setNumberFormat('@').setValue(String(ligne.compte));

      feuille.getRange(numero, 7, 1, 2).setValues([
        [1, arrondirMontantMixte_(ligne.montant)]
      ]);

      feuille.getRange(numero, 9).setFormula(
        '=IF(OR($A' + numero + '="",$F' + numero + '=""),"",$G' +
        numero + '*$H' + numero + ')'
      );

      feuille.getRange(numero, 10).setValue(nomCompte);
      feuille.getRange(numero, 11).setValue(ligne.programme || '');
      feuille.getRange(numero, 12).setValue(ligne.projet || '');

      feuille.getRange(numero, 13).setFormula(
        '=IF($A' + numero + '="","",SUMIFS($I$6:$I$' +
        derniereLigneRepartition + ',$A$6:$A$' + derniereLigneRepartition +
        ',$A' + numero + ',$Q$6:$Q$' + derniereLigneRepartition +
        ',"<>Annulée*"))'
      );
      feuille.getRange(numero, 14).setFormula(
        '=IF($A' + numero + '="","",$D' + numero + '-$M' + numero + ')'
      );
      feuille.getRange(numero, 15).setFormula(
        '=IF($A' + numero + '="","",IF(ABS($N' + numero +
        ')<0.005,"Prêt",IF($N' + numero + '<0,"Dépassement","À compléter")))'
      );

      feuille.getRange(numero, 16).setValue('Revenu comptable direct');
      feuille.getRange(numero, 17).setValue(
        conserverHistorique ? noteRevision : ''
      );
    }
  } catch (erreur) {
    lignesUtilisees.forEach(function(numero) {
      feuille.getRange(numero, 1, 1, 17).clearContent();
      const validationOriginale = validationsOriginales[numero];
      if (validationOriginale) {
        feuille.getRange(numero, 6).setDataValidation(validationOriginale);
      }
    });
    throw erreur;
  }

  importBancaireModeMixte_(valeursImport[0]);
  SpreadsheetApp.flush();
}

function finaliserRevenuDirectMixteParId_(idImport, options) {
  const optionsFinalisation = options || {};
  const idGroupeForce = String(optionsFinalisation.idGroupeForce || '').trim();
  const idRevisionSource = String(optionsFinalisation.idRevisionSource || '').trim();
  const exclus = ['4000', '4010', '4020'];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const repartition = obtenirFeuilleMixte_(ss, 'Répartition');
  const importBancaire = obtenirFeuilleMixte_(ss, 'Import bancaire');
  const transactions = obtenirFeuilleMixte_(ss, 'Transactions');
  const journal = obtenirFeuilleMixte_(ss, 'Journal');
  const configuration = obtenirFeuilleMixte_(ss, 'Configuration');

  SpreadsheetApp.flush();

  const ligneImport = trouverLigneParValeurMixte_(
    importBancaire, 1, idImport, 6
  );

  if (!ligneImport) {
    throw new Error(
      'La transaction ' + idImport + ' est introuvable dans « Import bancaire ».'
    );
  }

  const valeursImport = importBancaire
    .getRange(ligneImport, 1, 1, 15)
    .getValues()[0];
  const idDejaLie = String(valeursImport[9] || '').trim();
  const statutImport = String(valeursImport[10] || '').trim();

  if (idDejaLie || statutImport === 'Classée') {
    throw new Error(
      'Cette transaction bancaire est déjà classée' +
      (idDejaLie ? ' sous ' + idDejaLie : '') +
      '.'
    );
  }

  const montantBancaire = arrondirMontantMixte_(Number(valeursImport[3] || 0));

  if (montantBancaire <= 0) {
    throw new Error(
      'La finalisation de revenu direct requiert un montant bancaire positif.'
    );
  }

  const idGroupeBase = creerIdGroupeMixte_(idImport);
  const forfaitsSheet = ss.getSheetByName('Forfaits');

  if (forfaitsSheet) {
    nettoyerTraitementPartielMixte_(
      transactions, journal, forfaitsSheet, idImport, idGroupeBase
    );
  }

  if (transactionBancaireDejaPresenteMixte_(transactions, idImport)) {
    throw new Error(
      'Une transaction comptable utilise déjà la référence bancaire ' +
      idImport + '.'
    );
  }

  const idGroupe = idGroupeForce ||
    creerIdGroupeDisponibleMixte_(transactions, idImport);

  if (idGroupeForce) {
    validerAbsenceConflitIdGroupeMixte_(transactions, idGroupe);
  }

  const lignesRepartition = lireLignesActivesRepartitionMixte_(
    repartition, idImport
  );

  if (lignesRepartition.length === 0) {
    throw new Error(
      'Aucune ligne de répartition n\'a été trouvée pour ' + idImport + '.'
    );
  }

  const lignesActives = lignesRepartition.filter(function(ligne) {
    return String(ligne.valeurs[5] || '').trim() !== '';
  });

  if (lignesActives.length === 0) {
    throw new Error('Aucun compte n\'a été sélectionné.');
  }

  const planComptable = chargerPlanComptableMixte_(configuration);
  const comptesActifsRevenu = chargerComptesActifsRevenuDirect_(configuration);

  if (!planComptable['1000']) {
    throw new Error(
      'Le compte bancaire 1000 est introuvable dans le plan comptable.'
    );
  }

  lignesActives.forEach(function(ligne) {
    const codeCompte = String(ligne.valeurs[5] || '').trim();
    const montantLigne = arrondirMontantMixte_(
      Number(ligne.valeurs[6] || 0) * Number(ligne.valeurs[7] || 0)
    );

    if (!codeCompte || montantLigne <= 0) {
      throw new Error(
        'La ligne ' + ligne.numero +
        ' contient un compte ou un montant invalide.'
      );
    }

    if (exclus.indexOf(codeCompte) !== -1) {
      throw new Error(
        'Le compte ' + codeCompte + ' est réservé au mode spécialisé.'
      );
    }

    if (!comptesActifsRevenu[codeCompte]) {
      throw new Error(
        'Le compte ' + codeCompte +
        ' n\'est pas un compte de revenu direct actif.'
      );
    }
  });

  const totalReparti = arrondirMontantMixte_(
    lignesActives.reduce(function(total, ligne) {
      return total + Number(ligne.valeurs[6] || 0) * Number(ligne.valeurs[7] || 0);
    }, 0)
  );

  if (Math.abs(montantBancaire - totalReparti) >= 0.005) {
    throw new Error(
      'Le total réparti (' +
      totalReparti.toFixed(2) +
      ' $) ne correspond pas au montant bancaire (' +
      montantBancaire.toFixed(2) +
      ' $).'
    );
  }

  const statutRepartition = String(
    lignesRepartition[0].valeurs[14] || ''
  ).trim();

  if (statutRepartition !== 'Prêt') {
    throw new Error(
      'Le statut de la répartition doit être « Prêt » avant l\'enregistrement.'
    );
  }

  const dateTransaction = valeursImport[1];
  const descriptionBancaire = String(valeursImport[2] || '').trim();
  const contact = extraireNomInteracMixte_(descriptionBancaire) || descriptionBancaire;
  const transactionsCreees = [];

  lignesActives.forEach(function(ligne, index) {
    const codeCompte = String(ligne.valeurs[5] || '').trim();
    const montantLigne = arrondirMontantMixte_(
      Number(ligne.valeurs[6] || 0) * Number(ligne.valeurs[7] || 0)
    );
    const programme = String(ligne.valeurs[10] || '').trim();
    const projet = String(ligne.valeurs[11] || '').trim();

    const idTransaction =
      idGroupe + '-' + String(index + 1).padStart(2, '0');

    const ligneTransaction = ecrireTransactionMixte_(
      transactions,
      idTransaction,
      dateTransaction,
      contact,
      descriptionBancaire,
      montantLigne,
      codeCompte,
      programme,
      projet,
      idImport
    );

    transactionsCreees.push({
      idTransaction: idTransaction,
      ligne: ligneTransaction
    });

    ecrirePaireJournalRevenuDirectMixte_(
      journal,
      planComptable,
      idTransaction,
      dateTransaction,
      {
        description: descriptionBancaire,
        montant: montantLigne,
        codeCompte: codeCompte,
        programme: programme,
        projet: projet
      }
    );
  });

  const ancienneNote = String(valeursImport[12] || '').trim();
  const noteRevision = idRevisionSource
    ? ' | Révision de ' + idRevisionSource + ' vers ' + idGroupe
    : '';
  const nouvelleNote =
    (ancienneNote ? ancienneNote + ' | ' : '') +
    'Transaction mixte enregistrée : ' + idGroupe + noteRevision;

  importBancaire.getRange(ligneImport, 8, 1, 2).clearContent();
  importBancaire.getRange(ligneImport, 10).setValue(idGroupe);
  importBancaire.getRange(ligneImport, 11).setValue('Classée');
  importBancaire.getRange(ligneImport, 13).setValue(nouvelleNote);
  importBancaire.getRange(ligneImport, 15).clearContent();
  importBancaire.getRange(ligneImport, 16, 1, 4).clearDataValidations();

  if (idRevisionSource) {
    transactionsCreees.forEach(function(item) {
      transactions
        .getRange(item.ligne, 15)
        .setNote('Révision créée à partir de ' + idRevisionSource);
    });
  }

  SpreadsheetApp.flush();

  return { idGroupe: idGroupe };
}

function ecrirePaireJournalRevenuDirectMixte_(
  journal,
  planComptable,
  idTransaction,
  dateTransaction,
  ecriture
) {
  const montant = arrondirMontantMixte_(ecriture.montant);

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-D',
    idTransaction,
    dateTransaction,
    '1000',
    montant,
    0,
    ecriture.programme,
    ecriture.projet,
    ecriture.description
  );

  ecrireLigneJournalMixte_(
    journal,
    planComptable,
    'ECR-' + idTransaction + '-C',
    idTransaction,
    dateTransaction,
    ecriture.codeCompte,
    0,
    montant,
    ecriture.programme,
    ecriture.projet,
    ecriture.description
  );
}
