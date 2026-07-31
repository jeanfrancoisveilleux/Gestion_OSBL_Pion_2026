const CONFIG_PIECES_JUSTIFICATIVES = {
  proprieteIdClasseur: 'PIECES_JUSTIFICATIVES_ID_CLASSEUR',
  sousDossierReçus: '01 - Reçus et factures',
  premiereLigne: 6,
  tailleMaxOctets: 15 * 1024 * 1024,
  typesDocuments: [
    'Reçu',
    'Facture',
    'Contrat',
    'Subvention',
    'Relevé bancaire',
    'Autre'
  ],
  entetes: {
    Documents: [
      'ID document',
      'Date',
      'Type',
      'Nom',
      'Contact',
      'ID transaction',
      'ID projet',
      'Lien Drive',
      'Échéance',
      'Statut',
      'Notes'
    ],
    Transactions: [
      'ID transaction',
      'Date',
      'Type',
      'Contact',
      'Description',
      'Montant',
      'Code compte',
      'Nom du compte',
      'Programme',
      'Projet',
      'Compte de paiement',
      'Lien Drive',
      'Source',
      'Référence bancaire',
      'Statut',
      'Mois',
      'Contrôle pièce',
      'ID fournisseur',
      'Fournisseur',
      'ID contact'
    ],
    Journal: [
      'ID écriture',
      'ID transaction',
      'Date',
      'Code compte',
      'Nom du compte',
      'Type compte',
      'Débit',
      'Crédit',
      'Programme',
      'Projet',
      'Description',
      'Lien Drive',
      'Verrouillée',
      'Mois',
      'ID fournisseur',
      'Fournisseur'
    ]
  }
};

/**
 * Installe la capture de pièces justificatives sans créer de déploiement Web.
 * Le déploiement doit être créé manuellement dans Apps Script.
 */
function installerGestionPiecesJustificatives() {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est en cours. Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        'Ouvrez le classeur Google Sheets avant de lancer cet installateur.'
      );
    }

    memoriserClasseurPieces_(ss);
    verifierStructurePieces_(ss);
    obtenirDossierPrincipalPieces_(ss);

    const urlApplication = String(
      ScriptApp.getService().getUrl() || ''
    ).trim();
    const lienInstalle = urlApplication
      ? installerLienCapturePieces_(ss, urlApplication)
      : false;

    let message =
      'Gestion des pièces justificatives installée.\n\n' +
      'Le formulaire de bureau est disponible dans le menu « Gestion OSBL > ' +
      'Pièces justificatives ».\n';

    if (urlApplication) {
      message +=
        '\nApplication mobile détectée :\n' + urlApplication + '\n\n' +
        (lienInstalle
          ? 'Le lien mobile a été ajouté dans Documents!A3.'
          : 'Documents!A3 contient déjà une autre valeur; aucun lien n’a été remplacé.');
    } else {
      message +=
        '\nPour utiliser la caméra du téléphone, déployez ensuite le projet comme ' +
        'application Web, avec un accès limité à votre compte, puis relancez cet installateur.';
    }

    SpreadsheetApp.getUi().alert(message);

    return {
      succes: true,
      urlApplication: urlApplication,
      lienInstalle: lienInstalle
    };
  } finally {
    verrou.releaseLock();
  }
}

/**
 * Ouvre le formulaire dans Google Sheets pour les essais sur ordinateur.
 */
function ouvrirCapturePieceJustificative() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Le classeur Google Sheets est introuvable.');
  }

  memoriserClasseurPieces_(ss);
  verifierStructurePieces_(ss);

  const html = HtmlService
    .createHtmlOutputFromFile('CapturePiece')
    .setWidth(760)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Ajouter une pièce justificative'
  );
}

/**
 * Point d’entrée de l’application Web mobile.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('CapturePiece')
    .setTitle('Pièces justificatives — Gestion OSBL')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Retourne les transactions validées disponibles dans le formulaire.
 */
function obtenirDonneesCapturePiece() {
  const ss = obtenirClasseurPieces_();
  verifierStructurePieces_(ss);

  return {
    transactions: chargerTransactionsCapturePieces_(ss),
    tailleMaxMo: Math.round(
      CONFIG_PIECES_JUSTIFICATIVES.tailleMaxOctets / 1024 / 1024
    ),
    typesDocuments: CONFIG_PIECES_JUSTIFICATIVES.typesDocuments.slice(),
    urlApplication: String(ScriptApp.getService().getUrl() || '').trim()
  };
}

/**
 * Reçoit le formulaire HTML. Le formulaire doit être l’unique paramètre de
 * google.script.run afin que les champs de fichier soient transmis comme Blob.
 */
function enregistrerPieceJustificative(formulaire) {
  const verrou = LockService.getScriptLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre pièce est en cours d’enregistrement. Réessayez dans quelques secondes.'
    );
  }

  let fichierCree = null;
  let feuilleDocuments = null;
  let ligneDocument = 0;
  let transactionModifiee = null;
  let journalModifie = [];

  try {
    const ss = obtenirClasseurPieces_();
    verifierStructurePieces_(ss);

    const idTransaction = String(
      formulaire && formulaire.idTransaction || ''
    ).trim();
    const typeDocument = String(
      formulaire && formulaire.typeDocument || ''
    ).trim();
    const nomSaisi = String(
      formulaire && formulaire.nomDocument || ''
    ).trim();
    const notesSaisies = String(
      formulaire && formulaire.notes || ''
    ).trim();

    if (!idTransaction) {
      throw new Error('Sélectionnez une transaction.');
    }

    if (
      CONFIG_PIECES_JUSTIFICATIVES.typesDocuments.indexOf(typeDocument) === -1
    ) {
      throw new Error('Le type de document sélectionné est invalide.');
    }

    const blob = extraireBlobPieces_(formulaire);
    const fichierValide = validerBlobPieces_(blob);
    const contexte = trouverContexteTransactionPieces_(ss, idTransaction);

    if (contexte.statut !== 'Validée') {
      throw new Error(
        'La transaction ' + idTransaction +
        ' n’est plus validée. Actualisez le formulaire avant de continuer.'
      );
    }

    const lignesJournal = trouverLignesJournalPieces_(ss, idTransaction);

    if (lignesJournal.length === 0) {
      throw new Error(
        'Aucune écriture active du Journal ne correspond à la transaction ' +
        idTransaction + '. Aucun fichier n’a été enregistré.'
      );
    }

    feuilleDocuments = ss.getSheetByName('Documents');
    const idDocument = genererIdDocumentPieces_(
      feuilleDocuments,
      contexte.date,
      ss.getSpreadsheetTimeZone()
    );
    const dossierTransaction = obtenirDossierTransactionPieces_(
      ss,
      contexte,
      idTransaction
    );
    const nomFichier = creerNomFichierPieces_(
      idDocument,
      typeDocument,
      contexte,
      fichierValide.extension
    );

    fichierCree = dossierTransaction.createFile(blob.setName(nomFichier));
    fichierCree.setDescription(
      'Pièce justificative liée à la transaction ' + idTransaction + '.'
    );

    ligneDocument = prochaineLigneDocumentPieces_(feuilleDocuments);
    preparerLigneDocumentPieces_(feuilleDocuments, ligneDocument);

    const idProjet = obtenirIdProjetPieces_(ss, contexte.projet);
    const auteur = String(Session.getActiveUser().getEmail() || '').trim();
    const horodatage = Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd HH:mm'
    );
    const noteAudit =
      'Ajoutée via la capture de pièces le ' + horodatage +
      (auteur ? ' par ' + auteur : '');
    const notesFinales = notesSaisies
      ? notesSaisies + ' | ' + noteAudit
      : noteAudit;
    const nomDocument = nomSaisi ||
      typeDocument + ' — ' + contexte.description;

    feuilleDocuments.getRange(ligneDocument, 1).setNumberFormat('@');
    feuilleDocuments.getRange(ligneDocument, 6, 1, 2).setNumberFormat('@');
    feuilleDocuments.getRange(ligneDocument, 1, 1, 11).setValues([[
      idDocument,
      contexte.date,
      typeDocument,
      nomDocument,
      contexte.contact || contexte.fournisseur || '',
      idTransaction,
      idProjet,
      fichierCree.getUrl(),
      '',
      'Actif',
      notesFinales
    ]]);

    const feuilleTransactions = ss.getSheetByName('Transactions');
    const celluleLienTransaction = feuilleTransactions.getRange(
      contexte.numeroLigne,
      12
    );
    const ancienLienTransaction = String(
      celluleLienTransaction.getValue() || ''
    ).trim();
    const lienTransaction = ancienLienTransaction || dossierTransaction.getUrl();

    transactionModifiee = {
      cellule: celluleLienTransaction,
      ancienneValeur: ancienLienTransaction,
      modifiee: !ancienLienTransaction
    };

    if (!ancienLienTransaction) {
      celluleLienTransaction.setValue(lienTransaction);
    }

    const feuilleJournal = ss.getSheetByName('Journal');
    journalModifie = lignesJournal.map(function(numeroLigne) {
      const cellule = feuilleJournal.getRange(numeroLigne, 12);
      const ancienneValeur = String(cellule.getValue() || '').trim();

      if (!ancienneValeur) {
        cellule.setValue(lienTransaction);
      }

      return {
        cellule: cellule,
        ancienneValeur: ancienneValeur,
        modifiee: !ancienneValeur
      };
    });

    SpreadsheetApp.flush();

    const controlePiece = String(
      feuilleTransactions.getRange(contexte.numeroLigne, 17).getDisplayValue() || ''
    ).trim();

    if (controlePiece !== 'OK') {
      throw new Error(
        'La pièce a été écrite, mais le contrôle de la transaction n’est pas passé à « OK ».'
      );
    }

    return {
      succes: true,
      idDocument: idDocument,
      idTransaction: idTransaction,
      nomFichier: nomFichier,
      urlFichier: fichierCree.getUrl(),
      urlDossier: dossierTransaction.getUrl(),
      controlePiece: controlePiece
    };
  } catch (erreur) {
    journalModifie.forEach(function(element) {
      if (element.modifiee) {
        element.cellule.setValue(element.ancienneValeur);
      }
    });

    if (transactionModifiee && transactionModifiee.modifiee) {
      transactionModifiee.cellule.setValue(
        transactionModifiee.ancienneValeur
      );
    }

    if (feuilleDocuments && ligneDocument) {
      feuilleDocuments.getRange(ligneDocument, 1, 1, 11).clearContent();
    }

    if (fichierCree) {
      try {
        fichierCree.setTrashed(true);
      } catch (erreurCorbeille) {
        console.error(
          'Impossible de placer le fichier partiel dans la corbeille :',
          erreurCorbeille.message
        );
      }
    }

    SpreadsheetApp.flush();
    throw erreur;
  } finally {
    verrou.releaseLock();
  }
}

function memoriserClasseurPieces_(ss) {
  PropertiesService.getScriptProperties().setProperty(
    CONFIG_PIECES_JUSTIFICATIVES.proprieteIdClasseur,
    ss.getId()
  );
}

function obtenirClasseurPieces_() {
  const actif = SpreadsheetApp.getActiveSpreadsheet();

  if (actif) {
    memoriserClasseurPieces_(actif);
    return actif;
  }

  const idClasseur = String(
    PropertiesService.getScriptProperties().getProperty(
      CONFIG_PIECES_JUSTIFICATIVES.proprieteIdClasseur
    ) || ''
  ).trim();

  if (!idClasseur) {
    throw new Error(
      'Le classeur n’est pas configuré pour l’application mobile. ' +
      'Exécutez d’abord « Installer la capture mobile » depuis Google Sheets.'
    );
  }

  return SpreadsheetApp.openById(idClasseur);
}

function verifierStructurePieces_(ss) {
  const noms = ['Documents', 'Transactions', 'Journal'];

  noms.forEach(function(nom) {
    const feuille = ss.getSheetByName(nom);
    const attendus = CONFIG_PIECES_JUSTIFICATIVES.entetes[nom];

    if (!feuille) {
      throw new Error('L’onglet « ' + nom + ' » est introuvable.');
    }

    if (feuille.getMaxColumns() < attendus.length) {
      throw new Error(
        'L’onglet « ' + nom + ' » ne contient pas toutes les colonnes requises.'
      );
    }

    const presents = feuille
      .getRange(5, 1, 1, attendus.length)
      .getDisplayValues()[0]
      .map(function(valeur) {
        return String(valeur || '').trim();
      });

    attendus.forEach(function(attendu, index) {
      if (presents[index] !== attendu) {
        throw new Error(
          'En-tête incompatible dans ' + nom + '!' +
          convertirColonnePieces_(index + 1) + '5 : « ' +
          presents[index] + ' » au lieu de « ' + attendu + ' ».'
        );
      }
    });
  });

  const transactions = ss.getSheetByName('Transactions');
  const formuleControle = String(
    transactions.getRange(6, 17).getFormula() || ''
  );

  if (!formuleControle || formuleControle.indexOf('$L6') === -1) {
    throw new Error(
      'La formule de contrôle des pièces est absente ou incompatible dans Transactions!Q6.'
    );
  }
}

function convertirColonnePieces_(numero) {
  let valeur = numero;
  let resultat = '';

  while (valeur > 0) {
    const reste = (valeur - 1) % 26;
    resultat = String.fromCharCode(65 + reste) + resultat;
    valeur = Math.floor((valeur - 1) / 26);
  }

  return resultat;
}

function installerLienCapturePieces_(ss, urlApplication) {
  const feuille = ss.getSheetByName('Documents');
  const cellule = feuille.getRange('A3');
  const texteActuel = String(cellule.getDisplayValue() || '').trim();
  const libelle = '📷 Ajouter une pièce justificative';

  if (texteActuel && texteActuel !== libelle) {
    return false;
  }

  const style = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setForegroundColor('#1155cc')
    .setUnderline(true)
    .build();
  const valeur = SpreadsheetApp.newRichTextValue()
    .setText(libelle)
    .setLinkUrl(urlApplication)
    .setTextStyle(0, libelle.length, style)
    .build();

  cellule.setRichTextValue(valeur);
  cellule.setNote(
    'Ouvre l’application mobile de capture de reçus et factures.'
  );
  return true;
}

function chargerTransactionsCapturePieces_(ss) {
  const feuille = ss.getSheetByName('Transactions');
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < CONFIG_PIECES_JUSTIFICATIVES.premiereLigne) {
    return [];
  }

  const valeurs = feuille.getRange(
    CONFIG_PIECES_JUSTIFICATIVES.premiereLigne,
    1,
    derniereLigne - CONFIG_PIECES_JUSTIFICATIVES.premiereLigne + 1,
    20
  ).getValues();
  const fuseau = ss.getSpreadsheetTimeZone();

  return valeurs
    .map(function(ligne) {
      const id = String(ligne[0] || '').trim();
      const type = String(ligne[2] || '').trim();
      const statut = String(ligne[14] || '').trim();

      if (!id || statut !== 'Validée') {
        return null;
      }

      const date = normaliserDatePieces_(ligne[1]);
      const lien = String(ligne[11] || '').trim();
      const controle = String(ligne[16] || '').trim();

      return {
        id: id,
        date: Utilities.formatDate(date, fuseau, 'yyyy-MM-dd'),
        type: type,
        contact: String(ligne[3] || '').trim(),
        description: String(ligne[4] || '').trim(),
        montant: Number(ligne[5] || 0),
        programme: String(ligne[8] || '').trim(),
        projet: String(ligne[9] || '').trim(),
        source: String(ligne[12] || '').trim(),
        referenceBancaire: String(ligne[13] || '').trim(),
        controlePiece: controle,
        pieceRequise: type === 'Dépense' && (!lien || controle === 'Pièce requise'),
        lienExistant: lien,
        fournisseur: String(ligne[18] || '').trim()
      };
    })
    .filter(function(element) {
      return element !== null;
    })
    .sort(function(a, b) {
      if (a.pieceRequise !== b.pieceRequise) {
        return a.pieceRequise ? -1 : 1;
      }

      if (a.date !== b.date) {
        return a.date < b.date ? 1 : -1;
      }

      return a.id.localeCompare(b.id);
    });
}

function trouverContexteTransactionPieces_(ss, idTransaction) {
  const feuille = ss.getSheetByName('Transactions');
  const resultat = feuille
    .getRange(
      CONFIG_PIECES_JUSTIFICATIVES.premiereLigne,
      1,
      feuille.getMaxRows() - CONFIG_PIECES_JUSTIFICATIVES.premiereLigne + 1,
      1
    )
    .createTextFinder(idTransaction)
    .matchEntireCell(true)
    .findNext();

  if (!resultat) {
    throw new Error(
      'La transaction ' + idTransaction + ' est introuvable.'
    );
  }

  const numeroLigne = resultat.getRow();
  const valeurs = feuille.getRange(numeroLigne, 1, 1, 20).getValues()[0];
  const idTrouve = String(valeurs[0] || '').trim();

  if (idTrouve !== idTransaction) {
    throw new Error('La transaction sélectionnée a changé. Actualisez le formulaire.');
  }

  return {
    numeroLigne: numeroLigne,
    id: idTrouve,
    date: normaliserDatePieces_(valeurs[1]),
    type: String(valeurs[2] || '').trim(),
    contact: String(valeurs[3] || '').trim(),
    description: String(valeurs[4] || '').trim(),
    montant: Number(valeurs[5] || 0),
    programme: String(valeurs[8] || '').trim(),
    projet: String(valeurs[9] || '').trim(),
    lienDrive: String(valeurs[11] || '').trim(),
    referenceBancaire: String(valeurs[13] || '').trim(),
    statut: String(valeurs[14] || '').trim(),
    controlePiece: String(valeurs[16] || '').trim(),
    idFournisseur: String(valeurs[17] || '').trim(),
    fournisseur: String(valeurs[18] || '').trim(),
    idContact: String(valeurs[19] || '').trim()
  };
}

function trouverLignesJournalPieces_(ss, idTransaction) {
  const feuille = ss.getSheetByName('Journal');
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < CONFIG_PIECES_JUSTIFICATIVES.premiereLigne) {
    return [];
  }

  const valeurs = feuille.getRange(
    CONFIG_PIECES_JUSTIFICATIVES.premiereLigne,
    1,
    derniereLigne - CONFIG_PIECES_JUSTIFICATIVES.premiereLigne + 1,
    2
  ).getDisplayValues();

  return valeurs.reduce(function(resultat, ligne, index) {
    const idEcriture = String(ligne[0] || '').trim();
    const idLie = String(ligne[1] || '').trim();

    if (
      idLie === idTransaction &&
      idEcriture &&
      idEcriture.indexOf('ANN-') !== 0
    ) {
      resultat.push(CONFIG_PIECES_JUSTIFICATIVES.premiereLigne + index);
    }

    return resultat;
  }, []);
}

function normaliserDatePieces_(valeur) {
  const date = valeur instanceof Date
    ? new Date(valeur.getTime())
    : new Date(valeur);

  if (isNaN(date.getTime())) {
    throw new Error('La date de la transaction est invalide.');
  }

  return date;
}

function extraireBlobPieces_(formulaire) {
  const candidats = [
    formulaire ? formulaire.photo : null,
    formulaire ? formulaire.fichier : null
  ].filter(function(element) {
    return element && typeof element.getBytes === 'function' &&
      element.getBytes().length > 0;
  });

  if (candidats.length === 0) {
    throw new Error('Prenez une photo ou sélectionnez un fichier.');
  }

  if (candidats.length > 1) {
    throw new Error('Sélectionnez une seule photo ou un seul fichier à la fois.');
  }

  return candidats[0];
}

function validerBlobPieces_(blob) {
  const taille = blob.getBytes().length;
  const nomOriginal = String(blob.getName() || '').trim();
  const typeMime = String(blob.getContentType() || '').toLowerCase().trim();
  const correspondanceExtension = nomOriginal.match(/\.([a-zA-Z0-9]+)$/);
  const extensionOriginale = correspondanceExtension
    ? correspondanceExtension[1].toLowerCase()
    : '';
  const extensionsAcceptees = [
    'jpg', 'jpeg', 'png', 'heic', 'heif', 'webp', 'pdf'
  ];
  const extensionsParMime = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/heic': 'heic',
    'image/heif': 'heif',
    'image/webp': 'webp',
    'application/pdf': 'pdf'
  };
  const mimeGenerique = !typeMime || typeMime === 'application/octet-stream';
  const extensionMime = extensionsParMime[typeMime] || '';
  const extensionAcceptee =
    extensionsAcceptees.indexOf(extensionOriginale) !== -1;

  if (!taille) {
    throw new Error('Le fichier sélectionné est vide.');
  }

  if (taille > CONFIG_PIECES_JUSTIFICATIVES.tailleMaxOctets) {
    throw new Error(
      'Le fichier dépasse la limite de ' +
      Math.round(CONFIG_PIECES_JUSTIFICATIVES.tailleMaxOctets / 1024 / 1024) +
      ' Mo.'
    );
  }

  if (!extensionMime && !mimeGenerique) {
    throw new Error(
      'Format non accepté. Utilisez JPG, PNG, HEIC, WEBP ou PDF.'
    );
  }

  if (mimeGenerique && !extensionAcceptee) {
    throw new Error(
      'Format non accepté. Utilisez JPG, PNG, HEIC, WEBP ou PDF.'
    );
  }

  const extensionFinale = extensionMime || extensionOriginale;

  return {
    taille: taille,
    typeMime: typeMime,
    extension: extensionFinale === 'jpeg' ? 'jpg' : extensionFinale
  };
}

function obtenirDossierPrincipalPieces_(ss) {
  const configuration = ss.getSheetByName('Configuration');

  if (!configuration) {
    throw new Error('L’onglet Configuration est introuvable.');
  }

  const ligne = trouverLigneConfiguration_(
    configuration,
    'Dossier Drive principal'
  );
  const dossier = obtenirDossierConfigure_(
    configuration.getRange(ligne, 12).getValue()
  );

  if (!dossier) {
    throw new Error(
      'Le dossier Drive principal n’est pas configuré. Exécutez d’abord ' +
      '« Initialiser les dossiers Drive ».'
    );
  }

  return dossier;
}

function obtenirDossierTransactionPieces_(ss, contexte, idTransaction) {
  const mois = [
    'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
    'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'
  ];
  const dossierPrincipal = obtenirDossierPrincipalPieces_(ss);
  const dossierReçus = trouverOuCreerSousDossier_(
    dossierPrincipal,
    CONFIG_PIECES_JUSTIFICATIVES.sousDossierReçus
  );
  const fuseau = ss.getSpreadsheetTimeZone();
  const annee = Utilities.formatDate(contexte.date, fuseau, 'yyyy');
  const moisNumero = Number(Utilities.formatDate(contexte.date, fuseau, 'M'));
  const nomMois = String(moisNumero).padStart(2, '0') +
    ' - ' + mois[moisNumero - 1];
  const dossierAnnee = trouverOuCreerSousDossier_(dossierReçus, annee);
  const dossierMois = trouverOuCreerSousDossier_(dossierAnnee, nomMois);
  const fournisseur = nettoyerNomFichierPieces_(
    contexte.fournisseur || contexte.contact || 'Sans fournisseur'
  );
  const nomTransaction = nettoyerNomFichierPieces_(
    idTransaction + ' - ' + fournisseur
  );

  return trouverOuCreerSousDossier_(dossierMois, nomTransaction);
}

function genererIdDocumentPieces_(feuille, date, fuseau) {
  const annee = Utilities.formatDate(date, fuseau, 'yyyy');
  const prefixe = 'DOC-' + annee + '-';
  const derniereLigne = feuille.getLastRow();
  let maximum = 0;

  if (derniereLigne >= CONFIG_PIECES_JUSTIFICATIVES.premiereLigne) {
    feuille.getRange(
      CONFIG_PIECES_JUSTIFICATIVES.premiereLigne,
      1,
      derniereLigne - CONFIG_PIECES_JUSTIFICATIVES.premiereLigne + 1,
      1
    ).getDisplayValues().forEach(function(ligne) {
      const id = String(ligne[0] || '').trim();
      const match = id.match(new RegExp('^' + prefixe + '(\\d+)$'));

      if (match) {
        maximum = Math.max(maximum, Number(match[1] || 0));
      }
    });
  }

  return prefixe + String(maximum + 1).padStart(4, '0');
}

function prochaineLigneDocumentPieces_(feuille) {
  const premiere = CONFIG_PIECES_JUSTIFICATIVES.premiereLigne;
  const nombre = feuille.getMaxRows() - premiere + 1;
  const ids = feuille.getRange(premiere, 1, nombre, 1).getDisplayValues();

  for (let index = 0; index < ids.length; index += 1) {
    if (String(ids[index][0] || '').trim() === '') {
      return premiere + index;
    }
  }

  const ancienneLimite = feuille.getMaxRows();
  feuille.insertRowsAfter(ancienneLimite, 100);
  return ancienneLimite + 1;
}

function preparerLigneDocumentPieces_(feuille, numeroLigne) {
  const modele = feuille.getRange(6, 1, 1, 11);
  const cible = feuille.getRange(numeroLigne, 1, 1, 11);

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
}

function obtenirIdProjetPieces_(ss, valeurProjet) {
  const projet = String(valeurProjet || '').trim();

  if (!projet) {
    return '';
  }

  const feuille = ss.getSheetByName('Projets');

  if (!feuille || feuille.getLastRow() < 6) {
    return projet;
  }

  const valeurs = feuille.getRange(
    6,
    1,
    feuille.getLastRow() - 5,
    2
  ).getDisplayValues();

  for (let index = 0; index < valeurs.length; index += 1) {
    const id = String(valeurs[index][0] || '').trim();
    const nom = String(valeurs[index][1] || '').trim();

    if (projet === id || projet === nom) {
      return id;
    }
  }

  return projet;
}

function creerNomFichierPieces_(idDocument, typeDocument, contexte, extension) {
  const fournisseur = nettoyerNomFichierPieces_(
    contexte.fournisseur || contexte.contact || 'Sans fournisseur'
  );
  const montant = Math.abs(Number(contexte.montant || 0)).toFixed(2);
  const base = nettoyerNomFichierPieces_(
    idDocument + ' - ' + typeDocument + ' - ' + fournisseur + ' - ' + montant
  );

  return base.slice(0, 180) + '.' + extension;
}

function nettoyerNomFichierPieces_(valeur) {
  return String(valeur || '')
    .replace(/[\\/:*?"<>|#%{}\[\]]/g, '-')
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .trim() || 'Document';
}
