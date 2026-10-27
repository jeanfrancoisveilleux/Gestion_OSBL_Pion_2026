const CONFIG_OSBL = {
  dossierPrincipal: 'GESTION_BUDGETAIRE_OSBL',
  sousDossiers: [
    'REÇUS_FACTURES',
    'CONTRATS',
    'RELEVÉS_BANCAIRES',
    'PROJETS',
    'SUBVENTIONS',
    'PARTENAIRES',
    'RAPPORTS',
    'DOCUMENTS_DE_FONCTIONNEMENT'
  ]
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('CaptureRecuV2')
    .setTitle('Nouveau reçu — Gestion OSBL')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Gestion OSBL')
    .addItem(
      'Saisir une transaction',
      'afficherSaisieTransaction'
    )
    .addItem(
      'Importer un relevé Desjardins',
      'afficherImportBancaire'
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Reçus')
        .addItem(
          'Nouveau reçu',
          'ouvrirCaptureRecuV2'
        )
        .addItem(
          'Relire OCR du reçu sélectionné',
          'relireOcrRecuSelectionneV2'
        )
        .addItem(
          'Installer le raccourci mobile V2',
          'installerCaptureRecuV2'
        )
    )
    .addSubMenu(
      ui.createMenu('Règles bancaires')
        .addItem(
          'Valider les règles',
          'validerReglesBancairesConfigurables'
        )
        .addItem(
          'Réappliquer aux transactions non traitées',
          'reappliquerReglesBancairesAuxTransactionsAClasser'
        )
    )
    .addSeparator()
    .addItem(
      'Initialiser les dossiers Drive',
      'initialiserDossiersDrive'
    )
    .addItem(
      'Vérifier le système V2',
      'verifierSysteme'
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Maintenance')
        .addItem(
          'Nettoyer les anciens déclencheurs',
          'nettoyerAnciensDeclencheursV2'
        )
    )
    .addToUi();
}

function afficherSaisieTransaction() {
  const interface = HtmlService
    .createHtmlOutputFromFile('Transaction')
    .setTitle('Nouvelle transaction');

  SpreadsheetApp.getUi().showSidebar(interface);
}

function obtenirOptionsTransaction() {
  const ss = SpreadsheetApp.getActive();
  const comptes = ss.getSheetByName('Comptes');
  const data = ss.getSheetByName('DATA');

  if (!comptes || !data) {
    throw new Error(
      'Les onglets Comptes ou DATA sont introuvables.'
    );
  }

  const listeComptes = comptes
    .getRange(6, 1, Math.max(comptes.getLastRow() - 5, 1), 5)
    .getDisplayValues()
    .filter(function(ligne) {
      return String(ligne[0] || '').trim() &&
        String(ligne[4] || '').trim() === 'Oui';
    })
    .map(function(ligne) {
      return {
        code: String(ligne[0] || '').trim(),
        nom: String(ligne[1] || '').trim(),
        type: String(ligne[2] || '').trim()
      };
    });

  const listeTiers = data
    .getRange(6, 1, Math.max(data.getLastRow() - 5, 1), 4)
    .getDisplayValues()
    .filter(function(ligne) {
      return String(ligne[0] || '').trim() &&
        String(ligne[1] || '').trim() &&
        String(ligne[3] || '').trim() === 'Oui';
    })
    .map(function(ligne) {
      return {
        id: String(ligne[0] || '').trim(),
        nom: String(ligne[1] || '').trim(),
        type: String(ligne[2] || '').trim()
      };
    })
    .sort(function(a, b) {
      return a.nom.localeCompare(b.nom, 'fr');
    });

  const listeProgrammes = data
    .getRange(6, 7, Math.max(data.getLastRow() - 5, 1), 3)
    .getDisplayValues()
    .filter(function(ligne) {
      return String(ligne[1] || '').trim() &&
        String(ligne[2] || '').trim() === 'Oui';
    })
    .map(function(ligne) {
      return {
        code: String(ligne[0] || '').trim(),
        nom: String(ligne[1] || '').trim()
      };
    });

  return {
    date: Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd'
    ),
    comptes: listeComptes,
    tiers: listeTiers,
    programmes: listeProgrammes
  };
}

function ajouterTiersTransaction(nom, type) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const data = ss.getSheetByName('DATA');
    const fournisseurs = ss.getSheetByName('Fournisseurs');

    if (!data) {
      throw new Error('L’onglet DATA est introuvable.');
    }

    const nomPropre = String(nom || '')
      .replace(/\s+/g, ' ')
      .trim();

    const typePropre = String(type || '').trim();

    if (!nomPropre) {
      throw new Error('Le nom du client ou fournisseur est obligatoire.');
    }

    if (!['Fournisseur', 'Client', 'Autre'].includes(typePropre)) {
      throw new Error('Le type de tiers est invalide.');
    }

    const normaliser = function(valeur) {
      return String(valeur || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    };

    const derniereLigne = Math.max(data.getLastRow(), 6);
    const existants = data
      .getRange(6, 1, derniereLigne - 5, 4)
      .getDisplayValues();

    const deja = existants.find(function(ligne) {
      return normaliser(ligne[1]) === normaliser(nomPropre);
    });

    if (deja) {
      return {
        id: String(deja[0] || '').trim(),
        nom: String(deja[1] || '').trim(),
        type: String(deja[2] || '').trim(),
        existait: true
      };
    }

    const prefixe =
      typePropre === 'Fournisseur'
        ? 'FOU'
        : typePropre === 'Client'
          ? 'CLI'
          : 'AUT';

    let numeroMax = 0;

    existants.forEach(function(ligne) {
      const id = String(ligne[0] || '').trim();
      const match = id.match(new RegExp('^' + prefixe + '-(\\d{4})$'));

      if (match) {
        numeroMax = Math.max(numeroMax, Number(match[1]));
      }
    });

    if (typePropre === 'Fournisseur' && fournisseurs) {
      const idsFournisseurs = fournisseurs
        .getRange(
          6,
          1,
          Math.max(fournisseurs.getLastRow() - 5, 1),
          1
        )
        .getDisplayValues()
        .flat();

      idsFournisseurs.forEach(function(id) {
        const match = String(id || '').match(/^FOU-(\d{4})$/);

        if (match) {
          numeroMax = Math.max(numeroMax, Number(match[1]));
        }
      });
    }

    const id =
      prefixe + '-' +
      String(numeroMax + 1).padStart(4, '0');

    const ligneData = prochaineLigneLibre_(data, 1, 6);

    data.getRange(ligneData, 1, 1, 5).setValues([[
      id,
      nomPropre,
      typePropre,
      'Oui',
      typePropre === 'Fournisseur'
        ? 'Fournisseurs'
        : 'Saisie transaction'
    ]]);

    if (typePropre === 'Fournisseur' && fournisseurs) {
      const ligneFournisseur =
        prochaineLigneLibre_(fournisseurs, 1, 6);

      fournisseurs.getRange(
        ligneFournisseur,
        1,
        1,
        12
      ).setValues([[
        id,
        nomPropre,
        nomPropre,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'Oui',
        'Créé depuis la saisie de transaction'
      ]]);
    }

    SpreadsheetApp.flush();

    return {
      id: id,
      nom: nomPropre,
      type: typePropre,
      existait: false
    };
  } finally {
    verrou.releaseLock();
  }
}

function enregistrerTransaction(donnees) {
  const verrou = LockService.getDocumentLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }

  let ligneTransaction = null;
  let ligneVentilation = null;

  try {
    const ss = SpreadsheetApp.getActive();
    const manuelles = ss.getSheetByName('Transactions manuelles');
    const ventilation = ss.getSheetByName('Ventilation');

    if (!manuelles || !ventilation) {
      throw new Error(
        'Transactions manuelles ou Ventilation est introuvable.'
      );
    }

    const type = String(donnees.type || '').trim();
    const description = String(
      donnees.description || ''
    ).trim();
    const codeCompte = String(
      donnees.compte || ''
    ).trim();
    const programme = String(
      donnees.programme || ''
    ).trim();
    const idTiers = String(
      donnees.idTiers || ''
    ).trim();
    const lienDocument = String(
      donnees.document || ''
    ).trim();
    const modePaiement = String(
      donnees.modePaiement || ''
    ).trim();

    // Un seul compte bancaire : tout passe par 1000,
    // sauf les paiements comptants qui passent par la petite caisse 1100.
    const compteTresorerie =
      modePaiement === 'Comptant'
        ? '1100'
        : '1000';

    const montant = Number(
      String(donnees.montant || '')
        .replace(/\s/g, '')
        .replace(',', '.')
        .replace('$', '')
    );

    if (!['Revenu', 'Dépense', 'Transfert'].includes(type)) {
      throw new Error('Le type de transaction est invalide.');
    }

    if (
      !donnees.date ||
      !description ||
      !codeCompte
    ) {
      throw new Error(
        'La date, la description et le compte comptable sont obligatoires.'
      );
    }

    if (!montant || montant <= 0) {
      throw new Error(
        'Le montant doit être supérieur à zéro.'
      );
    }

    if (codeCompte === compteTresorerie) {
      throw new Error(
        'Le compte comptable choisi ne peut pas être le même que le compte de trésorerie utilisé automatiquement.'
      );
    }

    verifierCompteV2_(ss, codeCompte);
    verifierCompteV2_(ss, compteTresorerie);

    const date = new Date(
      String(donnees.date) + 'T12:00:00'
    );

    if (isNaN(date.getTime())) {
      throw new Error('La date est invalide.');
    }

    const identifiant =
      'T-MAN-' +
      Utilities.formatDate(
        new Date(),
        ss.getSpreadsheetTimeZone(),
        'yyyyMMdd-HHmmss'
      ) +
      '-' +
      Utilities.getUuid()
        .slice(0, 4)
        .toUpperCase();

    ligneTransaction = prochaineLigneLibre_(
      manuelles,
      2,
      6
    );

    manuelles
      .getRange(ligneTransaction, 1, 1, 7)
      .setValues([[
        identifiant,
        date,
        type,
        description,
        montant,
        compteTresorerie,
        idTiers
      ]]);

    if (lienDocument) {
      manuelles
        .getRange(ligneTransaction, 10)
        .setValue(lienDocument);
    }

    manuelles
      .getRange(ligneTransaction, 11)
      .setValue('Manuel');

    if (modePaiement) {
      manuelles
        .getRange(ligneTransaction, 17)
        .setValue('Mode paiement : ' + modePaiement);
    }

    manuelles
      .getRange(ligneTransaction, 2)
      .setNumberFormat('yyyy-mm-dd');

    manuelles
      .getRange(ligneTransaction, 5)
      .setNumberFormat(
        '$#,##0.00;[Red]-$#,##0.00'
      );

    ligneVentilation = prochaineLigneLibre_(
      ventilation,
      4,
      6
    );

    ventilation
      .getRange(ligneVentilation, 2)
      .setValue(identifiant);

    ventilation
      .getRange(ligneVentilation, 4)
      .setValue(codeCompte);

    if (programme) {
      ventilation
        .getRange(ligneVentilation, 6)
        .setValue(programme);
    }

    ventilation
      .getRange(ligneVentilation, 13)
      .setValue(montant)
      .setNumberFormat(
        '$#,##0.00;[Red]-$#,##0.00'
      );

    ventilation
      .getRange(ligneVentilation, 14)
      .setValue('Saisie manuelle V2');

    SpreadsheetApp.flush();

    const piece = extrairePieceTransactionV2_(donnees);

    let resultatRecu = null;

    if (piece) {
      const idFournisseur =
        idTiers.indexOf('FOU-') === 0
          ? idTiers
          : '';

      resultatRecu = enregistrerRecuV2({
        fichier: piece,
        dateRecu: String(donnees.date || ''),
        idFournisseur: idFournisseur,
        total: montant,
        modePaiement: modePaiement,
        idTransactionLiee: identifiant,
        notes: 'Pièce ajoutée depuis la saisie de transaction'
      });

      manuelles
        .getRange(ligneTransaction, 9)
        .setValue(resultatRecu.idRecu);

      manuelles
        .getRange(ligneTransaction, 10)
        .setValue(resultatRecu.urlFichier);
    }

    SpreadsheetApp.flush();

    return {
      succes: true,
      identifiant: identifiant,
      idRecu: resultatRecu ? resultatRecu.idRecu : '',
      urlFichier: resultatRecu ? resultatRecu.urlFichier : '',
      ocr: resultatRecu ? resultatRecu.ocr : null
    };
  } catch (erreur) {
    // Si l'ajout de la pièce échoue après la création de la transaction,
    // on retire les saisies manuelles pour ne pas laisser une opération partielle.
    try {
      if (ligneTransaction) {
        const manuelles = SpreadsheetApp
          .getActive()
          .getSheetByName('Transactions manuelles');

        if (manuelles) {
          manuelles
            .getRange(ligneTransaction, 1, 1, 7)
            .clearContent();
          manuelles
            .getRange(ligneTransaction, 9, 1, 3)
            .clearContent();
          manuelles
            .getRange(ligneTransaction, 17)
            .clearContent();
        }
      }

      if (ligneVentilation) {
        const ventilation = SpreadsheetApp
          .getActive()
          .getSheetByName('Ventilation');

        if (ventilation) {
          [
            2, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14
          ].forEach(function(colonne) {
            ventilation
              .getRange(ligneVentilation, colonne)
              .clearContent();
          });
        }
      }
    } catch (erreurNettoyage) {
      console.error(
        'Erreur pendant le nettoyage de la transaction partielle :',
        erreurNettoyage.message
      );
    }

    throw erreur;
  } finally {
    verrou.releaseLock();
  }
}

function extrairePieceTransactionV2_(donnees) {
  const candidats = [
    donnees ? donnees.photoTransaction : null,
    donnees ? donnees.fichierTransaction : null
  ].filter(function(element) {
    return element &&
      typeof element.getBytes === 'function' &&
      element.getBytes().length > 0;
  });

  if (candidats.length > 1) {
    throw new Error(
      'Utilisez soit la photo, soit le fichier, mais pas les deux en même temps.'
    );
  }

  return candidats.length ? candidats[0] : null;
}

function verifierCompteV2_(ss, code) {
  const feuille = ss.getSheetByName('Comptes');

  if (!feuille) {
    throw new Error(
      'L’onglet Comptes est introuvable.'
    );
  }

  const codes = feuille
    .getRange(
      6,
      1,
      Math.max(feuille.getLastRow() - 5, 1),
      1
    )
    .getDisplayValues()
    .flat()
    .map(function(valeur) {
      return String(valeur || '').trim();
    });

  if (!codes.includes(String(code).trim())) {
    throw new Error(
      'Le compte ' + code + ' est introuvable.'
    );
  }
}

function prochaineLigneLibre_(
  feuille,
  colonne,
  premiereLigne
) {
  const nombre =
    feuille.getMaxRows() -
    premiereLigne +
    1;

  const valeurs = feuille
    .getRange(
      premiereLigne,
      colonne,
      nombre,
      1
    )
    .getDisplayValues();

  for (
    let index = 0;
    index < valeurs.length;
    index++
  ) {
    if (!String(valeurs[index][0] || '').trim()) {
      return premiereLigne + index;
    }
  }

  feuille.insertRowsAfter(
    feuille.getMaxRows(),
    100
  );

  return feuille.getMaxRows() - 99;
}

function nettoyerAnciensDeclencheursV2() {
  const gestionnairesLegacy = new Set([
    'ajouterMenuTransactionsMixtesAuDemarrage',
    'ouvrirInterfaceMixteDepuisImport',
    'creerOuMettreAJourForfaitDepuisBanque',
    'ajouterMenuBudgetAnnuelAuDemarrage'
  ]);

  const declencheurs =
    ScriptApp.getProjectTriggers();

  const supprimes = [];

  declencheurs.forEach(function(declencheur) {
    const gestionnaire =
      declencheur.getHandlerFunction();

    if (
      gestionnairesLegacy.has(gestionnaire)
    ) {
      ScriptApp.deleteTrigger(declencheur);
      supprimes.push(gestionnaire);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Nettoyage terminé.\n\n' +
    'Déclencheurs supprimés : ' +
    (
      supprimes.length
        ? supprimes.join(', ')
        : 'Aucun'
    )
  );

  return supprimes;
}

function initialiserDossiersDrive() {
  const ss = SpreadsheetApp.getActive();
  const configuration =
    ss.getSheetByName('Configuration');

  if (!configuration) {
    throw new Error(
      'L’onglet Configuration est introuvable.'
    );
  }

  const ligneConfiguration =
    trouverLigneConfiguration_(
      configuration,
      'Dossier Drive principal'
    );

  let dossierPrincipal =
    obtenirDossierConfigure_(
      configuration
        .getRange(
          ligneConfiguration,
          12
        )
        .getValue()
    );

  if (!dossierPrincipal) {
    dossierPrincipal = DriveApp.createFolder(
      CONFIG_OSBL.dossierPrincipal
    );
  }

  CONFIG_OSBL.sousDossiers.forEach(
    function(nom) {
      trouverOuCreerSousDossier_(
        dossierPrincipal,
        nom
      );
    }
  );

  configuration
    .getRange(
      ligneConfiguration,
      12
    )
    .setValue(
      dossierPrincipal.getUrl()
    );

  configuration
    .getRange(
      ligneConfiguration,
      14
    )
    .setValue('Confirmé');

  ss.toast(
    'Les dossiers Drive sont prêts.',
    'Gestion OSBL',
    6
  );
}

function trouverLigneConfiguration_(
  feuille,
  parametre
) {
  const valeurs = feuille
    .getRange(
      1,
      11,
      feuille.getLastRow(),
      1
    )
    .getValues();

  for (
    let index = 0;
    index < valeurs.length;
    index++
  ) {
    if (
      String(valeurs[index][0] || '').trim() ===
      parametre
    ) {
      return index + 1;
    }
  }

  throw new Error(
    'Le paramètre « ' +
    parametre +
    ' » est introuvable.'
  );
}

function obtenirDossierConfigure_(url) {
  if (!url) {
    return null;
  }

  const resultat = String(url)
    .match(/\/folders\/([^/?]+)/);

  if (!resultat) {
    return null;
  }

  try {
    return DriveApp.getFolderById(
      resultat[1]
    );
  } catch (erreur) {
    return null;
  }
}

function trouverOuCreerSousDossier_(
  dossierParent,
  nom
) {
  const dossiers =
    dossierParent.getFoldersByName(nom);

  if (dossiers.hasNext()) {
    return dossiers.next();
  }

  return dossierParent.createFolder(nom);
}

function verifierSysteme() {
  const ss = SpreadsheetApp.getActive();

  const ongletsRequis = [
    'Accueil',
    'Import bancaire',
    'Transactions V2',
    'Transactions manuelles',
    'Ventilation',
    'Journal V2',
    'Reçus V2',
    'Comptes',
    'Programmes',

    'Fournisseurs',
    'Inventaire',
    'Budgets',
    'Vue financière',
    'Rapprochement V2',
    'Règles bancaires',
    'Soldes ouverture'
  ];

  const manquants = ongletsRequis
    .filter(function(nom) {
      return !ss.getSheetByName(nom);
    });

  if (manquants.length) {
    SpreadsheetApp.getUi().alert(
      'Onglets manquants : ' +
      manquants.join(', ')
    );
    return;
  }

  const journal =
    ss.getSheetByName('Journal V2');

  const debit = Number(
    journal.getRange('H6:H1000')
      .getValues()
      .flat()
      .reduce(function(total, valeur) {
        return total + (Number(valeur) || 0);
      }, 0)
  );

  const credit = Number(
    journal.getRange('I6:I1000')
      .getValues()
      .flat()
      .reduce(function(total, valeur) {
        return total + (Number(valeur) || 0);
      }, 0)
  );

  const ecart =
    Math.round(
      (debit - credit) * 100
    ) / 100;

  SpreadsheetApp.getUi().alert(
    'Vérification V2 terminée\n\n' +
    'Les ' +
    ongletsRequis.length +
    ' onglets requis sont présents.\n' +
    'Écart débit/crédit Journal V2 : ' +
    ecart.toFixed(2) +
    ' $'
  );
}

