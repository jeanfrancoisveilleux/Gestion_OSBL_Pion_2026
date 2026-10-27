const CONFIG_RECUS_V2 = {
  proprieteIdClasseur: 'RECUS_V2_ID_CLASSEUR',
  feuille: 'Reçus V2',
  feuilleFournisseurs: 'Fournisseurs',
  feuilleConfiguration: 'Configuration',
  premiereLigne: 6,
  dossierRecus: 'REÇUS_FACTURES',
  tailleMaxOctets: 15 * 1024 * 1024,
  langueOcr: 'fr'
};

/**
 * Ouvre la capture V2 depuis Google Sheets.
 */
function ouvrirCaptureRecuV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Le classeur Google Sheets est introuvable.');
  }

  memoriserClasseurRecusV2_(ss);
  verifierStructureRecusV2_(ss);

  const html = HtmlService
    .createHtmlOutputFromFile('CaptureRecuV2')
    .setWidth(520)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Nouveau reçu');
}

/**
 * Configure le raccourci mobile.
 * Le déploiement Web doit déjà exister.
 */
function installerCaptureRecuV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Ouvrez le classeur avant de lancer l’installation.');
  }

  memoriserClasseurRecusV2_(ss);
  verifierStructureRecusV2_(ss);
  obtenirDossierRecusV2_(ss);

  const urlApplication = String(ScriptApp.getService().getUrl() || '').trim();

  if (!urlApplication) {
    SpreadsheetApp.getUi().alert(
      'Le module Reçus V2 est prêt.\n\n' +
      'Créez ou mettez à jour le déploiement comme application Web, puis relancez ' +
      '« Installer le raccourci mobile V2 ».'
    );

    return {
      succes: true,
      urlApplication: '',
      lienInstalle: false
    };
  }

  const feuille = ss.getSheetByName(CONFIG_RECUS_V2.feuille);
  const cellule = feuille.getRange('A3');
  const libelle = '📷 Nouveau reçu';

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
    'Ouvre la capture mobile Reçus V2. Sur téléphone, ajoutez cette application Web à l’écran d’accueil.'
  );

  SpreadsheetApp.getUi().alert(
    'Capture mobile Reçus V2 installée.\n\n' +
    'Lien : ' + urlApplication + '\n\n' +
    'Le raccourci a aussi été ajouté dans Reçus V2!A3.'
  );

  return {
    succes: true,
    urlApplication: urlApplication,
    lienInstalle: true
  };
}

/**
 * Données légères utilisées par l’interface mobile.
 */
function obtenirDonneesCaptureRecuV2() {
  const ss = obtenirClasseurRecusV2_();
  verifierStructureRecusV2_(ss);

  return {
    dateAujourdhui: Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd'
    ),
    fournisseurs: chargerFournisseursRecusV2_(ss),
    transactions: chargerTransactionsRecusV2_(ss),
    tailleMaxMo: Math.round(CONFIG_RECUS_V2.tailleMaxOctets / 1024 / 1024),
    urlApplication: String(ScriptApp.getService().getUrl() || '').trim()
  };
}

/**
 * Enregistre une photo/PDF sans exiger de transaction bancaire existante.
 * Le formulaire HTML doit être transmis directement afin que le fichier arrive comme Blob.
 */
function enregistrerRecuV2(formulaire) {
  const verrou = LockService.getScriptLock();

  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Un autre reçu est en cours d’enregistrement. Réessayez dans quelques secondes.'
    );
  }

  let fichierCree = null;

  try {
    const ss = obtenirClasseurRecusV2_();
    verifierStructureRecusV2_(ss);

    const feuille = ss.getSheetByName(CONFIG_RECUS_V2.feuille);
    const fuseau = ss.getSpreadsheetTimeZone();

    const dateTexteSaisie = String(
      formulaire && formulaire.dateRecu || ''
    ).trim();

    const dateAujourdhui = Utilities.formatDate(
      new Date(),
      fuseau,
      'yyyy-MM-dd'
    );

    const autoriserCorrectionDateOcr =
      !dateTexteSaisie || dateTexteSaisie === dateAujourdhui;

    const dateRecu = convertirDateRecuV2_(dateTexteSaisie);
    const heure = String(formulaire && formulaire.heure || '').trim();
    const idFournisseur = String(
      formulaire && formulaire.idFournisseur || ''
    ).trim();
    const idTransactionLiee = String(
      formulaire && formulaire.idTransactionLiee || ''
    ).trim();
    const modePaiement = String(
      formulaire && formulaire.modePaiement || ''
    ).trim();
    const quatreChiffres = String(
      formulaire && formulaire.quatreChiffres || ''
    ).trim().replace(/\D/g, '').slice(-4);
    const notes = String(formulaire && formulaire.notes || '').trim();

    const sousTotal = nombreOptionnelRecuV2_(
      formulaire && formulaire.sousTotal
    );
    const tps = nombreOptionnelRecuV2_(formulaire && formulaire.tps);
    const tvq = nombreOptionnelRecuV2_(formulaire && formulaire.tvq);
    const autresTaxes = nombreOptionnelRecuV2_(
      formulaire && formulaire.autresTaxes
    );
    const pourboire = nombreOptionnelRecuV2_(
      formulaire && formulaire.pourboire
    );
    const total = nombreOptionnelRecuV2_(formulaire && formulaire.total);

    const transactionLiee = idTransactionLiee
      ? obtenirTransactionRecuV2_(ss, idTransactionLiee)
      : null;

    const blob = extraireBlobRecuV2_(formulaire);
    const fichierValide = validerBlobRecuV2_(blob);

    const ligne = prochaineLigneRecuV2_(feuille);

    feuille.getRange(ligne, 2).setValue(dateRecu);
    SpreadsheetApp.flush();

    const idRecu = String(
      feuille.getRange(ligne, 1).getDisplayValue() || ''
    ).trim();

    if (!idRecu) {
      throw new Error(
        'Impossible de générer l’identifiant du reçu dans Reçus V2.'
      );
    }

    const idFournisseurFinal = idFournisseur ||
      (transactionLiee ? transactionLiee.idFournisseur : '');

    const fournisseur = idFournisseurFinal
      ? trouverNomFournisseurRecuV2_(ss, idFournisseurFinal)
      : (transactionLiee ? transactionLiee.fournisseur : '');

    const dossier = obtenirDossierRecusV2_(ss, dateRecu);
    const nomFichier = creerNomFichierRecuV2_(
      idRecu,
      dateRecu,
      fournisseur,
      fichierValide.extension,
      fuseau
    );

    fichierCree = dossier.createFile(blob.setName(nomFichier));
    fichierCree.setDescription(
      'Reçu V2 ' + idRecu + ' — en attente de jumelage bancaire.'
    );

    if (heure) {
      feuille.getRange(ligne, 3).setValue(heure);
    }

    if (idFournisseurFinal) {
      feuille.getRange(ligne, 4).setValue(idFournisseurFinal);
    }

    if (idTransactionLiee) {
      feuille.getRange(ligne, 25).setValue(idTransactionLiee);
    }

    ecrireNombreOptionnelRecuV2_(feuille, ligne, 6, sousTotal);
    ecrireNombreOptionnelRecuV2_(feuille, ligne, 7, tps);
    ecrireNombreOptionnelRecuV2_(feuille, ligne, 8, tvq);
    ecrireNombreOptionnelRecuV2_(feuille, ligne, 9, autresTaxes);
    ecrireNombreOptionnelRecuV2_(feuille, ligne, 10, pourboire);

    if (total !== null) {
      feuille.getRange(ligne, 11).setValue(total);
    } else if (transactionLiee && transactionLiee.montant !== null) {
      feuille.getRange(ligne, 11).setValue(transactionLiee.montant);
    }

    if (modePaiement) {
      feuille.getRange(ligne, 12).setValue(modePaiement);
    }

    if (quatreChiffres) {
      feuille.getRange(ligne, 13).setNumberFormat('@').setValue(quatreChiffres);
    }

    feuille.getRange(ligne, 14).setValue(fichierCree.getUrl());
    feuille.getRange(ligne, 15).setValue(nomFichier);

    if (notes) {
      feuille.getRange(ligne, 20).setValue(notes);
    }

    SpreadsheetApp.flush();

    const resultatOcr = tenterLectureOcrRecuV2_(
      ss,
      feuille,
      ligne,
      fichierCree,
      autoriserCorrectionDateOcr
    );

    SpreadsheetApp.flush();

    return {
      succes: true,
      idRecu: idRecu,
      urlFichier: fichierCree.getUrl(),
      nomFichier: nomFichier,
      idTransactionLiee: idTransactionLiee,
      etat: String(
        feuille.getRange(ligne, 19).getDisplayValue() || 'À jumeler'
      ).trim(),
      ocr: resultatOcr
    };
  } catch (erreur) {
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

    throw erreur;
  } finally {
    verrou.releaseLock();
  }
}

function memoriserClasseurRecusV2_(ss) {
  PropertiesService.getScriptProperties().setProperty(
    CONFIG_RECUS_V2.proprieteIdClasseur,
    ss.getId()
  );
}

function obtenirClasseurRecusV2_() {
  const actif = SpreadsheetApp.getActiveSpreadsheet();

  if (actif) {
    memoriserClasseurRecusV2_(actif);
    return actif;
  }

  const idClasseur = String(
    PropertiesService.getScriptProperties().getProperty(
      CONFIG_RECUS_V2.proprieteIdClasseur
    ) || ''
  ).trim();

  if (!idClasseur) {
    throw new Error(
      'La capture mobile Reçus V2 n’est pas configurée. ' +
      'Exécutez « Installer le raccourci mobile V2 » depuis Google Sheets.'
    );
  }

  return SpreadsheetApp.openById(idClasseur);
}

function verifierStructureRecusV2_(ss) {
  const feuille = ss.getSheetByName(CONFIG_RECUS_V2.feuille);

  if (!feuille) {
    throw new Error('L’onglet « Reçus V2 » est introuvable.');
  }

  const attendus = [
    'ID reçu',
    'Date reçu',
    'Heure',
    'ID fournisseur',
    'Fournisseur',
    'Sous-total',
    'TPS',
    'TVQ',
    'Autres taxes',
    'Pourboire',
    'Total',
    'Mode paiement',
    '4 derniers chiffres',
    'Lien photo / PDF',
    'Nom fichier',
    'ID import bancaire lié',
    'Candidat bancaire',
    'Confiance',
    'État',
    'Notes',
    'Lecture auto',
    'Texte OCR',
    'Lu le',
    'Résultat OCR',
    'ID transaction liée'
  ];

  const presents = feuille
    .getRange(5, 1, 1, attendus.length)
    .getDisplayValues()[0]
    .map(function(valeur) {
      return String(valeur || '').trim();
    });

  attendus.forEach(function(attendu, index) {
    if (presents[index] !== attendu) {
      throw new Error(
        'Structure Reçus V2 incompatible à la colonne ' +
        (index + 1) + ' : « ' + presents[index] +
        ' » au lieu de « ' + attendu + ' ».'
      );
    }
  });
}

function chargerFournisseursRecusV2_(ss) {
  const feuille = ss.getSheetByName(CONFIG_RECUS_V2.feuilleFournisseurs);

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  return feuille
    .getRange(6, 1, feuille.getLastRow() - 5, 2)
    .getDisplayValues()
    .map(function(ligne) {
      return {
        id: String(ligne[0] || '').trim(),
        nom: String(ligne[1] || '').trim()
      };
    })
    .filter(function(element) {
      return element.id && element.nom && element.id !== 'FOU-0000';
    })
    .sort(function(a, b) {
      return a.nom.localeCompare(b.nom, 'fr');
    });
}


function chargerTransactionsRecusV2_(ss) {
  const feuille = ss.getSheetByName('Transactions V2');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const nombreLignes = Math.min(
    feuille.getLastRow() - 5,
    feuille.getMaxRows() - 5
  );

  const valeurs = feuille
    .getRange(6, 1, nombreLignes, 17)
    .getDisplayValues();

  return valeurs
    .filter(function(ligne) {
      return String(ligne[0] || '').trim() &&
        String(ligne[1] || '').trim();
    })
    .map(function(ligne) {
      return {
        id: String(ligne[0] || '').trim(),
        date: String(ligne[1] || '').trim(),
        type: String(ligne[2] || '').trim(),
        description: String(ligne[3] || '').trim(),
        montant: convertirMontantAfficheRecuV2_(ligne[4]),
        idFournisseur: String(ligne[6] || '').trim(),
        fournisseur: String(ligne[7] || '').trim(),
        source: String(ligne[10] || '').trim()
      };
    })
    .sort(function(a, b) {
      return String(b.date).localeCompare(String(a.date));
    })
    .slice(0, 250);
}

function obtenirTransactionRecuV2_(ss, idTransaction) {
  const id = String(idTransaction || '').trim();

  if (!id) {
    return null;
  }

  const transaction = chargerTransactionsRecusV2_(ss).find(
    function(item) {
      return item.id === id;
    }
  );

  if (!transaction) {
    throw new Error(
      'La transaction sélectionnée est introuvable dans Transactions V2.'
    );
  }

  return transaction;
}

function convertirMontantAfficheRecuV2_(valeur) {
  const texte = String(valeur == null ? '' : valeur)
    .replace(/\s/g, '')
    .replace(/\$/g, '')
    .replace(/,/g, '.')
    .replace(/[^\d.-]/g, '');

  if (!texte) {
    return null;
  }

  const nombre = Number(texte);

  return isFinite(nombre)
    ? Math.round(Math.abs(nombre) * 100) / 100
    : null;
}

function trouverNomFournisseurRecuV2_(ss, idFournisseur) {
  const fournisseurs = chargerFournisseursRecusV2_(ss);
  const trouve = fournisseurs.find(function(element) {
    return element.id === idFournisseur;
  });

  return trouve ? trouve.nom : '';
}

function prochaineLigneRecuV2_(feuille) {
  const nb = feuille.getMaxRows() - CONFIG_RECUS_V2.premiereLigne + 1;
  const valeurs = feuille
    .getRange(CONFIG_RECUS_V2.premiereLigne, 2, nb, 1)
    .getDisplayValues();

  for (let i = 0; i < valeurs.length; i++) {
    if (!String(valeurs[i][0] || '').trim()) {
      return CONFIG_RECUS_V2.premiereLigne + i;
    }
  }

  feuille.insertRowsAfter(feuille.getMaxRows(), 100);
  return feuille.getMaxRows() - 99;
}

function convertirDateRecuV2_(valeur) {
  const texte = String(valeur || '').trim();
  const date = texte ? new Date(texte + 'T12:00:00') : new Date();

  if (isNaN(date.getTime())) {
    throw new Error('La date du reçu est invalide.');
  }

  date.setHours(12, 0, 0, 0);
  return date;
}

function nombreOptionnelRecuV2_(valeur) {
  const texte = String(valeur == null ? '' : valeur)
    .trim()
    .replace(/\s/g, '')
    .replace(',', '.');

  if (!texte) {
    return null;
  }

  const nombre = Number(texte);

  if (!isFinite(nombre) || nombre < 0) {
    throw new Error('Un montant du reçu est invalide.');
  }

  return Math.round(nombre * 100) / 100;
}

function ecrireNombreOptionnelRecuV2_(feuille, ligne, colonne, valeur) {
  if (valeur !== null) {
    feuille.getRange(ligne, colonne).setValue(valeur);
  }
}

function extraireBlobRecuV2_(formulaire) {
  const candidats = [
    formulaire ? formulaire.photo : null,
    formulaire ? formulaire.fichier : null
  ].filter(function(element) {
    return element &&
      typeof element.getBytes === 'function' &&
      element.getBytes().length > 0;
  });

  if (candidats.length === 0) {
    throw new Error('Prenez une photo ou sélectionnez un fichier.');
  }

  if (candidats.length > 1) {
    throw new Error('Sélectionnez une seule photo ou un seul fichier.');
  }

  return candidats[0];
}

function validerBlobRecuV2_(blob) {
  const taille = blob.getBytes().length;

  if (!taille) {
    throw new Error('Le fichier sélectionné est vide.');
  }

  if (taille > CONFIG_RECUS_V2.tailleMaxOctets) {
    throw new Error('Le fichier dépasse la limite de 15 Mo.');
  }

  const type = String(blob.getContentType() || '').toLowerCase();
  const nom = String(blob.getName() || '').trim();
  const extensionNom = nom.indexOf('.') !== -1
    ? nom.split('.').pop().toLowerCase()
    : '';

  const extensions = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/heic': 'heic',
    'image/heif': 'heif',
    'application/pdf': 'pdf'
  };

  let extension = extensions[type] || extensionNom;

  if (['jpg', 'jpeg', 'png', 'webp', 'heic', 'heif', 'pdf'].indexOf(extension) === -1) {
    throw new Error(
      'Format non accepté. Utilisez une photo JPG/PNG/HEIC/WEBP ou un PDF.'
    );
  }

  if (extension === 'jpeg') {
    extension = 'jpg';
  }

  return {
    extension: extension,
    taille: taille
  };
}

function obtenirDossierRecusV2_(ss, dateRecu) {
  const configuration = ss.getSheetByName(CONFIG_RECUS_V2.feuilleConfiguration);

  if (!configuration) {
    throw new Error('L’onglet Configuration est introuvable.');
  }

  const valeurs = configuration
    .getRange(1, 11, Math.max(1, configuration.getLastRow()), 2)
    .getValues();

  let urlDossier = '';

  for (let i = 0; i < valeurs.length; i++) {
    if (String(valeurs[i][0] || '').trim() === 'Dossier Drive principal') {
      urlDossier = String(valeurs[i][1] || '').trim();
      break;
    }
  }

  const match = urlDossier.match(/\/folders\/([^/?]+)/);

  if (!match) {
    throw new Error(
      'Le dossier Drive principal n’est pas configuré dans Configuration.'
    );
  }

  const principal = DriveApp.getFolderById(match[1]);
  const recus = trouverOuCreerDossierRecuV2_(
    principal,
    CONFIG_RECUS_V2.dossierRecus
  );

  const annee = String(
    (dateRecu instanceof Date ? dateRecu : new Date()).getFullYear()
  );

  return trouverOuCreerDossierRecuV2_(recus, annee + '-FACTURES');
}

function trouverOuCreerDossierRecuV2_(parent, nom) {
  const dossiers = parent.getFoldersByName(nom);

  if (dossiers.hasNext()) {
    return dossiers.next();
  }

  return parent.createFolder(nom);
}

function creerNomFichierRecuV2_(
  idRecu,
  dateRecu,
  fournisseur,
  extension,
  fuseau
) {
  const dateTexte = Utilities.formatDate(dateRecu, fuseau, 'yyyy-MM-dd');
  const fournisseurNettoye = String(fournisseur || 'Recu')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);

  return [
    dateTexte,
    idRecu,
    fournisseurNettoye || 'Recu'
  ].join('_') + '.' + extension;
}

/**
 * Relance l'OCR sur le reçu de la ligne active.
 */
function relireOcrRecuSelectionneV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss && ss.getActiveSheet();

  if (!ss || !feuille || feuille.getName() !== CONFIG_RECUS_V2.feuille) {
    throw new Error(
      'Ouvrez l’onglet Reçus V2 et sélectionnez une cellule de la ligne à relire.'
    );
  }

  const ligne = feuille.getActiveRange().getRow();

  if (ligne < CONFIG_RECUS_V2.premiereLigne) {
    throw new Error('Sélectionnez une ligne de reçu.');
  }

  const idRecu = String(feuille.getRange(ligne, 1).getDisplayValue() || '').trim();
  const url = String(feuille.getRange(ligne, 14).getDisplayValue() || '').trim();

  if (!idRecu || !url) {
    throw new Error('Cette ligne ne contient pas un reçu complet avec une pièce jointe.');
  }

  const idFichier = extraireIdFichierDriveRecuV2_(url);

  if (!idFichier) {
    throw new Error('Impossible de reconnaître le fichier Drive lié à ce reçu.');
  }

  const fichier = DriveApp.getFileById(idFichier);
  feuille.getRange(ligne, 22, 1, 3).clearContent();

  const resultat = tenterLectureOcrRecuV2_(
    ss,
    feuille,
    ligne,
    fichier,
    false
  );

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    resultat.succes
      ? 'Lecture OCR terminée.\n\n' + resultat.resume
      : 'Le reçu est conservé, mais la lecture OCR a échoué.\n\n' + resultat.message
  );

  return resultat;
}

/**
 * OCR gratuit via Google Drive.
 * Requiert le service avancé « Drive API » dans Apps Script.
 */
function tenterLectureOcrRecuV2_(
  ss,
  feuille,
  ligne,
  fichier,
  autoriserCorrectionDate
) {
  try {
    if (typeof Drive === 'undefined' || !Drive.Files) {
      throw new Error(
        'Le service avancé Drive API n’est pas activé dans Apps Script.'
      );
    }

    const idRecu = String(
      feuille.getRange(ligne, 1).getDisplayValue() || ''
    ).trim();

    const texte = lireTexteOcrDriveV2_(fichier, idRecu);

    if (!texte) {
      throw new Error(
        'Google Drive n’a détecté aucun texte lisible sur cette pièce.'
      );
    }

    const analyse = analyserTexteRecuV2_(texte, ss);

    appliquerAnalyseOcrRecuV2_(
      feuille,
      ligne,
      analyse,
      texte,
      autoriserCorrectionDate
    );

    const resume = construireResumeOcrRecuV2_(analyse);

    feuille.getRange(ligne, 23).setValue(new Date());
    feuille.getRange(ligne, 24).setValue(resume);

    return {
      succes: true,
      resume: resume,
      champs: analyse
    };
  } catch (erreur) {
    const message = String(
      erreur && erreur.message ? erreur.message : erreur
    );

    feuille.getRange(ligne, 23).setValue(new Date());
    feuille.getRange(ligne, 24).setValue('Échec OCR — ' + message);

    return {
      succes: false,
      message: message
    };
  }
}

function lireTexteOcrDriveV2_(fichier, idRecu) {
  const blob = fichier.getBlob();
  let documentOcr = null;

  try {
    documentOcr = Drive.Files.create(
      {
        name: 'OCR_TEMP_' + idRecu,
        mimeType: 'application/vnd.google-apps.document'
      },
      blob,
      {
        ocrLanguage: CONFIG_RECUS_V2.langueOcr,
        fields: 'id,name'
      }
    );

    if (!documentOcr || !documentOcr.id) {
      throw new Error('Google Drive n’a pas retourné de document OCR.');
    }

    let texte = '';

    for (let tentative = 0; tentative < 4; tentative++) {
      try {
        Utilities.sleep(tentative === 0 ? 250 : 750);
        texte = DocumentApp
          .openById(documentOcr.id)
          .getBody()
          .getText();
      } catch (erreurLecture) {
        texte = '';
      }

      if (String(texte || '').trim()) {
        break;
      }
    }

    return String(texte || '').trim();
  } finally {
    if (documentOcr && documentOcr.id) {
      try {
        DriveApp.getFileById(documentOcr.id).setTrashed(true);
      } catch (erreurCorbeille) {
        console.error(
          'Impossible de supprimer le document OCR temporaire :',
          erreurCorbeille.message
        );
      }
    }
  }
}

function analyserTexteRecuV2_(texte, ss) {
  const sousTotal = trouverMontantParLibellesRecuV2_(
    texte,
    ['sous total', 'subtotal', 'sous-total']
  );

  const tps = trouverMontantParLibellesRecuV2_(texte, ['tps', 'gst']);
  const tvq = trouverMontantParLibellesRecuV2_(texte, ['tvq', 'qst']);
  const pourboire = trouverMontantParLibellesRecuV2_(
    texte,
    ['pourboire', 'tip', 'gratuity']
  );

  const total = trouverTotalRecuV2_(texte);
  const date = trouverDateRecuV2_(texte);
  const paiement = trouverPaiementRecuV2_(texte);
  const quatreChiffres = trouverQuatreChiffresRecuV2_(texte);
  const fournisseur = trouverFournisseurDansOcrRecuV2_(texte, ss);

  return {
    idFournisseur: fournisseur ? fournisseur.id : '',
    fournisseur: fournisseur ? fournisseur.nom : '',
    sousTotal: sousTotal,
    tps: tps,
    tvq: tvq,
    autresTaxes: null,
    pourboire: pourboire,
    total: total,
    date: date,
    modePaiement: paiement,
    quatreChiffres: quatreChiffres
  };
}

function appliquerAnalyseOcrRecuV2_(
  feuille,
  ligne,
  analyse,
  texte,
  autoriserCorrectionDate
) {
  if (
    autoriserCorrectionDate &&
    analyse.date &&
    dateOcrPlausibleRecuV2_(analyse.date)
  ) {
    feuille.getRange(ligne, 2).setValue(analyse.date);
  }

  if (
    analyse.idFournisseur &&
    !String(feuille.getRange(ligne, 4).getDisplayValue() || '').trim()
  ) {
    feuille.getRange(ligne, 4).setValue(analyse.idFournisseur);
  }

  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 6, analyse.sousTotal);
  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 7, analyse.tps);
  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 8, analyse.tvq);
  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 9, analyse.autresTaxes);
  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 10, analyse.pourboire);
  ecrireNombreOcrSiVideRecuV2_(feuille, ligne, 11, analyse.total);

  if (
    analyse.modePaiement &&
    !String(feuille.getRange(ligne, 12).getDisplayValue() || '').trim()
  ) {
    feuille.getRange(ligne, 12).setValue(analyse.modePaiement);
  }

  if (
    analyse.quatreChiffres &&
    !String(feuille.getRange(ligne, 13).getDisplayValue() || '').trim()
  ) {
    feuille
      .getRange(ligne, 13)
      .setNumberFormat('@')
      .setValue(analyse.quatreChiffres);
  }

  feuille.getRange(ligne, 22).setValue(texte);
}

function ecrireNombreOcrSiVideRecuV2_(feuille, ligne, colonne, valeur) {
  if (valeur === null || valeur === undefined || valeur === '') {
    return;
  }

  const cellule = feuille.getRange(ligne, colonne);
  const existant = cellule.getValue();

  if (existant === '' || existant === null) {
    cellule.setValue(valeur);
  }
}

function construireResumeOcrRecuV2_(analyse) {
  const parties = [];

  if (analyse.fournisseur) {
    parties.push('Fournisseur: ' + analyse.fournisseur);
  }
  if (analyse.total !== null) {
    parties.push('Total: ' + formaterMontantResumeRecuV2_(analyse.total));
  }
  if (analyse.sousTotal !== null) {
    parties.push('Sous-total: ' + formaterMontantResumeRecuV2_(analyse.sousTotal));
  }
  if (analyse.tps !== null) {
    parties.push('TPS: ' + formaterMontantResumeRecuV2_(analyse.tps));
  }
  if (analyse.tvq !== null) {
    parties.push('TVQ: ' + formaterMontantResumeRecuV2_(analyse.tvq));
  }
  if (analyse.pourboire !== null) {
    parties.push('Pourboire: ' + formaterMontantResumeRecuV2_(analyse.pourboire));
  }
  if (analyse.modePaiement) {
    parties.push('Paiement: ' + analyse.modePaiement);
  }
  if (analyse.quatreChiffres) {
    parties.push('Carte: •••• ' + analyse.quatreChiffres);
  }
  if (analyse.date) {
    parties.push(
      'Date lue: ' +
      Utilities.formatDate(
        analyse.date,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      )
    );
  }

  return parties.length
    ? parties.join(' | ')
    : 'Lecture terminee, aucun champ reconnu.';
}

function extraireIdFichierDriveRecuV2_(url) {
  const texte = String(url || '').trim();

  if (!texte) {
    return '';
  }

  const motifs = [
    /\/d\/([A-Za-z0-9_-]{20,})/,
    /[?&]id=([A-Za-z0-9_-]{20,})/,
    /^([A-Za-z0-9_-]{20,})$/
  ];

  for (let i = 0; i < motifs.length; i++) {
    const resultat = texte.match(motifs[i]);

    if (resultat) {
      return resultat[1];
    }
  }

  return '';
}

function normaliserTexteOcrRecuV2_(texte) {
  return String(texte || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function convertirMontantOcrRecuV2_(valeur) {
  let texte = String(valeur == null ? '' : valeur)
    .trim()
    .replace(/\u00a0/g, ' ')
    .replace(/\$/g, '')
    .replace(/\s/g, '');

  if (!texte) {
    return null;
  }

  const negatifParentheses =
    /^\(.+\)$/.test(texte);

  texte = texte
    .replace(/[()]/g, '')
    .replace(',', '.')
    .replace(/[^0-9.-]/g, '');

  const nombre = Number(texte);

  if (!isFinite(nombre)) {
    return null;
  }

  const final = negatifParentheses
    ? -Math.abs(nombre)
    : nombre;

  return Math.round(Math.abs(final) * 100) / 100;
}

function extraireMontantsLigneOcrRecuV2_(ligne) {
  const texte = String(ligne || '');

  const correspondances = texte.match(
    /(?:\d{1,3}(?:[ \u00a0]\d{3})+|\d+)[,.]\d{2}/g
  ) || [];

  return correspondances
    .map(convertirMontantOcrRecuV2_)
    .filter(function(valeur) {
      return valeur !== null;
    });
}

function trouverMontantParLibellesRecuV2_(
  texte,
  libelles
) {
  const lignes = String(texte || '')
    .split(/\r?\n/);

  const libellesNormalises = (libelles || [])
    .map(normaliserTexteOcrRecuV2_);

  for (let i = 0; i < lignes.length; i++) {
    const ligneNormalisee =
      normaliserTexteOcrRecuV2_(lignes[i]);

    const correspond = libellesNormalises.some(
      function(libelle) {
        return libelle &&
          ligneNormalisee.indexOf(libelle) >= 0;
      }
    );

    if (!correspond) {
      continue;
    }

    const montants =
      extraireMontantsLigneOcrRecuV2_(
        lignes[i]
      );

    if (montants.length) {
      return montants[
        montants.length - 1
      ];
    }
  }

  const aplati = String(texte || '')
    .replace(/\r?\n/g, ' ');

  for (
    let index = 0;
    index < libellesNormalises.length;
    index++
  ) {
    const libelle = libellesNormalises[index]
      .replace(/[.*+?^$()|[\]\\]/g, '\\$&')
      .replace(/\s+/g, '\\s+');

    const motif = new RegExp(
      libelle +
      '[^0-9]{0,25}' +
      '((?:\\d{1,3}(?:[ \\u00a0]\\d{3})+|\\d+)[,.]\\d{2})',
      'i'
    );

    const resultat = aplati.match(motif);

    if (resultat) {
      return convertirMontantOcrRecuV2_(
        resultat[1]
      );
    }
  }

  return null;
}

function trouverTotalRecuV2_(texte) {
  const lignes = String(texte || '')
    .split(/\r?\n/);

  for (let i = lignes.length - 1; i >= 0; i--) {
    const normalisee =
      normaliserTexteOcrRecuV2_(
        lignes[i]
      );

    if (
      !/\btotal\b|amount due|balance due|montant a payer/.test(
        normalisee
      )
    ) {
      continue;
    }

    if (
      /sous total|subtotal|tax total|total taxe/.test(
        normalisee
      )
    ) {
      continue;
    }

    const montants =
      extraireMontantsLigneOcrRecuV2_(
        lignes[i]
      );

    if (montants.length) {
      return montants[
        montants.length - 1
      ];
    }
  }

  return trouverMontantParLibellesRecuV2_(
    texte,
    [
      'montant a payer',
      'amount due',
      'balance due'
    ]
  );
}

function creerDateOcrRecuV2_(
  annee,
  mois,
  jour
) {
  const date = new Date(
    Number(annee),
    Number(mois) - 1,
    Number(jour),
    12,
    0,
    0
  );

  if (
    date.getFullYear() !== Number(annee) ||
    date.getMonth() !== Number(mois) - 1 ||
    date.getDate() !== Number(jour)
  ) {
    return null;
  }

  return date;
}

function trouverDateRecuV2_(texte) {
  const source = String(texte || '');

  let resultat = source.match(
    /\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/
  );

  if (resultat) {
    return creerDateOcrRecuV2_(
      resultat[1],
      resultat[2],
      resultat[3]
    );
  }

  resultat = source.match(
    /\b(\d{1,2})[-\/.](\d{1,2})[-\/.](20\d{2})\b/
  );

  if (resultat) {
    return creerDateOcrRecuV2_(
      resultat[3],
      resultat[2],
      resultat[1]
    );
  }

  resultat = source.match(
    /\b(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2})\b/
  );

  if (resultat) {
    return creerDateOcrRecuV2_(
      2000 + Number(resultat[3]),
      resultat[2],
      resultat[1]
    );
  }

  const normalise = normaliserTexteOcrRecuV2_(
    source
  );

  const mois = {
    jan: 1,
    fev: 2,
    feb: 2,
    mar: 3,
    avr: 4,
    apr: 4,
    mai: 5,
    may: 5,
    jun: 6,
    jui: 7,
    jul: 7,
    aou: 8,
    aug: 8,
    sep: 9,
    oct: 10,
    nov: 11,
    dec: 12
  };

  resultat = normalise.match(
    /\b(\d{1,2})\s+(jan|fev|feb|mar|avr|apr|mai|may|jun|jui|jul|aou|aug|sep|oct|nov|dec)[a-z]*\s+(20\d{2})\b/
  );

  if (resultat) {
    return creerDateOcrRecuV2_(
      resultat[3],
      mois[resultat[2]],
      resultat[1]
    );
  }

  return null;
}

function trouverPaiementRecuV2_(texte) {
  const normalise =
    normaliserTexteOcrRecuV2_(texte);

  if (
    /\bvisa\b|\bmastercard\b|\bmaster card\b|\bamex\b|american express|carte de credit|credit card/.test(
      normalise
    )
  ) {
    return 'Crédit';
  }

  if (
    /\binterac\b|carte de debit|debit card|\bdebit\b/.test(
      normalise
    )
  ) {
    return 'Débit';
  }

  if (
    /\bcomptant\b|\bcash\b|\bespeces\b/.test(
      normalise
    )
  ) {
    return 'Comptant';
  }

  if (
    /\bcheque\b|\bcheck\b|virement/.test(
      normalise
    )
  ) {
    return 'Autre';
  }

  return '';
}

function trouverQuatreChiffresRecuV2_(texte) {
  const source = String(texte || '');

  let resultat = source.match(
    /(?:\*|x|X|#){2,}\s*(\d{4})\b/
  );

  if (resultat) {
    return resultat[1];
  }

  resultat = source.match(
    /(?:visa|master\s*card|mastercard|amex|interac|carte|card)[^\n\r]{0,50}?(\d{4})\b/i
  );

  return resultat
    ? resultat[1]
    : '';
}

function trouverFournisseurDansOcrRecuV2_(
  texte,
  ss
) {
  const contenu =
    normaliserTexteOcrRecuV2_(texte);

  if (!contenu) {
    return null;
  }

  const fournisseurs =
    chargerFournisseursRecusV2_(ss)
      .map(function(fournisseur) {
        return {
          id: fournisseur.id,
          nom: fournisseur.nom,
          cle: normaliserTexteOcrRecuV2_(
            fournisseur.nom
          )
        };
      })
      .filter(function(fournisseur) {
        return fournisseur.cle.length >= 4 &&
          contenu.indexOf(
            fournisseur.cle
          ) >= 0;
      })
      .sort(function(a, b) {
        return b.cle.length -
          a.cle.length;
      });

  return fournisseurs.length
    ? {
        id: fournisseurs[0].id,
        nom: fournisseurs[0].nom
      }
    : null;
}

function dateOcrPlausibleRecuV2_(date) {
  if (
    !(date instanceof Date) ||
    isNaN(date.getTime())
  ) {
    return false;
  }

  const minimum = new Date(
    2015,
    0,
    1,
    0,
    0,
    0
  );

  const maximum = new Date();
  maximum.setDate(
    maximum.getDate() + 7
  );
  maximum.setHours(
    23,
    59,
    59,
    999
  );

  return date >= minimum &&
    date <= maximum;
}

function formaterMontantResumeRecuV2_(valeur) {
  const nombre = Number(valeur);

  if (!isFinite(nombre)) {
    return '';
  }

  return Utilities
    .formatString(
      '%.2f',
      Math.round(
        Math.abs(nombre) * 100
      ) / 100
    )
    .replace('.', ',') +
    ' $';
}

