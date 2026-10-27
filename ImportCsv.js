function afficherImportBancaire() {
  const interface = HtmlService
    .createHtmlOutputFromFile('ImportCSV')
    .setWidth(620)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(
    interface,
    'Importer un relevé Desjardins'
  );
}
/**
 * Point d'entrée universel.
 * Accepte CSV, PDF, XLS/XLSX ou un lien Google Sheets/Drive.
 */
function importerReleveBancaire(entree) {
  entree = entree || {};
  const lien = String(entree.lienGoogle || '').trim();
  if (lien) {
    const sourceLien = normaliserLienBancaire_(lien);
    return importerTransactionsBancaires_(
      sourceLien.transactions,
      sourceLien.nomSource
    );
  }
  const nomFichier = String(entree.nomFichier || '').trim();
  const mimeType = String(entree.mimeType || '').trim();
  const contenuBase64 = String(entree.contenuBase64 || '').trim();
  if (!nomFichier || !contenuBase64) {
    throw new Error(
      'Sélectionnez un fichier bancaire ou collez un lien Google Sheets.'
    );
  }
  const octets = Utilities.base64Decode(contenuBase64);
  const blob = Utilities.newBlob(
    octets,
    mimeType || 'application/octet-stream',
    nomFichier
  );
  const source = normaliserFichierBancaire_(
    blob,
    nomFichier,
    mimeType
  );
  return importerTransactionsBancaires_(
    source.transactions,
    source.nomSource
  );
}
/**
 * Compatibilité avec l'ancien formulaire CSV.
 */
function importerCsvDesjardins(contenu, nomFichier) {
  const texte = String(contenu || '');
  if (!texte) {
    throw new Error('Le fichier CSV est vide.');
  }
  const transactions = normaliserCsvDesjardins_(
    texte,
    nomFichier || 'Relevé Desjardins.csv'
  );
  return importerTransactionsBancaires_(
    transactions,
    nomFichier || 'Relevé Desjardins.csv'
  );
}
function importerTransactionsPdfClient(
  transactions,
  nomSource
) {
  if (
    !Array.isArray(transactions) ||
    !transactions.length
  ) {
    throw new Error(
      'Aucune transaction PDF exploitable n’a été reçue.'
    );
  }

  const normalisees = transactions.map(
    function(transaction, index) {
      const date = convertirDateImport_(
        transaction.date,
        new Date().getFullYear()
      );

      if (!date) {
        throw new Error(
          'Date invalide dans le PDF à la transaction ' +
          (index + 1) +
          '.'
        );
      }

      return {
        date: date,
        code: String(
          transaction.code || ''
        ).trim(),
        description: String(
          transaction.description || ''
        )
          .replace(/\s+/g, ' ')
          .trim(),
        retrait: Math.abs(
          Number(transaction.retrait) || 0
        ),
        depot: Math.abs(
          Number(transaction.depot) || 0
        ),
        solde:
          transaction.solde === null ||
          transaction.solde === '' ||
          typeof transaction.solde === 'undefined'
            ? null
            : Number(transaction.solde),
        reference: String(
          transaction.reference ||
          ('PDF-' + (index + 1))
        ).trim()
      };
    }
  );

  return importerTransactionsBancaires_(
    normalisees,
    String(
      nomSource || 'Relevé PDF'
    ).trim()
  );
}

function normaliserFichierBancaire_(
  blob,
  nomFichier,
  mimeType
) {
  const nom = String(nomFichier || blob.getName() || '').trim();
  const type = String(
    mimeType || blob.getContentType() || ''
  ).toLowerCase();
  const extension = nom.indexOf('.') >= 0
    ? nom.split('.').pop().toLowerCase()
    : '';
  if (
    extension === 'csv' ||
    extension === 'txt' ||
    type.indexOf('csv') >= 0 ||
    type === 'text/plain'
  ) {
    const texte = decoderTexteCsv_(blob);
    return {
      nomSource: nom || 'Relevé CSV',
      transactions: normaliserCsvDesjardins_(texte, nom)
    };
  }
  if (
    extension === 'pdf' ||
    type === 'application/pdf'
  ) {
    return {
      nomSource: nom || 'Relevé PDF',
      transactions: normaliserPdfDesjardins_(blob, nom)
    };
  }
  if (
    extension === 'xls' ||
    extension === 'xlsx' ||
    type.indexOf('excel') >= 0 ||
    type.indexOf('spreadsheetml') >= 0
  ) {
    return {
      nomSource: nom || 'Relevé Excel',
      transactions: normaliserExcelDesjardins_(blob, nom)
    };
  }
  throw new Error(
    'Format non pris en charge : ' +
    (extension || type || 'inconnu') +
    '. Utilisez CSV, PDF, XLS ou XLSX.'
  );
}
function normaliserLienBancaire_(lien) {
  const id = extraireIdDriveImport_(lien);
  if (!id) {
    throw new Error(
      'Le lien Google Drive / Sheets n’est pas reconnu.'
    );
  }
  const fichier = DriveApp.getFileById(id);
  const mime = String(fichier.getMimeType() || '');
  const nom = fichier.getName() || 'Fichier Drive';
  if (
    mime === 'application/vnd.google-apps.spreadsheet'
  ) {
    return {
      nomSource: nom,
      transactions: normaliserClasseurGoogle_(
        SpreadsheetApp.openById(id),
        nom
      )
    };
  }
  if (mime === 'application/pdf') {
    return {
      nomSource: nom,
      transactions: normaliserPdfDesjardins_(
        fichier.getBlob(),
        nom
      )
    };
  }
  if (
    mime.indexOf('excel') >= 0 ||
    mime.indexOf('spreadsheetml') >= 0
  ) {
    return {
      nomSource: nom,
      transactions: normaliserExcelDesjardins_(
        fichier.getBlob(),
        nom
      )
    };
  }
  if (
    mime.indexOf('csv') >= 0 ||
    mime === 'text/plain'
  ) {
    return {
      nomSource: nom,
      transactions: normaliserCsvDesjardins_(
        decoderTexteCsv_(fichier.getBlob()),
        nom
      )
    };
  }
  throw new Error(
    'Le fichier Drive lié n’est pas un Google Sheet, CSV, Excel ou PDF.'
  );
}
function extraireIdDriveImport_(lien) {
  const texte = String(lien || '').trim();
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
function decoderTexteCsv_(blob) {
  let texte = blob.getDataAsString('UTF-8');
  if (
    texte.indexOf('\uFFFD') >= 0 ||
    texte.indexOf('�') >= 0
  ) {
    texte = blob.getDataAsString('ISO-8859-1');
  }
  return texte;
}
function detecterSeparateurCsv_(texte) {
  const premiere = String(texte || '')
    .split(/\r?\n/)
    .find(function(ligne) {
      return String(ligne || '').trim();
    }) || '';
  const candidats = [',', ';', '\t'];
  let meilleur = ',';
  let meilleurScore = -1;
  candidats.forEach(function(separateur) {
    const score = premiere.split(separateur).length;
    if (score > meilleurScore) {
      meilleur = separateur;
      meilleurScore = score;
    }
  });
  return meilleur;
}
function normaliserCsvDesjardins_(texte, nomSource) {
  const separateur = detecterSeparateurCsv_(texte);
  const lignes = Utilities
    .parseCsv(String(texte || ''), separateur)
    .filter(function(ligne) {
      return ligne.some(function(valeur) {
        return String(valeur || '').trim() !== '';
      });
    });
  if (!lignes.length) {
    throw new Error(
      'Aucune donnée bancaire n’a été trouvée dans le fichier CSV.'
    );
  }
  const premiere = lignes[0];
  // Format CSV Desjardins historique déjà utilisé dans l'application.
  if (
    premiere.length >= 14 &&
    /^\d{4}\/\d{2}\/\d{2}$/.test(
      String(premiere[3] || '').trim()
    )
  ) {
    return lignes
      .map(function(ligne, index) {
        const texteDate = String(ligne[3] || '').trim();
        const morceaux = texteDate.split('/').map(Number);
        if (
          morceaux.length !== 3 ||
          !morceaux[0]
        ) {
          return null;
        }
        const retrait = convertirMontantImport_(ligne[7]);
        const depot = convertirMontantImport_(ligne[8]);
        const solde = convertirMontantImport_(ligne[13]);
        if (
          (!retrait && !depot) ||
          (retrait && depot)
        ) {
          return null;
        }
        return {
          date: new Date(
            morceaux[0],
            morceaux[1] - 1,
            morceaux[2],
            12,
            0,
            0
          ),
          code: String(ligne[2] || '').trim(),
          description: String(ligne[5] || '').trim(),
          retrait: retrait,
          depot: depot,
          solde: solde,
          reference: String(
            ligne[4] || ('CSV-' + (index + 1))
          ).trim()
        };
      })
      .filter(Boolean);
  }
  return normaliserTableauBancaire_(
    lignes,
    nomSource || 'CSV'
  );
}
function normaliserExcelDesjardins_(blob, nomSource) {
  if (
    typeof Drive === 'undefined' ||
    !Drive.Files
  ) {
    throw new Error(
      'Le service avancé Drive API doit être activé pour importer Excel.'
    );
  }
  let temporaire = null;
  try {
    temporaire = Drive.Files.create(
      {
        name: 'IMPORT_TEMP_' + Date.now(),
        mimeType: 'application/vnd.google-apps.spreadsheet'
      },
      blob,
      {
        fields: 'id,name'
      }
    );
    if (!temporaire || !temporaire.id) {
      throw new Error(
        'Google Drive n’a pas pu convertir le fichier Excel.'
      );
    }
    let classeur = null;
    for (let tentative = 0; tentative < 4; tentative++) {
      try {
        if (tentative > 0) {
          Utilities.sleep(500);
        }
        classeur = SpreadsheetApp.openById(temporaire.id);
        break;
      } catch (erreur) {
        classeur = null;
      }
    }
    if (!classeur) {
      throw new Error(
        'Le classeur Excel converti ne peut pas être ouvert.'
      );
    }
    return normaliserClasseurGoogle_(
      classeur,
      nomSource || blob.getName() || 'Excel'
    );
  } finally {
    if (temporaire && temporaire.id) {
      try {
        DriveApp.getFileById(temporaire.id).setTrashed(true);
      } catch (erreurCorbeille) {
        console.error(erreurCorbeille.message);
      }
    }
  }
}
function normaliserClasseurGoogle_(classeur, nomSource) {
  const transactions = [];
  classeur.getSheets().forEach(function(feuille) {
    const plage = feuille.getDataRange();
    if (
      plage.getNumRows() < 2 ||
      plage.getNumColumns() < 2
    ) {
      return;
    }
    try {
      const trouvees = normaliserTableauBancaire_(
        plage.getValues(),
        nomSource + ' / ' + feuille.getName()
      );
      trouvees.forEach(function(transaction) {
        transactions.push(transaction);
      });
    } catch (erreur) {
      // Une feuille non bancaire n'empêche pas les autres d'être analysées.
    }
  });
  if (!transactions.length) {
    throw new Error(
      'Aucune table bancaire reconnue dans le classeur.'
    );
  }
  return transactions;
}
function normaliserTableauBancaire_(valeurs, nomSource) {
  const tableau = valeurs || [];
  if (!tableau.length) {
    throw new Error(
      'Le tableau bancaire est vide.'
    );
  }
  let ligneEntete = -1;
  let indexDate = -1;
  let indexDescription = -1;
  let indexCode = -1;
  let indexFrais = -1;
  let indexRetrait = -1;
  let indexDepot = -1;
  let indexMontant = -1;
  let indexSolde = -1;
  const maximum = Math.min(tableau.length, 30);
  for (let ligne = 0; ligne < maximum; ligne++) {
    const normalises = tableau[ligne].map(normaliserEnteteImport_);
    const date = trouverIndexEntete_(
      normalises,
      ['date']
    );
    const description = trouverIndexEntete_(
      normalises,
      ['description', 'libelle', 'détails', 'details']
    );
    const retrait = trouverIndexEntete_(
      normalises,
      ['retrait', 'debit', 'débit']
    );
    const depot = trouverIndexEntete_(
      normalises,
      ['depot', 'dépôt', 'credit', 'crédit']
    );
    const montant = trouverIndexEntete_(
      normalises,
      ['montant']
    );
    if (
      date >= 0 &&
      description >= 0 &&
      (
        retrait >= 0 ||
        depot >= 0 ||
        montant >= 0
      )
    ) {
      ligneEntete = ligne;
      indexDate = date;
      indexDescription = description;
      indexRetrait = retrait;
      indexDepot = depot;
      indexMontant = montant;
      indexCode = trouverIndexEntete_(
        normalises,
        ['code', 'type']
      );
      indexFrais = trouverIndexEntete_(
        normalises,
        ['frais']
      );
      indexSolde = trouverIndexEntete_(
        normalises,
        ['solde', 'balance']
      );
      break;
    }
  }
  if (ligneEntete < 0) {
    throw new Error(
      'Les colonnes Date / Description / Montant n’ont pas été reconnues dans ' +
      nomSource +
      '.'
    );
  }
  const anneeIndice = extraireAnneeTableau_(
    tableau,
    nomSource
  );
  const transactions = [];
  for (
    let ligne = ligneEntete + 1;
    ligne < tableau.length;
    ligne++
  ) {
    const valeursLigne = tableau[ligne] || [];
    const date = convertirDateImport_(
      valeursLigne[indexDate],
      anneeIndice
    );
    if (!date) {
      continue;
    }
    const description = String(
      valeursLigne[indexDescription] || ''
    )
      .replace(/\s+/g, ' ')
      .trim();
    if (!description) {
      continue;
    }
    let retrait = indexRetrait >= 0
      ? convertirMontantImport_(valeursLigne[indexRetrait])
      : 0;
    let depot = indexDepot >= 0
      ? convertirMontantImport_(valeursLigne[indexDepot])
      : 0;
    const frais = indexFrais >= 0
      ? convertirMontantImport_(valeursLigne[indexFrais])
      : 0;
    if (frais) {
      retrait += frais;
    }
    if (
      indexMontant >= 0 &&
      !retrait &&
      !depot
    ) {
      const montant = convertirMontantSigneImport_(
        valeursLigne[indexMontant]
      );
      if (montant < 0) {
        retrait = Math.abs(montant);
      } else if (montant > 0) {
        depot = montant;
      }
    }
    if (!retrait && !depot) {
      continue;
    }
    transactions.push({
      date: date,
      code: indexCode >= 0
        ? String(valeursLigne[indexCode] || '').trim()
        : '',
      description: description,
      retrait: retrait,
      depot: depot,
      solde: indexSolde >= 0
        ? convertirMontantImport_(valeursLigne[indexSolde])
        : null,
      reference: 'TAB-' + String(ligne + 1)
    });
  }
  if (!transactions.length) {
    throw new Error(
      'Aucune transaction exploitable n’a été trouvée dans ' +
      nomSource +
      '.'
    );
  }
  return transactions;
}
function normaliserEnteteImport_(valeur) {
  return String(valeur || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
function trouverIndexEntete_(entetes, choix) {
  const choixNormalises = choix.map(normaliserEnteteImport_);
  for (let i = 0; i < entetes.length; i++) {
    if (
      choixNormalises.indexOf(entetes[i]) >= 0
    ) {
      return i;
    }
  }
  return -1;
}
function extraireAnneeTableau_(tableau, nomSource) {
  const texte = [
    nomSource || ''
  ]
    .concat(
      tableau
        .slice(0, 20)
        .map(function(ligne) {
          return ligne.join(' ');
        })
    )
    .join(' ');
  const resultat = texte.match(/\b(20\d{2})\b/);
  return resultat
    ? Number(resultat[1])
    : new Date().getFullYear();
}
function convertirDateImport_(valeur, anneeIndice) {
  if (
    Object.prototype.toString.call(valeur) === '[object Date]' &&
    !isNaN(valeur.getTime())
  ) {
    const copie = new Date(valeur);
    copie.setHours(12, 0, 0, 0);
    return copie;
  }
  const texte = String(valeur || '').trim();
  if (!texte) {
    return null;
  }
  let resultat = texte.match(
    /^(20\d{2})[\/-](\d{1,2})[\/-](\d{1,2})$/
  );
  if (resultat) {
    return new Date(
      Number(resultat[1]),
      Number(resultat[2]) - 1,
      Number(resultat[3]),
      12,
      0,
      0
    );
  }
  resultat = texte.match(
    /^(\d{1,2})[\/-](\d{1,2})[\/-](20\d{2})$/
  );
  if (resultat) {
    return new Date(
      Number(resultat[3]),
      Number(resultat[2]) - 1,
      Number(resultat[1]),
      12,
      0,
      0
    );
  }
  resultat = texte.match(
    /^(\d{1,2})\s+([A-Za-zÀ-ÿ]{3,})\s*(20\d{2})?$/
  );
  if (resultat) {
    const mois = numeroMoisImport_(resultat[2]);
    if (mois >= 0) {
      return new Date(
        Number(resultat[3] || anneeIndice),
        mois,
        Number(resultat[1]),
        12,
        0,
        0
      );
    }
  }
  const dateNative = new Date(texte);
  if (!isNaN(dateNative.getTime())) {
    dateNative.setHours(12, 0, 0, 0);
    return dateNative;
  }
  return null;
}
function numeroMoisImport_(mois) {
  const cle = String(mois || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .slice(0, 3);
  const moisMap = {
    JAN: 0,
    FEV: 1,
    MAR: 2,
    AVR: 3,
    MAI: 4,
    JUN: 5,
    JUI: 6,
    AOU: 7,
    SEP: 8,
    OCT: 9,
    NOV: 10,
    DEC: 11
  };
  return Object.prototype.hasOwnProperty.call(moisMap, cle)
    ? moisMap[cle]
    : -1;
}
function convertirMontantSigneImport_(valeur) {
  if (typeof valeur === 'number') {
    return isFinite(valeur) ? valeur : 0;
  }
  let texte = String(valeur || '')
    .trim()
    .replace(/\s/g, '')
    .replace(/\$/g, '');
  if (!texte) {
    return 0;
  }
  const negatifParentheses =
    /^\(.+\)$/.test(texte);
  texte = texte
    .replace(/[()]/g, '')
    .replace(',', '.')
    .replace(/[^0-9.\-]/g, '');
  const nombre = Number(texte) || 0;
  return negatifParentheses
    ? -Math.abs(nombre)
    : nombre;
}
function convertirMontantImport_(valeur) {
  return Math.abs(
    convertirMontantSigneImport_(valeur)
  );
}
function normaliserPdfDesjardins_(blob, nomSource) {
  let documentTemporaireId = '';
  try {
    // 1) Conversion sans OCR : les relevés Desjardins téléchargés
    // contiennent normalement déjà du texte.
    documentTemporaireId = creerDocumentPdfV2_(
      blob,
      nomSource,
      false
    );
    let transactions =
      parserDocumentPdfDesjardins_(
        documentTemporaireId,
        nomSource
      );
    if (transactions.length) {
      return transactions;
    }
    // Secours : essayer aussi le texte continu du document converti.
    let texte = lireTexteDocumentPdfTemporaire_(
      documentTemporaireId
    );
    transactions = parserPdfDesjardins_(
      texte,
      nomSource
    );
    if (transactions.length) {
      return transactions;
    }
    supprimerDocumentPdfTemporaire_(
      documentTemporaireId
    );
    documentTemporaireId = '';
    // 2) Dernier recours : OCR pour un PDF réellement scanné.
    documentTemporaireId = creerDocumentPdfV2_(
      blob,
      nomSource,
      true
    );
    transactions =
      parserDocumentPdfDesjardins_(
        documentTemporaireId,
        nomSource
      );
    if (transactions.length) {
      return transactions;
    }
    texte = lireTexteDocumentPdfTemporaire_(
      documentTemporaireId
    );
    transactions = parserPdfDesjardins_(
      texte,
      nomSource
    );
    if (!transactions.length) {
      throw new Error(
        'Aucune transaction Desjardins n’a été reconnue dans ce PDF.'
      );
    }
    return transactions;
  } finally {
    supprimerDocumentPdfTemporaire_(
      documentTemporaireId
    );
  }
}
function parserDocumentPdfDesjardins_(
  documentId,
  nomSource
) {
  if (!documentId) {
    return [];
  }
  const document = DocumentApp.openById(
    documentId
  );
  const corps = document.getBody();
  const texteGeneral = (
    corps.getText() +
    ' ' +
    String(nomSource || '')
  );
  const anneeResultat =
    texteGeneral.match(
      /\b(20\d{2})\b/
    );
  const annee = anneeResultat
    ? Number(anneeResultat[1])
    : new Date().getFullYear();
  const transactions = [];
  let soldePrecedent = null;
  for (
    let indexElement = 0;
    indexElement < corps.getNumChildren();
    indexElement++
  ) {
    const element = corps.getChild(
      indexElement
    );
    if (
      element.getType() !==
      DocumentApp.ElementType.TABLE
    ) {
      continue;
    }
    const table = element.asTable();
    let indexDate = -1;
    let indexCode = -1;
    let indexDescription = -1;
    let indexFrais = -1;
    let indexRetrait = -1;
    let indexDepot = -1;
    let indexSolde = -1;
    let ligneEntete = -1;
    for (
      let ligne = 0;
      ligne < table.getNumRows();
      ligne++
    ) {
      const row = table.getRow(ligne);
      const cellules = [];
      for (
        let colonne = 0;
        colonne < row.getNumCells();
        colonne++
      ) {
        cellules.push(
          normaliserEnteteImport_(
            row
              .getCell(colonne)
              .getText()
          )
        );
      }
      const date = trouverIndexEntete_(
        cellules,
        ['date']
      );
      const description =
        trouverIndexEntete_(
          cellules,
          ['description']
        );
      const solde = trouverIndexEntete_(
        cellules,
        ['solde']
      );
      if (
        date >= 0 &&
        description >= 0 &&
        solde >= 0
      ) {
        ligneEntete = ligne;
        indexDate = date;
        indexDescription =
          description;
        indexSolde = solde;
        indexCode =
          trouverIndexEntete_(
            cellules,
            ['code']
          );
        indexFrais =
          trouverIndexEntete_(
            cellules,
            ['frais']
          );
        indexRetrait =
          trouverIndexEntete_(
            cellules,
            ['retrait', 'debit']
          );
        indexDepot =
          trouverIndexEntete_(
            cellules,
            ['depot', 'credit']
          );
        break;
      }
    }
    if (ligneEntete < 0) {
      continue;
    }
    for (
      let ligne = ligneEntete + 1;
      ligne < table.getNumRows();
      ligne++
    ) {
      const row = table.getRow(ligne);
      const valeurCellule = function(index) {
        if (
          index < 0 ||
          index >= row.getNumCells()
        ) {
          return '';
        }
        return String(
          row
            .getCell(index)
            .getText() || ''
        )
          .replace(/\s+/g, ' ')
          .trim();
      };
      const texteDate =
        valeurCellule(indexDate);
      const description =
        valeurCellule(
          indexDescription
        );
      const texteSolde =
        valeurCellule(indexSolde);
      if (
        !texteDate &&
        /solde\s+report/i.test(
          description
        )
      ) {
        const soldeReporte =
          convertirMontantImport_(
            texteSolde
          );
        if (soldeReporte) {
          soldePrecedent =
            soldeReporte;
        }
        continue;
      }
      const date =
        convertirDateImport_(
          texteDate,
          annee
        );
      if (!date) {
        continue;
      }
      const code =
        valeurCellule(indexCode);
      const frais =
        convertirMontantImport_(
          valeurCellule(indexFrais)
        );
      let retrait =
        convertirMontantImport_(
          valeurCellule(indexRetrait)
        );
      let depot =
        convertirMontantImport_(
          valeurCellule(indexDepot)
        );
      if (frais) {
        retrait += frais;
      }
      const solde =
        convertirMontantImport_(
          texteSolde
        );
      if (
        !retrait &&
        !depot &&
        solde &&
        soldePrecedent !== null
      ) {
        const variation =
          Math.round(
            (
              solde -
              soldePrecedent
            ) * 100
          ) / 100;
        if (variation < 0) {
          retrait = Math.abs(
            variation
          );
        } else if (variation > 0) {
          depot = variation;
        }
      }
      if (
        (!retrait && !depot) ||
        (retrait && depot)
      ) {
        if (solde) {
          soldePrecedent = solde;
        }
        continue;
      }
      transactions.push({
        date: date,
        code: code,
        description: description,
        retrait: retrait,
        depot: depot,
        solde:
          solde || null,
        reference:
          'PDF-' +
          String(
            transactions.length + 1
          ).padStart(3, '0')
      });
      if (solde) {
        soldePrecedent = solde;
      }
    }
  }
  return transactions;
}
function extraireTextePdfImport_(blob, nomSource) {
  let documentTemporaireId = '';
  try {
    // 1) Premier choix : conversion du PDF texte SANS OCR.
    // Les relevés Desjardins téléchargés contiennent généralement déjà
    // une couche texte exploitable. Cette voie n'utilise pas le quota OCR.
    documentTemporaireId = creerDocumentPdfV2_(
      blob,
      nomSource,
      false
    );
    let texte = lireTexteDocumentPdfTemporaire_(
      documentTemporaireId
    );
    if (String(texte || '').trim()) {
      return texte;
    }
    // Le document sans OCR est vide : nettoyer avant la tentative suivante.
    supprimerDocumentPdfTemporaire_(
      documentTemporaireId
    );
    documentTemporaireId = '';
    // 2) Dernier recours : OCR pour un PDF réellement scanné.
    documentTemporaireId = creerDocumentPdfV2_(
      blob,
      nomSource,
      true
    );
    texte = lireTexteDocumentPdfTemporaire_(
      documentTemporaireId
    );
    if (!String(texte || '').trim()) {
      throw new Error(
        'Aucun texte lisible n’a été extrait du PDF ' +
        (nomSource || '') +
        '.'
      );
    }
    return texte;
  } finally {
    supprimerDocumentPdfTemporaire_(
      documentTemporaireId
    );
  }
}
function lireTexteDocumentPdfTemporaire_(documentId) {
  if (!documentId) {
    return '';
  }
  let texte = '';
  for (let tentative = 0; tentative < 6; tentative++) {
    try {
      if (tentative > 0) {
        Utilities.sleep(700);
      }
      texte = DocumentApp
        .openById(documentId)
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
}
function supprimerDocumentPdfTemporaire_(documentId) {
  if (!documentId) {
    return;
  }
  try {
    DriveApp
      .getFileById(documentId)
      .setTrashed(true);
  } catch (erreurCorbeille) {
    console.error(
      'Impossible de supprimer le document PDF temporaire : ' +
      erreurCorbeille.message
    );
  }
}
function creerDocumentPdfV2_(
  blob,
  nomSource,
  utiliserOcr
) {
  const token = ScriptApp.getOAuthToken();
  if (!token) {
    throw new Error(
      'Impossible d’obtenir l’autorisation Google Drive pour lire le PDF.'
    );
  }
  const boundary =
    '===============OSBL_' +
    Utilities.getUuid().replace(/-/g, '');
  const metadata = JSON.stringify({
    title:
      'IMPORT_PDF_TEMP_' +
      Date.now() +
      '_' +
      nettoyerNomImport_(nomSource || 'releve.pdf')
  });
  const debut =
    '--' +
    boundary +
    '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    metadata +
    '\r\n' +
    '--' +
    boundary +
    '\r\n' +
    'Content-Type: application/pdf\r\n\r\n';
  const fin =
    '\r\n--' +
    boundary +
    '--';
  const octets = []
    .concat(
      Utilities.newBlob(debut).getBytes(),
      blob.getBytes(),
      Utilities.newBlob(fin).getBytes()
    );
  let url =
    'https://www.googleapis.com/upload/drive/v2/files' +
    '?uploadType=multipart' +
    '&convert=true';
  if (utiliserOcr) {
    url +=
      '&ocr=true' +
      '&ocrLanguage=fr';
  } else {
    url += '&ocr=false';
  }
  const reponse = UrlFetchApp.fetch(
    url,
    {
      method: 'post',
      contentType:
        'multipart/related; boundary=' +
        boundary,
      headers: {
        Authorization:
          'Bearer ' +
          token
      },
      payload: octets,
      muteHttpExceptions: true
    }
  );
  const code = reponse.getResponseCode();
  const contenu = reponse.getContentText();
  if (code < 200 || code >= 300) {
    let message = contenu;
    try {
      const erreur = JSON.parse(contenu);
      message =
        erreur &&
        erreur.error &&
        erreur.error.message
          ? erreur.error.message
          : contenu;
    } catch (erreurJson) {
      // Conserver la réponse brute.
    }
    if (
      utiliserOcr &&
      code === 403 &&
      /rate limit exceeded for OCR/i.test(
        String(message || '')
      )
    ) {
      throw new Error(
        'Le PDF semble nécessiter l’OCR, mais Google Drive a atteint sa limite OCR pour ce compte. ' +
        'Essayez plutôt le CSV/Excel du relevé, ou un PDF contenant déjà du texte.'
      );
    }
    throw new Error(
      (
        utiliserOcr
          ? 'Google Drive OCR PDF'
          : 'Google Drive conversion PDF'
      ) +
      ' a refusé le fichier (' +
      code +
      ') : ' +
      message
    );
  }
  let resultat = null;
  try {
    resultat = JSON.parse(contenu);
  } catch (erreurJson) {
    throw new Error(
      'Réponse inattendue de Google Drive pendant la conversion du PDF.'
    );
  }
  const id = String(
    resultat &&
    resultat.id
      ? resultat.id
      : ''
  ).trim();
  if (!id) {
    throw new Error(
      'Google Drive n’a pas retourné de document temporaire pour ce PDF.'
    );
  }
  return id;
}
function nettoyerNomImport_(nom) {
  return String(nom || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
}
function parserPdfDesjardins_(texte, nomSource) {
  let contenu = String(texte || '')
    .replace(/\r/g, '\n')
    .replace(/\u00A0/g, ' ');
  const debut = contenu.search(
    /EPARGNE\s+AVEC\s+OPERATIONS/i
  );
  if (debut >= 0) {
    const suite = contenu.slice(debut + 1);
    const finRelative = suite.search(
      /COMPTE\s+D['’]EPARGNE\s+ET\s+DE\s+PLACEMENT/i
    );
    contenu = finRelative >= 0
      ? contenu.slice(
          debut,
          debut + 1 + finRelative
        )
      : contenu.slice(debut);
  }
  const anneeResultat = (
    String(texte || '') +
    ' ' +
    String(nomSource || '')
  ).match(/\b(20\d{2})\b/);
  const annee = anneeResultat
    ? Number(anneeResultat[1])
    : new Date().getFullYear();
  // On aplati les retours de ligne : Google Drive peut réorganiser
  // les colonnes PDF sur plusieurs lignes lors de la conversion.
  const aplati = contenu
    .replace(/\s+/g, ' ')
    .trim();
  const soldeReporteResultat = aplati.match(
    /Solde\s+report\w*[\s\S]{0,250}?((?:\d{1,3}(?:[ ]\d{3})+|\d+)[.,]\d{2})/i
  );
  let soldePrecedent = soldeReporteResultat
    ? convertirMontantImport_(soldeReporteResultat[1])
    : null;
  const motifTransaction = new RegExp(
    '(?:^|\\s)' +
    '(\\d{1,2})\\s+' +
    '([A-Za-zÀ-ÿ]{3,9})\\s+' +
    '([A-Z0-9]{1,5})' +
    '(?=\\s)',
    'g'
  );
  const reperes = [];
  let resultat = null;
  while (
    (resultat = motifTransaction.exec(aplati)) !== null
  ) {
    const mois = numeroMoisImport_(resultat[2]);
    if (mois < 0) {
      continue;
    }
    reperes.push({
      index: resultat.index,
      fin: motifTransaction.lastIndex,
      jour: Number(resultat[1]),
      mois: mois,
      libelleMois: resultat[2],
      code: String(resultat[3] || '').trim()
    });
  }
  if (!reperes.length) {
    throw new Error(
      'Le texte du PDF a bien été lu, mais aucune date de transaction Desjardins n’a été reconnue.'
    );
  }
  const transactions = [];
  for (let i = 0; i < reperes.length; i++) {
    const repere = reperes[i];
    const prochain = reperes[i + 1];
    const bloc = aplati
      .slice(
        repere.fin,
        prochain
          ? prochain.index
          : aplati.length
      )
      .trim();
    if (!bloc) {
      continue;
    }
    const montantsTexte = bloc.match(
      /(?:\d{1,3}(?:[ ]\d{3})+|\d+)[.,]\d{2}/g
    ) || [];
    if (!montantsTexte.length) {
      continue;
    }
    const montants = montantsTexte.map(
      convertirMontantImport_
    );
    const analyse = analyserMontantsBlocPdf_(
      montants,
      soldePrecedent
    );
    if (!analyse) {
      continue;
    }
    const premierMontant = bloc.search(
      /(?:\d{1,3}(?:[ ]\d{3})+|\d+)[.,]\d{2}/
    );
    const description = (
      premierMontant >= 0
        ? bloc.slice(0, premierMontant)
        : bloc
    )
      .replace(/\s+/g, ' ')
      .replace(/\s*\/\s*/g, ' / ')
      .trim();
    if (!description) {
      continue;
    }
    transactions.push({
      date: new Date(
        annee,
        repere.mois,
        repere.jour,
        12,
        0,
        0
      ),
      code: repere.code,
      description: description,
      retrait:
        analyse.montantSigne < 0
          ? Math.abs(analyse.montantSigne)
          : 0,
      depot:
        analyse.montantSigne > 0
          ? analyse.montantSigne
          : 0,
      solde: analyse.solde,
      reference:
        'PDF-' +
        String(transactions.length + 1)
          .padStart(3, '0')
    });
    soldePrecedent = analyse.solde;
  }
  if (!transactions.length) {
    throw new Error(
      reperes.length +
      ' ligne(s) de transaction ont été repérées dans le PDF, mais aucune n’a pu être reconstruite.'
    );
  }
  return transactions;
}
function analyserMontantsBlocPdf_(
  montants,
  soldePrecedent
) {
  if (!montants || !montants.length) {
    return null;
  }
  // Si on connaît le solde précédent, on cherche le montant qui peut être
  // le nouveau solde et dont la variation correspond à un autre montant du bloc.
  if (soldePrecedent !== null) {
    let meilleur = null;
    for (let indexSolde = 0; indexSolde < montants.length; indexSolde++) {
      const solde = Number(montants[indexSolde]) || 0;
      const variation =
        Math.round(
          (solde - Number(soldePrecedent)) * 100
        ) / 100;
      if (!variation) {
        continue;
      }
      const variationAbsolue = Math.abs(variation);
      let correspondance = false;
      for (let j = 0; j < montants.length; j++) {
        if (j === indexSolde) {
          continue;
        }
        if (
          Math.abs(
            Number(montants[j]) -
            variationAbsolue
          ) < 0.011
        ) {
          correspondance = true;
          break;
        }
      }
      const score =
        (correspondance ? 100 : 0) +
        indexSolde;
      if (
        !meilleur ||
        score > meilleur.score
      ) {
        meilleur = {
          score: score,
          solde: solde,
          montantSigne: variation
        };
      }
    }
    if (meilleur) {
      return {
        solde: meilleur.solde,
        montantSigne:
          Math.round(
            meilleur.montantSigne * 100
          ) / 100
      };
    }
  }
  // Secours : le dernier montant du bloc est normalement le solde,
  // l’avant-dernier est la transaction.
  if (montants.length >= 2) {
    const solde =
      Number(
        montants[
          montants.length - 1
        ]
      ) || 0;
    const montant =
      Number(
        montants[
          montants.length - 2
        ]
      ) || 0;
    return {
      solde: solde,
      montantSigne: montant
    };
  }
  return null;
}
function importerTransactionsBancaires_(
  transactions,
  nomSource
) {
  const liste = (transactions || [])
    .filter(function(transaction) {
      return transaction &&
        transaction.date &&
        (
          Number(transaction.retrait) ||
          Number(transaction.depot)
        );
    });
  if (!liste.length) {
    throw new Error(
      'Aucune transaction bancaire exploitable n’a été trouvée.'
    );
  }
  const ss = SpreadsheetApp.getActive();
  const feuille = ss.getSheetByName('Import bancaire');
  const soldes = ss.getSheetByName('Soldes ouverture');
  if (!feuille || !soldes) {
    throw new Error(
      'Import bancaire ou Soldes ouverture est introuvable.'
    );
  }
  const fuseau = ss.getSpreadsheetTimeZone();
  const derniereLigneDonnees =
    trouverDerniereLigneDonneesImport_(feuille);

  const dernierIndex = Math.max(
    derniereLigneDonnees,
    6
  );
  const existantes = feuille
    .getRange(
      6,
      2,
      dernierIndex - 5,
      4
    )
    .getValues();
  const clesExistantes = new Set();
  existantes.forEach(function(ligne) {
    const date = ligne[0];
    const description = String(ligne[1] || '');
    const montant = Number(ligne[2]) || 0;
    const solde = ligne[3];
    if (!date && !description && !montant) {
      return;
    }
    clesExistantes.add(
      creerCleTransactionImport_(
        date,
        description,
        montant,
        solde,
        fuseau
      )
    );
  });
  const maintenant = new Date();
  const nouvellesLignes = [];
  let doublons = 0;
  let rejetees = 0;
  let totalRetraits = 0;
  let totalDepots = 0;
  liste.forEach(function(transaction, index) {
    const date = transaction.date;
    if (
      !(date instanceof Date) ||
      isNaN(date.getTime())
    ) {
      rejetees++;
      return;
    }
    const retrait =
      Math.round(
        (Number(transaction.retrait) || 0) * 100
      ) / 100;
    const depot =
      Math.round(
        (Number(transaction.depot) || 0) * 100
      ) / 100;
    if (
      (!retrait && !depot) ||
      (retrait && depot)
    ) {
      rejetees++;
      return;
    }
    const montant = depot - retrait;
    const solde = transaction.solde === null ||
      transaction.solde === '' ||
      typeof transaction.solde === 'undefined'
      ? ''
      : Math.round(
          Number(transaction.solde) * 100
        ) / 100;
    const description = String(
      transaction.description || ''
    )
      .replace(/\s+/g, ' ')
      .trim();
    const cle = creerCleTransactionImport_(
      date,
      description,
      montant,
      solde,
      fuseau
    );
    if (clesExistantes.has(cle)) {
      doublons++;
      return;
    }
    clesExistantes.add(cle);
    const empreinte = calculerEmpreinte_(cle);
    totalRetraits += retrait;
    totalDepots += depot;
    const texteDate = Utilities.formatDate(
      date,
      fuseau,
      'yyyy/MM/dd'
    );
    const code = String(
      transaction.code || ''
    ).trim();
    const reference = String(
      transaction.reference ||
      ('L' + (index + 1))
    ).trim();
    const identifiant =
      'B-' +
      texteDate.replace(/\//g, '') +
      '-' +
      empreinte.slice(0, 10).toUpperCase();
    nouvellesLignes.push([
      identifiant,
      dateVersNumeroSerieSheets_(date, fuseau),
      description,
      montant,
      solde === '' ? '' : solde,
      [code, reference]
        .filter(Boolean)
        .join('-'),
      empreinte,
      '',
      '',
      '',
      'À classer',
      maintenant,
      nomSource || 'Relevé bancaire'
    ]);
  });
  if (!nouvellesLignes.length) {
    return {
      source: nomSource || 'Relevé bancaire',
      importees: 0,
      doublons: doublons,
      rejetees: rejetees,
      message:
        'Aucune nouvelle transaction à importer.'
    };
  }
  const verrou = LockService.getDocumentLock();
  if (!verrou.tryLock(30000)) {
    throw new Error(
      'Une autre opération est déjà en cours. Réessayez dans quelques secondes.'
    );
  }
  try {
    const ligneDepart = Math.max(
      6,
      derniereLigneDonnees + 1
    );
    const ligneFinale =
      ligneDepart +
      nouvellesLignes.length -
      1;
    if (
      ligneFinale >
      feuille.getMaxRows()
    ) {
      const ancienneMax =
        feuille.getMaxRows();
      feuille.insertRowsAfter(
        ancienneMax,
        ligneFinale - ancienneMax
      );
    }
    feuille
      .getRange(
        ligneDepart,
        1,
        nouvellesLignes.length,
        13
      )
      .setValues(nouvellesLignes);

    // Toujours preparer les colonnes V2 W:AG pour chaque nouveau lot importe.
    // La ligne 6 sert de modele pour les formules, validations et formats.
    // Cette copie doit se faire meme si la feuille possede deja assez de lignes.
    feuille
      .getRange(
        6,
        23,
        1,
        11
      )
      .copyTo(
        feuille.getRange(
          ligneDepart,
          23,
          nouvellesLignes.length,
          11
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );
    feuille
      .getRange(
        ligneDepart,
        2,
        nouvellesLignes.length,
        1
      )
      .setNumberFormat('yyyy-mm-dd');
    feuille
      .getRange(
        ligneDepart,
        4,
        nouvellesLignes.length,
        2
      )
      .setNumberFormat(
        '$#,##0.00;[Red]-$#,##0.00'
      );
    feuille
      .getRange(
        ligneDepart,
        12,
        nouvellesLignes.length,
        1
      )
      .setNumberFormat(
        'yyyy-mm-dd hh:mm'
      );
    const resultRegles =
      appliquerReglesBancairesV2Import_(
        feuille,
        ligneDepart,
        nouvellesLignes.length,
        ss
      );
    const premiereTransaction =
      nouvellesLignes[0];
    const derniereTransaction =
      nouvellesLignes[
        nouvellesLignes.length - 1
      ];
    let soldeInitialCalcule = null;
    let soldeFinal = null;
    let ecartOuverture = null;
    if (
      premiereTransaction[4] !== '' &&
      derniereTransaction[4] !== ''
    ) {
      soldeInitialCalcule =
        Number(premiereTransaction[4]) -
        Number(premiereTransaction[3]);
      soldeFinal =
        Number(derniereTransaction[4]);
      const soldeOuverture =
        Number(
          soldes
            .getRange('C6')
            .getValue()
        ) || 0;
      ecartOuverture =
        Math.round(
          (
            soldeInitialCalcule -
            soldeOuverture
          ) * 100
        ) / 100;
    }
    SpreadsheetApp.flush();
    return {
      source: nomSource || 'Relevé bancaire',
      importees: nouvellesLignes.length,
      doublons: doublons,
      rejetees: rejetees,
      totalRetraits: totalRetraits,
      totalDepots: totalDepots,
      soldeInitial: soldeInitialCalcule,
      soldeFinal: soldeFinal,
      ecartOuverture: ecartOuverture,
      avertissementRegles:
        resultRegles &&
        resultRegles.avertissement
          ? resultRegles.avertissement
          : null
    };
  } finally {
    verrou.releaseLock();
  }
}

function trouverDerniereLigneDonneesImport_(feuille) {
  const premiereLigne = 6;
  const nombreLignes =
    feuille.getMaxRows() - premiereLigne + 1;

  if (nombreLignes <= 0) {
    return premiereLigne - 1;
  }

  const valeurs = feuille
    .getRange(
      premiereLigne,
      1,
      nombreLignes,
      1
    )
    .getDisplayValues();

  for (
    let index = valeurs.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(valeurs[index][0] || '').trim()
    ) {
      return premiereLigne + index;
    }
  }

  return premiereLigne - 1;
}

function dateVersNumeroSerieSheets_(date, fuseau) {
  const texte = Utilities.formatDate(
    date,
    fuseau || Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );

  const morceaux = texte
    .split('-')
    .map(Number);

  return (
    Date.UTC(
      morceaux[0],
      morceaux[1] - 1,
      morceaux[2]
    ) /
      86400000 +
    25569
  );
}

function creerCleTransactionImport_(
  date,
  description,
  montant,
  solde,
  fuseau
) {
  let texteDate = '';
  if (
    date instanceof Date &&
    !isNaN(date.getTime())
  ) {
    texteDate = Utilities.formatDate(
      date,
      fuseau || Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  } else {
    texteDate = String(date || '').trim();
  }
  const montantNormalise =
    Math.round(
      (Number(montant) || 0) * 100
    ) / 100;
  const soldeDisponible =
    solde !== '' &&
    solde !== null &&
    typeof solde !== 'undefined' &&
    isFinite(Number(solde));
  if (soldeDisponible) {
    return [
      texteDate,
      montantNormalise.toFixed(2),
      (
        Math.round(
          Number(solde) * 100
        ) / 100
      ).toFixed(2)
    ].join('|');
  }
  return [
    texteDate,
    montantNormalise.toFixed(2),
    normaliserEnteteImport_(description)
  ].join('|');
}
function calculerEmpreinte_(texte) {
  return Utilities
    .computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(texte || ''),
      Utilities.Charset.UTF_8
    )
    .map(function(octet) {
      const valeur =
        octet < 0
          ? octet + 256
          : octet;
      return (
        '0' +
        valeur.toString(16)
      ).slice(-2);
    })
    .join('');
}

