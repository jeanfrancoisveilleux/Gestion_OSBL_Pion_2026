function afficherImportBancaire() {
  const interface = HtmlService
    .createHtmlOutputFromFile('ImportCSV')
    .setWidth(520)
    .setHeight(470);

  SpreadsheetApp.getUi().showModalDialog(
    interface,
    'Importer un relevé Desjardins'
  );
}

function importerCsvDesjardins(contenu, nomFichier) {
  if (!contenu) {
    throw new Error('Le fichier CSV est vide.');
  }

  const classeur = SpreadsheetApp.getActive();
  const feuille = classeur.getSheetByName('Import bancaire');
  const soldes = classeur.getSheetByName('Soldes ouverture');

  const lignes = Utilities.parseCsv(contenu, ',')
    .filter(function(ligne) {
      return ligne.some(function(valeur) {
        return String(valeur).trim() !== '';
      });
    });

  if (!lignes.length) {
    throw new Error('Aucune transaction n’a été trouvée.');
  }

  const premiere = lignes[0];

  if (
    premiere.length < 14 ||
    !/^\d{4}\/\d{2}\/\d{2}$/.test(
      String(premiere[3]).trim()
    )
  ) {
    throw new Error(
      'Le fichier ne correspond pas au format CSV Desjardins attendu.'
    );
  }

  const dernierIndex = Math.max(feuille.getLastRow(), 6);

  const empreintesExistantes = new Set(
    feuille
      .getRange(6, 7, dernierIndex - 5, 1)
      .getDisplayValues()
      .flat()
      .filter(String)
  );

  const maintenant = new Date();
  const nouvellesLignes = [];

  let doublons = 0;
  let rejetees = 0;
  let totalRetraits = 0;
  let totalDepots = 0;

  lignes.forEach(function(ligne) {
    const texteDate = String(ligne[3] || '').trim();
    const morceauxDate = texteDate.split('/').map(Number);

    const retrait = convertirMontantCsv_(ligne[7]);
    const depot = convertirMontantCsv_(ligne[8]);
    const solde = convertirMontantCsv_(ligne[13]);

    if (
      morceauxDate.length !== 3 ||
      !morceauxDate[0] ||
      (!retrait && !depot) ||
      (retrait && depot)
    ) {
      rejetees++;
      return;
    }

    const date = new Date(
      morceauxDate[0],
      morceauxDate[1] - 1,
      morceauxDate[2],
      12,
      0,
      0
    );

    const description = String(ligne[5] || '').trim();
    const sequence = String(ligne[4] || '').trim();
    const code = String(ligne[2] || '').trim();
    const montant = depot - retrait;

    const sourceEmpreinte = [
      texteDate,
      sequence,
      description,
      retrait.toFixed(2),
      depot.toFixed(2),
      solde.toFixed(2)
    ].join('|');

    const empreinte = calculerEmpreinte_(sourceEmpreinte);

    if (empreintesExistantes.has(empreinte)) {
      doublons++;
      return;
    }

    empreintesExistantes.add(empreinte);
    totalRetraits += retrait;
    totalDepots += depot;

    const identifiant =
      'B-' +
      texteDate.replace(/\//g, '') +
      '-' +
      sequence +
      '-' +
      empreinte.slice(0, 6).toUpperCase();

    nouvellesLignes.push([
      identifiant,
      date,
      description,
      montant,
      solde,
      code + '-' + sequence,
      empreinte,
      '',
      '',
      '',
      'À classer',
      maintenant,
      nomFichier || 'Relevé Desjardins'
    ]);
  });

  if (!nouvellesLignes.length) {
    return {
      importees: 0,
      doublons: doublons,
      rejetees: rejetees,
      message: 'Aucune nouvelle transaction à importer.'
    };
  }

  const ligneDepart = feuille.getLastRow() + 1;
  const ligneFinale =
    ligneDepart + nouvellesLignes.length - 1;

  if (ligneFinale > feuille.getMaxRows()) {
    feuille.insertRowsAfter(
      feuille.getMaxRows(),
      ligneFinale - feuille.getMaxRows()
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
    .setNumberFormat('$#,##0.00;[Red]-$#,##0.00');

  feuille
    .getRange(
      ligneDepart,
      12,
      nouvellesLignes.length,
      1
    )
    .setNumberFormat('yyyy-mm-dd hh:mm');

  // Appliquer les règles bancaires aux nouvelles lignes (colonnes P:S)
  appliquerReglesBancairesImport_(
    feuille,
    ligneDepart,
    nouvellesLignes.length,
    classeur
  );

  const premiereTransaction = nouvellesLignes[0];
  const derniereTransaction =
    nouvellesLignes[nouvellesLignes.length - 1];

  const premierMontant =
    Number(premiereTransaction[3]) || 0;

  const premierSolde =
    Number(premiereTransaction[4]) || 0;

  const soldeInitialCalcule =
    premierSolde - premierMontant;

  const soldeOuverture =
    Number(soldes.getRange('C6').getValue()) || 0;

  const ecartOuverture =
    Math.round(
      (soldeInitialCalcule - soldeOuverture) * 100
    ) / 100;

  SpreadsheetApp.flush();

  return {
    importees: nouvellesLignes.length,
    doublons: doublons,
    rejetees: rejetees,
    totalRetraits: totalRetraits,
    totalDepots: totalDepots,
    soldeInitial: soldeInitialCalcule,
    soldeFinal: derniereTransaction[4],
    ecartOuverture: ecartOuverture
  };
}

function convertirMontantCsv_(valeur) {
  const texte = String(valeur || '')
    .replace(/\s/g, '')
    .replace(',', '.')
    .replace('$', '');

  return Number(texte) || 0;
}

function calculerEmpreinte_(texte) {
  return Utilities
    .computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      texte,
      Utilities.Charset.UTF_8
    )
    .map(function(octet) {
      const valeur = octet < 0 ? octet + 256 : octet;
      return ('0' + valeur.toString(16)).slice(-2);
    })
    .join('');
}

// Applique les règles bancaires aux nouvelles lignes importées (suggestions P:S).
// Ne modifie pas les colonnes H, I ni les lignes déjà comptabilisées.
// Ne fait jamais échouer l'import si les onglets ou la zone T:Z ne sont pas installés.
function appliquerReglesBancairesImport_(feuille, ligneDepart, nombreLignes, classeur) {
  if (!nombreLignes) return;

  try {
    const ss = classeur || SpreadsheetApp.getActive();
    const donnees = feuille
      .getRange(ligneDepart, 1, nombreLignes, 13)
      .getValues();

    for (let i = 0; i < nombreLignes; i++) {
      const statut = String(donnees[i][10] || '').trim();
      if (statut !== 'À classer') continue;

      const description = String(donnees[i][2] || '').trim();
      const montant = Number(donnees[i][3]) || 0;

      const regle = rechercherRegleBancaire_(ss, description, montant);
      if (!regle) continue;

      const ligneSheet = ligneDepart + i;

      if (regle.idFournisseur) {
        feuille.getRange(ligneSheet, 16).setNumberFormat('@').setValue(regle.idFournisseur);
        feuille.getRange(ligneSheet, 17).setValue(regle.nomFournisseur || '');
      }
      if (regle.idContact) {
        feuille.getRange(ligneSheet, 18).setNumberFormat('@').setValue(regle.idContact);
        feuille.getRange(ligneSheet, 19).setValue(regle.nomContact || '');
      }
    }
  } catch (e) {
    // Onglet Fournisseurs, Contacts ou Configuration!T:Z absent : ne pas bloquer l'import
  }
}