const CONFIG_REGLES_V2_ = {
  feuille: 'Règles bancaires',
  premiereLigne: 6,
  nombreColonnes: 13
};

function chargerReglesBancairesActivesV2_(ss) {
  const feuille = ss.getSheetByName(CONFIG_REGLES_V2_.feuille);

  if (!feuille || feuille.getLastRow() < CONFIG_REGLES_V2_.premiereLigne) {
    return [];
  }

  return feuille
    .getRange(
      CONFIG_REGLES_V2_.premiereLigne,
      1,
      feuille.getLastRow() - CONFIG_REGLES_V2_.premiereLigne + 1,
      CONFIG_REGLES_V2_.nombreColonnes
    )
    .getDisplayValues()
    .map(function(ligne, index) {
      return {
        ordre: index,
        idRegle: String(ligne[0] || '').trim(),
        actif: String(ligne[1] || '').trim(),
        priorite: Number(ligne[2] || 0),
        sens: String(ligne[3] || '').trim(),
        correspondance: String(ligne[4] || '').trim(),
        texte: String(ligne[5] || '').trim(),
        typeClassement: String(ligne[6] || '').trim(),
        composante: String(ligne[7] || '').trim(),
        idFournisseur: String(ligne[8] || '').trim(),
        idContact: String(ligne[9] || '').trim(),
        codeCompte: String(ligne[10] || '').trim(),
        programme: String(ligne[11] || '').trim(),
        notes: String(ligne[12] || '').trim()
      };
    })
    .filter(function(regle) {
      return regle.idRegle &&
        regle.actif === 'Oui' &&
        regle.texte;
    })
    .sort(function(a, b) {
      return b.priorite - a.priorite ||
        a.ordre - b.ordre;
    });
}

function normaliserRegleBancaireV2_(valeur) {
  return String(valeur || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function rechercherRegleBancaireV2_(regles, description, montant) {
  const sens = Number(montant) >= 0 ? 'Entrée' : 'Sortie';
  const texteDescription = normaliserRegleBancaireV2_(description);

  for (let index = 0; index < regles.length; index++) {
    const regle = regles[index];

    if (regle.sens && regle.sens !== sens) {
      continue;
    }

    const texteRegle = normaliserRegleBancaireV2_(regle.texte);

    const correspond =
      regle.correspondance === 'Exacte'
        ? texteDescription === texteRegle
        : texteDescription.indexOf(texteRegle) !== -1;

    if (correspond) {
      return regle;
    }
  }

  return null;
}

function dictionnaireV2_(ss, nomFeuille, colonneCle, colonneValeur) {
  const feuille = ss.getSheetByName(nomFeuille);
  const resultat = {};

  if (!feuille || feuille.getLastRow() < 6) {
    return resultat;
  }

  const largeur = Math.max(colonneCle, colonneValeur);

  feuille
    .getRange(6, 1, feuille.getLastRow() - 5, largeur)
    .getDisplayValues()
    .forEach(function(ligne) {
      const cle = String(ligne[colonneCle - 1] || '').trim();
      const valeur = String(ligne[colonneValeur - 1] || '').trim();

      if (cle) {
        resultat[cle] = valeur;
      }
    });

  return resultat;
}

function chargerComposantesActivesReglesV2_(ss) {
  const configuration = ss.getSheetByName('Configuration');
  const resultat = {};

  if (!configuration || configuration.getLastRow() < 50) {
    return resultat;
  }

  configuration
    .getRange(50, 28, configuration.getLastRow() - 49, 3)
    .getDisplayValues()
    .forEach(function(ligne) {
      const actif = String(ligne[0] || '').trim();
      const libelle = String(ligne[2] || '').trim();

      if (actif === 'Oui' && libelle) {
        resultat[libelle] = true;
      }
    });

  return resultat;
}

function appliquerReglesBancairesV2Import_(feuille, ligneDepart, nombreLignes, ss) {
  if (!nombreLignes) {
    return { avertissement: null };
  }

  const regles = chargerReglesBancairesActivesV2_(ss);

  if (!regles.length) {
    return { avertissement: null };
  }

  // Structure actuelle Import bancaire : A:AE = 31 colonnes.
  const donnees = feuille
    .getRange(ligneDepart, 1, nombreLignes, 31)
    .getValues();

  const comptes = dictionnaireV2_(ss, 'Comptes', 1, 2);
  const fournisseurs = dictionnaireV2_(ss, 'Fournisseurs', 1, 2);
  const contacts = dictionnaireV2_(ss, 'Contacts', 1, 4);

  const sortieHI = [];
  const sortiePV = [];

  donnees.forEach(function(ligne) {
    const idImport = String(ligne[0] || '').trim();
    const description = String(ligne[2] || '').trim();
    const montant = Number(ligne[3] || 0);
    const traitement = String(ligne[25] || '').trim();

    if (!idImport || traitement) {
      sortieHI.push([
        ligne[7] || '',
        ligne[8] || ''
      ]);

      sortiePV.push([
        ligne[13] || '',
        ligne[14] || '',
        ligne[15] || '',
        ligne[16] || '',
        ligne[17] || '',
        ligne[18] || '',
        ligne[19] || ''
      ]);
      return;
    }

    const regle = rechercherRegleBancaireV2_(
      regles,
      description,
      montant
    );

    if (!regle) {
      sortieHI.push(['', '']);
      sortiePV.push(['', '', '', '', '', '', '']);
      return;
    }

    sortieHI.push([
      regle.codeCompte
        ? (comptes[regle.codeCompte] || '')
        : '',
      regle.programme || ''
    ]);

    sortiePV.push([
      regle.idFournisseur || '',
      regle.idFournisseur
        ? (fournisseurs[regle.idFournisseur] || '')
        : '',
      regle.idContact || '',
      regle.idContact
        ? (contacts[regle.idContact] || '')
        : '',
      regle.typeClassement || '',
      regle.composante || '',
      regle.idRegle || ''
    ]);
  });

  feuille
    .getRange(ligneDepart, 8, nombreLignes, 2)
    .setNumberFormat('@')
    .setValues(sortieHI);

  feuille
    .getRange(ligneDepart, 14, nombreLignes, 7)
    .setNumberFormat('@')
    .setValues(sortiePV);

  return { avertissement: null };
}

function reappliquerReglesBancairesAuxTransactionsAClasser() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName('Import bancaire');

  if (!feuille) {
    throw new Error('L’onglet Import bancaire est introuvable.');
  }

  const derniereLigne = feuille.getLastRow();

  if (derniereLigne < 6) {
    SpreadsheetApp.getUi().alert('Aucune transaction bancaire.');
    return;
  }

  appliquerReglesBancairesV2Import_(
    feuille,
    6,
    derniereLigne - 5,
    ss
  );

  SpreadsheetApp.getUi().alert(
    'Les règles ont été réappliquées aux transactions non traitées.'
  );
}

function validerReglesBancairesConfigurables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName('Règles bancaires');

  if (!feuille) {
    throw new Error('L’onglet Règles bancaires est introuvable.');
  }

  const comptes = dictionnaireV2_(ss, 'Comptes', 1, 2);
  const programmes = dictionnaireV2_(ss, 'Programmes', 2, 1);
  const fournisseurs = dictionnaireV2_(ss, 'Fournisseurs', 1, 2);
  const contacts = dictionnaireV2_(ss, 'Contacts', 1, 4);
  const composantes = chargerComposantesActivesReglesV2_(ss);

  const erreurs = [];
  const ids = new Set();

  if (feuille.getLastRow() >= 6) {
    feuille
      .getRange(6, 1, feuille.getLastRow() - 5, 13)
      .getDisplayValues()
      .forEach(function(ligne, index) {
        const numeroLigne = index + 6;
        const id = String(ligne[0] || '').trim();
        const actif = String(ligne[1] || '').trim();

        if (!id && actif !== 'Oui') {
          return;
        }

        if (!id) {
          erreurs.push('Ligne ' + numeroLigne + ' : ID règle manquant.');
          return;
        }

        if (ids.has(id)) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : ID règle en double (' + id + ').'
          );
        }
        ids.add(id);

        const sens = String(ligne[3] || '').trim();
        const correspondance = String(ligne[4] || '').trim();
        const texte = String(ligne[5] || '').trim();
        const typeClassement = String(ligne[6] || '').trim();
        const composante = String(ligne[7] || '').trim();
        const fournisseur = String(ligne[8] || '').trim();
        const contact = String(ligne[9] || '').trim();
        const compte = String(ligne[10] || '').trim();
        const programme = String(ligne[11] || '').trim();

        if (
          actif === 'Oui' &&
          !['Entrée', 'Sortie'].includes(sens)
        ) {
          erreurs.push('Ligne ' + numeroLigne + ' : sens invalide.');
        }

        if (
          actif === 'Oui' &&
          !['Contient', 'Exacte'].includes(correspondance)
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne + ' : correspondance invalide.'
          );
        }

        if (actif === 'Oui' && !texte) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : texte à reconnaître manquant.'
          );
        }

        if (
          typeClassement &&
          ![
            'Dépense',
            'Revenu comptable direct',
            'Entrées, forfaits ou marchandises'
          ].includes(typeClassement)
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : type de classement invalide.'
          );
        }

        if (compte && comptes[compte] === undefined) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : compte ' + compte + ' introuvable.'
          );
        }

        if (
          programme &&
          programmes[programme] === undefined
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : programme « ' + programme + ' » introuvable.'
          );
        }

        if (
          fournisseur &&
          fournisseurs[fournisseur] === undefined
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : fournisseur ' + fournisseur + ' introuvable.'
          );
        }

        if (
          contact &&
          contacts[contact] === undefined
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : contact ' + contact + ' introuvable.'
          );
        }

        if (
          composante &&
          composantes[composante] !== true
        ) {
          erreurs.push(
            'Ligne ' + numeroLigne +
            ' : composante « ' + composante +
            ' » absente ou inactive.'
          );
        }
      });
  }

  SpreadsheetApp.getUi().alert(
    erreurs.length
      ? (
          'Validation terminée avec ' +
          erreurs.length +
          ' problème(s).\n\n' +
          erreurs.slice(0, 20).join('\n')
        )
      : (
          'Validation terminée.\n\n' +
          'Toutes les règles actives sont cohérentes avec les référentiels V2.'
        )
  );

  return erreurs;
}
