// ─── Installer public ────────────────────────────────────────────────────────

function installerGestionFournisseursContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  installerOngletFournisseurs_(ss);
  installerOngletContacts_(ss);
  installerPlanComptableSupplementaire_(ss);
  installerReglesBancaires_(ss);
  installerColonnesFournisseurTransactions_(ss);
  installerColonnesFournisseurJournal_(ss);
  installerColonnesFournisseurImportBancaire_(ss);

  const lignesReparees = repairerSyncFournisseurImportClassee_(ss);

  SpreadsheetApp.getUi().alert(
    'Gestion des fournisseurs et contacts installée avec succès.\n\n' +
    'Onglets créés : Fournisseurs, Contacts.\n' +
    'Comptes ajoutés : 4080, 6065 (si absents).\n' +
    'Règles bancaires configurées dans Configuration!T:Z.\n' +
    'Colonnes fournisseur ajoutées dans Transactions (R:T), Journal (O:P) ' +
    'et Import bancaire (P:S).\n' +
    (lignesReparees > 0
      ? lignesReparees + ' ligne(s) de l\'import bancaire réparée(s) avec le fournisseur confirmé.'
      : 'Aucune ligne à réparer dans l\'import bancaire.')
  );
}

// ─── Onglet Fournisseurs ──────────────────────────────────────────────────────

function installerOngletFournisseurs_(ss) {
  let feuille = ss.getSheetByName('Fournisseurs');

  if (!feuille) {
    feuille = ss.insertSheet('Fournisseurs');
    feuille.setHiddenGridlines(true);
  }

  // En-tête de section (ligne 1)
  if (!String(feuille.getRange(1, 1).getValue() || '').trim()) {
    feuille.getRange('A1:L1').merge()
      .setValue('Fournisseurs')
      .setBackground('#e8f0fe')
      .setFontWeight('bold');
  }

  // En-têtes de colonnes (ligne 5)
  const entetes = [
    'ID fournisseur', 'Fournisseur', 'Nom légal', 'Courriel général',
    'Téléphone', 'Adresse', 'Ville', 'Province', 'Pays', 'Code postal',
    'Actif', 'Notes'
  ];

  const cellEntetes = feuille.getRange(5, 1, 1, entetes.length);

  if (!String(cellEntetes.getValues()[0][0] || '').trim()) {
    cellEntetes
      .setValues([entetes])
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  // Fournisseurs initiaux
  const fournisseursInitiaux = [
    ['FOU-0000', 'À déterminer'],
    ['FOU-0001', 'Desjardins'],
    ['FOU-0002', 'Infini-jeux'],
    ['FOU-0003', 'Cartier Resto-Bar'],
    ['FOU-0004', 'Tim Hortons'],
    ['FOU-0005', "Boutique L'Imaginaire"],
    ['FOU-0006', 'Amusement Jacques-Cartier'],
    ['FOU-0007', 'Au Chalet en bois rond']
  ];

  const idsExistants = new Set();
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne >= 6) {
    feuille.getRange(6, 1, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const id = String(ligne[0] || '').trim();
        if (id) { idsExistants.add(id); }
      });
  }

  fournisseursInitiaux.forEach(function(f) {
    if (!idsExistants.has(f[0])) {
      const lg = prochaineLigneLibre_(feuille, 1, 6);
      feuille.getRange(lg, 1).setNumberFormat('@').setValue(f[0]);
      feuille.getRange(lg, 2).setValue(f[1]);
      feuille.getRange(lg, 11).setValue('Oui');
    }
  });

  if (!feuille.getFrozenRows()) {
    feuille.setFrozenRows(5);
  }

  const largeurs = [
    130, 200, 200, 175, 120, 180, 120, 90, 100, 100, 70, 200
  ];
  largeurs.forEach(function(largeur, index) {
    feuille.setColumnWidth(index + 1, largeur);
  });
}

// ─── Onglet Contacts ──────────────────────────────────────────────────────────

function installerOngletContacts_(ss) {
  let feuille = ss.getSheetByName('Contacts');

  if (!feuille) {
    feuille = ss.insertSheet('Contacts');
    feuille.setHiddenGridlines(true);
  }

  // En-tête de section (ligne 1)
  if (!String(feuille.getRange(1, 1).getValue() || '').trim()) {
    feuille.getRange('A1:O1').merge()
      .setValue('Contacts')
      .setBackground('#e8f0fe')
      .setFontWeight('bold');
  }

  // En-têtes de colonnes (ligne 5)
  const entetes = [
    'ID contact', 'ID fournisseur', 'Fournisseur', 'Nom du contact',
    'Fonction', 'Courriel', 'Téléphone', 'Adresse',
    'Ville', 'Province', 'Pays', 'Code postal', 'Contact principal',
    'Actif', 'Notes'
  ];

  const cellEntetes = feuille.getRange(5, 1, 1, entetes.length);

  if (!String(cellEntetes.getValues()[0][0] || '').trim()) {
    cellEntetes
      .setValues([entetes])
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  // Contact initial CON-0001
  const idsExistants = new Set();
  const derniereLigne = feuille.getLastRow();

  if (derniereLigne >= 6) {
    feuille.getRange(6, 1, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const id = String(ligne[0] || '').trim();
        if (id) { idsExistants.add(id); }
      });
  }

  if (!idsExistants.has('CON-0001')) {
    const lg = prochaineLigneLibre_(feuille, 1, 6);
    feuille.getRange(lg, 1).setNumberFormat('@').setValue('CON-0001');
    feuille.getRange(lg, 2).setNumberFormat('@').setValue('FOU-0006');
    feuille.getRange(lg, 3).setValue('Amusement Jacques-Cartier');
    feuille.getRange(lg, 4).setValue('Jean-François Bertrand');
    feuille.getRange(lg, 13).setValue('Oui'); // Contact principal
    feuille.getRange(lg, 14).setValue('Oui'); // Actif
  }

  if (!feuille.getFrozenRows()) {
    feuille.setFrozenRows(5);
  }

  const largeurs = [
    110, 120, 200, 200, 150, 175, 120, 175, 120, 90, 100, 100,
    120, 70, 200
  ];
  largeurs.forEach(function(largeur, index) {
    feuille.setColumnWidth(index + 1, largeur);
  });
}

// ─── Plan comptable supplémentaire ───────────────────────────────────────────

function installerPlanComptableSupplementaire_(ss) {
  const config = ss.getSheetByName('Configuration');

  if (!config) {
    throw new Error("L'onglet Configuration est introuvable.");
  }

  const comptes = [
    ['4080', "Revenus d'intérêts et ristournes", 'Revenu', 'Autres', 'Oui'],
    ['6065', 'Hébergement et nuitées', 'Dépense', 'Administration', 'Oui']
  ];

  const derniereLigne = config.getLastRow();
  const codesExistants = new Set();

  if (derniereLigne >= 6) {
    config.getRange(6, 1, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const code = String(ligne[0] || '').trim();
        if (code) { codesExistants.add(code); }
      });
  }

  comptes.forEach(function(compte) {
    if (!codesExistants.has(compte[0])) {
      const lg = prochaineLigneLibre_(config, 1, 6);
      config.getRange(lg, 1, 1, 5).setValues([compte]);
      config.getRange(lg, 1).setNumberFormat('@').setValue(compte[0]);
    }
  });
}

// ─── Règles bancaires (Configuration!T:Z) ────────────────────────────────────

function installerReglesBancaires_(ss) {
  const config = ss.getSheetByName('Configuration');

  if (!config) {
    throw new Error("L'onglet Configuration est introuvable.");
  }

  // T4 = identifiant de section
  if (!String(config.getRange(4, 20).getValue() || '').trim()) {
    config.getRange(4, 20).setValue('Règles bancaires')
      .setBackground('#e8f0fe')
      .setFontWeight('bold');
  }

  // En-têtes ligne 5 (T5:Z5)
  const entetes = [
    'Sens', 'Texte à reconnaître', 'ID fournisseur',
    'ID contact par défaut', 'Code compte', 'Programme', 'Actif'
  ];

  const cellEntetes = config.getRange(5, 20, 1, 7);

  if (!String(cellEntetes.getValues()[0][0] || '').trim()) {
    cellEntetes
      .setValues([entetes])
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  const regles = [
    ['Sortie', 'Virement Interac à /JEAN-FRANCOIS /', 'FOU-0002', '', '6000', 'Pion joues-tu?', 'Oui'],
    ['Sortie', 'Achat /CARTIER RESTO B', 'FOU-0003', '', '6130', 'Pion joues-tu? – Cartier', 'Oui'],
    ['Sortie', 'Achat /TIM HORTONS', 'FOU-0004', '', '6130', '', 'Oui'],
    ['Sortie', "Achat /L'IMAGINAIRE", 'FOU-0005', '', '', '', 'Oui'],
    ['Sortie', 'Virement Interac à /Amusement Jacq/', 'FOU-0006', 'CON-0001', '6010', 'Pion joues-tu?', 'Oui'],
    ['Sortie', "Frais fixes d'utilisation", 'FOU-0001', '', '6100', 'Administration générale', 'Oui'],
    ["Sortie", "Frais d'utilisation", 'FOU-0001', '', '6100', 'Administration générale', 'Oui'],
    ['Sortie', 'Achat /MS AUCHALETENBOISRON', 'FOU-0007', '', '6020', 'Pion des bois', 'Oui'],
    ['Entrée', 'Ristourne', 'FOU-0001', '', '4080', 'Administration générale', 'Oui']
  ];

  // Lire les textes existants (col U = col 21, index 1 dans la plage T:Z)
  const derniereLigne = config.getLastRow();
  const textesExistants = new Set();

  if (derniereLigne >= 6) {
    config.getRange(6, 21, derniereLigne - 5, 1)
      .getDisplayValues()
      .forEach(function(ligne) {
        const texte = String(ligne[0] || '').trim().toLowerCase();
        if (texte) { textesExistants.add(texte); }
      });
  }

  // Trouver la première ligne vide dans la section T:Z
  let ligneEcriture = 6;

  if (derniereLigne >= 6) {
    const valeursT = config.getRange(6, 20, derniereLigne - 5, 1).getDisplayValues();
    for (let i = 0; i < valeursT.length; i += 1) {
      if (!String(valeursT[i][0] || '').trim()) {
        ligneEcriture = i + 6;
        break;
      }
      ligneEcriture = i + 7;
    }
  }

  regles.forEach(function(regle) {
    const texteNorm = regle[1].toLowerCase();

    if (!textesExistants.has(texteNorm)) {
      config.getRange(ligneEcriture, 20, 1, 7).setValues([regle]);
      config.getRange(ligneEcriture, 22).setNumberFormat('@'); // ID fournisseur
      config.getRange(ligneEcriture, 23).setNumberFormat('@'); // ID contact
      config.getRange(ligneEcriture, 24).setNumberFormat('@'); // Code compte
      textesExistants.add(texteNorm);
      ligneEcriture += 1;
    }
  });
}

// ─── Colonnes de synchronisation ─────────────────────────────────────────────

function installerColonnesFournisseurTransactions_(ss) {
  const feuille = ss.getSheetByName('Transactions');

  if (!feuille) {
    return;
  }

  // R5 : ID fournisseur, S5 : Fournisseur
  if (!String(feuille.getRange(5, 18).getValue() || '').trim()) {
    feuille.getRange(5, 18).setValue('ID fournisseur')
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  if (!String(feuille.getRange(5, 19).getValue() || '').trim()) {
    feuille.getRange(5, 19).setValue('Fournisseur')
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  if (!String(feuille.getRange(5, 20).getValue() || '').trim()) {
    feuille.getRange(5, 20).setValue('ID contact')
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  feuille.setColumnWidth(18, 110);
  feuille.setColumnWidth(19, 200);
  feuille.setColumnWidth(20, 110);
}

function installerColonnesFournisseurJournal_(ss) {
  const feuille = ss.getSheetByName('Journal');

  if (!feuille) {
    return;
  }

  // O5 : ID fournisseur, P5 : Fournisseur
  if (!String(feuille.getRange(5, 15).getValue() || '').trim()) {
    feuille.getRange(5, 15).setValue('ID fournisseur')
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  if (!String(feuille.getRange(5, 16).getValue() || '').trim()) {
    feuille.getRange(5, 16).setValue('Fournisseur')
      .setBackground('#f1f3f4')
      .setFontWeight('bold')
      .setWrap(true);
  }

  feuille.setColumnWidth(15, 110);
  feuille.setColumnWidth(16, 200);
}

function installerColonnesFournisseurImportBancaire_(ss) {
  const feuille = ss.getSheetByName('Import bancaire');

  if (!feuille) {
    return;
  }

  // Correspondance colonne → lettre pour les messages d'erreur
  const lettreColonne = { 16: 'P', 17: 'Q', 18: 'R', 19: 'S' };

  // P5:S5 corrigées cellule par cellule (idempotent)
  const attendus = [
    { col: 16, valeur: 'ID fournisseur suggéré', largeur: 130 },
    { col: 17, valeur: 'Fournisseur suggéré',    largeur: 180 },
    { col: 18, valeur: 'ID contact suggéré',     largeur: 130 },
    { col: 19, valeur: 'Contact suggéré',         largeur: 180 }
  ];

  attendus.forEach(function(entete) {
    const cellule = feuille.getRange(5, entete.col);
    const valeurActuelle = String(cellule.getValue() || '').trim();

    if (valeurActuelle === '') {
      // Cellule vide : écrire l'en-tête avec la même mise en forme que les autres en-têtes de la ligne 5
      cellule
        .setValue(entete.valeur)
        .setBackground('#f1f3f4')
        .setFontWeight('bold')
        .setWrap(true);
      feuille.setColumnWidth(entete.col, entete.largeur);
    } else if (valeurActuelle === entete.valeur) {
      // Déjà correct : ne rien faire
    } else {
      throw new Error(
        'Import bancaire : la cellule ' +
        lettreColonne[entete.col] + '5 contient « ' + valeurActuelle +
        ' » au lieu de « ' + entete.valeur +
        ' ». Corrigez manuellement avant de relancer l\'installeur.'
      );
    }
  });

  // Retirer les validations héritées de O dans P6:S (idempotent, ne touche pas aux valeurs)
  const derniereLigne = feuille.getMaxRows();
  if (derniereLigne >= 6) {
    feuille.getRange(6, 16, derniereLigne - 5, 4).clearDataValidations();
  }
}

// ─── Réparation des imports classés ──────────────────────────────────────────

function repairerSyncFournisseurImportClassee_(ss) {
  const importBancaire = ss.getSheetByName('Import bancaire');
  const transactions = ss.getSheetByName('Transactions');

  if (!importBancaire || !transactions) return 0;

  const derniereLigneImport = importBancaire.getLastRow();
  const derniereLigneTransactions = transactions.getLastRow();

  if (derniereLigneImport < 6 || derniereLigneTransactions < 6) return 0;

  // Index : referenceImport → infos fournisseur/contact (première transaction active trouvée)
  const indexFournisseurs = {};

  transactions
    .getRange(6, 1, derniereLigneTransactions - 5, 20)
    .getValues()
    .forEach(function(ligne) {
      const referenceImport = String(ligne[13] || '').trim(); // col N
      const statut = String(ligne[14] || '').trim();          // col O
      const idFournisseur = String(ligne[17] || '').trim();   // col R
      const nomFournisseur = String(ligne[18] || '').trim();  // col S
      const idContact = String(ligne[19] || '').trim();       // col T
      const nomContact = String(ligne[3] || '').trim();       // col D

      if (!referenceImport) return;
      if (statut === 'Annulée' || statut === 'Exemple') return;
      if (!idFournisseur) return;
      if (indexFournisseurs[referenceImport]) return; // premier actif suffit

      indexFournisseurs[referenceImport] = {
        idFournisseur: idFournisseur,
        nomFournisseur: nomFournisseur,
        idContact: idContact,
        nomContact: nomContact
      };
    });

  const valeursImport = importBancaire
    .getRange(6, 1, derniereLigneImport - 5, 19)
    .getValues();

  let compteur = 0;

  valeursImport.forEach(function(ligne, index) {
    const referenceImport = String(ligne[0] || '').trim();  // col A
    const statut = String(ligne[10] || '').trim();          // col K
    const pActuel = String(ligne[15] || '').trim();         // col P
    const rActuel = String(ligne[17] || '').trim();         // col R

    if (statut !== 'Classée') return;
    if (!referenceImport) return;

    const infos = indexFournisseurs[referenceImport];
    if (!infos) return;

    // Déjà synchronisé (fournisseur ET contact corrects)
    if (pActuel === infos.idFournisseur && rActuel === infos.idContact) return;

    const ligneSheet = index + 6;

    // Retirer la validation de P:S sur cette ligne uniquement (jamais O)
    importBancaire.getRange(ligneSheet, 16, 1, 4).clearDataValidations();

    // P:Q = fournisseur confirmé
    importBancaire.getRange(ligneSheet, 16).setNumberFormat('@').setValue(infos.idFournisseur);
    importBancaire.getRange(ligneSheet, 17).setValue(infos.nomFournisseur);

    // R:S = contact si présent, sinon vider
    if (infos.idContact) {
      importBancaire.getRange(ligneSheet, 18).setNumberFormat('@').setValue(infos.idContact);
      importBancaire.getRange(ligneSheet, 19).setValue(infos.nomContact);
    } else {
      importBancaire.getRange(ligneSheet, 18, 1, 2).clearContent();
    }

    compteur += 1;
  });

  return compteur;
}

// ─── Accès aux données (utilisé par FinalisationMixtes.js et ImportCsv.js) ───

function obtenirIndexEntetesGestionFournisseurs_(feuille, entetesRequises) {
  const derniereColonne = feuille.getLastColumn();

  if (derniereColonne < 1) {
    throw new Error(
      "L'onglet « " + feuille.getName() + " » ne contient aucun en-tête."
    );
  }

  const entetes = feuille
    .getRange(5, 1, 1, derniereColonne)
    .getDisplayValues()[0]
    .map(function(entete) {
      return String(entete || '').trim();
    });
  const index = {};

  entetesRequises.forEach(function(enteteRequise) {
    const position = entetes.indexOf(enteteRequise);

    if (position === -1) {
      throw new Error(
        "L'en-tête « " + enteteRequise + " » est introuvable dans l'onglet « " +
        feuille.getName() + ' ».'
      );
    }

    index[enteteRequise] = position;
  });

  return {
    index: index,
    nombreColonnes: derniereColonne
  };
}

function chargerFournisseursActifs_(ss) {
  const feuille = ss.getSheetByName('Fournisseurs');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const structure = obtenirIndexEntetesGestionFournisseurs_(
    feuille,
    ['ID fournisseur', 'Fournisseur', 'Actif']
  );

  return feuille
    .getRange(
      6,
      1,
      feuille.getLastRow() - 5,
      structure.nombreColonnes
    )
    .getDisplayValues()
    .filter(function(ligne) {
      return (
        String(ligne[structure.index['ID fournisseur']] || '').trim() &&
        String(ligne[structure.index.Actif] || '').trim() === 'Oui'
      );
    })
    .map(function(ligne) {
      return {
        id: String(ligne[structure.index['ID fournisseur']] || '').trim(),
        nom: String(ligne[structure.index.Fournisseur] || '').trim()
      };
    });
}

function chargerTousContacts_(ss) {
  const feuille = ss.getSheetByName('Contacts');

  if (!feuille || feuille.getLastRow() < 6) {
    return [];
  }

  const structure = obtenirIndexEntetesGestionFournisseurs_(
    feuille,
    ['ID contact', 'ID fournisseur', 'Nom du contact', 'Actif']
  );

  return feuille
    .getRange(
      6,
      1,
      feuille.getLastRow() - 5,
      structure.nombreColonnes
    )
    .getDisplayValues()
    .filter(function(ligne) {
      return (
        String(ligne[structure.index['ID contact']] || '').trim() &&
        String(ligne[structure.index.Actif] || '').trim() === 'Oui'
      );
    })
    .map(function(ligne) {
      return {
        id: String(ligne[structure.index['ID contact']] || '').trim(),
        idFournisseur:
          String(ligne[structure.index['ID fournisseur']] || '').trim(),
        nom: String(ligne[structure.index['Nom du contact']] || '').trim()
      };
    });
}

function obtenirNomFournisseur_(ss, idFournisseur) {
  if (!idFournisseur) {
    return '';
  }

  const feuille = ss.getSheetByName('Fournisseurs');

  if (!feuille || feuille.getLastRow() < 6) {
    return idFournisseur;
  }

  const valeurs = feuille
    .getRange(6, 1, feuille.getLastRow() - 5, 2)
    .getDisplayValues();

  for (let i = 0; i < valeurs.length; i += 1) {
    if (String(valeurs[i][0] || '').trim() === idFournisseur) {
      return String(valeurs[i][1] || '').trim();
    }
  }

  return idFournisseur;
}

function obtenirNomContact_(ss, idContact) {
  if (!idContact) {
    return '';
  }

  const feuille = ss.getSheetByName('Contacts');

  if (!feuille || feuille.getLastRow() < 6) {
    return '';
  }

  const valeurs = feuille
    .getRange(6, 1, feuille.getLastRow() - 5, 4)
    .getDisplayValues();

  for (let i = 0; i < valeurs.length; i += 1) {
    if (String(valeurs[i][0] || '').trim() === idContact) {
      return String(valeurs[i][3] || '').trim();
    }
  }

  return '';
}

function normaliserTexteRechercheRegle_(texte) {
  return String(texte || '')
    .toLowerCase()
    .replace(/[àâä]/g, 'a')
    .replace(/[éèêë]/g, 'e')
    .replace(/[îï]/g, 'i')
    .replace(/[ôö]/g, 'o')
    .replace(/[ùûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/\s+/g, ' ')
    .trim();
}

function rechercherRegleBancaire_(ss, description, montant) {
  const config = ss.getSheetByName('Configuration');

  if (!config || config.getLastRow() < 6) {
    return null;
  }

  // Règles en T6:Z<derniereLigne> (cols 20-26, 7 colonnes)
  const derniereLigne = config.getLastRow();

  if (derniereLigne < 6) {
    return null;
  }

  const plage = config
    .getRange(6, 20, derniereLigne - 5, 7)
    .getValues();

  const sensRecherche = montant < 0 ? 'Sortie' : 'Entrée';
  const descNormalisee = normaliserTexteRechercheRegle_(description);

  for (let i = 0; i < plage.length; i += 1) {
    const ligne = plage[i];
    const sens = String(ligne[0] || '').trim();
    const texte = String(ligne[1] || '').trim();
    const actif = String(ligne[6] || '').trim();

    if (!actif || actif !== 'Oui') {
      continue;
    }

    if (sens !== sensRecherche) {
      continue;
    }

    if (!texte) {
      continue;
    }

    if (descNormalisee.indexOf(normaliserTexteRechercheRegle_(texte)) === -1) {
      continue;
    }

    const idFournisseur = String(ligne[2] || '').trim();
    const idContact = String(ligne[3] || '').trim();
    const codeCompte = String(ligne[4] || '').trim();
    const programme = String(ligne[5] || '').trim();

    return {
      idFournisseur: idFournisseur,
      nomFournisseur: obtenirNomFournisseur_(ss, idFournisseur),
      idContact: idContact,
      nomContact: obtenirNomContact_(ss, idContact),
      codeCompte: codeCompte,
      programme: programme
    };
  }

  return null;
}

function validerFournisseurEtContact_(ss, idFournisseur, idContact) {
  if (!idFournisseur) {
    throw new Error('Le fournisseur est obligatoire.');
  }

  const feuilleF = ss.getSheetByName('Fournisseurs');

  if (!feuilleF || feuilleF.getLastRow() < 6) {
    throw new Error("L'onglet Fournisseurs est introuvable ou vide.");
  }

  const structureF = obtenirIndexEntetesGestionFournisseurs_(
    feuilleF,
    ['ID fournisseur', 'Fournisseur', 'Actif']
  );
  const valeursF = feuilleF
    .getRange(
      6,
      1,
      feuilleF.getLastRow() - 5,
      structureF.nombreColonnes
    )
    .getDisplayValues();

  let fournisseurTrouve = null;

  for (let i = 0; i < valeursF.length; i += 1) {
    const id = String(
      valeursF[i][structureF.index['ID fournisseur']] || ''
    ).trim();

    if (id !== idFournisseur) {
      continue;
    }

    const actif = String(
      valeursF[i][structureF.index.Actif] || ''
    ).trim();

    if (actif !== 'Oui') {
      throw new Error(
        'Le fournisseur ' + idFournisseur + ' n\'est pas actif.'
      );
    }

    fournisseurTrouve = {
      id: id,
      nom: String(
        valeursF[i][structureF.index.Fournisseur] || ''
      ).trim()
    };

    break;
  }

  if (!fournisseurTrouve) {
    throw new Error(
      'Le fournisseur ' + idFournisseur + ' est introuvable.'
    );
  }

  let contactTrouve = { id: '', nom: '' };

  if (idContact) {
    const feuilleC = ss.getSheetByName('Contacts');

    if (!feuilleC || feuilleC.getLastRow() < 6) {
      throw new Error("L'onglet Contacts est introuvable ou vide.");
    }

    const structureC = obtenirIndexEntetesGestionFournisseurs_(
      feuilleC,
      ['ID contact', 'ID fournisseur', 'Nom du contact', 'Actif']
    );
    const valeursC = feuilleC
      .getRange(
        6,
        1,
        feuilleC.getLastRow() - 5,
        structureC.nombreColonnes
      )
      .getDisplayValues();

    let contactValide = false;

    for (let i = 0; i < valeursC.length; i += 1) {
      const idC = String(
        valeursC[i][structureC.index['ID contact']] || ''
      ).trim();

      if (idC !== idContact) {
        continue;
      }

      const idF = String(
        valeursC[i][structureC.index['ID fournisseur']] || ''
      ).trim();

      if (idF !== idFournisseur) {
        throw new Error(
          'Le contact ' + idContact +
          ' n\'appartient pas au fournisseur ' + idFournisseur + '.'
        );
      }

      const actifC = String(
        valeursC[i][structureC.index.Actif] || ''
      ).trim();

      if (actifC !== 'Oui') {
        throw new Error('Le contact ' + idContact + ' n\'est pas actif.');
      }

      contactTrouve = {
        id: idC,
        nom: String(
          valeursC[i][structureC.index['Nom du contact']] || ''
        ).trim()
      };

      contactValide = true;
      break;
    }

    if (!contactValide) {
      throw new Error('Le contact ' + idContact + ' est introuvable.');
    }
  }

  return {
    idFournisseur: fournisseurTrouve.id,
    nomFournisseur: fournisseurTrouve.nom,
    idContact: contactTrouve.id,
    nomContact: contactTrouve.nom
  };
}
