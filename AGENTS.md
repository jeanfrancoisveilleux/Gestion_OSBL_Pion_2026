# Gestion OSBL Pion 2026

- Répondre et documenter le projet en français.
- Ce projet Apps Script est lié au Google Sheet de gestion comptable de l’OSBL « Les événements Pion joues-tu? ».
- Utiliser Import bancaire comme source des opérations bancaires.
- Transactions et Journal conservent la piste d’audit comptable.
- Ne jamais supprimer une transaction comptabilisée : utiliser une annulation avec écritures inverses.
- Toute annulation ou modification doit synchroniser Transactions, Journal, Import bancaire, Répartition, Forfaits, Inventaire et les rapports concernés.
- Les rapports comptables reposent sur Journal.
- Les anciennes transactions annulées doivent rester visibles avec le statut Annulée.
- Les nouvelles versions d’une transaction utilisent un identifiant de révision, par exemple -R2-.
- Les codes de comptes doivent être écrits comme du texte afin de respecter les validations Google Sheets.
- FinalisationMixtes.js contient la logique des transactions mixtes.
- RepartitionMixte.html contient l’interface de répartition.
- Préserver les données existantes et les modifications sans rapport avec la tâche.
- Inspecter le code concerné avant toute modification.
- Vérifier la syntaxe JavaScript après chaque correction.
- Avant clasp push, résumer les modifications et demander une confirmation explicite.
- Ne jamais exécuter clasp pull lorsqu’il existe des modifications locales non sauvegardées.
- Lorsqu’un script est corrigé dans une réponse, fournir le fichier complet et non seulement un extrait.
