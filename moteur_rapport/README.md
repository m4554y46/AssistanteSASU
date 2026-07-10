# Moteur autonome de rapport Assistant SASU

Ce dossier contient un moteur local qui produit un rapport Word dynamique.

## Lancement

Double-cliquer sur `lancer_rapport.bat` ou lancer :

```powershell
C:\Users\micas\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe run_weekly_report.py
```

## Fonctionnement

1. Lit `config.json`.
2. Lance des recherches web publiques via flux de résultats Bing RSS.
3. Déduplique et vérifie les URLs.
4. Score les articles et opportunités selon les mots-clés et la qualité de source.
5. Sélectionne des méthodologies depuis la bibliothèque configurée.
6. Génère un rapport Word dans `C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU`.
7. Produit aussi un `runlog.json` avec les éléments collectés et scorés.

## Limites assumées

- Le moteur ne contourne pas les logins, paywalls ou protections de plateformes.
- LinkedIn, Malt et certaines plateformes peuvent limiter l'accès ; le rapport le signale via les notes de vérification.
- Les recherches sont dynamiques, mais dépendent de la disponibilité des résultats publics au moment du lancement.
- Le scoring est transparent et modifiable dans `config.json`.
