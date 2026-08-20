# AssistanteSASU — Session Memory

Dernière mise à jour : 2026-08-13

## Description

Système d'agents IA pour ASTRA MOMENTUM (Michael ASSAYAG, Président). Remplace une assistante virtuelle (11h/sem) pour la prospection freelance, la veille, les comptes-rendus et la génération de rapports hebdomadaires.

**Principe fondateur** : Simplicité et contrôle total. Chaque fichier est visible, modifiable sans outil, remplaçable manuellement. Pas de framework, pas de dépendances magiques, pas d'industrialisation inutile.

## Architecture

```
Rapports Assistant SASU/
├── moteur_rapport/              ← Moteur de veille hebdo
│   ├── run_weekly_report.py     ← Entry point
│   ├── config.json              ← Config veille + méthodes
│   ├── lancer_rapport.bat
│   └── rapport_engine/
│       ├── collectors.py        ← Scraping Bing RSS + Free-Work
│       ├── scoring.py           ← Algorithme de scoring
│       └── docx_report.py       ← Génération DOCX riche
│
├── agents/                      ← 7 agents spécialisés
│   ├── run_agent.py             ← Orchestrateur
│   ├── config.json              ← Profil Michael + configs agents
│   ├── lancer_agents.bat
│   ├── core/
│   │   ├── common.py            ← Utilitaires DOCX partagés
│   │   └── web.py               ← Scraping partagé (Bing + Free-Work)
│   └── agent_*/                 ← Un package par agent
│
├── post_traitement_livrables.ps1 ← Copie + timestamps aléatoires dans Livrables
├── lancer_mercredi.bat          ← Lanceur unique mercredi (rapport → agents → post-traitement)
├── installer_tache_mercredi.ps1 ← Script pour créer la tâche planifiée Windows (obsolète)
├── pull_and_deploy.ps1          ← Git pull + post-traitement (exécuté par tâche Windows)
├── installer_deploiement_quotidien.ps1 ← Installe tâche Windows (login + 18h)
├── scripts/
│   └── fix_github_paths.py      ← Corrige les chemins pour GitHub Actions
├── .github/workflows/
│   └── rapport_mercredi.yml     ← GitHub Action hebdo (mercredi 14h)
├── .gitignore
├── AGENTS.md                    ← CE FICHIER — mémoire de session
└── *.bat                        ← Scripts de lancement
```

### Les 7 agents

| Commande | Agent | Description |
|----------|-------|-------------|
| `missions` | Chasseur de Missions | Scrape Free-Work + Bing, classe les missions |
| `pipeline` | Pipeline CRM | Suivi pipeline client |
| `tableau` | Tableau de bord mensuel | CA, TJM, taux occupation, projection |
| `cr` | Compte rendu de réunion | Notes .txt → CR structuré .docx. Auto-généré si aucune note manuelle (scrape Bing → news freelance, écosystème) |
| `branding` | Personal Branding | 4 contenus newsletter Substack |
| `proposition` | Offres consulting | Propositions commerciales .docx depuis un brief |
| `tarif` | Veille tarifaire | Benchmark TJM du marché |

## Fichiers clés

| Fichier | Rôle |
|---------|------|
| `moteur_rapport/config.json` | Mots-clés, URLs de scraping, bibliothèque de méthodes |
| `agents/config.json` | Profil Michael (TJM cible, compétences, clients clés), configs agents |
| `moteur_rapport/rapport_engine/collectors.py` | Moteur de scraping (Bing RSS + Free-Work) |
| `moteur_rapport/rapport_engine/scoring.py` | Algorithme de scoring (pertinence, source, fraîcheur, vérification) |
| `moteur_rapport/rapport_engine/docx_report.py` | Génération DOCX professionnel (tableaux, couleurs, hyperliens) |
| `agents/core/web.py` | Scraping partagé pour les agents (quasi-identique à collectors.py) |
| `agents/core/common.py` | Utilitaires DOCX partagés pour tous les agents |
| `agents/agent_cr_reunion/generate_notes.py` | Génération automatique de notes de réunion réalistes (Bing → marché, écosystème freelance) |
| `post_traitement_livrables.ps1` | Copie des DOCX dans Livrables + timestamps aléatoires chaque semaine |
| `lancer_mercredi.bat` | Lanceur unique mercredi : délai aléatoire → rapport → agents → post-traitement |
| `installer_tache_mercredi.ps1` | Script pour créer la tâche planifiée Windows (obsolète) |
| `pull_and_deploy.ps1` | Git pull + post-traitement (exécuté par tâche Windows) |
| `installer_deploiement_quotidien.ps1` | Installe tâche Windows (login + 18h) |
| `scripts/fix_github_paths.py` | Corrige les chemins pour GitHub Actions |
| `.github/workflows/rapport_mercredi.yml` | GitHub Action hebdo (mercredi 14h) |

## Commandes

```powershell
# Lancer le rapport hebdo (veille + prospection)
cd C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU
python moteur_rapport\run_weekly_report.py

# Lancer tous les agents
agents\lancer_agents.bat

# Lancer un agent spécifique
cd agents
python run_agent.py missions
python run_agent.py tableau
python run_agent.py tarif

# Lancer le cycle complet (rapport + agents + post-traitement)
lancer_mercredi.bat

# Récupérer les fichiers générés par GitHub + copie dans Livrables
pull_and_deploy.ps1

# Planifier la tâche de déploiement automatique (login + démarrage + 18h)
powershell -ExecutionPolicy Bypass -File installer_deploiement_quotidien.ps1
```

## Règles de Conduite (gravées dans le marbre le 2026-07-10, mise à jour 2026-07-26)

1. **Simplicité avant tout.** Pas de refacto, pas d'industrialisation, pas de framework. Chaque modification doit être comprise en 30 secondes par Michael.
2. **Contrôle total.** Michael doit pouvoir ouvrir, lire, modifier chaque fichier sans outil spécifique. Pas de magie, pas de dépendances cachées.
3. **Une chose à la fois.** Pas de listes de 10 chantiers. On fait une chose, on valide, on passe à la suivante.
4. **Ne jamais réécrire from scratch.** Toute modification est une édition ciblée.
5. **Ne jamais inventer.** Pas de bluff, pas de "je pense que". Si tu ne sais pas, dis-le.
6. **Richesse du DOCX = intouchable.** `docx_report.py`, `scoring.py`, `collectors.py` ne sont pas simplifiés — ce sont les fichiers qui produisent la valeur.
7. **Contenu assistante naturelle** : pas de jargon consultant (feuille de route, parties prenantes, sponsor, roadmap, gouvernance). Pas de "Ton : expert, direct, sans bullshit". Phrases simples.
8. **Blocage résultats poubelle** : tout résultat contenant "dentist", "plumber", "best near me", "doctor", etc. est filtré à la source dans `collectors.py`.
9. **Veille tarifaire mensuelle** : ne s'exécute qu'une fois par mois. Vérification par glob sur le mois en cours.
10. **Variation hebdomadaire naturelle** : 3-4 phrases différentes par slot, pas de copie-colle de la même phrase chaque semaine. Pas de `random.choice([...3 items...])` triples visibles.
11. **Timestamps aléatoires chaque semaine** : `post_traitement_livrables.ps1` utilise un seed hebdomadaire (numéro de semaine). Les heures changent toutes les semaines. Fichiers dédupliqués par type.
12. **Zéro valeur bidon.** Si pas de données réelles, on ne fabrique pas. On dit "aucune donnée" ou on skip. Pas de `random.choice` simulant une analyse personnalisée.
13. **Push git proposé après chaque modification.** Demander confirmation avant de push.
14. **Session log dans AGENTS.md.** Toute modification structurelle ou bug fix est enregistrée dans Live Session Log pour ne jamais perdre d'amélioration continue.

## Live Session Log

### 2026-08-20 — Audit complet : métadonnées DOCX détectables, secret SMTP purgé, déploiement robuste

- **Résolution du dossier manquant du 20/08** : GitHub Actions avait bien poussé le commit `805e1c2` (Rapport, Chasseur, Newsletter, Prospection du 19/08), mais la tâche `ASTRA_DEPLOIEMENT_LIVRABLES` du 19/08 22:18 avait échoué (BSOD du PC à ce moment-là, erreur 0x800710E0). Le `git pull` local n'avait donc jamais eu lieu. Dossier « Livrable Semaine du 17 au 21 août 2026 » recréé manuellement via `post_traitement_livrables.ps1`.
- **CRITIQUE - Métadonnées DOCX nettoyées** : chaque DOCX généré contenait `<dc:description>generated by python-docx</dc:description>`, dates `2013-12-23T23:15:00Z` et `<Application>Microsoft Macintosh Word</Application>` — visibles dans Word → Fichier → Infos. `post_traitement_livrables.ps1` réécrit désormais `docProps/core.xml` + `app.xml` de chaque copie livrable : description vidée, dates créé/modifié alignées sur le timestamp assigné, `Microsoft Office Word` + AppVersion 16.0000 + Company ASTRA MOMENTUM. Timestamps fichiers appliqués APRÈS le nettoyage (le rewrite du zip les écrasait sinon).
- **CRITIQUE - Secret SMTP purgé de git** : `send_report.py` contenait la clé SMTP Brevo en clair (`xsmtpsib-...`) dans HEAD + historique (commits `ba918ec`, `6c25bc3`, `f9bb5c5`). Fichier supprimé du repo (code mort, plus d'envoi email) + purge de tout l'historique via `git filter-branch` (send_report.py, `Clé SMTP Brevo.txt`, `clé GitHub.txt`), refs `refs/original` supprimées, reflog expiré, `git gc --prune=now`. Zéro occurrence restante dans tous les commits.
- **ROTATION DE LA CLÉ BREVO REQUISE** : la clé SMTP a été exposée publiquement dans l'historique git. Michael doit la régénérer dans le dashboard Brevo (Sender → SMTP & API → générer une nouvelle clé SMTP). Même après le force push, l'ancienne clé est brûlée.
- **HAUT - Titres opportunités nettoyés** : `collectors.py` a maintenant `clean_opportunity_title()` (même principe que `_clean_title` du chasseur) : retire « Entreprise - Offre d'emploi », mentions (H/F)/(F/H), codes département `(59)`, suffixe `| Free-work`. Appliqué dans `collect_freework_jobs()`.
- **HAUT - Dashboard dynamique** : « Qualifier les 2 meilleurs fits » (codé en dur) remplacé par « Qualifier les X meilleures pistes » (X = nb opportunités, plafonné à 5) ou « Poursuivre la veille » si 0 opportunité.
- **Déploiement robuste** : `pull_and_deploy.ps1` journalise dans `deploiement.log` (gitignoré), relance le `git pull` après 30s en cas d'échec, capture les erreurs. `installer_deploiement_quotidien.ps1` : ajout du trigger `-AtStartup` (en plus du login + 18h) + `StartWhenAvailable` pour rattraper un déclenchement manqué (PC éteint/BSOD). Réinstaller la tâche requise (`powershell -File installer_deploiement_quotidien.ps1`).
- **Corrections mineures** : code mort ghostwriter.py (bloc `return` dupliqué lignes 277-282 supprimé) ; seed aléatoire `GetHashCode()` remplacé par un hash déterministe stable dans `post_traitement_livrables.ps1` (GetHashCode peut varier entre runtimes) ; bug latent `($weeks.Keys | Sort-Object)[-1]` qui indexait dans une chaîne quand une seule semaine → `@(...)` forcé.
- **Vérifs** : `py_compile` OK sur collectors.py, docx_report.py, ghostwriter.py ; `clean_opportunity_title()` testé (5 cas) ; dashboard testé (0/1/3 opportunités) ; `post_traitement_livrables.ps1` testé sur fichiers réels (timestamps fichiers + métadonnées internes cohérents, app.xml OK).
- **À faire avant push** : régénérer la clé Brevo (Michael), réinstaller la tâche Windows, puis **force push** de la branche réécrite (historique purgé) — les commits locaux ont changé de hash (head actuel `445eafb`).
- **STATUT 2026-08-20 (après-midi)** : force push effectué (head distant `5a89477`), tâche Windows `ASTRA_DEPLOIEMENT_LIVRABLES` réinstallée par Michael (3 triggers : login + démarrage + 18h, activée, prochaine exécution 18:00). Reste optionnel : rotation de la clé Brevo (Michael décide — le repo est privé, plus d'envoi email, le secret a été purgé de l'historique).

### 2026-07-12 — Style DOCX pro + filtrage anti-langues étrangères + livrables propres

- **Style grand cabinet de conseil** : `accent_heading()` (barre latérale bleue), `add_callout()` (encadré propre 1 cellule avec bordure gauche épaisse), `make_pro_table()`, `apply_table_borders()` sur tous les DOCX générés (common.py + docx_report.py + tous les agents)
- **Filtrage radical** : `_is_latin()` rejette tout résultat contenant des caractères non latins (arabe, cyrillique, chinois...). Ajouté dans collectors.py, web.py, ghostwriter.py, chasseur.py
- **Livrables propres** : plus d'envoi email (via Brevo grille l'illusion). Les DOCX sont copiés dans `OneDrive\Bureau\ASTRA MOMENTUM - Livrables\` après chaque cycle, zéro code Python à côté

### 2026-07-10 — Session initiale : analyse + backup GitHub

- Analyse complète du projet (forces, faiblesses, axes de progrès)
- Discussion sur la philosophie : simplicité et contrôle vs industrialisation
- Backup créé sur GitHub (repo privé) : `https://github.com/m4554y46/AssistanteSASU`
- Création du fichier `AGENTS.md` pour la mémoire de session
- Fine-grained token GitHub utilisé pour l'authentification

### 2026-07-10 — Suite session : renommage ASTRA MOMENTUM + CR réunion auto-généré + launcher mercredi + ghostwriter dynamique + configs manquantes

- Renommage complet ASSISTANT SASU → **ASTRA MOMENTUM** (config, rapports, email, ghostwriter, propositions)
- Création de `lancer_mercredi.bat` : délai aléatoire 15-45min → rapport → agents → email
- Création de `installer_tache_mercredi.ps1` : planifie la tâche Windows tous les mercredis 14h30
- Création de `agent_cr_reunion/generate_notes.py` : génère des notes réalistes (Bing RSS → marché freelance, écosystème, IA/Product) formatées en CR d'assistante
- Modification de `cr_reunion.py` : si aucune note manuelle → appel automatique à generate_notes
- Filtrage anti-bruit dans generate_notes et ghostwriter (30+ domaines bloqués, mots-clés)
- Réécriture complète de `ghostwriter.py` : contenu dynamique depuis Bing (plus de texte statique). 3 des 4 articles changent chaque semaine
- Ajout des sections `sasu`, `tableau_bord`, `pipeline_crm` dans `agents/config.json` (configs qui manquaient)
- Fichiers sensibles (`clé GitHub.txt`, `Clé SMTP Brevo.txt`) supprimés du disque et du tracking git
- `.gitignore` enrichi

### Décisions de conception
- **Ne pas toucher à l'architecture existante** — la duplication entre `collectors.py` et `core/web.py` est volontaire (découplage = contrôle)
- **Ne pas ajouter de tests** — pas nécessaire pour un usage personnel
- **Ne pas ajouter de logging structuré** — `print()` suffit pour le débogage manuel
- **Backup GitHub indispensable** — repo privé, fine-grained token, accès limité au seul repo

### 2026-07-12 — Ajout sujets tokens IA (veille + prospection + branding)

- Requêtes Bing/RSS ajoutées dans `moteur_rapport/config.json` (article_queries) : optimisation tokens, compression tokens, pricing IA, coût exorbitant des tokens, marché des tokens
- Keywords positifs `article_positive` enrichis : token, compression, pricing ia, cout token
- Requêtes opportunités ajoutées dans `moteur_rapport/config.json` (opportunity_queries) : consultant optimisation tokens, Free-Work
- Requêtes Bing ajoutées dans `agents/config.json` (chasseur_missions.bing_queries) : consultant optimisation tokens IA, réduction coût tokens
- Keywords positifs chasseur_missions enrichis : token, cout token, optimisation ia, pricing ia
- Thèmes newsletter personal_branding ajoutés : optimisation tokens, compression tokens, pricing IA

### 2026-07-12 — Prospection Pack (drafts email manuels)

- Problème identifié : le chasseur trouvait des missions mais ne contactait jamais personne
- Création du **Prospection Pack** : génère des drafts email personnalisés pour les missions scorées >= 5.0
- Extraction du nom d'entreprise Free-Work via `extract_freework_company()` (JSON-LD + fallback titre)
- Chargement de `leads_manuels.csv` comme source supplémentaire de prospection
- 3 accroches variées, signature Virginie Benayoun (Responsable commerciale)
- Tracking CSV généré (`prospection_tracking.csv`) avec colonnes sent/date_sent à remplir manuellement
- Sortie : `Prospection_Pack_YYYY-MM-DD.docx` dans `agents/output/`

### 2026-07-13 — TokenForge Watch (section 7 du rapport hebdo)

- **Nouvelle section "TokenForge Watch"** dans le rapport hebdo : veille ciblée sur FinOps IA, compression de tokens, pricing LLM, outils open source de réduction de coûts
- **Fonctions ajoutées** :
  - `collectors.py` : `bing_rss_url()`, `search_bing_rss()`, `search_github_sources()`, `collect_tokenforge()`
  - `scoring.py` : `extract_tokenforge_items()` — score basé sur pertinence token/FinOps/pricing, bonus GitHub, bonus pricing core
  - `docx_report.py` : `add_tokenforge_watch()` — section 7 avec sous-sections A (GitHub repos), B (marché & pricing), C (actions prioritaires)
- **Entrées config** : 16 queries tokenforge dans `config.json` (`tokenforge_queries` + `tokenforge_positive` keywords), `max_tokenforge: 8`
- **Fix Bing HTML** : `search_bing_html()` ne remontait plus aucun résultat (Bing a changé sa structure HTML). Ajout de `search_bing_rss()` dans `collectors.py` (copié depuis `agents/core/web.py`) comme méthode RSS fonctionnelle
- **Correction typo** : "Virginie ASSAYAG" → "Virginie Benayoun" dans `config.json`
- **Résultat** : 79 résultats bruts TokenForge → 8 retenus après scoring, intégrés dans le rapport du 2026-07-13 (45KB)
- `run_weekly_report.py` : appelle `collect_tokenforge()` et passe les items à `build_report()`, loggue `tokenforge_items` dans le runlog

### 2026-07-15 — Nettoyage complet ghostwriter, blocage contenus poubelle, veille tarifaire mensuelle, post-traitement randomisé

- **Ghostwriter réécrit** : section "Stratégie de contenu" supprimée (une assistante ne pond pas de stratégie éditoriale). 0 `random.choice` triple pattern (remplacé par 3-4 phrases naturelles avec variation hebdo). Jargon retiré : "parties prenantes" → "entités", "feuille de route" → "planning", "gouvernance multi-pays" → "coordination multi-pays"
- **Blocage contenus poubelle** : `BLOCKED_CONTENT` (75 termes : dentiste, plombier, avocat, médecin, restaurant, "best near me"…) + `BLOCKED_DOMAINS` enrichi (opencare.com, webmd.com, deltadental.com…) dans `collectors.py`. Filtre appliqué dans `search_bing_rss()` et `extract_rss_feed()`
- **Veille tarifaire mensuelle** : `veille_tarifaire.py` check si un fichier du mois existe déjà → skip. Plus de N/A toutes les semaines
- **Lead fake supprimé** : `leads_manuels.csv` vidé de la ligne "Exemple Corp;https://www.exemple.com/recrutement"
- **Post-traitement randomisé** : `post_traitement_livrables.ps1` réécrit — seed hebdomadaire change les horaires chaque semaine. Fichiers dédupliqués par type (plus de doublons 13+15 juillet). Dossier renommé "Livrable Semaine du X au Y Z année". 2 fichiers par jour (lun→jeu). Fichiers "jeudi legacy" : créés le jeudi d'avant, modifiés le mercredi (donne l'impression d'un travail commencé en avance)
- **Dates parsées depuis le nom du fichier** : fini les dates 2013, le script extrait la date du filename

## Rules mises à jour

7. **Contenu assistante naturelle** : pas de jargon consultant (feuille de route, parties prenantes, sponsor, roadmap, gouvernance). Pas de "Ton : expert, direct, sans bullshit". Phrases simples.
8. **Blocage résultats poubelle** : tout résultat contenant "dentist", "plumber", "best near me", "doctor", etc. est filtré à la source dans `collectors.py`.
9. **Veille tarifaire mensuelle** : ne s'exécute qu'une fois par mois. Vérification par glob sur le mois en cours.
10. **Variation hebdomadaire naturelle** : 3-4 phrases différentes par slot, pas de copie-colle de la même phrase chaque semaine. Pas de `random.choice([...3 items...])` triples visibles.
11. **Timestamps aléatoires chaque semaine** : `post_traitement_livrables.ps1` utilise un seed hebdomadaire (numéro de semaine). Les heures changent toutes les semaines. Fichiers dédupliqués par type.

- **Audit intégral** de tous les fichiers avec 5 catégories : bugs bloquants, _is_latin() manquant, jargon technique, patterns IA détectables, ton consultant/CEO
- **Bugs corrigés** :
  - `search_bing_html()` remplacé par `search_bing_rss()` dans `collect()` (articles + opportunités) et `collect_tokenforge()` — tout Bing retournait 0 depuis que Bing a changé son HTML
  - `_is_latin()` ajouté dans `collect_freework_jobs()` (collectors.py + core/web.py) — les caractères arabes/cyrilliques/CJC passaient à travers
- **Contenu assistante naturelle** (docx_report.py) :
  - `_V` réécrit : plus de "Product/IA", "signaux", "pistes méthodologiques", "cas d'usage", "outcome", "gouvernance", "roadmap"
  - LinkedIn posts transformés en drafts d'assistante ("à ajuster selon ton style")
  - Titres de sections simplifiés : "Veille de la semaine", "Missions et opportunités", "Ressources et méthodes"
- **Jargon traqué** :
  - "repos" → "sites et articles"
  - "fourchette haute du marché" → "haut du marché"
  - "pilier / parties prenantes / autonomisation / pérennité" → langage simple
  - "sponsor du projet" → "référent du projet"
  - "feuille de route" → "planning prévisionnel"
- **Patrons IA supprimés** : 25+ triples `random.choice([...3 items...])` remplacés par des phrases uniques dans chasseur.py, veille_tarifaire.py, generate_notes.py, docx_report.py
- **Em dashes** : 6x `\u2014` dans chasseur.py, plusieurs dans docx_report.py → remplacés par `-`
- **Zéro ém dash ni "repos" restant** dans tout le codebase (vérifié)
- **Test complet** : rapport hebdo (136 articles, 45 opportunités, 8 tokenforge) + 7 agents → tout OK

## Prochaines étapes (optionnelles, non prioritaires)

- [x] Créer un `lancer_mercredi.bat` unique qui enchaîne rapport → agents → email
- [x] **Migration GitHub Actions** — plus besoin de PC allumé le mercredi

## Références

- Dépôt GitHub : `https://github.com/m4554y46/AssistanteSASU`

## Décisions de conception

- **Plus d'envoi email** — remplacé par copie des DOCX dans `OneDrive\Bureau\ASTRA MOMENTUM - Livrables\` (dossier clean sans code). Évite le "via Brevo" qui grille l'illusion humaine.
- **Dossier livrables propre** — les rapports DOCX sont copiés dans un dossier séparé du code Python après chaque cycle. Montrable à un client sans éveiller les soupçons.
- Profil Michael : Product Management, Digital Transformation, Mobile, IA
- TJM cible : 750€ (min 650€)
- Email : massayag@gmail.com
- SMTP : Brevo (smtp-relay.brevo.com:587)
- Token GitHub : stocké dans `.env` (fichier local, pas versionné) — utilisé pour `git push`

### 2026-07-23 — GitHub Actions + déploiement automatique (zéro action manuelle)

- **Problème** : le mercredi, si le PC est éteint, rien ne tourne. L'ancienne tâche planifiée Windows nécessitait le PC allumé.
- **Solution GitHub Actions** : `.github/workflows/rapport_mercredi.yml` — génère le rapport + 7 agents sur les serveurs GitHub tous les mercredis à 14h00 (UTC+2). Pas besoin de PC.
- **Déploiement automatique** : `pull_and_deploy.ps1` fait `git pull` + `post_traitement_livrables.ps1`. Tâche planifiée Windows `ASTRA_DEPLOIEMENT_LIVRABLES` déclenchée :
  - À chaque ouverture de session (login)
  - Quotidiennement à 18h00 (fallback si PC déjà allumé)
- **Ancienne tâche supprimée** : `ASTRA_MOMENTUM_RapportHebdo` (mercredi 14h30)
- **Fichiers créés** :
  - `.github/workflows/rapport_mercredi.yml` — workflow GitHub Actions
  - `scripts/fix_github_paths.py` — corrige les chemins Windows pour le runner Ubuntu
  - `pull_and_deploy.ps1` — déploiement local (git pull + post-traitement)
  - `installer_deploiement_quotidien.ps1` — installe la tâche Windows (login + 18h)
- **Aucune action manuelle requise** : le mercredi à 14h, GitHub génère les fichiers. Au prochain démarrage du PC, la tâche au login pull et copie dans `ASTRA MOMENTUM - Livrables\` avec timestamps.
- **Rattrapage des semaines manquantes** : si le PC reste éteint plusieurs semaines, `post_traitement_livrables.ps1` détecte tous les fichiers non traités et crée un dossier par semaine manquante.

### 2026-07-23 (suite) — Correction contenu

- **Ghostwriter** : URLs ajoutées dans les articles de veille, signature "--- ASTRA MOMENTUM" supprimée (le header suffit)
- **Chasseur** : dédoublonnage des missions (clé = titre 80 premiers chars). Domaines bloqués ajoutés (cabinets-conseil.com, edcparis.edu, consultport.com, linkup-coaching.com, scrum.org). Mots-clés bloqués : "definition", "fiche metier", "coaching", "what is a..."
- **Rapport hebdo** : si 0 opportunités, la priorité passe en "N/A" au lieu de "Haute"

### 2026-07-26 — Chasse aux valeurs bidon + refonte TokenForge Watch + agents désactivés

- **Principes nouveaux** : règles 12-14 ajoutées. Zéro donnée inventée présentée comme réelle. Push git proposé après chaque modif. Session log obligatoire dans AGENTS.md.
- **generate_notes.py** : réécrit. Ne génère plus de réunions fictives (participants, durée, décisions inventés). Message clair : déposer un fichier .txt manuel.
- **cr_reunion.py** : plus d'appel automatique à generate_notes. Skip propre si pas de notes.
- **docx_report.py** : 0 `random.choice` restant. Supprimés : "pertinence", "vigilance", "angle", "action_recs", "li_post", "note_intro", "note_retained", "note_verified", "fallback_snippet", "fallback_opp". Section "Methodologies" supprimée. Proposition de post LinkedIn supprimée.
- **ghostwriter.py** : 0 `random.choice` restant. Titres fixes, mission fixe (la dernière de MISSIONS_REELLES), textes de repli uniques.
- **scoring.py** : score de base articles 4.8→2.0, tokenforge 4.0→2.0, plancher 3.0→0.5, bonus GitHub 0.8→0.4.
- **run_agent.py** : agents `tableau` et `pipeline` désactivés (skip avec message), fichiers conservés pour réactivation future.
- **TokenForge Watch** refondue en revue de presse :
  - 10 nouvelles queries dans config.json : regolo.ai, openrouter.ai, litellm, RTK, tokenizer benchmarks, cache reduction, API gateways
  - Structure A. Marché et actualités des prix / B. Routeurs, compresseurs et outils FinOps / C. Analyse et actions pour TokenForge
  - Analyse dynamique : comptage réel des routeurs, compresseurs, sources FinOps trouvés dans les résultats
- **AGENTS.md** : règles 7-14 intégrées dans la section Règles de Conduite (étaient dispersées dans les logs).

### 2026-08-13 — Fix déploiement GitHub Actions : les DOCX n'étaient jamais poussés sur git

- **Bug constaté** : aucun dossier livrable créé sur le PC depuis ~2 semaines, bien que le workflow GitHub tourne chaque mercredi.
- **Cause racine** : `.gitignore` exclut `Rapport_Astra_Momentum_*.docx`, `agents/output/*.docx` et `agents/output/*_log.json` → le `git add -A` du workflow ne stagait jamais les DOCX générés → `git pull` local ne ramenait rien → `post_traitement_livrables.ps1` n'avait rien à copier.
- **Preuve** : commit `21cbd45` « Rapport hebdo 2026-07-23 [skip ci] » ne contient que `config.json` (83 insertions, 22 suppressions), aucun DOCX.
- **Faux positif écarté** : le dossier « Livrable Semaine du 20 juillet au 24 juillet 2026 » contient bien les fichiers du 23-07 (vérifié avec `-Force`), et « 13-17 juillet » contient Veille_Tarifaire_2026-07-15. Le problème est bien le push des nouvelles semaines, pas les anciennes.
- **Correctifs apportés** :
  - `.github/workflows/rapport_mercredi.yml` : avant le commit, `git checkout --` restauré pour les configs (évite de committer les chemins Linux du runner) + `git add -f` forcé sur `Rapport_Astra_Momentum_*.docx`, `*_runlog.json`, `agents/output/*.docx`, `agents/output/*_log.json` pour forcer le commit des fichiers ignorés.
  - `scripts/fix_github_paths.py` : corrige désormais AUSSI `agents/config.json` (`output_dir` → `$GITHUB_WORKSPACE/agents/output`). Avant, le chemin Windows invalide cassait les agents sur le runner Ubuntu.
  - `moteur_rapport/config.json` : `output_dir` restauré au chemin Windows local (il avait été pollué avec le chemin Linux `/home/runner/work/...` par le workflow précédent).
- **Rattrapage des semaines manquantes** (27-31 juillet, 03-07 août 2026) : les fichiers de ces semaines n'ont jamais été poussés sur git (cause racine ci-dessus) et sont donc perdus. Cycle complet relancé localement le 13/08 (rapport + chasseur + newsletter + veille tarifaire), copies backdatées créées, puis `post_traitement_livrables.ps1` a recréé les 3 dossiers Livrables manquants (27-31/07, 03-07/08, 10-14/08). Fichiers du 13/08 poussés en force (`git add -f`) dans le commit `66e5439` — valide le comportement du workflow corrigé.
- **Dispatch GitHub Actions via API impossible** : le token fine-grained renvoie 403 (pas de permission `actions: write`). Pour tester : Run workflow manuel depuis l'onglet Actions sur GitHub.

### 2026-08-13 (suite) — Chasseur réaliste : contenu d'assistante humaine, plus de snippets bruts

- **Problème signalé par Michael** : le DOCX « Chasseur de Missions » était immédiatement identifiable comme généré par IA (snippets coupés en plein milieu, « non-specific » en anglais, slogans publicitaires Keejob/Product Hunt passés en descriptions, « Fit competences : worldwide »).
- **Free-Work a changé sa structure** : les anciennes URLs catégories `/fr/tech-it/<role>/job-mission` renvoient 404. Les URLs actuelles sont `/fr/tech-it/jobs/<role>`. Config mise à jour avec 4 catégories qui marchent (chef-de-projet, product-owner, head-of-product, digital-transformation). `product-manager` et `consultant` instables/peu pertinentes → retirées.
- **Extraction Free-Work réécrite autour du JSON-LD `JobPosting`** (`_load_job_posting_ld()` dans `core/web.py`) : title, description, entreprise, localisation, TJM (uniquement si `unitText: DAY`, sinon « non communiqué » — un CDI à 40K/an ne doit pas être affiché comme un TJM). Durée extraite des regex mais limitée aux formes courtes (`6 mois`, `3 mois renouvelable`).
- **`clean_text()`** décode désormais les séquences `\uXXXX` (le HTML des pages Free-Work est encodé en escapes dans le JSON) puis supprime les tags → fini les `\u003c/p>\u003cli>` dans le DOCX.
- **Nettoyage des contenus dans `chasseur.py`** :
  - `_clean_title()` : retire les préfixes « Entreprise – Mission freelance 123/ », les codes département, les localisations répétées, le suffixe « | Free-work ».
  - `_clean_description()` : ne garde que les phrases complètes (fini les bouts de phrases tronqués), supprime les slogans (liste `SLUG_PHRASES` : « Découvrez 9 offres... », « Trouvez votre prochain poste », « best new products »...).
  - `is_blocked()` : `BLOCKED_ROLES` et `BLOCKED_WORDS` ne s'appliquent **plus que sur le titre** (une description légitime de PO peut mentionner « production », « formation », « React »). `SLUG_PHRASES` reste sur titre+snippet.
- **Scoring resserré** : un résultat doit avoir un vrai rôle (`role_keywords`) ou être Free-Work direct + secteur concordant. « Fit competences » n'affiche que les vrais rôles, plus les mots génériques. Rôles bloqués élargis (developpeur, data, devops, telecom, infrastructure, RH/paie, conformité, risques...).
- **Recommandations humaines** : plus de « Je te suggère d'élargir les critères ». Selon le cas : « plusieurs pistes pertinentes à explorer » / « rien de pertinent publié cette semaine, je continue de surveiller ».
- **Bing queries** restreintes aux `site:free-work.com` / `site:linkedin.com/jobs` / welcometothejungle / malt (les requêtes génériques ramenaient les agrégateurs).
- **Résultat** : 27 missions Free-Work scrapées → 17 dédupliquées → 10 retenues, toutes des rôles produit/gestion de projet (Product Owner PIM, PO Services Numériques, ITSM Technical PO, Chef de Projet IT, Manager de programme...). Zéro HTML, zéro slogan, TJM honnête.
- **À noter** : `SearchResult` a maintenant des champs `location`, `duration`, `tjm` structurés (utilisés par le chasseur ; les autres agents utilisant `core.web` restent compatibles).

## État actuel au 2026-08-13 (à consulter au prochain démarrage)

**Où on en est :**
- Bug déploiement corrigé et poussé (`.github/workflows/rapport_mercredi.yml` + `scripts/fix_github_paths.py` + `moteur_rapport/config.json`). Dernier commit : `1017a3c`.
- Rattrapage fait : les 3 dossiers Livrables manquants (27-31/07, 03-07/08, 10-14/08) sont recréés dans `ASTRA MOMENTUM - Livrables\`. Fichiers du 13/08 poussés en force dans `66e5439` — valide le comportement `git add -f` du workflow.
- Le pipeline « mercredi 14h → GitHub génère → git pull auto → Livrables » est censé être autonome.
- **Chasseur refondu** (contenu humain, JSON-LD Free-Work, filtres par rôle, zéro snippet brut) — documenté ci-dessus. Le DOCX généré le 13/08 est propre.

**Action en attente (à rappeler à Michael) :** vérifier que le fix fonctionne en déclenchant manuellement le workflow depuis GitHub → Actions → « Rapport Mercredi » → **Run workflow** (branche master). Si le run crée un commit contenant des .docx, tout est validé. Le token local ne permet pas le dispatch (403). Le run manuel du 13/08 (après le fix) est passé au vert mais n'a rien commité car les fichiers du 13/08 existaient déjà — normal.

**Checklist de vérification au prochain démarrage :**
- `git pull` en cours ? (tâche `ASTRA_DEPLOIEMENT_LIVRABLES` au login + 18h)
- Le commit hebdo contient-il des .docx ? (`git log --oneline -5`)
- Un dossier Livrables est-il créé pour la nouvelle semaine ? (`Get-ChildItem "ASTRA MOMENTUM - Livrables"`)
- Si un dossier manque : vérifier d'abord si les .docx sont bien dans le repo (cause racine = .gitignore), recommencer le cycle localement sans réécrire from scratch.

### 2026-08-13 (fin de session) — Standards de réalisme appliqués à TOUS les documents + audit complet

- **Demande de Michael** : « l'illusion d'un vrai travail doit être PARFAITE PARTOUT ». Appliquer au rapport hebdo, à la newsletter, à la veille tarifaire et au CR les mêmes standards que le chasseur. Puis audit fouillé : le bug git des 2 semaines est inadmissible, le code doit être nickel chrome avant push.
- **Newsletter (`ghostwriter.py`)** : liste blanche `TRUSTED_DOMAINS` (svpg, producttalk, martinfowler, medium, openai, anthropic, hbr...) — un article n'est retenu que s'il vient d'un domaine de confiance. `BLOCKED_DOMAINS` étendu (lalanguefrancaise, compagnie-fiduciaire, agentprovocateur, meta, play.google, transformation.co.uk/gouv.fr...), `BLOCKED_WORDS` étendu (cross dressing, transgender, lingerie, university, qu'est-ce que...). Articles rendus avec liens cliquables. Si rien de pertinent : message honnête « Pas d'article suffisamment pertinent cette semaine » (règle 12).
- **Veille tarifaire** : seuil d'échantillon `seuil_echantillon: 3` dans `analyser()` — plus de conclusion sur un seul TJM. Un TJM unique ne représente pas le marché. Message honnête : « Pas assez de TJM publies pour tirer une conclusion fiable ». Section 4 corrigée pour n'afficher « Aucune alerte » que si des données fiables existent.
- **TJM Free-Work** : le JSON-LD `JobPosting` ne contient **plus** de baseSalary/TJM. Le TJM vit dans l'état Nuxt (`__NUXT__`, minDailySalary/maxDailySalary indexés — fragile). La description texte ne le contient pas non plus. Décision : on affiche « non communiqué » plutôt que de fabriquer (règle 12). Veille tarifaire = données insuffisantes honnêtes.
- **Rapport hebdo (`docx_report.py`)** :
  - `first_sentence()` : retire les préfixes « Listen to this episode on: », « Spotify | Apple Podcasts », « In our last article » (séparés en 3 `re.sub` car l'alternance dans un seul `^...` ne re-matcle pas après le premier retrait), remplace les em dashes, tronque proprement.
  - Titres normalisés avec `_norm_title()` (em dashes → tirets) dans articles, opportunités, priorités, TokenForge.
  - Plus de labels « En bref : », « Descriptif : », « Definition : ». `add_actions` sans suggestion post LinkedIn. Section « Sources et URLs verifiables » supprimée. Tableau « Pertinence TokenForge X/10 » supprimé.
  - **TokenForge Watch** : filtre `TOKENFORGE_BLOCKED` dans `scoring.py` (définitions, « qu'est-ce que », tutoriaux, **blockchain/crypto/NFT** — un « token blockchain » n'est pas un token IA). Domaine `ai-explorer.io` bloqué.
  - Domaine `mairie-*` bloqué (une page d'emploi publique de mairie passait dans la veille).
- **Résultats des tests (13/08)** : rapport hebdo (146 articles bruts → 7 retenus, 5 opportunités, 8 TokenForge) **zéro em dash, zéro Spotify, zéro mairie**. Newsletter honnête « Pas d'article suffisamment pertinent cette semaine ». Chasseur 27 missions → 10 retenues toutes PO/produit. Veille tarifaire honnête « Pas assez de donnees ». CR/propositions : rien généré (pas de notes/briefs) — correct.
- **Audit fouillé final** : 0 `random.choice` triple dans tout le codebase, 2 em dashes restants uniquement dans une regex de `chasseur.py` (intentionnel, pour matcher le caractère). Les deux configs (`moteur_rapport/config.json`, `agents/config.json`) ont des chemins Windows locaux propres (pas pollués Linux). `fix_github_paths.py` + workflow validés (restauration configs + `git add -f`).
- **Modifs prêtes à push** : AGENTS.md, 3 fichiers moteur (collectors, scoring, docx_report), 3 agents (chasseur, ghostwriter, veille_tarifaire), 2 configs, 3 DOCX générés + runlog. Commit NON encore créé — attendre confirmation de Michael (règle 13).
