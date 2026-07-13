# AssistanteSASU — Session Memory

Dernière mise à jour : 2026-07-12

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
├── send_report.py               ← Envoi email via Brevo SMTP
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
| `send_report.py` | Envoi du rapport par email via Brevo SMTP |
| `lancer_mercredi.bat` | Lanceur unique mercredi : délai aléatoire → rapport → agents → email |
| `installer_tache_mercredi.ps1` | Script pour créer la tâche planifiée Windows |

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

# Envoyer le dernier rapport par email
python send_report.py

# Lancer le cycle complet (rapport + agents + email)
lancer_mercredi.bat

# Planifier la tâche Windows (mercredi 14h30)
pwsh -ExecutionPolicy Bypass -File installer_tache_mercredi.ps1
```

## Règles de Conduite (gravées dans le marbre le 2026-07-10)

1. **Simplicité avant tout.** Pas de refacto, pas d'industrialisation, pas de framework. Chaque modification doit être comprise en 30 secondes par Michael.
2. **Contrôle total.** Michael doit pouvoir ouvrir, lire, modifier chaque fichier sans outil spécifique. Pas de magie, pas de dépendances cachées.
3. **Une chose à la fois.** Pas de listes de 10 chantiers. On fait une chose, on valide, on passe à la suivante.
4. **Ne jamais réécrire from scratch.** Toute modification est une édition ciblée.
5. **Ne jamais inventer.** Pas de bluff, pas de "je pense que". Si tu ne sais pas, dis-le.
6. **Richesse du DOCX = intouchable.** `docx_report.py`, `scoring.py`, `collectors.py` ne sont pas simplifiés — ce sont les fichiers qui produisent la valeur.

## Live Session Log

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

## Prochaines étapes (optionnelles, non prioritaires)

- [x] Créer un `lancer_mercredi.bat` unique qui enchaîne rapport → agents → email

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
