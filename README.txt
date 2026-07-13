=====================================
  ASTRA MOMENTUM - Agents autonomes
=====================================

Ce dossier contient le systeme d'agents autonomes pour
la prospection de missions freelance et la generation de rapports.

Auteur : Michael ASSAYAG - Président ASTRA MOMENTUM


POUR ROUVRIR LE PROJET DANS OPENCODE
=====================================

1. Ouvrir OpenCode
2. Dans le terminal OpenCode, taper :

   cd "C:\Users\micas\OneDrive\Bureau\Rapports Assistant SASU"

3. Quand le chat s'ouvre, dire simplement :
    "continue le projet ASTRA MOMENTUM"


LANCER LES AGENTS
=====================================

Double-clic sur le fichier :

   agents\lancer_agents.bat

Cela lance tous les agents et genere les rapports dans :

   agents\output\

OU en ligne de commande (dans le dossier agents\) :

   lancer_agents.bat missions   -> juste la chasse aux missions
   lancer_agents.bat pipeline   -> suivi pipeline CRM
   lancer_agents.bat tableau    -> tableau de bord mensuel
   lancer_agents.bat cr         -> compte rendu de reunion
   lancer_agents.bat branding   -> contenus newsletter
   lancer_agents.bat proposition -> propositions commerciales
   lancer_agents.bat tarif      -> veille tarifaire
   lancer_agents.bat all        -> tous les agents


CE QUI EXISTE
=====================================

Agent 1 : Chasseur de Missions
  - Scrape 8 pages Free-Work + Bing
  - Extrait les TJM des missions
  - Classe les missions / POSTULER / A QUALIFIER / SURVEILLER
  - Possibilite d'ajouter ses missions perso dans :
    agents\agent_chasseur_leads\missions_manuelles.csv

Agent 2 : Pipeline CRM
  - Suivi du pipeline client

Agent 3 : Tableau de bord mensuel SASU
  - CA, TJM, taux occupation, projection annuelle

Agent 4 : Compte rendu de reunion
  - Notes .txt -> CR structure .docx

Agent 5 : Personal Branding
  - 4 contenus editoriaux pour newsletter Substack
  - Positionnement freelance senior Product & IA

Agent 6 : Offres consulting
  - Propositions commerciales .docx depuis un brief

Agent 7 : Veille tarifaire
  - Benchmark TJM du marche
  - Generation automatique de drafts email de prospection


CE QUI RESTE A FAIRE
=====================================

[ ] Envoi automatique des rapports par email
[ ] Planification automatique (tous les mercredis 15h)
[ ] Ajouter d'autres sources de scraping
[ ] Ameliorer le scoring des missions
[ ] Creer une landing page Substack automatique
