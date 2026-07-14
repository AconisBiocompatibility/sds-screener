# 🔬 SDS Biocompatibility Screener — SIGVARIS GROUP

Application Streamlit d'analyse automatique des Fiches de Données de Sécurité.
**Cadre réglementaire** : MDR 2017/745 / ISO 10993-1:2025

---

## Structure du dossier

```
sds_screener/
├── app.py                    ← Application principale
├── agent_prompt_generic.txt  ← Prompt de l'agent IA (générique, avec placeholders)
├── requirements.txt          ← Dépendances Python
└── README.md                 ← Ce fichier
```

---

## Étape 1 — Vérifier le fichier agent_prompt_generic.txt

Le fichier `agent_prompt_generic.txt` est déjà inclus. Il contient des placeholders :
- `<<CLIENT_NAME>>` — nom du client
- `<<DEVICE_DESCRIPTION>>` — description du dispositif médical
- `<<CONTACT_SCENARIO>>` — scénario de contact ISO 10993-1
- `<<PRODUCT_TYPE_VOCABULARY>>` — tableau des types de produits

Ces placeholders sont remplacés automatiquement par l'application via la **barre latérale de configuration client**.

---

## Étape 2 — Déployer sur Streamlit Cloud (gratuit)

### 2a. Créer un compte GitHub (si pas déjà fait)
Allez sur [github.com](https://github.com) et créez un compte gratuit.

### 2b. Créer un dépôt GitHub
1. Cliquez **"New repository"**
2. Nommez-le `sds-screener` (ou autre nom de votre choix)
3. Laissez-le **privé** (Private) — vos fichiers ne seront pas publics
4. Cliquez **"Create repository"**

### 2c. Uploader les fichiers
Dans votre nouveau dépôt GitHub :
1. Cliquez **"uploading an existing file"**
2. Glissez-déposez ces 4 fichiers :
   - `app.py`
   - `agent_prompt.txt`
   - `requirements.txt`
   - `README.md`
3. Cliquez **"Commit changes"**

### 2d. Déployer sur Streamlit Cloud
1. Allez sur [share.streamlit.io](https://share.streamlit.io)
2. Connectez-vous avec GitHub
3. Cliquez **"New app"**
4. Sélectionnez votre dépôt `sds-screener`
5. Fichier principal : `app.py`
6. Cliquez **"Deploy"**

✅ Votre application sera accessible via une URL publique (ou privée selon vos paramètres).

---

## Utilisation

1. **Ouvrez l'URL** de votre application Streamlit
2. **Barre latérale gauche** :
   - Entrez votre clé API Anthropic (`sk-ant-...`)
   - Chargez le fichier `SIGVARIS FDS.xlsx` (votre fichier avec les bases de données)
3. **Zone principale** : déposez une ou plusieurs FDS au format PDF
4. Cliquez **"Analyser"**
5. Attendez 1-2 minutes par FDS
6. **Téléchargez** le fichier Excel mis à jour

---

## Coûts estimés (API Anthropic)

| Volume | Coût estimé |
|--------|-------------|
| 5 FDS/mois | ~0.50 € |
| 20 FDS/mois | ~2 € |
| 100 FDS/mois | ~10 € |

*Basé sur Claude Sonnet 4.6 — tarif approximatif, variable selon la taille des FDS.*

---

## Limites connues (v1.0)

- Les PDF **scannés en image** (non searchable) ne peuvent pas être lus automatiquement
- Les bases de données REACH Annex XVII, TEDX, ED Assessment ne sont pas encore
  intégrées en lookup direct — Claude utilise sa connaissance intégrée pour ces tables
- La colonne **AF (Toxicologist Comment)** est toujours laissée vide (réservée à l'expert)
- La colonne **AG (Corrected Alert Level)** est laissée vide pour le relecteur humain

---

## Support

Pour toute question, contactez Élodie Saudrais.
