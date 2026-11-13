# Serveur Web — Gestionnaire de Licences IATF

Ce dossier contient **l'application web a installer sur un serveur** pour gerer les licences de l'add-in Excel IATF Consultant Assistant.

## Structure

```
serveur-web/
├── web/                 Interface web d'administration (React + Vite)
│   ├── App.tsx          Application principale
│   ├── main.tsx         Point d'entree
│   ├── index.html       Page HTML
│   └── styles.css       Styles
├── supabase/            Backend (base de donnees + API)
│   ├── config.toml      Configuration Supabase
│   ├── functions/
│   │   └── license-api/ API REST de gestion des licences
│   └── migrations/      Migrations SQL (schema de base)
├── package.json         Dependances Node.js
├── tsconfig.json        Configuration TypeScript
├── vite.config.ts       Configuration Vite
└── .env                 Variables d'environnement (URL + cles Supabase)
```

## Fonctionnalites

- **Generer des licences** (Standard ou Premium) avec date d'expiration
- **Lister et rechercher** toutes les licences
- **Activer / Revoquer / Prolonger** les licences
- **Suivre les activations** par machine
- **Statistiques** (total, actives, expirees, premium, etc.)
- **API REST** pour la verification depuis Excel

## Installation sur serveur

### 1. Installer Node.js

```bash
# Telechargez Node.js 18+ depuis https://nodejs.org/
# Puis dans le dossier serveur-web/ :
npm install
```

### 2. Configurer les variables d'environnement

Le fichier `.env` contient :
- `SUPABASE_URL` : URL de votre projet Supabase
- `SUPABASE_ANON_KEY` : cle publique (anon)
- `SUPABASE_SERVICE_ROLE_KEY` : cle privee (service role)

Ces valeurs sont deja configurees pour votre projet Supabase.

### 3. Deployer la base de donnees et l'API

La migration SQL et l'edge function sont deja deployees sur Supabase.
Pour verifier :

```bash
# L'API est accessible a :
# https://VOTRE-PROJET.supabase.co/functions/v1/license-api
```

### 4. Lancer l'interface web

```bash
npm run dev
```

L'interface est accessible sur `http://localhost:5174`

### 5. Build pour la production

```bash
npm run build
```

Les fichiers statiques sont generes dans `dist/` et peuvent etre heberges sur n'importe quel serveur web (Nginx, Apache, Vercel, Netlify, etc.)

## API REST

L'API est accessible via l'edge function Supabase :

| Endpoint | Methode | Description |
|----------|---------|-------------|
| `/verify?key=XXX&machineId=YYY` | GET | Verifier une licence |
| `/activate` | POST | Activer une licence sur une machine |
| `/generate` | POST | Generer une nouvelle licence (admin) |
| `/list?page=1&search=XXX&status=ACTIVE` | GET | Lister les licences |
| `/stats` | GET | Statistiques globales |
| `/revoke/:id` | PUT | Revoquer une licence |
| `/extend/:id` | PUT | Prolonger une licence |
| `/:id` | DELETE | Supprimer une licence |
| `/activations/:id` | GET | Lister les activations d'une licence |

## Lier avec Excel

Dans le dossier `excel-local/`, le fichier `src/modules/M_LicenceWeb.bas` contient deux constantes a modifier :

```vba
Private Const LICENCE_API_URL As String = "https://VOTRE-PROJET.supabase.co/functions/v1/license-api"
Private Const LICENCE_API_KEY As String = "VOTRE_ANON_KEY"
```

Remplacez par les valeurs de votre `.env` (SUPABASE_URL et SUPABASE_ANON_KEY).
