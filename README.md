# git-days-report

Compte, pour chaque dépôt git trouvé sous une ou plusieurs racines, le
nombre de **jours distincts avec au moins un commit** sur une plage de mois
donnée, et écrit le résultat dans un fichier Excel (`.xlsx`) — une ligne par
projet, une colonne par mois, avec totaux.

- Ne regarde que les dépôts **git** (les dépôts SVN, sans `.git`, sont
  ignorés naturellement).
- **Déduplique** les projets qui ont deux clones du même remote (ex: un
  dossier `svn_xxx` laissé par un ancien pont, à côté du clone actif, ou le
  même projet cloné sous deux racines) — un seul est compté.
- Les racines à scanner et le filtre auteur se configurent dans
  **`config.json`** (à côté du script, jamais en dur dans le code).

## Installation (une seule fois)

```bash
cd ~/scripts/git-days-report
python3 -m venv .venv
.venv/bin/pip install openpyxl
```

## Configuration

`config.json` :

```json
{
  "roots": [
    "~/work"
  ],
  "authors": [
    "j.soffichiti@highconnexion.com"
  ]
}
```

- **`roots`** : un ou plusieurs dossiers racines à scanner. Liste vide ou
  clé absente = `~/work` par défaut.
- **`authors`** : une ou plusieurs adresses email dont les commits doivent
  être comptés (OR entre elles). Liste vide ou clé absente = aucun filtre,
  tous les commits comptent.

## Utilisation

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09
```

Scanne les racines de `config.json` et écrit
`rapport_commits_2026-06_a_2026-09.xlsx` dans le dossier courant.

### Options

| Option | Défaut | Description |
|---|---|---|
| `--start` | *(requis)* | Mois de début — `YYYY-MM` ou `MM/YYYY` (ex: `2026-06` ou `06/2026`) |
| `--end` | *(requis)* | Mois de fin, même format |
| `--root` | `roots` dans config.json, sinon `~/work` | Racine à scanner. Répétable (`--root A --root B`) pour en remplacer plusieurs |
| `--config` | `config.json` à côté du script | Fichier JSON (racines + adresses email à filtrer) |
| `--author` | *(vide)* | Motif (regex git) qui remplace le fichier de config pour ce run. `--author ""` pour ne filtrer personne |
| `--out` | `rapport_commits_<début>_a_<fin>.xlsx` | Chemin du fichier Excel en sortie |
| `--include-empty` | désactivé | Inclut aussi les projets sans aucune activité sur la période |

### Exemples

Rapport sur toute l'année, racines et auteurs pris dans `config.json` :

```bash
.venv/bin/python git_days_report.py --start 01/2026 --end 12/2026 --out rapport_2026.xlsx
```

Scanner plusieurs racines pour ce run, sans toucher au fichier de config :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 \
  --root ~/work --root ~/autres-projets
```

Rapport pour tous les auteurs d'un projet (pas seulement ceux de `config.json`) :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 --author ""
```

Filtrer sur une adresse précise sans toucher au fichier de config :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 \
  --author j.dupont@highconnexion.com
```

Utiliser un fichier de config différent (ex: pour un autre collègue) :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 \
  --config config-jdupont.json
```

## Sortie

Un fichier `.xlsx` avec une feuille "Jours de commit" :
- une ligne par projet (chemin relatif à sa racine — préfixé du nom de la
  racine si plusieurs racines sont scannées), triée par activité décroissante,
- une colonne par mois de la plage demandée,
- une colonne et une ligne "Total",
- en-têtes figés et filtre automatique activés.

## Note méthodologique

Un jour compte de la même façon qu'il contienne 1 ou 10 commits — c'est un
indicateur d'activité, pas une mesure de temps passé.
