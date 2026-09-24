# git-days-report

Compte, pour chaque dépôt git trouvé sous un dossier racine, le nombre de
**jours distincts avec au moins un commit** sur une plage de mois donnée, et
écrit le résultat dans un fichier Excel (`.xlsx`) — une ligne par projet, une
colonne par mois, avec totaux.

- Ne regarde que les dépôts **git** (les dépôts SVN, sans `.git`, sont
  ignorés naturellement).
- **Déduplique** les projets qui ont deux clones du même remote (ex: un
  dossier `svn_xxx` laissé par un ancien pont, à côté du clone actif) — un
  seul est compté.
- Par défaut, ne compte que les commits dont l'auteur correspond
  **exactement** à ton email git global (`git config --global user.email`).

## Installation (une seule fois)

```bash
cd ~/work/scripts/git-days-report
python3 -m venv .venv
.venv/bin/pip install openpyxl
```

## Utilisation

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09
```

Ça scanne `~/work` (défaut) et écrit
`rapport_commits_2026-06_a_2026-09.xlsx` dans le dossier courant.

### Options

| Option | Défaut | Description |
|---|---|---|
| `--start` | *(requis)* | Mois de début — `YYYY-MM` ou `MM/YYYY` (ex: `2026-06` ou `06/2026`) |
| `--end` | *(requis)* | Mois de fin, même format |
| `--root` | `~/work` | Dossier racine à scanner |
| `--author` | email git global | Motif (regex) filtrant l'auteur des commits. `--author ""` pour ne filtrer personne |
| `--out` | `rapport_commits_<début>_a_<fin>.xlsx` | Chemin du fichier Excel en sortie |
| `--include-empty` | désactivé | Inclut aussi les projets sans aucune activité sur la période |

### Exemples

Rapport sur toute l'année pour un autre dossier :

```bash
.venv/bin/python git_days_report.py --start 01/2026 --end 12/2026 \
  --root ~/work --out rapport_2026.xlsx
```

Rapport pour tous les auteurs d'un projet (pas seulement toi) :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 --author ""
```

Filtrer sur un autre auteur précis :

```bash
.venv/bin/python git_days_report.py --start 2026-06 --end 2026-09 \
  --author j.dupont@highconnexion.com
```

## Sortie

Un fichier `.xlsx` avec une feuille "Jours de commit" :
- une ligne par projet (chemin relatif à `--root`), triée par activité décroissante,
- une colonne par mois de la plage demandée,
- une colonne et une ligne "Total",
- en-têtes figés et filtre automatique activés.

## Note méthodologique

Un jour compte de la même façon qu'il contienne 1 ou 10 commits — c'est un
indicateur d'activité, pas une mesure de temps passé.
