#!/usr/bin/env python3
"""
Rapport de jours de commit par projet et par mois.

Parcourt récursivement une ou plusieurs racines, découvre tous les dépôts git
(les dépôts SVN sont ignorés naturellement : on ne cherche que des ".git"),
compte pour chaque dépôt les jours distincts avec au moins un commit d'un ou
plusieurs auteurs sur une plage de mois, et écrit le résultat dans un fichier
Excel (.xlsx).

Les racines à scanner et la liste des adresses email à filtrer se
configurent dans config.json (à côté de ce script) — rien en dur dans le
code. Voir README.md.

Usage:
    .venv/bin/python git_days_report.py --start 2026-06 --end 2026-09

    .venv/bin/python git_days_report.py --start 06/2026 --end 09/2026 \
        --root ~/work --root ~/other-projects --out rapport.xlsx

Options utiles:
    --root        Racine à scanner, répétable (défaut: 'roots' dans config.json,
                  sinon ~/work)
    --config      Fichier de config JSON (racines + adresses email à filtrer)
                  (défaut: config.json à côté de ce script)
    --author      Motif (regex) filtrant l'auteur des commits, remplace le
                  fichier de config pour ce run. Passer --author "" pour ne
                  filtrer aucun auteur.
    --out         Chemin du fichier .xlsx en sortie
    --include-empty  Inclure aussi les projets sans aucune activité sur la période
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit(
        "openpyxl n'est pas installé dans cet interpréteur.\n"
        "Utilise le venv du projet, par ex. :\n"
        "  .venv/bin/python git_days_report.py ...\n"
        "(créé avec: python3 -m venv .venv && .venv/bin/pip install openpyxl)"
    )

MOIS_FR = [
    "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]

# Dossiers jamais descendus pendant le scan (bruit / volumineux / non pertinents).
SKIP_DIR_NAMES = {
    "vendor", "node_modules", ".idea", ".vscode", "dist", "build",
    "target", ".venv", "venv", "__pycache__", ".DS_Store", ".next",
    ".cache", "coverage",
}


def parse_month(value: str) -> tuple[int, int]:
    """Accepte 'YYYY-MM' ou 'MM/YYYY' -> (year, month)."""
    m = re.match(r"^(\d{4})-(\d{1,2})$", value)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
    else:
        m = re.match(r"^(\d{1,2})/(\d{4})$", value)
        if not m:
            raise argparse.ArgumentTypeError(
                f"Format de mois invalide: {value!r} (attendu YYYY-MM ou MM/YYYY)"
            )
        month, year = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise argparse.ArgumentTypeError(f"Mois invalide: {value!r}")
    return year, month


def month_range(start: tuple[int, int], end: tuple[int, int]) -> list[tuple[int, int]]:
    (sy, sm), (ey, em) = start, end
    if (sy, sm) > (ey, em):
        raise argparse.ArgumentTypeError("La date de début doit précéder la date de fin")
    months = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        months.append((y, m))
        m += 1
        if m == 13:
            m = 1
            y += 1
    return months


def discover_git_repos(root: Path) -> list[Path]:
    """Retourne la liste des dossiers contenant un .git, triés par chemin relatif."""
    repos: list[Path] = []
    for dirpath, dirnames, _filenames in _walk_pruned(root):
        if ".git" in dirnames:
            repos.append(Path(dirpath))
            dirnames.remove(".git")  # inutile de descendre dedans
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIR_NAMES]
    repos.sort(key=lambda p: str(p.relative_to(root)).lower())
    return repos


def remote_url(repo: Path) -> str:
    """URL normalisée du remote 'origin', ou '' si absent/local."""
    try:
        out = subprocess.run(
            ["git", "-C", str(repo), "remote", "get-url", "origin"],
            capture_output=True, text=True, timeout=5,
        )
    except Exception:
        return ""
    url = out.stdout.strip().lower()
    if url.endswith(".git"):
        url = url[:-4]
    return url


def repo_label(root: Path, repo: Path, multi_root: bool) -> str:
    """Étiquette d'un dépôt : chemin relatif à sa racine, préfixé par le nom
    de la racine seulement quand plusieurs racines sont scannées (pour éviter
    toute ambiguïté entre deux racines qui partageraient un sous-dossier)."""
    rel = str(repo.relative_to(root))
    return f"{root.name}/{rel}" if multi_root else rel


def discover_all_repos(roots: list[Path]) -> list[tuple[Path, Path]]:
    """Retourne une liste de (racine, dépôt) pour toutes les racines données."""
    pairs: list[tuple[Path, Path]] = []
    for root in roots:
        for repo in discover_git_repos(root):
            pairs.append((root, repo))
    return pairs


def dedupe_repos(pairs: list[tuple[Path, Path]], multi_root: bool) -> list[tuple[Path, Path]]:
    """Certains projets ont deux clones du même remote (ex: un dossier 'svn_xxx'
    laissé par un ancien pont git-svn, à côté du clone git actif — ou le même
    projet cloné sous deux racines différentes). On ne garde qu'un dépôt par
    remote, en préférant celui dont le chemin ne contient pas 'svn' — c'est le
    clone actif. Les dépôts sans remote (locaux/POC) sont tous conservés, ils
    ne peuvent pas être des doublons."""
    groups: dict[str, list[tuple[Path, Path]]] = {}
    for root, repo in pairs:
        url = remote_url(repo)
        key = url if url else f"__local__:{repo}"
        groups.setdefault(key, []).append((root, repo))

    kept: list[tuple[Path, Path]] = []
    for key, group in groups.items():
        if len(group) == 1:
            kept.append(group[0])
            continue
        non_svn = [(r, p) for r, p in group if "svn" not in repo_label(r, p, multi_root).lower()]
        chosen = sorted(non_svn or group, key=lambda rp: repo_label(rp[0], rp[1], multi_root))[0]
        for r, p in group:
            if (r, p) != chosen:
                print(f"  = doublon ignoré (même remote que {repo_label(*chosen, multi_root)}): {repo_label(r, p, multi_root)}")
        kept.append(chosen)

    kept.sort(key=lambda rp: repo_label(rp[0], rp[1], multi_root).lower())
    return kept


def _walk_pruned(root: Path):
    import os

    for dirpath, dirnames, filenames in os.walk(root):
        yield dirpath, dirnames, filenames


DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent / "config.json"


def _load_config(config_path: Path) -> dict:
    if not config_path.is_file():
        return {}
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        sys.exit(f"Impossible de lire {config_path} : {exc}")
    if not isinstance(raw, dict):
        sys.exit(f"{config_path} : le contenu doit être un objet JSON")
    return raw


def load_config_authors(config_path: Path) -> list[str]:
    """Lit la liste 'authors' (adresses email) depuis un fichier JSON.
    Retourne [] si le fichier est absent ou la liste vide (= pas de filtre)."""
    authors = _load_config(config_path).get("authors", [])
    if not isinstance(authors, list):
        sys.exit(f"{config_path} : la clé 'authors' doit être une liste d'adresses email")
    return [a.strip() for a in authors if isinstance(a, str) and a.strip()]


def load_config_roots(config_path: Path) -> list[str]:
    """Lit la liste 'roots' (dossiers racines à scanner) depuis un fichier JSON.
    Retourne [] si le fichier est absent ou la liste vide."""
    roots = _load_config(config_path).get("roots", [])
    if not isinstance(roots, list):
        sys.exit(f"{config_path} : la clé 'roots' doit être une liste de chemins")
    return [r.strip() for r in roots if isinstance(r, str) and r.strip()]


def commit_days(repo: Path, authors: list[str], since: str, until: str) -> set[str]:
    """Jours distincts (YYYY-MM-DD) avec un commit de l'un des `authors`
    (OR — git combine plusieurs --author) dans [since, until)."""
    cmd = [
        "git", "-C", str(repo), "log", "--all",
        f"--since={since}", f"--until={until}",
        "--format=%ad", "--date=format:%Y-%m-%d",
    ]
    for author in authors:
        cmd.insert(4, f"--author={author}")
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except Exception as exc:
        print(f"  ! {repo}: erreur git ({exc})", file=sys.stderr)
        return set()
    if out.returncode != 0:
        print(f"  ! {repo}: {out.stderr.strip()}", file=sys.stderr)
        return set()
    return {line for line in out.stdout.splitlines() if line}


def build_report(roots: list[Path], months: list[tuple[int, int]], authors: list[str],
                  include_empty: bool) -> tuple[list[str], dict[str, list[int]]]:
    since = f"{months[0][0]:04d}-{months[0][1]:02d}-01"
    last_y, last_m = months[-1]
    if last_m == 12:
        until_y, until_m = last_y + 1, 1
    else:
        until_y, until_m = last_y, last_m + 1
    until = f"{until_y:04d}-{until_m:02d}-01T00:00:00"

    col_labels = [f"{MOIS_FR[m]} {y}" for (y, m) in months]

    multi_root = len(roots) > 1
    pairs = discover_all_repos(roots)
    print(f"Dépôts git trouvés sous {', '.join(str(r) for r in roots)} : {len(pairs)}")
    pairs = dedupe_repos(pairs, multi_root)
    print(f"Dépôts retenus après déduplication (même remote) : {len(pairs)}")

    data: dict[str, list[int]] = {}
    for root, repo in pairs:
        label = repo_label(root, repo, multi_root)
        days = commit_days(repo, authors, since, until)
        if not days:
            if include_empty:
                data[label] = [0] * len(months)
            continue
        counts = [0] * len(months)
        for i, (y, m) in enumerate(months):
            prefix = f"{y:04d}-{m:02d}"
            counts[i] = sum(1 for d in days if d.startswith(prefix))
        if include_empty or any(counts):
            data[label] = counts

    return col_labels, data


def write_xlsx(out_path: Path, col_labels: list[str], data: dict[str, list[int]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Jours de commit"

    header_fill = PatternFill("solid", fgColor="1F2933")
    header_font = Font(bold=True, color="FFFFFF")
    total_font = Font(bold=True)
    total_fill = PatternFill("solid", fgColor="EDEFF2")

    headers = ["Projet"] + col_labels + ["Total"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center" if cell.column > 1 else "left")

    # trie par total décroissant, ordre alphabétique en cas d'égalité
    rows = sorted(
        data.items(), key=lambda kv: (-sum(kv[1]), kv[0].lower())
    )

    col_totals = [0] * len(col_labels)
    for label, counts in rows:
        total = sum(counts)
        ws.append([label] + counts + [total])
        for i in range(len(col_labels)):
            col_totals[i] += counts[i]

    total_row = ["Total"] + col_totals + [sum(col_totals)]
    ws.append(total_row)
    last_row = ws.max_row
    for cell in ws[last_row]:
        cell.font = total_font
        cell.fill = total_fill

    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row - 1}"

    ws.column_dimensions["A"].width = max(28, max((len(r[0]) for r in rows), default=10) + 2)
    for i in range(len(col_labels) + 1):
        ws.column_dimensions[get_column_letter(2 + i)].width = 12

    wb.save(out_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--start", required=True, type=parse_month, help="Mois de début, ex: 2026-06 ou 06/2026")
    parser.add_argument("--end", required=True, type=parse_month, help="Mois de fin, ex: 2026-09 ou 09/2026")
    parser.add_argument("--root", action="append", default=None,
                         help="Racine à scanner, répétable (défaut: 'roots' dans config.json, sinon ~/work)")
    parser.add_argument("--config", default=None,
                         help=f"Fichier de config JSON (racines + adresses email à filtrer) "
                              f"(défaut: {DEFAULT_CONFIG_PATH.name} à côté de ce script)")
    parser.add_argument("--author", default=None,
                         help="Motif (regex git) filtrant l'auteur, remplace le fichier de config "
                              "pour ce run. Passer --author \"\" pour ne filtrer aucun auteur.")
    parser.add_argument("--out", default=None, help="Fichier .xlsx en sortie")
    parser.add_argument("--include-empty", action="store_true",
                         help="Inclure aussi les projets sans aucune activité sur la période")
    args = parser.parse_args()

    config_path = Path(args.config).expanduser() if args.config else DEFAULT_CONFIG_PATH

    if args.root:
        raw_roots = args.root
        roots_source = "--root"
    else:
        raw_roots = load_config_roots(config_path)
        roots_source = f"{config_path.name}" if raw_roots else None
        if not raw_roots:
            raw_roots = ["~/work"]
            roots_source = "défaut"

    roots: list[Path] = []
    for r in raw_roots:
        p = Path(r).expanduser().resolve()
        if not p.is_dir():
            print(f"! Racine introuvable, ignorée : {p}", file=sys.stderr)
            continue
        roots.append(p)
    if not roots:
        sys.exit("Aucune racine valide à scanner.")
    print(f"Racines scannées ({roots_source}) : {', '.join(str(r) for r in roots)}")

    months = month_range(args.start, args.end)

    if args.author is not None:
        authors = [args.author] if args.author else []
        print(f"Filtre auteur (--author) : {authors[0]!r}" if authors else "Aucun filtre auteur (--author \"\")")
    else:
        authors = load_config_authors(config_path)
        if authors:
            print(f"Filtre auteur ({config_path.name}) : {', '.join(authors)}")
        else:
            print(f"Aucun filtre auteur — {config_path} absent ou vide, tous les commits sont comptés")
        # adresses littérales issues du config : échappées pour éviter que
        # '.' ou '+' ne soient interprétés comme une regex par `git --author`
        authors = [re.escape(a) for a in authors]

    col_labels, data = build_report(roots, months, authors, args.include_empty)

    if args.out:
        out_path = Path(args.out).expanduser()
    else:
        sy, sm = months[0]
        ey, em = months[-1]
        out_path = Path(f"rapport_commits_{sy:04d}-{sm:02d}_a_{ey:04d}-{em:02d}.xlsx")

    write_xlsx(out_path, col_labels, data)
    print(f"\n{len(data)} projet(s) avec activité écrit(s) dans : {out_path.resolve()}")


if __name__ == "__main__":
    main()
