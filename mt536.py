#!/usr/bin/env python3
"""
Cœur de la réconciliation MT536 (relevé des mouvements titres) CTRL / TEST.

Pur Python (bibliothèque standard), réutilise le parseur générique de swift_recon.py.
L'intégration Oracle / SQLite / Excel est décrite dans docs/mt536-integration.md.

Principe :
  1. Chaque message est découpé en trois niveaux :
       - ENTETE      : bloc GENL (+ ADDINFO), comparé par compte/période ;
       - INSTRUMENT  : champs SUBSAFE/FIN hors transactions, comparés par compte/sous-compte/ISIN ;
       - TRANSACTION : un objet par bloc TRAN, avec tous ses champs (LINK, TRANSDET, SETPRTY…).
  2. Les transactions CTRL et TEST sont appariées en plusieurs passes, à l'intérieur
     d'un même groupe compte / sous-compte / période / ISIN :
       - REFERENCE (facultative) : références identiques (ex. 20C:RELA) ;
       - FORTE   : sens, paiement, type, quantité, montant, dates de négociation et de règlement ;
       - FAIBLE  : sens seul, puis appariement par similarité (part de champs identiques),
                   pour qu'une quantité modifiée sorte en ECART et pas en MANQUANT + EN TROP.
  3. Les transactions appariées sont comparées champ à champ.

Usage autonome (messages dans des dossiers, un message par fichier) :
  python mt536.py demo536/ctrl demo536/test -o demo536 --ignore "28E,20C:SEME,98C:PREP,20C:PREV,20C:RELA"
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

from swift_recon import GENERIC_RE, message_type, normalize, raw_fields, security_id

# Blocs répétables d'une transaction, étiquetés par le qualifiant d'un champ repère :
# LINK[RELA] est comparé à LINK[RELA], SETPRTY[DEAG] à SETPRTY[DEAG], quel que soit l'ordre.
LABEL_TAGS = {"LINK": "20C", "SETPRTY": "95"}

# Champs exclus par défaut : n° de page, référence du message, horodatage, lien relevé précédent.
DEFAULT_IGNORE = ["28E", "20C:SEME", "98A:PREP", "98C:PREP", "20C:PREV"]

# Clés d'appariement : 'TAG:QUAL', TAG pouvant être un préfixe ('98' = 98A ou 98C).
# Pour les tags 98x, seule la date (8 premiers caractères) entre dans la clé.
STRONG_KEY = ["22H:REDE", "22H:PAYM", "22F:TRAN", "36B:PSTA", "19A:PSTA", "98:TRAD", "98:ESET"]
WEAK_KEY = ["22H:REDE"]
MIN_SIMILARITY = 0.5

STATUS_DIFF, STATUS_MISSING, STATUS_EXTRA, STATUS_OK = (
    "ECART", "MANQUANT DANS TEST", "EN TROP DANS TEST", "OK")

FieldRow = tuple[str, str, str, str]          # (bloc, tag, qualifiant, valeur normalisée)


# =============================================================================
# 1. Parsing
# =============================================================================

class _Block:
    """Bloc 16R ouvert. L'étiquette est fixée au premier champ repère rencontré."""
    __slots__ = ("name", "label")

    def __init__(self, name: str):
        self.name, self.label = name, None

    def __str__(self) -> str:
        return f"{self.name}[{self.label}]" if self.label else self.name


@dataclass
class Txn:
    side: str
    msg_id: str
    idx: int                 # rang de la transaction dans le message (1..n)
    account: str
    sub: str                 # sous-compte (97A::SAFE du SUBSAFE), '-' si absent
    period: str              # 69A/69B::STAT
    isin: str                # ISIN ou /XX/…, '?' si 35B illisible
    fields: list[FieldRow] = field(default_factory=list)   # bloc relatif au TRAN

    @property
    def label(self) -> str:
        return f"{self.msg_id}#{self.idx}"

    @property
    def group(self) -> tuple[str, str, str, str]:
        return (self.account, self.sub, self.period, self.isin)

    def value(self, spec: str) -> str:
        """Valeur du premier champ 'TAG:QUAL' ('' si absent). Dates 98x réduites à AAAAMMJJ."""
        tag, _, qual = spec.partition(":")
        for _, t, q, v in self.fields:
            if t.startswith(tag) and (not qual or q == qual):
                return v[:8] if t.startswith("98") else v
        return ""


@dataclass
class Statement536:
    side: str
    msg_id: str
    account: str | None
    period: str | None
    page: str | None
    header: list[FieldRow]                              # GENL, ADDINFO…
    instr: list[tuple[str, str, FieldRow]]              # (sous-compte, isin, champ) hors TRAN
    txns: list[Txn]
    error: str | None = None


def explode_536(text: str, side: str = "", msg_id: str = "") -> Statement536:
    stack: list[_Block] = []
    raw: list[tuple] = []          # (blocs, tag, qual, valeur, sous-compte, isin, txn)
    account = period = page = None
    sub, isin, txn = "-", "-", None
    txns: list[Txn] = []

    for tag, value in raw_fields(text):
        if tag == "16R":
            b = _Block(value.strip())
            stack.append(b)
            if b.name == "SUBSAFE":
                sub = "-"
            elif b.name == "FIN":
                isin = "-"
            elif b.name == "TRAN":
                txn = Txn(side, msg_id, len(txns) + 1, "", sub, "", isin)
                txns.append(txn)
            continue
        if tag == "16S":
            b = stack.pop() if stack else None
            if b is not None and b.name == "TRAN":
                txn = None
            elif b is not None and b.name == "FIN":
                isin = "-"
            elif b is not None and b.name == "SUBSAFE":
                sub = "-"
            continue

        qual = ""
        m = GENERIC_RE.match(value)
        if m:
            qual, value = m.group(1), m.group(2).lstrip("/")
        names = [b.name for b in stack]
        top = stack[-1] if stack else None
        if (top is not None and top.label is None and qual
                and top.name in LABEL_TAGS and tag.startswith(LABEL_TAGS[top.name])):
            top.label = qual

        if names == ["GENL"]:
            if tag[:2] == "97" and qual == "SAFE" and account is None:
                account = value.split("\n")[0].strip()
            elif tag[:2] == "69" and qual == "STAT" and period is None:
                period = normalize(value)
            elif tag == "28E":
                page = value.strip()
        elif names == ["SUBSAFE"] and tag[:2] == "97" and qual == "SAFE":
            sub = value.split("\n")[0].strip()
        elif names[-1:] == ["FIN"] and tag == "35B":
            isin = security_id(value) or "?"

        raw.append((tuple(stack), tag, qual, normalize(value), sub, isin, txn))

    # Les étiquettes de blocs sont définitives : on calcule les chemins.
    header, instr = [], []
    for blocks, tag, qual, val, f_sub, f_isin, f_txn in raw:
        if f_txn is not None:
            names = [b.name for b in blocks]
            i = names.index("TRAN")
            rel = "/".join(map(str, blocks[i + 1:])) or "TRAN"
            f_txn.fields.append((rel, tag, qual, val))
        elif blocks and blocks[0].name == "SUBSAFE":
            instr.append((f_sub, f_isin, ("/".join(map(str, blocks)), tag, qual, val)))
        else:
            header.append(("/".join(map(str, blocks)), tag, qual, val))

    for t in txns:
        t.account, t.period = account or "?", period or "?"
    error = None
    if not raw:
        error = "aucun champ trouvé"
    elif not account or not period:
        error = "compte (97A::SAFE) ou période (69a::STAT) introuvable"
    return Statement536(side, msg_id, account, period, page, header, instr, txns, error)


# =============================================================================
# 2. Comparaison
# =============================================================================

def ignore_filter(rules: Iterable[str]) -> Callable[[str, str], bool]:
    """'28E' exclut un tag, '20C:SEME' un tag + qualifiant."""
    tags, pairs = set(), set()
    for r in rules:
        t, _, q = r.strip().partition(":")
        if t:
            (pairs.add((t, q)) if q else tags.add(t))
    return lambda tag, qual: tag in tags or (tag, qual) in pairs


def diff_fields(ctrl: Iterable[FieldRow], test: Iterable[FieldRow]) -> list[tuple]:
    """Compare deux ensembles de champs. Les valeurs identiques s'annulent d'abord,
    les restes de même (bloc, tag, qualifiant) sont mis face à face."""
    cb, tb = defaultdict(Counter), defaultdict(Counter)
    for b, t, q, v in ctrl:
        cb[(b, t, q)][v] += 1
    for b, t, q, v in test:
        tb[(b, t, q)][v] += 1
    out = []
    for k in sorted(cb.keys() | tb.keys()):
        c, t = cb.get(k, Counter()), tb.get(k, Counter())
        common = c & t
        cr, tr = sorted((c - common).elements()), sorted((t - common).elements())
        for i in range(max(len(cr), len(tr))):
            cv = cr[i] if i < len(cr) else None
            tv = tr[i] if i < len(tr) else None
            st = STATUS_DIFF if cv is not None and tv is not None else (
                STATUS_MISSING if tv is None else STATUS_EXTRA)
            out.append((st, *k, cv, tv))
    return out


def _similarity(a: Counter, b: Counter) -> float:
    n = max(sum(a.values()), sum(b.values()))
    return sum((a & b).values()) / n if n else 1.0


def _pair(a: list[Txn], b: list[Txn], sig: dict[int, Counter], min_sim: float):
    """Appariement glouton par similarité décroissante."""
    scored = sorted(((-_similarity(sig[id(x)], sig[id(y)]), i, j)
                     for i, x in enumerate(a) for j, y in enumerate(b)))
    used_a, used_b, pairs = set(), set(), []
    for neg, i, j in scored:
        if -neg < min_sim:
            break
        if i in used_a or j in used_b:
            continue
        used_a.add(i)
        used_b.add(j)
        pairs.append((a[i], b[j]))
    return (pairs,
            [x for i, x in enumerate(a) if i not in used_a],
            [y for j, y in enumerate(b) if j not in used_b])


def match_transactions(ctrl: Sequence[Txn], test: Sequence[Txn],
                       passes: Sequence[tuple[str, Sequence[str]]],
                       ignored: Callable[[str, str], bool],
                       min_similarity: float = MIN_SIMILARITY):
    """Retourne ([(txn_ctrl, txn_test, passe)], orphelins CTRL, orphelins TEST).
    Toutes les passes restent à l'intérieur du groupe compte/sous-compte/période/ISIN.
    Seule la dernière passe applique le seuil de similarité ; dans les autres, une
    transaction dont tous les champs de clé sont vides est reportée à la passe suivante."""
    sig = {id(x): Counter(f for f in x.fields if not ignored(f[1], f[2]))
           for x in (*ctrl, *test)}
    rest_c, rest_t = list(ctrl), list(test)
    matched = []
    for n, (name, spec) in enumerate(passes):
        last = n == len(passes) - 1
        gc, gt = defaultdict(list), defaultdict(list)
        next_c, next_t = [], []
        for src, groups, nxt in ((rest_c, gc, next_c), (rest_t, gt, next_t)):
            for x in src:
                key = tuple(x.value(s) for s in spec)
                if not last and not any(key):
                    nxt.append(x)
                else:
                    groups[(x.group, key)].append(x)
        for k in sorted(gc.keys() | gt.keys()):
            pairs, oc, ot = _pair(gc.get(k, []), gt.get(k, []), sig,
                                  min_similarity if last else 0.0)
            matched += [(c, t, name) for c, t in pairs]
            next_c += oc
            next_t += ot
        rest_c, rest_t = next_c, next_t
    return matched, rest_c, rest_t


TXN_HEADER = ["statut", "compte", "sous_compte", "periode", "isin", "passe",
              "tx_ctrl", "tx_test", "sens", "quantite", "date_reglement",
              "reference_ctrl", "reference_test", "nb_ecarts", "champs_en_ecart"]
DIFF_HEADER = ["statut", "niveau", "compte", "sous_compte", "periode", "isin",
               "tx_ctrl", "tx_test", "bloc", "tag", "qualifiant", "valeur_ctrl", "valeur_test"]
ERROR_HEADER = ["base", "msg_id", "page", "erreur"]


@dataclass
class Result536:
    transactions: list[tuple] = field(default_factory=list)   # TXN_HEADER
    diffs: list[tuple] = field(default_factory=list)          # DIFF_HEADER
    errors: list[tuple] = field(default_factory=list)         # ERROR_HEADER
    stats: dict[str, int] = field(default_factory=dict)


def _txn_row(status, c: Txn | None, t: Txn | None, passe, diffs) -> tuple:
    x = c or t
    tags = sorted({f"{d[2]}::{d[3]}" if d[3] else d[2] for d in diffs})   # d = (statut, bloc, tag, qual, …)
    return (status, *x.group, passe, c.label if c else None, t.label if t else None,
            x.value("22H:REDE"), x.value("36B:PSTA"), x.value("98:ESET"),
            c.value("20C") if c else None, t.value("20C") if t else None,
            len(diffs), ", ".join(tags) or None)


def reconcile(ctrl_msgs: Iterable[tuple[str, str]], test_msgs: Iterable[tuple[str, str]],
              ignore: Sequence[str] = DEFAULT_IGNORE, ref_keys: Sequence[str] = (),
              strong_key: Sequence[str] = STRONG_KEY, weak_key: Sequence[str] = WEAK_KEY,
              min_similarity: float = MIN_SIMILARITY) -> Result536:
    """Compare deux flux de (msg_id, texte MT536). Les messages d'un autre type
    (bloc 2 lisible) sont ignorés. Tout est chargé en mémoire : pour de gros volumes,
    appeler reconcile compte par compte (voir docs/mt536-integration.md)."""
    ignored = ignore_filter(ignore)
    res = Result536()
    stmts: dict[str, list[Statement536]] = {"CTRL": [], "TEST": []}

    for side, msgs in (("CTRL", ctrl_msgs), ("TEST", test_msgs)):
        for msg_id, text in msgs:
            mt = message_type(text)
            if mt and mt != "536":
                continue
            try:
                st = explode_536(text, side, msg_id)
            except Exception as exc:  # un message illisible n'arrête pas le traitement
                res.errors.append((side, msg_id, None, f"{type(exc).__name__}: {exc}"))
                continue
            if st.error:
                res.errors.append((side, msg_id, st.page, st.error))
            stmts[side].append(st)

    # Entête : valeurs distinctes par compte/période (répétées sur chaque page)
    head = {s: defaultdict(set) for s in stmts}
    instr = {s: defaultdict(set) for s in stmts}
    txns: dict[str, list[Txn]] = {s: [] for s in stmts}
    for side, lst in stmts.items():
        for st in lst:
            key = (st.account or "?", st.period or "?")
            head[side][key].update(f for f in st.header if not ignored(f[1], f[2]))
            for sub, isin, f in st.instr:
                if not ignored(f[1], f[2]):
                    instr[side][(*key[:1], sub, key[1], isin)].add(f)
            txns[side] += st.txns

    for key in sorted(head["CTRL"].keys() | head["TEST"].keys()):
        for d in diff_fields(head["CTRL"].get(key, ()), head["TEST"].get(key, ())):
            res.diffs.append((d[0], "ENTETE", key[0], "-", key[1], "-", None, None, *d[1:]))
    for key in sorted(instr["CTRL"].keys() | instr["TEST"].keys()):
        for d in diff_fields(instr["CTRL"].get(key, ()), instr["TEST"].get(key, ())):
            res.diffs.append((d[0], "INSTRUMENT", *key, None, None, *d[1:]))

    passes = ([("REFERENCE", list(ref_keys))] if ref_keys else []) + [
        ("FORTE", list(strong_key)), ("FAIBLE", list(weak_key))]
    matched, orphan_c, orphan_t = match_transactions(
        txns["CTRL"], txns["TEST"], passes, ignored, min_similarity)

    def keep(fs):
        return [f for f in fs if not ignored(f[1], f[2])]

    for c, t, passe in matched:
        diffs = diff_fields(keep(c.fields), keep(t.fields))
        res.transactions.append(_txn_row(STATUS_DIFF if diffs else STATUS_OK, c, t, passe, diffs))
        res.diffs += [(d[0], "TRANSACTION", *c.group, c.label, t.label, *d[1:]) for d in diffs]
    for status, lst in ((STATUS_MISSING, orphan_c), (STATUS_EXTRA, orphan_t)):
        for x in lst:
            c, t = (x, None) if status == STATUS_MISSING else (None, x)
            res.transactions.append(_txn_row(status, c, t, None, []))

    res.transactions.sort(key=lambda r: (r[0] == STATUS_OK, r[1:5], r[6] or "", r[7] or ""))
    res.diffs.sort(key=lambda r: (r[2], r[4], r[1] != "ENTETE", r[5], r[3], r[6] or "", r[7] or "", r[8:11]))
    res.stats = {
        "messages_ctrl": len(stmts["CTRL"]), "messages_test": len(stmts["TEST"]),
        "transactions_ctrl": len(txns["CTRL"]), "transactions_test": len(txns["TEST"]),
        **Counter(f"transactions_{r[0]}" for r in res.transactions),
        **Counter(f"appariees_passe_{p}" for _, _, p in matched),
        "ecarts_entete": sum(r[1] == "ENTETE" for r in res.diffs),
        "ecarts_instrument": sum(r[1] == "INSTRUMENT" for r in res.diffs),
        "ecarts_transaction": sum(r[1] == "TRANSACTION" for r in res.diffs),
        "erreurs": len(res.errors),
    }
    return res


# =============================================================================
# 3. Utilisation autonome (dossiers -> CSV)
# =============================================================================

def _read_dir(folder: Path, pattern: str) -> list[tuple[str, str]]:
    if not folder.is_dir():
        sys.exit(f"Dossier introuvable : {folder}")
    return [(p.name, p.read_text(encoding="utf-8", errors="replace"))
            for p in sorted(folder.glob(pattern))]


def write_csv(path: Path, header: list[str], rows: Iterable[tuple]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(header)
        w.writerows(rows)


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(description="Réconciliation MT536 CTRL / TEST (dossiers de messages)")
    ap.add_argument("ctrl_dir")
    ap.add_argument("test_dir")
    ap.add_argument("-o", "--out", default=".", help="dossier de sortie des CSV")
    ap.add_argument("--pattern", default="*.txt")
    ap.add_argument("--ignore", default=",".join(DEFAULT_IGNORE),
                    help="champs exclus, ex. '28E,20C:SEME,20C:RELA'")
    ap.add_argument("--ref", default="", help="clés de la passe REFERENCE, ex. '20C:RELA'")
    ap.add_argument("--min-sim", type=float, default=MIN_SIMILARITY)
    args = ap.parse_args(argv)

    res = reconcile(_read_dir(Path(args.ctrl_dir), args.pattern),
                    _read_dir(Path(args.test_dir), args.pattern),
                    ignore=[r for r in args.ignore.split(",") if r.strip()],
                    ref_keys=[r.strip() for r in args.ref.split(",") if r.strip()],
                    min_similarity=args.min_sim)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    write_csv(out / "mt536_transactions.csv", TXN_HEADER, res.transactions)
    write_csv(out / "mt536_ecarts.csv", DIFF_HEADER, res.diffs)
    write_csv(out / "mt536_erreurs.csv", ERROR_HEADER, res.errors)
    for k, v in res.stats.items():
        print(f"{k:<40} {v}")
    print(f"CSV écrits dans {out.resolve()}")


if __name__ == "__main__":
    main()
