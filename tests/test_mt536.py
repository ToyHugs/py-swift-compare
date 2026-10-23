from samples536 import DE5, FR2, LU1, fin, msg536, tx, demo_messages

from mt536 import (DEFAULT_IGNORE, explode_536, ignore_filter, match_transactions,
                   reconcile, STRONG_KEY, WEAK_KEY)

IGNORE_REF = DEFAULT_IGNORE + ["20C:RELA"]


def _run(ctrl: dict, test: dict, **kw):
    return reconcile(list(ctrl.items()), list(test.items()), **kw)


def test_parse_structure():
    st = explode_536(msg536("S", "1/ONLY", "ACC9", [
        fin(FR2, tx("R1", "RECE", "50,", "2100,", trad="20261013", party="BRK")),
        fin("/XX/AVQ1\nSTRUCTURE", tx("R2", "DELI", "1,")),
    ]), "CTRL", "m1")
    assert st.error is None
    assert (st.account, st.period, st.page) == ("ACC9", "20261001/20261020", "1/ONLY")
    assert [t.isin for t in st.txns] == ["FR0000000002", "/XX/AVQ1"]
    t = st.txns[0]
    assert t.value("36B:PSTA") == "UNIT/50"
    assert t.value("98:TRAD") == "20261013"
    blocs = {f[0] for f in t.fields}
    assert {"LINK[RELA]", "TRANSDET", "TRANSDET/SETPRTY[DEAG]", "TRANSDET/SETPRTY[PSET]"} <= blocs


def test_identical_except_pagination_order_format_and_refs():
    a = tx("C1", "RECE", "50,", "2100,", eset="20261015")
    b = tx("C2", "DELI", "20,", "840,", eset="20261016")
    ctrl = {"c1": msg536("X1", "1/FIRS", "ACC1", [fin(FR2, a)]),
            "c2": msg536("X2", "2/LAST", "ACC1", [fin(FR2, b), fin(LU1, a)])}
    b2 = tx("T2", "DELI", "20,00", "840,00", eset="20261016")
    a2 = tx("T1", "RECE", "50,", "2100,", eset="20261015")
    test = {"t1": msg536("Y1", "1/ONLY", "ACC1", [fin(LU1, a2), fin(FR2, b2, a2)])}
    res = _run(ctrl, test, ignore=IGNORE_REF)
    assert res.diffs == []
    assert {r[0] for r in res.transactions} == {"OK"}
    assert len(res.transactions) == 3


def test_refs_differ_are_reported_when_not_ignored():
    ctrl = {"c": msg536("X", "1/ONLY", "ACC1", [fin(FR2, tx("C1", "RECE", "5,"))])}
    test = {"t": msg536("Y", "1/ONLY", "ACC1", [fin(FR2, tx("T1", "RECE", "5,"))])}
    res = _run(ctrl, test)
    assert [(d[9], d[10], d[11], d[12]) for d in res.diffs] == [("20C", "RELA", "C1", "T1")]


def test_reference_pass_disambiguates_duplicates():
    # Deux réceptions identiques sauf la contrepartie ; références stables.
    c = [tx("R1", "RECE", "5,", party="A"), tx("R2", "RECE", "5,", party="B")]
    t = [tx("R2", "RECE", "5,", party="B"), tx("R1", "RECE", "5,", party="C")]
    ctrl = {"c": msg536("X", "1/ONLY", "ACC1", [fin(FR2, *c)])}
    test = {"t": msg536("Y", "1/ONLY", "ACC1", [fin(FR2, *t)])}
    res = _run(ctrl, test, ref_keys=["20C:RELA"])
    assert {r[5] for r in res.transactions} == {"REFERENCE"}
    diffs = [(d[11], d[12]) for d in res.diffs]
    assert diffs == [("A", "C")]


def test_weak_pass_pairs_by_similarity():
    ctrl = {"c": msg536("X", "1/ONLY", "ACC1", [fin(FR2,
            tx("C1", "DELI", "20,", "840,", trad="20261014", eset="20261016", party="P1"),
            tx("C2", "DELI", "10,", "420,", trad="20261016", eset="20261018", party="P2"))])}
    test = {"t": msg536("Y", "1/ONLY", "ACC1", [fin(FR2,
            tx("T2", "DELI", "10,", "420,", trad="20261016", eset="20261019", party="P2"),
            tx("T1", "DELI", "25,", "1050,", trad="20261014", eset="20261016", party="P1"))])}
    res = _run(ctrl, test, ignore=IGNORE_REF)
    rows = {r[9]: r for r in res.transactions}         # par quantité CTRL
    assert rows["UNIT/20"][5] == "FAIBLE" and rows["UNIT/20"][14] == "19A::PSTA, 36B::PSTA"
    assert rows["UNIT/10"][5] == "FAIBLE" and rows["UNIT/10"][14] == "98A::ESET"


def test_too_different_stays_missing_and_extra():
    ctrl = {"c": msg536("X", "1/ONLY", "ACC1", [fin(FR2,
            tx("C1", "RECE", "1,", "10,", paym="FREE", typ="CORP", trad="20261001", eset="20261002", party="P1"))])}
    test = {"t": msg536("Y", "1/ONLY", "ACC1", [fin(FR2,
            tx("T1", "RECE", "9,", "99,", trad="20261011", eset="20261012", party="P9"))])}
    res = _run(ctrl, test, ignore=IGNORE_REF)
    assert sorted(r[0] for r in res.transactions) == ["EN TROP DANS TEST", "MANQUANT DANS TEST"]


def test_groups_never_cross_isin_or_account():
    a = tx("C1", "RECE", "5,")
    ctrl = {"c": msg536("X", "1/ONLY", "ACC1", [fin(FR2, a)])}
    test = {"t": msg536("Y", "1/ONLY", "ACC2", [fin(DE5, a)])}
    res = _run(ctrl, test, ignore=IGNORE_REF)
    assert sorted(r[0] for r in res.transactions) == ["EN TROP DANS TEST", "MANQUANT DANS TEST"]


def test_other_message_types_are_skipped():
    m535 = msg536("X", "1/ONLY", "ACC1", [fin(FR2, tx("C1", "RECE", "5,"))]).replace("I536", "I535")
    res = _run({"c": m535}, {})
    assert res.stats["messages_ctrl"] == 0


def test_missing_account_is_an_error():
    bad = msg536("X", "1/ONLY", "ACC1", [fin(FR2, tx("C1", "RECE", "5,"))]).replace(":97A::SAFE//ACC1\n", "")
    res = _run({"c": bad}, {})
    assert res.errors and "introuvable" in res.errors[0][3]


def test_demo_dataset():
    ctrl, test = demo_messages()
    res = _run(ctrl, test, ignore=IGNORE_REF)
    s = res.stats
    assert s["transactions_OK"] == 2
    assert s["transactions_ECART"] == 3
    assert s["transactions_MANQUANT DANS TEST"] == 1
    assert s["transactions_EN TROP DANS TEST"] == 1
    assert s["ecarts_entete"] == 1
    assert s["erreurs"] == 0


def test_match_transactions_directly():
    st_c = explode_536(msg536("X", "1/ONLY", "A", [fin(FR2, tx("C", "RECE", "5,"))]), "CTRL", "c")
    st_t = explode_536(msg536("Y", "1/ONLY", "A", [fin(FR2, tx("T", "RECE", "5,"))]), "TEST", "t")
    passes = [("FORTE", STRONG_KEY), ("FAIBLE", WEAK_KEY)]
    matched, oc, ot = match_transactions(st_c.txns, st_t.txns, passes, ignore_filter(IGNORE_REF))
    assert len(matched) == 1 and matched[0][2] == "FORTE" and not oc and not ot
