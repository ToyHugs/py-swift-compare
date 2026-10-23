"""Construction de MT536 de démonstration / test (pas de 93B, comme chez nous)."""


def tx(ref, sens, qty, amt=None, paym="APMT", typ="SETT", trad=None, eset="20261015",
       party=None, party_qual=None):
    """Un bloc TRAN. party_qual par défaut : DEAG en réception, REAG en livraison."""
    lines = [":16R:TRAN", ":16R:LINK", f":20C::RELA//{ref}", ":16S:LINK",
             ":16R:TRANSDET", f":36B::PSTA//UNIT/{qty}"]
    if amt:
        lines.append(f":19A::PSTA//EUR{amt}")
    lines += [f":22F::TRAN//{typ}", f":22H::REDE//{sens}", f":22H::PAYM//{paym}"]
    if trad:
        lines.append(f":98A::TRAD//{trad}")
    lines.append(f":98A::ESET//{eset}")
    if party:
        q = party_qual or ("DEAG" if sens == "RECE" else "REAG")
        lines += [":16R:SETPRTY", f":95P::{q}//{party}", ":16S:SETPRTY",
                  ":16R:SETPRTY", ":95P::PSET//NCSDXX21", ":16S:SETPRTY"]
    lines += [":16S:TRANSDET", ":16S:TRAN"]
    return "\n".join(lines)


def fin(isin_line, *txns):
    return "\n".join([":16R:FIN", f":35B:{isin_line}", *txns, ":16S:FIN"])


def msg536(seme, page, acc, fins, code="COMP", period="20261001/20261020"):
    body = "\n".join(fins)
    return f"""{{1:F01BANKLULLAXXX0000000000}}{{2:I536CLIENTXXXXXXXN}}{{4:
:16R:GENL
:28E:{page}
:13A::STAT//001
:20C::SEME//{seme}
:23G:NEWM
:98C::PREP//20261021063000
:69A::STAT//{period}
:22F::SFRE//DAIL
:22F::CODE//{code}
:97A::SAFE//{acc}
:17B::ACTI//Y
:17B::CONS//N
:16S:GENL
:16R:SUBSAFE
{body}
:16S:SUBSAFE
-}}"""


FR2 = "ISIN FR0000000002\nACTION BETA SA"
LU1 = "ISIN LU0000000001\nFONDS ALPHA"
XS3 = "ISIN XS0000000003\nOBLIG GAMMA 3% 2030"
DE5 = "ISIN DE0000000005\nACTION EPSILON AG"


def demo_messages():
    """Jeu de démo : CTRL sur 2 pages, TEST sur 1 page, ordre différent,
    références régénérées (20C:RELA différentes), écarts volontaires :
      - FR2 livraison C-2 : quantité 20 -> 25 (et montant)       -> ECART via passe FAIBLE
      - FR2 livraison C-5 : date de règlement 18 -> 19             -> ECART via passe FAIBLE
      - LU1 C-4          : contrepartie DEAG BROKERA -> BROKERB    -> ECART via passe FORTE
      - XS3 C-6          : absente de TEST                         -> MANQUANT
      - DE5              : nouvelle dans TEST                      -> EN TROP
      - entête           : 22F::CODE COMP -> DELT                  -> ECART ENTETE
    """
    t1 = lambda ref: tx(ref, "RECE", "50,", "2100,", trad="20261013", eset="20261015", party="BROKERA")
    ctrl = {
        "c536_001.txt": msg536("C536P1", "1/FIRS", "ACC1001", [
            fin(FR2, t1("C-1"),
                tx("C-2", "DELI", "20,", "840,", trad="20261014", eset="20261016", party="BROKERA"),
                t1("C-3")),
            fin(LU1, tx("C-4", "RECE", "100,", paym="FREE", eset="20261010", party="BROKERA")),
        ]),
        "c536_002.txt": msg536("C536P2", "2/LAST", "ACC1001", [
            fin(FR2, tx("C-5", "DELI", "10,", "420,", trad="20261016", eset="20261018", party="BROKERB")),
            fin(XS3, tx("C-6", "RECE", "1000,", "987500,", trad="20261015", eset="20261017")),
        ]),
    }
    test = {
        "t536_001.txt": msg536("T536P1", "1/ONLY", "ACC1001", [
            fin(FR2,
                tx("T-9", "DELI", "10,00", "420,00", trad="20261016", eset="20261019", party="BROKERB"),
                t1("T-7"),
                tx("T-8", "DELI", "25,", "1050,", trad="20261014", eset="20261016", party="BROKERA"),
                t1("T-6")),
            fin(LU1, tx("T-5", "RECE", "100,", paym="FREE", eset="20261010", party="BROKERB")),
            fin(DE5, tx("T-4", "RECE", "5,", "50,", trad="20261019", eset="20261021")),
        ], code="DELT"),
    }
    return ctrl, test
