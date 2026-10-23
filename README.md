# Réconciliation SWIFT MT535 / MT536 — CTRL / TEST

Compare **tous les champs** des messages SWIFT de deux bases d'un progiciel bancaire (ancienne version = CTRL,
nouvelle version = TEST).

- `swift_recon.py` — MT535 (relevé de positions) : l'entête `GENL` et chaque champ de chaque
  position (ISIN ou `/XX/…`).
- `mt536.py` — MT536 (relevé de mouvements) : appariement des transactions entre CTRL et TEST
  puis comparaison champ à champ (entête, instrument, transaction).

## Installation

```
pip install -r requirements.txt
```

`python-oracledb` en mode « thin » : pas besoin d'Instant Client.

## Configuration

Copier `config.ini.example` en `config.ini`, puis renseigner pour `[ctrl]` et `[test]` :
DSN, utilisateur, table, colonnes id / CLOB / type, et la colonne de dates (`"TIMESTAMP"` par défaut).

Mots de passe : variables `SWIFT_CTRL_PASSWORD` / `SWIFT_TEST_PASSWORD`, sinon saisis au lancement.

## Lancement

```
# Plage commune aux deux bases (date de fin incluse)
python swift_recon.py config.ini --from 2026-10-20 --to 2026-10-20

# Plage différente par base (ex. ancienne version tournée le 20, nouvelle le 22)
python swift_recon.py config.ini --ctrl-from 2026-10-20 --ctrl-to 2026-10-20 \
                                 --test-from "2026-10-22 08:00" --test-to "2026-10-22 18:30"

# Rejouer la comparaison après avoir modifié la liste "ignore", sans relire Oracle
python swift_recon.py config.ini --compare-only
```

Formats de date : `AAAA-MM-JJ`, `AAAA-MM-JJ HH:MM`, `AAAA-MM-JJ HH:MM:SS`. La date de fin est incluse
(`--to 2026-10-20` prend toute la journée du 20).

## Résultat (Excel)

| Onglet    | Contenu |
|-----------|---------|
| Synthese  | Périodes, volumes, nombre d'écarts par statut |
| Comptes   | Une ligne par compte/date : statut, nb de messages et de positions de chaque côté |
| Positions | Une ligne par compte/ISIN : OK, ECART (avec les champs concernés), MANQUANT, EN TROP |
| Ecarts    | Une ligne par champ différent : bloc, tag, qualifiant, valeur CTRL, valeur TEST, msg_id |
| Erreurs   | Messages illisibles ou sans compte / date de relevé |

## Règles de comparaison

- Clé : compte (`97A::SAFE`) + date de relevé (`98a::STAT`) + ISIN + bloc + tag + qualifiant.
- La pagination (`28E`) est ignorée : un relevé sur 84 pages est comparé comme un tout,
  et un découpage différent entre CTRL et TEST ne crée pas de faux écarts.
- Valeurs normalisées : `1000,00` = `1000,`, espaces et retours ligne ignorés.
- Champs exclus par défaut : `28E, 20C:SEME, 98A:PREP, 98C:PREP` (clé `ignore` du `.ini`).
  Ajouter tout champ qui change légitimement à chaque génération.

## Démo sans Oracle

```
python swift_recon.py demo/config_demo.ini
```

Lit des messages exemples dans `demo/ctrl` et `demo/test` (écarts volontaires : prix modifié,
ISIN manquant, ISIN en trop, entête modifié, pagination et formats différents).

## MT536 (mouvements titres)

`mt536.py` fonctionne pour l'instant en mode autonome (dossiers de messages → CSV), sans
connexion Oracle :

```
python mt536.py demo536/ctrl demo536/test -o demo536 --ignore "28E,20C:SEME,98C:PREP,20C:PREV,20C:RELA"
```

Produit `mt536_transactions.csv`, `mt536_ecarts.csv` et `mt536_erreurs.csv` dans le dossier indiqué
par `-o`. Les transactions sont appariées par référence (facultatif), puis par clé forte, puis par
similarité, avant la comparaison champ à champ.

Tests : `pytest` (voir `pytest.ini` et `tests/`).
