# Baza Wyników Żużlowych WZDB

Repozytorium przechowuje automatycznie budowaną bazę `db/latest.wzdb` dla aplikacji **Wyniki Żużlowe v4**. Źródłem jest publiczny skoroszyt `PL2.xlsm` udostępniony w Google Drive.

## Jak działa automat

Workflow [`.github/workflows/update-db.yml`](.github/workflows/update-db.yml) uruchamia się raz na godzinę (w 17. minucie) oraz ręcznie przez `workflow_dispatch`.

1. Pobiera `PL2.xlsm` przez `gdown`; awaryjnie korzysta bezpośrednio z `drive.usercontent.google.com`.
2. Oblicza SHA-256 źródła i generatora oraz porównuje je z hashami w `db/version.json`.
3. Jeśli oba hashe są identyczne, kończy pracę bez przebudowy i bez commita.
4. Jeśli źródło lub generator się zmieniły, uruchamia `scripts/build_wzdb.py`, aktualizuje pliki w `db/`, commituje je i pushuje do `main`. Ręczne uruchomienie pozwala też zaznaczyć `force_rebuild`.

Workflow ma tylko wymagane uprawnienie `contents: write`. Źródłowy `PL2.xlsm` jest plikiem tymczasowym, znajduje się w `.gitignore` i nie jest zapisywany w repozytorium.

## Format WZDB v4

`latest.wzdb` to JSON UTF-8 skompresowany GZIP. Obiekt główny zawiera:

- `version`, `source`, `built`, `strings`, `players`, `years`, `stats`, `events`,
- `strings` — wspólną tablicę tekstów, do której rekordy odwołują się indeksami,
- `players` — rekordy `[nazwa, indeks_narodowości, data_urodzenia, nazwa_znormalizowana]`,
- `years` — rekordy wyników pogrupowane według sezonu,
- `events` — indeks wydarzeń dla każdego sezonu; wpis ma postać `[indeks_pierwszego_rekordu, liczba_rekordów]`.

`db/version.json` zawiera hash źródła, jego skróconą wersję, hash wynikowego WZDB, czas budowy i statystyki kontrolne.

## Budowa lokalna

Wymagany jest Python 3.12 (lub nowszy zgodny) i zależności z `requirements.txt`:

```bash
python -m pip install -r requirements.txt
python scripts/build_wzdb.py /ścieżka/do/PL2.xlsm
```

Przy weryfikacji konkretnego wydania można włączyć twarde oczekiwania:

```bash
python scripts/build_wzdb.py /ścieżka/do/PL2.xlsm \
  --expect-rows 312955 \
  --expect-players 5442 \
  --expect-seasons 17 \
  --expect-events 34519 \
  --expect-from 2010 \
  --expect-to 2026
```

Konwerter zapisuje pliki atomowo i nie modyfikuje źródłowego skoroszytu.
