# Baza Wyników Żużlowych WZDB

Repozytorium przechowuje automatycznie budowaną bazę `db/latest.wzdb` dla aplikacji **Wyniki Żużlowe v4**. Źródłem jest publiczny skoroszyt `PL2.xlsm` udostępniony w Google Drive.

## Jak działa automat

Workflow [`.github/workflows/update-db.yml`](.github/workflows/update-db.yml) uruchamia się co 15 minut oraz ręcznie przez `workflow_dispatch`. GitHub może opóźnić start zadania cyklicznego, dlatego aplikacja nadal korzysta z ostatniej poprawnie opublikowanej bazy do czasu zakończenia synchronizacji.

1. Pobiera `PL2.xlsm` bezpośrednio z `drive.usercontent.google.com`, używając unikalnego parametru żądania i nagłówków `no-cache`; awaryjnie korzysta z `gdown`.
2. Oblicza SHA-256 źródła oraz wszystkich wejść generatora (kod i zatwierdzona mapa dat) i porównuje je z hashami w `db/version.json`.
3. Jeśli hashe źródła, generatora i mapy dat są identyczne, kończy pracę bez przebudowy i bez commita.
4. Jeśli źródło, generator albo mapa dat się zmieniły, uruchamia `scripts/build_wzdb.py`, aktualizuje pliki w `db/`, commituje je i pushuje do `main`. Ręczne uruchomienie pozwala też zaznaczyć `force_rebuild`.

Workflow ma tylko wymagane uprawnienie `contents: write`. Źródłowy `PL2.xlsm` jest plikiem tymczasowym, znajduje się w `.gitignore` i nie jest zapisywany w repozytorium.

## Format WZDB v4

`latest.wzdb` to JSON UTF-8 skompresowany GZIP. Obiekt główny zawiera:

- `version`, `source`, `built`, `strings`, `players`, `years`, `stats`, `events`,
- `strings` — wspólną tablicę tekstów, do której rekordy odwołują się indeksami,
- `players` — rekordy `[nazwa, indeks_narodowości, data_urodzenia, nazwa_znormalizowana]`,
- `years` — rekordy wyników pogrupowane według sezonu,
- `events` — indeks wydarzeń dla każdego sezonu; starszy wpis ma postać `[indeks_pierwszego_rekordu, liczba_rekordów]`, a gdy znana jest rzeczywista data: `[indeks_pierwszego_rekordu, liczba_rekordów, fragmentCount, teams, indeks_daty_w_strings]`.

Opcjonalny piąty element wpisu `events` jest rozszerzeniem zgodnym z WZDB v4. Data ma format ISO `YYYY-MM-DD`, jest internowana raz w globalnej tablicy `strings` i nie jest duplikowana w rekordach zawodników. Brak piątego elementu oznacza nieznaną datę. Rekord wyniku zachowuje dotychczasowe `row[0]..row[13]`; opcjonalny numer startowy z kolumny A (`Nr`) jest zapisany na końcu jako `row[14]`. Kolumna N (`Sezon`) nie jest zapisywana jako numer startowy.

Zatwierdzone automatycznie daty sezonu 2026 znajdują się w `data/event_dates_2026.json`. Skrypt `scripts/match_event_dates.py` odtwarza tę mapę z dostarczonego kalendarza i zapisuje audyt dopasowań, niejednoznaczności oraz braków w katalogu `reports/`.

`db/version.json` zawiera hash źródła, jego skróconą wersję, hash wejść generatora, hash zatwierdzonej mapy dat, hash wynikowego WZDB, czas budowy, dostępny nagłówek `Last-Modified` z Google Drive i statystyki kontrolne. SHA-256 zawartości pozostaje głównym identyfikatorem wersji; data modyfikacji służy diagnostyce.

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
