# PrintCorrectSize dla MicroStation V8i

Makro VBA do MicroStation V8i, które zapisuje każdy model z pliku DGN do
**osobnego PDF-a z poprawnym rozmiarem strony**.

Rozwiązuje typowy problem przy drukowaniu zbiorczym: gdy w projekcie są
arkusze A4, A3 i o niestandardowych wymiarach, wybranie jednego rozmiaru
papieru powoduje, że mniejsze arkusze są pomniejszone, a większe obcięte.
To makro mierzy ramkę w każdym modelu i ustawia rozmiar strony w PDF.

Podstawowy scenariusz: jeden plik DGN, wiele **modeli rysunkowych**, a każdy
model ma własną ramkę na levelu, którego nazwa zawiera `RAMKA` (np. `M_RAMKA`
albo ` _RAMKA`).

> Status: kod źródłowy jest gotowy do testów na prawdziwych rysunkach.
> Najpierw przetestuj go na kopii pliku DGN. Repozytorium celowo zawiera
> źródło, a nie gotowy plik `.mvba`, żeby kod był czytelny i możliwy do
> sprawdzenia.

## Co robi

- Drukuje wszystkie modele do druku z aktywnego DGN, każdy do osobnego PDF-a.
- Mierzy ramkę na levelu `RAMKA` w każdym modelu.
- Rozpoznaje rozmiar strony w kolejności:
  1. definicja arkusza w sheet modelu (jeśli występuje),
  2. elementy na levelu ramki/obrysu (`RAMKA`, `M_RAMKA`, ` _RAMKA`),
  3. zaznaczenie elementów (fallback, ręczne wymuszenie rozmiaru),
  4. największy zamknięty element,
  5. extents wszystkich elementów graficznych.
- Jeśli zmierzona ramka odbiega od A4 albo A3 o maksymalnie 5 mm, strona jest
  zaokrąglana do dokładnego A4/A3. Inne rozmiary, w tym pozostałe A/B,
  zachowują rzeczywisty zmierzony wymiar.
- Ustawia stronę PDF na wykrytą szerokość × wysokość (także niestandardowe
  wymiary, np. 350 × 500 mm).
- Korzysta z wbudowanego sterownika Bentley PDF (`pdf.pltcfg`), więc nie
  trzeba instalować zewnętrznego PDF Creatora.
- Zapisuje `PrintCorrectSize_<dgn>.log` z rozmiarami i wynikami.
- Ma tryb podglądu, który tylko pokazuje wykryte rozmiary.

## Wymagania

- MicroStation V8i (SELECTseries) z VBA.
- Modele rysunkowe z ramką na levelu, którego nazwa zawiera `RAMKA`
  (`M_RAMKA`, ` _RAMKA`). Inne nazwy leveli można skonfigurować.
- Standardowy sterownik Bentley PDF `pdf.pltcfg` dostępny przez
  `MS_PLTCFG_PATH`.
- Sterownik musi pozwalać na edycję rozmiaru papieru. Sterowniki Bentley PDF
  zwykle to umożliwiają.

## Instalacja

1. Pobierz [`src/PrintCorrectSize.bas`](src/PrintCorrectSize.bas).
2. W MicroStation otwórz **Utilities > Macro > Project Manager**
   (w niektórych wersjach V8i: **Tools > VBA > Project Manager**).
3. Utwórz nowy projekt VBA o nazwie `PrintCorrectSize` i zapisz go jako
   `PrintCorrectSize.mvba` w folderze, który MicroStation przeszukuje, np.:
   `...\Workspace\Standards\Macros\vba`.
4. Otwórz edytor VBA i wybierz **File > Import File...**.
5. Wskaż `PrintCorrectSize.bas`.
6. Zapisz projekt VBA.
7. Jeśli projekt nie jest załadowany, załaduj go w Project Managerze.

Szczegółowa instrukcja: [`docs/INSTALL.md`](docs/INSTALL.md).

## Użycie

1. Otwórz plik DGN do wydruku.
2. Zapisz plik.
3. Najpierw uruchom podgląd:

   ```text
   vba run [PrintCorrectSize]PrintCorrectSize.Preview
   ```

   Sprawdź listę modeli i wykryte rozmiary stron.

4. Uruchom główne makro:

   ```text
   vba run [PrintCorrectSize]PrintCorrectSize.Main
   ```

   Oba makra są też dostępne w **Utilities > Macros**.

5. PDF-y trafiają do podfolderu `PDF` obok pliku DGN:

   ```text
   <nazwa-dgn>_<nazwa-modelu>.pdf
   ```

   Jeśli drukowany jest tylko jeden model, nazwa to `<nazwa-dgn>.pdf`.
   Spacje i znaki niedozwolone w nazwach plików Windows są zamieniane na `_`.

## Jak wykrywany jest rozmiar

Wygrywa pierwsza pasująca reguła:

| Kolejność | Źródło | Kiedy |
|---|---|---|
| 1 | `SHEET` | Model jest sheet modelem i ma poprawną definicję arkusza. |
| 2 | `BORDER LEVEL` | Na levelu, którego nazwa zawiera frazę z `BORDER_LEVEL_HINTS`, są elementy. Domyślne `RAMKA` pasuje do `M_RAMKA` i ` _RAMKA`. |
| 3 | `SELECTION` | Nie znaleziono ramki, ale coś jest zaznaczone w modelu. Zaznacz ramkę lub elementy wyznaczające stronę. |
| 4 | `CLOSED ELEMENT` | Używany jest największy zamknięty element (zwykle prostokąt). |
| 5 | `EXTENTS` | Fallback: extents wszystkich elementów graficznych. |

Dla modelu rysunkowego z ramką na `M_RAMKA` albo ` _RAMKA` mierzona jest
ramka i to jest strona. Jeśli ramka jest zamkniętym kształtem, makro bierze
największy zamknięty element na levelu ramki; jeśli ramka jest z linii,
bierze sumę elementów innych niż tekst z tego levelu. Tekst i wymiary na
levelu ramki są ignorowane, żeby notatki poza ramką nie powiększały strony.

Dla sheet modelu rozmiar z MicroStation jest nadrzędny. Zaznaczenie jest
tylko fallbackiem dla nietypowych modeli bez ramki: zaznacz obrys strony
przed uruchomieniem makra.

Domyślnie `PRINT_ONLY_SHEETS_IF_PRESENT` ma wartość `False`, więc drukowane
są wszystkie modele, gdy nie ma sheet modeli. Jeśli w DGN są sheet modele i
chcesz drukować tylko je, ustaw tę stałą na `True`.

## Konfiguracja

Wszystkie ustawienia to stałe na początku `PrintCorrectSize.bas`:

| Stała | Domyślnie | Znaczenie |
|---|---|---|
| `PDF_DRIVER` | `pdf.pltcfg` | Sterownik Bentley PDF. |
| `CHECK_PDF_DRIVER` | `True` | Sprawdź, czy sterownik istnieje, przed drukowaniem. |
| `OUTPUT_SUBFOLDER` | `PDF` | Podfolder na PDF-y; `""` zapisuje obok DGN. |
| `BORDER_LEVEL_HINTS` | `RAMKA` | Fragmenty nazw leveli, oddzielone przecinkami, bez rozróżniania wielkości liter. Pasuje do `M_RAMKA` i ` _RAMKA`. |
| `USE_SELECTION_AS_PAGE` | `True` | Użyj zaznaczenia, gdy nie znaleziono ramki. |
| `PRINT_ONLY_SHEETS_IF_PRESENT` | `False` | Jeśli są sheet modele, drukuj tylko je. |
| `STANDARD_SIZES_TO_SNAP` | `A4,A3` | Rozmiary zaokrąglane do dokładnych wymiarów ISO. |
| `STANDARD_SIZE_TOLERANCE_MM` | `5` | Maksymalna różnica od rozmiaru standardowego, na każdym boku. |
| `MIN_PAGE_MM` | `40` | Odrzucaj wykryte strony mniejsze niż to. |
| `MAX_PAGE_MM` | `10000` | Odrzucaj wykryte strony większe niż to. |
| `PDF_TIMEOUT_SEC` | `180` | Czas oczekiwania na jeden PDF. |
| `WRITE_LOG` | `True` | Zapisz log obok PDF-ów. |

## Ważne

- Makro zmienia aktywny model oraz tworzy/usuwa fence w MicroStation.
  Zapisz plik przed uruchomieniem.
- Makro nie zmienia geometrii rysunku. Zmienia tylko ustawienia druku i
  tymczasowy fence.
- Makro działa na aktywnym pliku DGN. Jeśli każdy obiekt to osobny plik DGN,
  uruchom makro w każdym pliku (bezpieczny tryb batch na folder może zostać
  dodany później).
- A4 i A3 są zaokrąglane do dokładnych wymiarów ISO, jeśli zmierzona ramka
  odbiega o maksymalnie 5 mm. Inne rozmiary, w tym A0–A2 i B, zachowują
  rzeczywisty zmierzony wymiar. Podgląd pokazuje stronę końcową oraz wymiar
  zmierzony.
- Sterownik PDF może dodać własne niewielkie marginesy, więc rysunek może być
  minimalnie mniejszy od strony.
- Niestandardowe rozmiary zależą od możliwości edycji papieru w sterowniku.
  Jeśli sterownik ignoruje `PRINT XFORM` / `PRINT YFORM`, użyj sterownika
  Bentley PDF, który pozwala edytować rozmiar papieru.

## Testowanie

Przed użyciem na produkcyjnych rysunkach przejdź przez
[`docs/TEST_CHECKLIST.md`](docs/TEST_CHECKLIST.md).

## Problemy

Zobacz [`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md).

Szybka lista:

- Brak PDF-a lub `FAIL` w logu: sprawdź, czy `pdf.pltcfg` jest zainstalowany,
  a `MS_PLTCFG_PATH` wskazuje jego folder.
- Zły rozmiar strony: uruchom `Preview`, sprawdź źródło w nawiasach i w razie
  potrzeby zaznacz obrys strony przed uruchomieniem makra.
- Dodatkowe PDF-y: w DGN są modele, których nie chcesz drukować. Domyślnie
  drukowane są wszystkie modele; usuń lub zmień nazwy niechcianych modeli
  albo ustaw `PRINT_ONLY_SHEETS_IF_PRESENT`, jeśli w pliku są sheet modele.
- Nie ma makra: projekt nie jest załadowany albo nazwa projektu nie jest
  `PrintCorrectSize`.

## Licencja

MIT. Zobacz [LICENSE](LICENSE).

MicroStation i Bentley są znakami towarowymi Bentley Systems, Incorporated.
Projekt nie jest powiązany z Bentley Systems ani przez nią autoryzowany.
