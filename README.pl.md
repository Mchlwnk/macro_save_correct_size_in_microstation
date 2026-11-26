# PrintCorrectSize dla MicroStation V8i

Makro VBA do MicroStation V8i, które zapisuje każdy model/arkusz z pliku DGN
do **osobnego PDF-a z poprawnym rozmiarem strony**.

Rozwiązuje typowy problem przy drukowaniu zbiorczym: gdy w projekcie są
arkusze A4, A3 i o niestandardowych wymiarach, wybranie jednego rozmiaru
papieru powoduje, że mniejsze arkusze są pomniejszone, a większe obcięte.
To makro sprawdza rozmiar każdego modelu osobno i ustawia go w PDF.

> Status: kod źródłowy jest gotowy do testów na prawdziwych rysunkach.
> Najpierw przetestuj go na kopii pliku DGN. Repozytorium celowo zawiera
> źródło, a nie gotowy plik `.mvba`, żeby kod był czytelny i możliwy do
> sprawdzenia.

## Co robi

- Drukuje wszystkie modele do druku z aktywnego DGN, każdy do osobnego PDF-a.
- Rozpoznaje rozmiar strony w kolejności:
  1. definicja arkusza w sheet modelu,
  2. zaznaczenie elementów (ręczne wymuszenie rozmiaru),
  3. elementy na levelu ramki/obrysu,
  4. największy zamknięty element,
  5. extents wszystkich elementów graficznych.
- Ustawia stronę PDF dokładnie na wykrytą szerokość × wysokość (także
  niestandardowe wymiary, np. 350 × 500 mm).
- Korzysta z wbudowanego sterownika Bentley PDF (`pdf.pltcfg`), więc nie
  trzeba instalować zewnętrznego PDF Creatora.
- Zapisuje `PrintCorrectSize.log` z rozmiarami i wynikami.
- Ma tryb podglądu, który tylko pokazuje wykryte rozmiary.

## Wymagania

- MicroStation V8i (SELECTseries) z VBA.
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

## Jak wykrywany jest rozmiar

Wygrywa pierwsza pasująca reguła:

| Kolejność | Źródło | Kiedy |
|---|---|---|
| 1 | `SHEET` | Model jest sheet modelem i ma poprawną definicję arkusza. |
| 2 | `SELECTION` | Coś jest zaznaczone w modelu. Zaznacz ramkę lub elementy wyznaczające stronę. |
| 3 | `BORDER LEVEL` | Na levelu, którego nazwa zawiera jedną z fraz `BORDER_LEVEL_HINTS`, są elementy. |
| 4 | `CLOSED ELEMENT` | Używany jest największy zamknięty element (zwykle prostokąt). |
| 5 | `EXTENTS` | Fallback: extents wszystkich elementów graficznych. |

Dla sheet modelu rozmiar z MicroStation jest nadrzędny. Zaznaczenie jest
przydatne dla modeli rysunkowych bez ramki na rozpoznanym levelu: zaznacz
obrys strony przed uruchomieniem makra.

Domyślnie, jeśli DGN zawiera sheet modele, drukowane są tylko one. Modele
rysunkowe w takim pliku to zwykle geometria źródłowa, a nie arkusze do
wydruku. Jeśli chcesz drukować wszystkie modele, zmień
`PRINT_ONLY_SHEETS_IF_PRESENT` w źródle.

## Konfiguracja

Wszystkie ustawienia to stałe na początku `PrintCorrectSize.bas`:

| Stała | Domyślnie | Znaczenie |
|---|---|---|
| `PDF_DRIVER` | `pdf.pltcfg` | Sterownik Bentley PDF. |
| `OUTPUT_SUBFOLDER` | `PDF` | Podfolder na PDF-y; `""` zapisuje obok DGN. |
| `BORDER_LEVEL_HINTS` | `PLOT,BORDER,SHEET,RAMKA,ARKUSZ,FORMAT,OBRAMOWANIE` | Fragmenty nazw leveli, oddzielone przecinkami, bez rozróżniania wielkości liter. |
| `USE_SELECTION_AS_PAGE` | `True` | Użyj zaznaczenia jako rozmiaru strony. |
| `PRINT_ONLY_SHEETS_IF_PRESENT` | `True` | Jeśli są sheet modele, drukuj tylko je. |
| `MIN_PAGE_MM` | `40` | Odrzucaj wykryte strony mniejsze niż to. |
| `MAX_PAGE_MM` | `10000` | Odrzucaj wykryte strony większe niż to. |
| `PDF_TIMEOUT_SEC` | `180` | Czas oczekiwania na jeden PDF. |
| `WRITE_LOG` | `True` | Zapisz `PrintCorrectSize.log`. |

## Ważne

- Makro zmienia aktywny model oraz tworzy/usuwa fence w MicroStation.
  Zapisz plik przed uruchomieniem.
- Makro nie zmienia geometrii rysunku. Zmienia tylko ustawienia druku i
  tymczasowy fence.
- Makro działa na aktywnym pliku DGN. Jeśli każdy obiekt to osobny plik DGN,
  uruchom makro w każdym pliku (bezpieczny tryb batch na folder może zostać
  dodany później).
- Rozmiar strony PDF jest dokładny. Sterownik PDF może jednak dodać własne
  niewielkie marginesy, więc rysunek może być minimalnie mniejszy od strony.
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
  preferowane są sheet modele; w innym wypadku usuń lub zmień nazwy
  niechcianych modeli.
- Nie ma makra: projekt nie jest załadowany albo nazwa projektu nie jest
  `PrintCorrectSize`.

## Licencja

MIT. Zobacz [LICENSE](LICENSE).

MicroStation i Bentley są znakami towarowymi Bentley Systems, Incorporated.
Projekt nie jest powiązany z Bentley Systems ani przez nią autoryzowany.
