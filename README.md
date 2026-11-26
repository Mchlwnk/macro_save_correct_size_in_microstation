# PrintCorrectSize dla MicroStation V8i

Makro VBA do MicroStation V8i, które zapisuje każdy model z aktywnego pliku
DGN do **osobnego PDF-a z poprawnym rozmiarem strony**.

Rozwiązuje typowy problem drukowania zbiorczego: gdy w projekcie są arkusze
A4, A3 i o niestandardowych wymiarach, wybranie jednego rozmiaru papieru
powoduje, że mniejsze arkusze są pomniejszone, a większe są obcinane. To makro
mierzy ramkę w każdym modelu i ustawia rozmiar strony w PDF.

Podstawowy scenariusz: jeden plik DGN, wiele **modeli rysunkowych**, a każdy
model ma własną ramkę na levelu, którego nazwa zawiera `RAMKA` — na przykład
`M_RAMKA` albo ` _RAMKA`.

> **Status:** to jest wersja źródłowa do testów na prawdziwych rysunkach.
> Najpierw sprawdź ją na kopii pliku DGN. Repozytorium celowo zawiera źródło
> `.bas`, a nie gotowy plik `.mvba`, żeby kod był czytelny i możliwy do
> sprawdzenia. Po udanym teście można zapisać projekt jako `.mvba` i dołączyć
> go do wydania (Releases), żeby inni nie musieli importować modułu ręcznie.

## Co robi makro

- Drukuje wszystkie modele wybrane do wydruku z aktywnego DGN, każdy do
  osobnego PDF-a.
- Mierzy ramkę na levelu `RAMKA` w modelach rysunkowych.
- Rozpoznaje rozmiar strony w tej kolejności:
  1. definicja arkusza w modelu arkusza (jeśli występuje),
  2. elementy na levelu ramki (`RAMKA`, `M_RAMKA`, ` _RAMKA`),
  3. zaznaczenie elementów (awaryjnie, gdy nie ma ramki),
  4. największy zamknięty element,
  5. extents wszystkich elementów graficznych.
- Jeśli zmierzona ramka odbiega od A4 albo A3 o **maksymalnie 5 mm na każdym
  boku**, strona jest zaokrąglana do dokładnego A4/A3.
- Pozostałe rozmiary, w tym A0–A2 i rozmiary B, zachowują **rzeczywisty
  zmierzony wymiar**.
- Ustawia stronę PDF na wykrytą szerokość × wysokość, także dla nietypowych
  wymiarów, np. 350 × 500 mm.
- Korzysta z wbudowanego sterownika Bentley PDF (`pdf.pltcfg`) i
  udokumentowanych key-inów `PRINT ...`, więc nie trzeba instalować
  zewnętrznego PDF Creatora.
- Zapisuje log `PrintCorrectSize_<dgn>.log` z rozmiarami i wynikami.
- Ma tryb podglądu, który tylko pokazuje wykryte rozmiary i nie tworzy PDF-ów.

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
   `...\Workspace\Standards\vba`.
4. Otwórz edytor VBA i wybierz **File > Import File...**.
5. Wskaż `PrintCorrectSize.bas`.
6. Zapisz projekt VBA.
7. Jeśli projekt nie jest załadowany, załaduj go w Project Managerze.

Jeśli zapiszesz projekt poza folderami z `MS_VBASEARCHDIRECTORIES`, przed
uruchomieniem załaduj go ręcznie:

```text
vba load "C:\...\PrintCorrectSize.mvba"
```

Pełna instrukcja, w tym alternatywne ścieżki i uruchamianie przez key-in:
[`docs/INSTALL.md`](docs/INSTALL.md).

## Uruchomienie

1. Otwórz plik DGN do wydruku.
2. Zapisz plik.
3. Najpierw uruchom podgląd:

   ```text
   vba run [PrintCorrectSize]PrintCorrectSize.Preview
   ```

   Sprawdź listę modeli, wykryte rozmiary i źródło w nawiasach. Przykład
   wpisu:

   ```text
   01. Mapa_A4  ->  210.0 x 297.0 mm  [LEVEL RAMKI]  ->  A4 (zmierzono 212.0 x 299.0 mm)
   02. Mapa_A3  ->  297.0 x 420.0 mm  [LEVEL RAMKI]  ->  A3 (zmierzono 299.0 x 421.0 mm)
   03. Mapa_X   ->  350.0 x 500.0 mm  [LEVEL RAMKI]
   ```

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
   Spacje i znaki niedozwolone w nazwach plików Windows są zamieniane na `_`,
   a polskie znaki diakrytyczne na litery bez ogonków (np. `Ł` → `L`).

## Jak wykrywany jest rozmiar strony

Wygrywa pierwsza pasująca reguła:

| Kolejność | Źródło w podglądzie | Kiedy jest używane |
|---|---|---|
| 1 | `ARKUSZ` | Model jest modelem arkusza (sheet model) i ma poprawną definicję arkusza. |
| 2 | `LEVEL RAMKI` | Na levelu, którego nazwa zawiera frazę z `BORDER_LEVEL_HINTS`, są elementy. Domyślne `RAMKA` pasuje do `M_RAMKA` i ` _RAMKA`. |
| 3 | `ZAZNACZENIE` | Nie znaleziono ramki, ale coś jest zaznaczone w modelu. Zaznacz ramkę lub elementy wyznaczające stronę. |
| 4 | `ZAMKNIETY ELEMENT` | Używany jest największy zamknięty element (zwykle prostokąt). |
| 5 | `GRANICE` | Rozwiązanie awaryjne: extents wszystkich elementów graficznych. |

Dla modelu rysunkowego z ramką na `M_RAMKA` albo ` _RAMKA` mierzona jest
ramka i to jest strona. Makro bierze większego z dwóch kandydatów:
największy zamknięty element na levelu ramki albo sumę elementów innych niż
tekst z tego levelu. Dzięki temu zamknięta ramka wygrywa, a gdy ramka jest
z linii — suma linii wygrywa z mniejszym blokiem tytułowym. Tekst i wymiary
na levelu ramki są ignorowane, żeby notatki poza ramką nie powiększały strony.

Dla modelu arkusza (sheet model) rozmiar z MicroStation jest nadrzędny.
Zaznaczenie jest tylko rozwiązaniem awaryjnym dla nietypowych modeli bez
ramki: zaznacz obrys strony przed uruchomieniem makra.

Domyślnie `PRINT_ONLY_SHEETS_IF_PRESENT` ma wartość `False`, więc drukowane
są wszystkie modele, gdy nie ma modeli arkuszy. Jeśli w DGN są modele arkuszy i
chcesz drukować tylko je, ustaw tę stałą na `True`.

## Dopasowanie A4/A3

Tolerancja dotyczy **każdego boku osobno** i domyślnie wynosi 5 mm. Przykłady:

| Zmierzona ramka | Wynik |
|---|---|
| 212 × 299 mm | A4, strona 210 × 297 mm |
| 300 × 423 mm | A3, strona 297 × 420 mm |
| 299 × 212 mm | A4 poziomo, strona 297 × 210 mm |
| 216 × 300 mm | wymiar rzeczywisty 216 × 300 mm (za daleko od A4) |
| 420 × 594 mm | wymiar rzeczywisty 420 × 594 mm (A2 nie jest domyślnie zaokrąglane) |
| 707 × 1000 mm | wymiar rzeczywisty 707 × 1000 mm (B1 nie jest domyślnie zaokrąglane) |

Jeśli chcesz zaokrąglać także inne rozmiary A/B, dopisz je do
`STANDARD_SIZES_TO_SNAP`, np. `A4,A3,A2,B1`. Każda nazwa musi istnieć w
funkcji `PaperDimensions` w źródle.

## Konfiguracja

Wszystkie ustawienia to stałe na początku `PrintCorrectSize.bas`:

| Stała | Domyślnie | Znaczenie |
|---|---|---|
| `PDF_DRIVER` | `pdf.pltcfg` | Sterownik Bentley PDF. |
| `CHECK_PDF_DRIVER` | `True` | Sprawdź, czy sterownik istnieje, przed drukowaniem. |
| `OUTPUT_SUBFOLDER` | `PDF` | Podfolder na PDF-y; `""` zapisuje obok DGN. |
| `BORDER_LEVEL_HINTS` | `RAMKA` | Fragmenty nazw leveli, oddzielone przecinkami, bez rozróżniania wielkości liter. Pasuje do `M_RAMKA` i ` _RAMKA`. |
| `USE_SELECTION_AS_PAGE` | `True` | Użyj zaznaczenia, gdy nie znaleziono ramki. |
| `PRINT_ONLY_SHEETS_IF_PRESENT` | `False` | Jeśli są modele arkuszy, drukuj tylko je. |
| `STANDARD_SIZES_TO_SNAP` | `A4,A3` | Rozmiary zaokrąglane do dokładnych wymiarów ISO. |
| `STANDARD_SIZE_TOLERANCE_MM` | `5` | Maksymalna różnica od rozmiaru standardowego, na każdym boku. |
| `MIN_PAGE_MM` | `40` | Odrzucaj wykryte strony mniejsze niż to. |
| `MAX_PAGE_MM` | `10000` | Odrzucaj wykryte strony większe niż to. |
| `PDF_TIMEOUT_SEC` | `180` | Czas oczekiwania na jeden PDF. |
| `WRITE_LOG` | `True` | Zapisz log obok PDF-ów. |

## Ważne uwagi

- Makro zmienia aktywny model i tworzy/usuwa fence w MicroStation. Zapisz
  plik przed uruchomieniem.
- Makro nie modyfikuje istniejących elementów rysunku. Tworzy tylko tymczasowy fence i
  zmienia ustawienia druku.
- Makro działa na aktywnym pliku DGN. Jeśli każdy obiekt to osobny plik DGN,
  uruchom makro w każdym pliku.
- Domyślnie drukowane są **wszystkie modele**. Jeśli któryś model nie jest
  obiektem do wydruku, będzie widoczny w podglądzie — usuń go, zmień nazwę
  albo ustaw `PRINT_ONLY_SHEETS_IF_PRESENT`.
- A4 i A3 są zaokrąglane do dokładnych wymiarów ISO, jeśli zmierzona ramka
  odbiega o maksymalnie 5 mm na każdym boku. Inne rozmiary, w tym A0–A2 i B,
  zachowują rzeczywisty zmierzony wymiar.
- Sterownik PDF może dodać własne niewielkie marginesy, więc rysunek może być
  minimalnie mniejszy od strony.
- Niestandardowe rozmiary zależą od możliwości edycji papieru w sterowniku.
  Jeśli sterownik ignoruje `PRINT XFORM` / `PRINT YFORM`, użyj sterownika
  Bentley PDF, który pozwala edytować rozmiar papieru.
- Ramka jest mierzona jako prostokąt wyrównany do osi. Obrócona ramka da
  prostokąt otaczający (bounding box), czyli większy rozmiar.
- Ramka w referencji nie jest widziana przez skan aktywnego modelu. Umieść
  ramkę w modelu głównym albo zaznacz elementy strony ręcznie.
- Model zawierający wyłącznie referencje i brak geometrii lokalnej zostanie
  pominięty jako „brak geometrii”.

## Testowanie

Przed użyciem na produkcyjnych rysunkach przejdź przez
[`docs/TEST_CHECKLIST.md`](docs/TEST_CHECKLIST.md).

## Problemy

Zobacz [`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md).

Szybka lista:

- Brak PDF-a lub `BLAD` w logu: sprawdź, czy `pdf.pltcfg` jest zainstalowany,
  a `MS_PLTCFG_PATH` wskazuje jego folder.
- Zły rozmiar strony: uruchom `Preview`, sprawdź źródło w nawiasach i w razie
  potrzeby zaznacz obrys strony przed uruchomieniem makra.
- Dodatkowe PDF-y: w DGN są modele, których nie chcesz drukować. Usuń je,
  zmień nazwy albo ustaw `PRINT_ONLY_SHEETS_IF_PRESENT`.
- Nie ma makra: projekt nie jest załadowany albo nazwa projektu nie jest
  `PrintCorrectSize`.
- PDF jest otwarty w Acrobat/Reader: zamknij go przed ponownym uruchomieniem,
  bo starego pliku nie da się nadpisać.

## Uwaga o polskich znakach w kodzie

Plik `PrintCorrectSize.bas` jest zapisany w ASCII, bez polskich znaków
diakrytycznych. To celowe: import źródeł VBA w MicroStation V8i może
interpretować plik w kodowaniu systemowym i uszkodzić znaki diakrytyczne.
Dlatego w podglądzie i logu zobaczysz np. `ZAMKNIETY ELEMENT`, `POMINIETY`
albo `BLAD`, a nie `ZAMKNIĘTY`, `POMINIĘTY`, `BŁĄD`. Dokumentacja używa
pełnej polszczyzny.

## Struktura repozytorium

```text
src/PrintCorrectSize.bas      źródło makra (import do projektu VBA)
docs/INSTALL.md               pełna instrukcja instalacji
docs/TEST_CHECKLIST.md        lista kontrolna testów na rysunkach
docs/TROUBLESHOOTING.md       typowe problemy i rozwiązania
README.md                     ten plik
LICENSE                       licencja MIT
```

## Licencja

MIT. Zobacz [LICENSE](LICENSE).

MicroStation i Bentley są znakami towarowymi Bentley Systems, Incorporated.
Projekt nie jest powiązany z Bentley Systems ani przez nią autoryzowany.
