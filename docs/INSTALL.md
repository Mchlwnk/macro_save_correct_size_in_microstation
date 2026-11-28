# Instalacja makra PrintCorrectSize

Ta instrukcja dotyczy MicroStation V8i (SELECTseries). Makro jest plikiem
źródłowym VBA `.bas`, który trzeba zaimportować do projektu VBA w MicroStation.
Projekt jest zapisywany jako plik `.mvba`.

Do drukowania potrzebny jest sterownik Bentley PDF `pdf.pltcfg` z włączoną
możliwością edycji rozmiaru papieru. Makro sprawdza obecność sterownika przed
rozpoczęciem.

## 1. Przygotuj plik źródłowy

Pobierz z repozytorium:

```text
src/PrintCorrectSize.bas
```

Zapisz go w łatwym do znalezienia miejscu, np. na pulpicie. Nie zmieniaj
kodowania pliku. Ustawienia stałych zmieniaj w edytorze VBA, nie w Notatniku —
plik jest przygotowany jako źródło VBA.

## 2. Wybierz folder na projekt VBA

MicroStation szuka projektów VBA w folderach zdefiniowanych przez zmienną
konfiguracyjną `MS_VBASEARCHDIRECTORIES`. Typowe miejsce:

```text
...\Workspace\Standards\vba
```

W domyślnej instalacji V8i może to być np.:

```text
C:\ProgramData\Bentley\MicroStation V8i (SELECTseries)\WorkSpace\Standards\vba
```

Dokładna ścieżka zależy od instalacji i konfiguracji workspace. Nie zgaduj
jej — sprawdź wartość `MS_VBASEARCHDIRECTORIES` w konfiguracji.

Jeśli folder nie istnieje, możesz go utworzyć. Możesz też zapisać projekt w
innym miejscu i ładować go ręcznie (patrz punkt 6).

Nowe projekty są domyślnie tworzone w folderze `MS_VBANEWPROJECTDIRECTORY`.
Możesz sprawdzić obie zmienne w:

**Workspace > Configuration > Configuration Variables**

lub w oknie konfiguracji workspace, zależnie od wersji V8i.

## 3. Utwórz projekt VBA

### Wariant A — przez Project Manager

1. W MicroStation otwórz **Utilities > Macro > Project Manager**.
   W niektórych wersjach V8i menu może być **Tools > VBA > Project Manager**.
2. Kliknij ikonę **New Project**.
3. W oknie zapisu ustaw nazwę pliku:

   ```text
   PrintCorrectSize.mvba
   ```

   i wskaż folder z punktu 2.
4. Projekt pojawi się na liście w Project Managerze jako `PrintCorrectSize`.
   Jeśli okno pozwala też wpisać nazwę projektu, wpisz `PrintCorrectSize`.

### Wariant B — przez key-in

Otwórz **Utilities > Key-in** i wpisz:

```text
vba create PrintCorrectSize.mvba
```

Jeśli nie podasz ścieżki, projekt zostanie utworzony w folderze
`MS_VBANEWPROJECTDIRECTORY` (a jeśli ta zmienna nie jest ustawiona — w katalogu
programu MicroStation). Projekt zostanie od razu utworzony i załadowany.

## 4. Zaimportuj moduł

1. Zaznacz projekt `PrintCorrectSize` w Project Managerze.
2. Kliknij **Edit** / **Show VBA Editor**, żeby otworzyć edytor Visual Basic.
   Możesz też użyć key-inu:

   ```text
   vba edit PrintCorrectSize
   ```

3. W edytorze VBA wybierz **File > Import File...**.
4. Wskaż plik `PrintCorrectSize.bas`.
5. Moduł `PrintCorrectSize` pojawi się w oknie Project Explorer.
6. Jeśli moduł o tej nazwie już istnieje, najpierw usuń stary moduł, a potem
   zaimportuj nowy plik.
7. Zapisz projekt: **File > Save** (Ctrl+S).
8. Zamknij edytor VBA. Jeśli MicroStation zapyta o zapis projektu, wybierz
   **Yes**.

Nazwa modułu jest ustawiona w pliku przez `Attribute VB_Name =
"PrintCorrectSize"`. Dzięki temu key-in ma postać:

```text
[PrintCorrectSize]PrintCorrectSize.Main
```

Pierwszy człon to nazwa **projektu**, drugi to nazwa **modułu**.

## 5. Załaduj projekt

Jeśli projekt nie jest załadowany:

1. Otwórz **Utilities > Macro > Project Manager**.
2. Zaznacz `PrintCorrectSize`.
3. Kliknij **Load Project**.

Możesz też użyć key-inu:

```text
vba load PrintCorrectSize
```

Jeśli projekt jest poza `MS_VBASEARCHDIRECTORIES`, podaj pełną ścieżkę:

```text
vba load "C:\...\PrintCorrectSize.mvba"
```

Cudzysłowy są potrzebne, gdy ścieżka zawiera spacje.

## 6. Uruchom makro

### SelfTest — sprawdzenie logiki

Najpierw uruchom test wbudowany. Nie wymaga otwartego DGN i niczego nie
drukuje:

```text
vba run [PrintCorrectSize]PrintCorrectSize.SelfTest
```

Powinien pojawić się komunikat `SelfTest OK`. Jeśli pojawi się `SelfTest:
BLAD`, makro zostało źle zaimportowane albo plik `.bas` został zmieniony.

### Podgląd — najpierw to uruchom

Otwórz **Utilities > Key-in** i wpisz:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Preview
```

Makro pokaże listę modeli, wykryte rozmiary stron i źródło rozmiaru:

- `ARKUSZ` — rozmiar z definicji modelu arkusza,
- `LEVEL RAMKI` — ramka na levelu `RAMKA`,
- `ZAZNACZENIE` — użyto zaznaczenia,
- `ZAMKNIETY ELEMENT` — największy zamknięty element,
- `GRANICE` — extents wszystkich elementów.

Jeśli coś jest nie tak, popraw level ramki albo zaznacz właściwe elementy
przed uruchomieniem.

### Właściwy druk

Gdy podgląd wygląda dobrze:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Main
```

Możesz też uruchomić oba makra z okna **Utilities > Macros**.

## 7. Pierwszy test na kopii DGN

1. Otwórz **kopię** prawdziwego pliku DGN, nie oryginał.
2. Zapisz plik.
3. Uruchom `Preview`.
4. Sprawdź:
   - czy lista modeli się zgadza,
   - czy każdy model ma właściwy rozmiar,
   - czy źródło w nawiasach ma sens.
5. Uruchom `Main`.
6. Otwórz PDF-y z podfolderu `PDF` i sprawdź:
   - rozmiar strony w **File > Properties**,
   - czy nic nie jest obcięte,
   - czy rysunek wypełnia stronę,
   - czy linie, tekst i referencje wyglądają poprawnie.
7. Sprawdź log `PrintCorrectSize_<dgn>.log` w folderze `PDF`.

Pełna lista kontrolna: [`TEST_CHECKLIST.md`](TEST_CHECKLIST.md).

## 8. Typowe problemy z instalacją

- **Nie ma makra na liście** — projekt nie jest załadowany albo moduł nie
  został zaimportowany.
- **`vba run [PrintCorrectSize]...` nie znajduje projektu** — projekt jest
  poza `MS_VBASEARCHDIRECTORIES`. Załaduj go ręcznie przez Project Manager
  albo przez `vba load "pełna ścieżka\PrintCorrectSize.mvba"`.
- **Nazwa projektu jest inna** — użyj własnej nazwy w key-inie, np.
  `vba run [MojaNazwa]PrintCorrectSize.Main`.
- **Import się nie udaje** — usuń stary moduł `PrintCorrectSize` z projektu i
  zaimportuj plik ponownie.
- **Brak sterownika PDF** — makro sprawdza `pdf.pltcfg` przez
  `MS_PLTCFG_PATH`. Szczegóły: [`TROUBLESHOOTING.md`](TROUBLESHOOTING.md).

## 9. Aktualizacja makra

Gdy pojawi się nowsza wersja `PrintCorrectSize.bas`:

1. Otwórz edytor VBA dla projektu `PrintCorrectSize`.
2. Usuń stary moduł `PrintCorrectSize`.
3. Zaimportuj nowy plik `.bas`.
4. Zapisz projekt `.mvba`.

Repozytorium zawiera tylko aktualną wersję makra.

## 10. Odinstalowanie

1. W Project Managerze odładuj projekt `PrintCorrectSize`.
2. Usuń plik `PrintCorrectSize.mvba` z folderu makr.

Makro nie zapisuje niczego poza folderem `PDF` obok pliku DGN; w tym folderze
może powstać log. Nie modyfikuje istniejących elementów rysunku; tworzy tylko tymczasowy
fence i zmienia ustawienia druku.
