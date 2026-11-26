# Rozwiązywanie problemów

## Nie ma makra na liście

- Sprawdź, czy projekt `PrintCorrectSize` jest załadowany w
  **Utilities > Macro > Project Manager**.
- Sprawdź, czy moduł `PrintCorrectSize` został zaimportowany do tego projektu.
- Jeśli zmieniłeś nazwę projektu, użyj jej w key-inie:
  `vba run [TwojaNazwa]PrintCorrectSize.Main`.
- Sprawdź, czy makra VBA nie są zablokowane w MicroStation.

## `vba run ...` mówi, że nie znaleziono makra

Składnia jest taka:

```text
vba run [NazwaProjektu]NazwaModulu.NazwaProcedury
```

Dla tego projektu:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Main
```

Pierwszy `PrintCorrectSize` to **projekt**, drugi to **moduł**. Jeśli projekt
jest poza `MS_VBASEARCHDIRECTORIES`, najpierw go załaduj:

```text
vba load "C:\...\PrintCorrectSize.mvba"
vba run [PrintCorrectSize]PrintCorrectSize.Main
```

## Nie powstaje PDF albo w logu jest `BLAD`

1. Sprawdź, czy sterownik Bentley PDF jest zainstalowany: `pdf.pltcfg`.
2. Sprawdź zmienną `MS_PLTCFG_PATH`. Powinna wskazywać folder ze sterownikiem.
   Domyślnie makro sprawdza to przed drukowaniem i nie zacznie, jeśli
   sterownika nie znajdzie.
3. Otwórz okno druku ręcznie, wybierz sterownik Bentley PDF i spróbuj
   wydrukować jeden model. Jeśli to też nie działa, problem jest w sterowniku,
   nie w makrze.
4. Sprawdź, czy folder `PDF` obok DGN jest zapisywalny.
5. Sprawdź log `PrintCorrectSize_<dgn>.log` w folderze `PDF`.
6. Zamknij PDF w Acrobat/Reader przed ponownym uruchomieniem. Zablokowanego
   pliku nie da się nadpisać; w logu pojawi się
   `nie można usunąć starego PDF-a`.

Timeout wynosi 180 sekund na jeden PDF. Przy bardzo dużych lub rastrowych
modelach zwiększ `PDF_TIMEOUT_SEC` w źródle.

Jeśli sterownik jest znajdowany przez nietypową ścieżkę, możesz ustawić
`CHECK_PDF_DRIVER = False` w źródle, ale wtedy makro nie sprawdzi sterownika
przed drukowaniem.

## Strona PDF ma standardowy rozmiar zamiast niestandardowego

Makro ustawia papier key-inami:

```text
PRINT XFORM <szerokość> MM
PRINT YFORM <wysokość> MM
```

To wymaga sterownika Bentley, który pozwala edytować rozmiar papieru. W oknie
druku:

1. W oknie druku otwórz edytor konfiguracji sterownika. W zależności od
   wersji V8i będzie to np. **File > Edit Printer Driver Configuration**,
   **File > Edit Printer Driver** albo osobna ikona konfiguracji sterownika.
2. Znajdź opcję w rodzaju **Allow paper size editing** i włącz ją.
3. Zapisz konfigurację sterownika.
4. Uruchom makro ponownie.

Jeśli organizacja używa własnego `.pltcfg` z zablokowanym rozmiarem papieru,
skopiuj standardowy `pdf.pltcfg`, włącz edycję papieru i ustaw `PDF_DRIVER` w
makrze na nazwę swojej kopii.

Bardzo duża strona niestandardowa może też przekroczyć maksymalny rozmiar
formularza sterownika. Wtedy sterownik może ograniczyć stronę albo odmówić
wydruku.

## Zły rozmiar strony

Najpierw uruchom `Preview`. Źródło jest pokazane w nawiasach:

- `[ARKUSZ]` — użyto definicji modelu arkusza. Sprawdź
  **Model Properties > Sheet > Size**.
- `[LEVEL RAMKI]` — znaleziono elementy na levelu pasującym do
  `BORDER_LEVEL_HINTS` (domyślnie `RAMKA`, więc `M_RAMKA` i ` _RAMKA`).
  Makro bierze większego kandydata: największy zamknięty element albo sumę
  elementów innych niż tekst z tego levelu. Sprawdź, czy na levelu jest tylko
  ramka, a nie dodatkowa geometria poza stroną.
- `[ZAZNACZENIE]` — nie znaleziono ramki, ale coś było zaznaczone. Wyczyść
  zaznaczenie albo zaznacz właściwe elementy strony.
- `[ZAMKNIETY ELEMENT]` — użyto największego zamkniętego elementu. Jeśli to
  nie jest ramka strony, umieść ramkę na rozpoznanym levelu albo zaznacz ją
  ręcznie.
- `[GRANICE]` — nie znaleziono ramki ani definicji arkusza, więc użyto
  extents całej geometrii. Elementy poza stroną powiększą stronę.

Ramka jest mierzona jako prostokąt wyrównany do osi. Obrócona ramka jest
mierzona przez prostokąt otaczający (bounding box), czyli większy prostokąt.
Dla dokładnego rozmiaru używaj ramki wyrównanej do osi.

### Ręczne wymuszenie rozmiaru

Dla modelu rysunkowego zaznacz ramkę (albo wszystkie elementy wyznaczające
stronę) przed uruchomieniem makra. Zaznaczenie jest używane, gdy nie ma ramki
na levelu.

Jeśli ramka jest w referencji, a nie w aktywnym modelu, skan jej nie widzi.
Umieść ramkę w modelu głównym albo zaznacz elementy strony ręcznie. Model
zawierający wyłącznie referencje i brak geometrii lokalnej zostanie pominięty
jako „brak geometrii”.

## Dopasowanie A4/A3

Zmierzona ramka jest zaokrąglana do dokładnego A4 albo A3 tylko wtedy, gdy
**oba boki** mieszczą się w tolerancji 5 mm. Przykłady:

- 212 × 299 mm → A4 (210 × 297 mm)
- 300 × 423 mm → A3 (297 × 420 mm)
- 216 × 300 mm → wymiar rzeczywisty 216 × 300 mm (więcej niż 5 mm od A4)
- 420 × 594 mm → wymiar rzeczywisty 420 × 594 mm (A2 nie jest domyślnie
  zaokrąglane)

Podgląd pokazuje zarówno stronę końcową, jak i wymiar zmierzony. Aby
zaokrąglać inne rozmiary, dopisz je do `STANDARD_SIZES_TO_SNAP`, np.
`A4,A3,A2,B1`. Każda nazwa musi istnieć w funkcji `PaperDimensions` w źródle.

## Powstają dodatkowe PDF-y

Makro drukuje wszystkie modele wybrane do druku. Domyślnie
`PRINT_ONLY_SHEETS_IF_PRESENT = False`, więc drukowane są wszystkie modele.

Jeśli dostajesz PDF-y dla modeli, których nie chcesz drukować:

- usuń lub zmień nazwy niechcianych modeli,
- ustaw `PRINT_ONLY_SHEETS_IF_PRESENT = True`, jeśli w pliku są modele arkuszy i
  chcesz drukować tylko je,
- albo dodaj filtr modeli w źródle.

## Każdy obiekt to osobny plik DGN

Obecna wersja drukuje **aktywny plik DGN**. Jeśli każdy obiekt jest osobnym
plikiem `.dgn`:

1. Otwórz plik.
2. Uruchom makro.
3. Powtórz dla następnego pliku.

Tryb batch na cały folder nie jest jeszcze dodany, ponieważ automatyczne
otwieranie plików może porzucić niezapisane zmiany.

## Rysunek jest minimalnie mniejszy od strony

Rozmiar strony PDF jest dokładny, ale sterownik PDF może dodać własne
niewielkie marginesy. To normalne dla niektórych konfiguracji sterownika.
Sprawdź marginesy w oknie konfiguracji sterownika.

Jeśli rysunek jest znacznie mniejszy, `Maximize` dopasowuje stronę, której
proporcje nie odpowiadają rysunkowi. Sprawdź, czy wykryty rozmiar jest
poprawny.

## Rysunek jest obcięty

Wykryta strona jest mniejsza niż rzeczywisty rysunek. Sprawdź podgląd i
źródło:

- dla modelu arkusza sprawdź rozmiar arkusza,
- dla levelu ramki sprawdź, czy level zawiera całą ramkę i czy nic nie jest
  na ukrytym/zablokowanym levelu,
- zaznacz pełną ramkę ręcznie.

## Makro zmienia fence albo aktywny model

Makro używa tymczasowego fence'a i aktywuje kolejne modele podczas drukowania.
Na końcu usuwa fence i wraca do modelu, który był aktywny przed uruchomieniem.
Fence, który istniał wcześniej, nie jest odtwarzany. Zapisz plik i nie
uruchamiaj makra, gdy potrzebujesz konkretnego fence'a.

Makro nie modyfikuje istniejących elementów rysunku. Tworzy tylko tymczasowy fence i
zmienia ustawienia druku.

## Plik DGN jest tylko do odczytu

Drukowanie powinno działać, ale makro tworzy folder `PDF` obok DGN. Jeśli
folder nie jest zapisywalny, makro się zatrzyma. Użyj zapisywalnej kopii DGN
albo zmień `OUTPUT_SUBFOLDER` na zapisywalne miejsce.

## Ścieżka DGN albo folderu zawiera spacje

Makro zwykle zamienia folder wyjściowy na krótką ścieżkę Windows
(`PROGRA~1`), zastępuje spacje w nazwie pliku znakiem `_` i zamienia polskie
znaki diakrytyczne w nazwie pliku na litery bez ogonków. To unika problemów
z parserem key-inów. Jeśli wolumin nie udostępnia krótkich ścieżek (niektóre
udziały sieciowe), makro ujmuje ścieżkę w cudzysłów. Jeśli to nadal nie
działa, do testu użyj folderu bez spacji i najlepiej bez polskich znaków
w ścieżce.

## Wersja MicroStation

Makro jest przeznaczone dla MicroStation V8i i korzysta z obiektów VBA V8i:

- `ModelReference.GetSheetDefinition`
- `Fence.DefineFromModelPoints`
- key-inów `PRINT ...`

MicroStation CONNECT ma inne API drukowania, więc ten sam kod nie jest
przewidziany dla CONNECT. Część odpowiedzialną za wykrywanie rozmiaru i pliki
można zaadaptować, ale wywołania druku trzeba by zastąpić.
