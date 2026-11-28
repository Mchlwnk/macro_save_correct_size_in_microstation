# Lista kontrolna testów

Użyj tej listy przed publikacją nowej wersji i przed pierwszym użyciem makra
na produkcyjnych rysunkach.

## Przygotowanie

- [ ] Pracuj na **kopii** prawdziwego pliku DGN.
- [ ] Zapisz plik.
- [ ] Zapisz sobie oczekiwane rozmiary arkuszy dla każdego modelu.
- [ ] Sprawdź, czy sterownik Bentley PDF `pdf.pltcfg` jest zainstalowany.
- [ ] Sprawdź, czy `MS_PLTCFG_PATH` wskazuje folder ze sterownikiem.
- [ ] Sprawdź, czy sterownik pozwala na edycję rozmiaru papieru.
- [ ] Sprawdź, czy ramki są na levelach `M_RAMKA` albo ` _RAMKA`.

## SelfTest

- [ ] Uruchom `PrintCorrectSize.SelfTest`.
- [ ] Komunikat to `SelfTest OK`.
- [ ] Liczba sprawdzonych testów jest większa od zera.

## Podgląd

- [ ] Uruchom `PrintCorrectSize.Preview`.
- [ ] Każdy oczekiwany model jest na liście.
- [ ] Nie ma modeli-źródeł, których nie chcesz drukować.
- [ ] Wykryte rozmiary zgadzają się z rzeczywistymi ramkami (z dokładnością
      do ułamka milimetra).
- [ ] Źródło w nawiasach ma sens:
      `ARKUSZ`, `LEVEL RAMKI`, `ZAZNACZENIE`, `ZAMKNIETY ELEMENT`, `GRANICE`.
- [ ] Żaden model nie jest oznaczony jako `POMINIETY` przez pomyłkę.

## Druk

- [ ] Uruchom `PrintCorrectSize.Main`.
- [ ] Potwierdź listę modeli.
- [ ] Makro kończy się bez okna błędu.
- [ ] W folderze `PDF` jest log `PrintCorrectSize_<dgn>.log`.
- [ ] Log pokazuje `OK` dla każdego oczekiwanego modelu.

## Weryfikacja PDF

Dla każdego PDF-a:

- [ ] Plik istnieje i nie ma rozmiaru 0 bajtów.
- [ ] Rozmiar strony w **File > Properties** zgadza się z wykrytym rozmiarem.
- [ ] Rysunek nie jest obcięty.
- [ ] Rysunek wypełnia stronę (z uwzględnieniem marginesów sterownika).
- [ ] Grubości linii, wypełnienia, tekst i referencje wyglądają poprawnie.
- [ ] PDF otwiera się w Acrobat/Reader bez błędów.

## Przypadki brzegowe

- [ ] Ramka A4 pionowa.
- [ ] Ramka A3 pozioma.
- [ ] Ramka w granicach 5 mm od A4, np. 212 × 299 mm → strona 210 × 297 mm.
- [ ] Ramka w granicach 5 mm od A3, np. 300 × 423 mm → strona 297 × 420 mm.
- [ ] Ramka 6 mm od A4, np. 216 × 300 mm → strona 216 × 300 mm.
- [ ] Ramka A2 albo B → strona zachowuje zmierzony wymiar.
- [ ] Ramka niestandardowa, np. 350 × 500 mm.
- [ ] Ramka na `M_RAMKA`.
- [ ] Ramka na ` _RAMKA` (z początkową spacją).
- [ ] Ramka z linii, nie z zamkniętego kształtu.
- [ ] Ramka z tekstem lub wymiarem poza obrysem — strona nie może się
      powiększyć.
- [ ] Model bez ramki, z zaznaczeniem elementów strony.
- [ ] DGN z modelami arkuszy i modelami źródłowymi.
- [ ] DGN w ścieżce ze spacjami.
- [ ] Nazwa modelu ze znakami niedozwolonymi w nazwie pliku.
- [ ] Model, który powinien zostać pominięty jako pusty.

## Po teście

- [ ] Zamknij PDF-y.
- [ ] Usuń testowe PDF-y i logi z folderu roboczego repozytorium (są ignorowane
      przez `.gitignore`).
- [ ] Nie commituj prawdziwych plików DGN ani PDF-ów.
