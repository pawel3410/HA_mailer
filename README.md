# HA_mailer

Integracja HACS dla Home Assistanta. W wybranych terminach (wyrażenia **cron**) robi migawki wartości encji, zapisuje je w plikach **Excel (.xlsx)** i wysyła mailem z jednego konta **Gmail** na wskazane adresy. Jedna integracja obsługuje **wiele raportów**, a każdy z nich ma własny plik, harmonogram, odbiorców, tytuł, treść, kolumny i wiersze.

## Funkcje

- **wiele raportów** w jednej integracji (wspólne konto Gmail), każdy z przełącznikiem włączającym i wyłączającym harmonogram,
- **konfigurowalny układ pliku**: wybór i kolejność kolumn, wiersze z encją, **atrybutem encji** lub z samym opisem, **wiele arkuszy** w jednym pliku,
- ujednolicony wygląd Excela, z oznaczaniem encji niedostępnych i brakujących,
- DW i UDW, własna treść maila ze znacznikami,
- raport powstaje **zawsze o czasie z harmonogramu**, w tle, a wysyłka (także zaległych plików) odbywa się osobno,
- **ponawianie** nieudanych wysyłek co zadany czas, tym samym plikiem,
- **archiwum plików** w wybranym folderze, z usuwaniem po zadanym czasie (na podstawie fizycznych plików w katalogu),
- jawne błędy: log, powiadomienie, wpis w **Naprawach** (Repairs), status w encjach,
- edycja wierszy graficznie lub tekstowo (YAML), konfiguracja także przez `configuration.yaml`,
- akcje `ha_mailer.send_now` i `ha_mailer.clear_pending`.

## Wymagania

- Home Assistant 2024.8 lub nowszy,
- konto Gmail z weryfikacją dwuetapową i **hasłem aplikacji**,
- dostęp HA do internetu (`smtp.gmail.com`, port 465) i zapis w wybranym folderze.

Biblioteki `openpyxl` i `croniter` Home Assistant instaluje sam.

## Instalacja

### Przez HACS

1. W HACS wybierz **⋮ → Własne repozytoria**, wklej adres repozytorium, kategoria **Integracja**.
2. Zainstaluj **HA_mailer** i **zrestartuj Home Assistanta**.
3. **Ustawienia → Urządzenia i usługi → Dodaj integrację → HA_mailer** (lub skonfiguruj przez YAML, patrz niżej).

### Ręcznie

Skopiuj folder `custom_components/ha_mailer` do `/config/custom_components/` i zrestartuj HA.

## Hasło aplikacji Gmail

Zwykłe hasło **nie zadziała**. Włącz weryfikację dwuetapową i utwórz hasło aplikacji: https://myaccount.google.com/apppasswords. Spacje w haśle są ignorowane.

## Konfiguracja w interfejsie

**Dodawanie integracji** to dwa kroki:

1. **Ustawienia wspólne** dla wszystkich raportów:

| Pole | Opis |
|---|---|
| Konto Gmail i hasło aplikacji | nadawca wszystkich raportów (logowanie jest sprawdzane przy zapisie) |
| Ponawianie co (minuty) | co ile ponawiać nieudaną wysyłkę |
| Przechowuj pliki (dni) | po ilu dniach usuwać pliki, `0` = nie usuwaj |
| Folder na pliki Excel | wspólny folder dla wszystkich raportów, domyślnie `HA_mailer_reports` w katalogu z `configuration.yaml`. Ścieżka względna jest liczona od tego katalogu, można też podać ścieżkę bezwzględną (np. `/media/raporty`). |

2. **Pierwszy raport**: nazwa, harmonogram (cron), odbiorca, DW, UDW, stały tytuł, własna treść maila, kolumny i encje.

**Zmiany**: przycisk **Konfiguruj** otwiera listę raportów z menu: *Ustawienia wspólne*, *Dodaj raport*, *Edytuj raport*, *Usuń raporty*, *Zapisz wszystko i zamknij* oraz *Odrzuć zmiany i zamknij*.

**Zapis jest tylko jeden.** Wszystkie zmiany (także w edytorze wierszy) trafiają od razu do kopii roboczej i są widoczne w trakcie sesji, ale trwale zapisuje je wyłącznie **Zapisz wszystko i zamknij** (dostępne na liście raportów i w menu raportu). **Odrzuć zmiany i zamknij** kasuje wszystko, co zmieniono w tej sesji. Z każdego podmenu można wrócić krok wstecz bez utraty zmian. Przed zapisem sprawdzana jest poprawność wszystkich raportów.

Po wybraniu raportu do edycji dostępne jest menu:

| Pozycja | Co robi |
|---|---|
| Ustawienia raportu | nazwa, harmonogram, adresy (do, DW, UDW), tytuł, treść, kolumny |
| Wiersze - edytor graficzny | dodaj wiersz z encją (opcjonalnie z atrybutem i opisem), z samym opisem lub nowy arkusz, edytuj, usuń, przesuń na wskazaną pozycję |
| Kolumny i wiersze - edycja tekstowa | cały układ raportu jako YAML w jednym polu |

Jeśli dodasz wiersz z samym opisem, a układ nie ma kolumny „Opis”, zostanie ona dodana automatycznie.

Przykład edycji tekstowej:

```yaml
columns: [lp, opis, encja, atrybut, nazwa, wartosc, jednostka, zmiana]
rows:
  - sheet: Salon                  # nowy arkusz Excela
  - description: Temperatury      # wiersz z samym opisem
  - entity: sensor.salon_temperatura
    description: Temperatura
  - entity: sensor.salon_czujnik
    attribute: battery_level      # wartość atrybutu zamiast stanu
  - sheet: Kuchnia
  - entity: sensor.kuchnia_temperatura
```

## Układ pliku Excel

**Kolumny** (klucz → nagłówek): `lp` → Lp., `opis` → Opis, `encja` → Encja, `atrybut` → Atrybut, `nazwa` → Nazwa, `wartosc` → Wartość, `jednostka` → Jednostka, `zmiana` → Ostatnia zmiana. Kolumny występują w podanej kolejności, a niewybrana kolumna nie trafia do pliku. Kolumny nie mogą się powtarzać.

**Wiersze** (każdy arkusz ma ten sam, stały układ):

| Wiersz arkusza | Zawartość |
|---|---|
| 1 | tytuł (nazwa raportu, w dodatkowych arkuszach także nazwa arkusza) |
| 2 | czas migawki oraz liczba encji niedostępnych lub brakujących |
| 4 | nagłówki wybranych kolumn |
| 5 i dalej | wiersze raportu, po kolei |

- **Wiersz z encją**: wartości ze stanu w chwili tworzenia pliku. Liczby są zapisywane jako liczby, teksty zawsze jako tekst (stan zaczynający się od „=” nie jest formułą). Brakująca encja zostaje na miejscu z napisem `brak encji`.
- **Wiersz z atrybutem**: w kolumnie „Wartość” jest wartość wskazanego atrybutu (brak atrybutu daje `brak atrybutu`), a „Nazwa” ma postać `Nazwa - atrybut`. Listy i słowniki są zapisywane jako tekst JSON, wartości logiczne jako `tak` / `nie`.
- **Wiersz z samym opisem**: pogrubiony tekst w kolumnie „Opis” na jasnoniebieskim tle (nagłówek lub separator).
- **Nowy arkusz** (`sheet`): wiersze od tego miejsca trafiają do nowego arkusza, a numeracja Lp. zaczyna się od 1. Nazwy są dostosowywane do ograniczeń Excela (31 znaków, bez znaków `[ ] : * ? / \`).
- **Ostrzeżenia**: komórka wartości encji o stanie `unavailable` lub `unknown`, brakującej encji i brakującego atrybutu jest podświetlona na czerwono, a liczba takich wierszy jest podana w wierszu 2 i w treści maila.
- **Wygląd**: ramki, wyrównanie (liczby po prawej), naprzemienne tło wierszy, zamrożony nagłówek, dopasowana szerokość kolumn, wydruk poziomy na jednej stronie szerokości z powtarzanym nagłówkiem.
- Pozycja wiersza w pliku odpowiada pozycji na liście. Usunięcie wiersza przesuwa kolejne w górę.

## Mail

- **Tytuł**: stały tytuł raportu plus data i godzina utworzenia pliku, np. `HA_mailer - raport 2026-11-03 08:00:00`.
- **Adresaci**: odbiorca (kilka adresów po przecinku), DW (widoczna kopia) i UDW (ukryta kopia, nie trafia do nagłówków).
- **Treść**: domyślnie krótka informacja z liczbą encji i ostrzeżeniem o encjach niedostępnych. Własną treść ustawisz w polu „Własna treść maila”, używając znaczników: `[raport]`, `[data]`, `[liczba_encji]`, `[ostrzezenia]`. W przypadku opóźnionej wysyłki do treści dopisywana jest adnotacja.

## Jak działa wysyłka

1. **O czasie z harmonogramu** raport jest tworzony w tle, od razu zapisywany w folderze i dodawany do kolejki. Tworzenie nie czeka na wysyłkę, więc migawka zawsze odpowiada terminowi, nawet gdy poczta nie działa.
2. **Wysyłka w tle** obejmuje nowy plik i wszystkie zaległe, w kolejności tworzenia.
3. Jeśli **błąd dotyczy połączenia lub logowania**, pozostałe pliki nie są dodatkowo męczone: kolejna próba całej kolejki następuje po interwale ponawiania. Błąd dotyczący tylko jednego raportu (np. odrzucony adres) nie blokuje pozostałych.
4. Ponowiony mail ma ten sam tytuł i załącznik (stan z chwili harmonogramu), a ponawianie trwa **do skutku**. Kolejka jest zapisywana na dysku, więc przetrwa restart HA.
5. Wyłączenie raportu przełącznikiem zatrzymuje tworzenie nowych plików z harmonogramu, a zaległe pliki nadal są wysyłane. Ręczna wysyłka (`send_now`) działa zawsze.
6. Zmiana ustawień przeładowuje integrację, ale najpierw czeka na dokończenie trwającej wysyłki, więc nie dochodzi do duplikatów.

Ponawianie nie ma limitu prób, np. zły adres lub hasło będą ponawiane bez końca. Aby przerwać, użyj akcji `ha_mailer.clear_pending`, a jeśli przyczyną były błędne dane, popraw je w ustawieniach, bo kolejka wyśle się przy najbliższej próbie.

**Folder niedostępny**: gdy plik nie da się zapisać (np. odłączony dysk), raport nie przepada. Jest wysyłany z pamięci, w Naprawach pojawia się ostrzeżenie, a przy kolejnych próbach dodatek ponawia zapis pliku. Plik utworzony tylko w pamięci nie przetrwa restartu HA.

## Pliki i retencja

Każdy raport tworzy plik `HA_mailer_<id raportu>_<data>_<godzina>.xlsx` we wspólnym folderze (zapis jest sprawdzany przy konfiguracji).

- Retencja opiera się na **fizycznych plikach** w katalogu i **dacie ich zapisu**, a nie na wewnętrznym indeksie. Co godzinę i przy starcie HA dodatek skanuje folder (oraz foldery używane wcześniej) i usuwa pliki starsze niż ustawiona liczba dni.
- Usuwane są wyłącznie pliki o nazwie w formacie tworzonym przez HA_mailer, więc pozostałe pliki w tym folderze są bezpieczne.
- Pliki czekające w kolejce wysyłki nie są usuwane.
- Po zmianie folderu stare pliki nadal podlegają retencji.
- Jeśli ustawisz folder wewnątrz `/config/www`, pliki będą dostępne publicznie pod `/local/` bez logowania. Nie rób tego dla danych wrażliwych.

## Konfiguracja przez configuration.yaml

```yaml
ha_mailer:
  id: glowny                       # opcjonalne, wymagane przy kilku blokach
  gmail_user: moj.adres@gmail.com
  gmail_app_password: !secret gmail_app_password
  retry_minutes: 10                # opcjonalne, domyślnie 10
  retention_days: 30               # opcjonalne, domyślnie 30 (0 = nie usuwaj)
  storage_path: HA_mailer_reports  # opcjonalne, domyślnie HA_mailer_reports
  reports:
    - name: Raport dzienny
      cron: "0 8 * * *"
      recipient: biuro@example.com
      cc: szef@example.com                       # opcjonalne
      bcc: archiwum@example.com                  # opcjonalne
      subject: "HA_mailer - raport dzienny"      # opcjonalne
      body: "Raport [raport] z [data], encji: [liczba_encji]."   # opcjonalne
      columns: [lp, opis, encja, nazwa, wartosc, jednostka, zmiana]   # opcjonalne
      rows:
        - description: Salon
        - entity: sensor.salon_temperatura
          description: Temperatura
        - entity: sensor.salon_czujnik
          attribute: battery_level
    - name: Raport tygodniowy
      id: tygodniowy                # opcjonalne, domyślnie z nazwy
      cron: "0 7 * * 1"
      recipient: "szef@example.com, ksiegowosc@example.com"
      columns: [encja, wartosc]
      rows:
        - sheet: Energia
        - entity: sensor.licznik_energii
```

Po restarcie HA blok jest importowany do wpisu widocznego w interfejsie (**HA_mailer (YAML)**). Stary, płaski format z wersji 1.2 (jeden raport na poziomie bloku) nadal działa i jest traktowany jako jeden raport o nazwie „Raport”.

**Synchronizacja YAML ↔ interfejs:**

- Zmiany z interfejsu **zostają**, dopóki nie zmienisz treści bloku w YAML.
- Gdy zmienisz blok w YAML i zrestartujesz HA, wpis jest **w całości nadpisywany** zawartością YAML, także zmiany zrobione w interfejsie.
- Usunięcie bloku z YAML nie usuwa wpisu, trzeba go skasować w interfejsie.
- Przy imporcie nie sprawdza się logowania do Gmaila ani zapisu w folderze. Błąd pojawi się przy pierwszej wysyłce.
- Błędy w YAML (zły cron, nieistniejąca kolumna, powtórzone ID raportu) zgłasza Home Assistant przy sprawdzaniu konfiguracji.

## Encje

Dla każdego raportu powstaje osobne urządzenie (o nazwie raportu), a dodatkowo urządzenie **HA_mailer** z licznikiem kolejki. Identyfikatory encji pochodzą z nazw, więc mogą się różnić od przykładów.

| Encja | Stan | Atrybuty |
|---|---|---|
| Wysyłka z harmonogramu (przełącznik) | wł. / wył. | `nastepna_wysylka` |
| Status wysyłki | `OK` / `Błąd` / `Brak danych` | `blad`, `ostatnia_proba`, `oczekujace_pliki`, `aktywny`, `odbiorca` |
| Ostatnia wysyłka | czas ostatniej **udanej** wysyłki | `ostatnia_proba` |
| Statystyki | liczba udanych wysyłek | `udane`, `udane_po_ponowieniu`, `nieudane_proby`, `razem_prob`, `skutecznosc_prob_proc`, `ostatni_blad`, `czas_ostatniego_bledu`, `rozmiar_ostatniego_pliku_kb`, `nastepna_wysylka`, `liczba_encji`, `liczba_wierszy`, `liczba_arkuszy` |
| Oczekujące wysyłki (urządzenie HA_mailer) | liczba plików w kolejce | `kolejka` (raport, plik, zapisany_na_dysku, utworzono, nieudane_proby, nastepna_proba, blad), `katalog_plikow`, `katalog_dostepny`, `pliki_w_katalogu`, `retencja_dni` |

Stan przełącznika jest zapamiętywany i przetrwa restart. `nieudane_proby` liczy każdą nieudaną próbę, także ponowienia. Dane są zapisywane w `/config/.storage/ha_mailer.<entry_id>`. Encje usuniętych raportów są usuwane automatycznie.

## Obsługa błędów

Każda nieudana próba: trafia do logu (filtr `ha_mailer`), tworzy powiadomienie w panelu HA i wpis w **Ustawienia → Naprawy** dla danego raportu (znikają po wysłaniu wszystkich jego zaległych plików), ustawia status `Błąd` z opisem w atrybucie `blad` i zwiększa licznik `nieudane_proby`. Niedostępny folder ma osobny wpis w Naprawach.

```yaml
automation:
  - alias: "HA_mailer - błąd wysyłki"
    triggers:
      - trigger: state
        entity_id: sensor.raport_dzienny_status_wysylki   # dostosuj do swojego raportu
        to: "Błąd"
    actions:
      - action: notify.notify
        data:
          title: "HA_mailer"
          message: "{{ state_attr('sensor.raport_dzienny_status_wysylki', 'blad') }}"
```

## Akcje

| Akcja | Opis |
|---|---|
| `ha_mailer.send_now` | natychmiast tworzy plik i wysyła go (poza harmonogramem), bez ponawiania; błąd wraca do wywołującego |
| `ha_mailer.clear_pending` | porzuca pliki czekające w kolejce (zostają w folderze do końca retencji) |

Obie akcje mają opcjonalne pola **Konfiguracja** i **Raport** (nazwa lub ID). Puste pola oznaczają wszystkie. Test: **Narzędzia deweloperskie → Akcje**.

## Przykłady harmonogramu

| Wyrażenie | Kiedy |
|---|---|
| `0 8 * * *` | codziennie o 8:00 |
| `0 7 * * 1` | w poniedziałki o 7:00 |
| `0 18 1 * *` | pierwszego dnia miesiąca o 18:00 |
| `0 6,18 * * 1-5` | pn-pt o 6:00 i 18:00 |
| `*/30 * * * *` | co 30 minut |

Format: `minuta godzina dzień-miesiąca miesiąc dzień-tygodnia`, strefa czasowa jak w HA.

## Aktualizacja z wcześniejszych wersji

Wpisy z wersji 1.x i 2.0 są migrowane automatycznie (dotychczasowa konfiguracja staje się raportem „Raport”, encje i statystyki zostają). Zmiany względem 2.0:

- układ arkusza ma teraz tytuł w wierszu 1, czas migawki w wierszu 2, nagłówki w wierszu 4 i dane od wiersza 5,
- domyślny folder to `HA_mailer_reports` w katalogu z `configuration.yaml` (wcześniej zapisany folder zostaje bez zmian),
- atrybut statusu `oczekujace_ponowienia` nazywa się teraz `oczekujace_pliki`, bo kolejka obejmuje też świeże, jeszcze niewysłane pliki,
- stare pliki w folderze, utworzone przez wersję 2.0, są objęte retencją po fizycznych plikach.

## Rozwiązywanie problemów

| Objaw | Przyczyna / rozwiązanie |
|---|---|
| „Logowanie do Gmaila nie powiodło się” | Użyto zwykłego hasła zamiast hasła aplikacji albo brak weryfikacji dwuetapowej. |
| „Nie można utworzyć folderu lub zapisać w nim plików” | Sprawdź ścieżkę i uprawnienia (np. `HA_mailer_reports`, `/media/...`, `/share/...`). |
| „Nie można połączyć się z smtp.gmail.com” | Brak internetu lub zablokowany port 465. |
| Kolejka oczekujących ciągle rośnie | Wysyłka stale się nie udaje. Sprawdź atrybut `blad` statusu i Naprawy, popraw dane i ewentualnie użyj `clear_pending`. |
| Naprawa „folder na raporty jest niedostępny” | Dysk lub udział sieciowy nie jest dostępny albo brakuje uprawnień. Do czasu naprawy raporty nie są archiwizowane. |
| „Masz wiersze bez encji - dodaj kolumnę Opis” | Wiersze z samym opisem wymagają kolumny „Opis” w układzie. |
| Wartość atrybutu nie pojawia się w pliku | Dodaj kolumnę „Atrybut”, jeśli chcesz widzieć nazwę atrybutu. Sama wartość trafia do kolumny „Wartość”. |
| Raport nie powstaje o czasie | Sprawdź przełącznik „Wysyłka z harmonogramu” i atrybut `nastepna_wysylka`. |
| Zmiany z interfejsu zniknęły po restarcie | Zmieniono blok `ha_mailer` w YAML, co nadpisuje wpis. |
| Mail nie przychodzi, status `OK` | Sprawdź spam i adres odbiorcy. |

## Ograniczenia

- Wysyłki, które wypadły w czasie wyłączenia HA, nie są tworzone wstecz, ale kolejka wysyłki działa po restarcie.
- Hasło aplikacji jest przechowywane w konfiguracji HA (`/config/.storage`).
- Obowiązują limity Gmaila (załącznik do 25 MB, dzienny limit wysyłek).
- Tekst wyglądający na liczbę (np. wersja `1.10`) jest zapisywany jako liczba.
- Interfejs konfiguracji jest po polsku.

## Struktura repozytorium

```
hacs.json
README.md
custom_components/ha_mailer/
├── __init__.py        # logika: raporty, harmonogramy, Excel, kolejka, archiwum, import YAML
├── config_flow.py     # konfiguracja, opcje, edytor graficzny i tekstowy
├── sensor.py          # encje: status, ostatnia wysyłka, statystyki, kolejka
├── switch.py          # przełącznik harmonogramu raportu
├── services.yaml      # opis akcji
├── manifest.json
└── translations/pl.json
```
