# HA_mailer

Integracja HACS dla Home Assistanta. W wybranych terminach (wyrażenia **cron**) robi migawki wartości encji, zapisuje je w plikach **Excel (.xlsx)** i wysyła mailem z jednego konta **Gmail** na wskazane adresy. Jedna konfiguracja obsługuje **wiele raportów**, a każdy z nich ma własny plik, harmonogram, odbiorców, tytuł, treść, kolumny i wiersze albo własny szablon Excel.

**Konfiguracja odbywa się wyłącznie w pliku `configuration.yaml`.** Integracja nie ma żadnych formularzy ani opcji w interfejsie.

## Funkcje

- **wiele raportów** (wspólne konto Gmail), każdy z przełącznikiem włączającym i wyłączającym harmonogram,
- **konfigurowalny układ pliku**: wybór i kolejność kolumn, wiersze z encją, **atrybutem encji** lub z samym opisem, **wiele arkuszy**,
- **własny szablon Excel** ze znacznikami `{{...}}` w komórkach: dodatek wypełnia go wartościami encji, zachowując Twoje formatowanie,
- ujednolicony wygląd Excela, z oznaczaniem encji niedostępnych i brakujących,
- DW i UDW, własna treść maila ze znacznikami,
- **język tekstów** generowanych przez dodatek (polski lub angielski) w treści maila i w pliku Excel,
- **tryb braku wartości** (`no_value`): encja bez wartości może być ostrzeżeniem, informacją albo być pomijana,
- raport powstaje **zawsze o czasie z harmonogramu**, w tle, a wysyłka (także zaległych plików) odbywa się osobno,
- **ponawianie** nieudanych wysyłek co zadany czas, tym samym plikiem,
- **archiwum plików** w wybranym folderze, z usuwaniem po zadanym czasie (na podstawie fizycznych plików w katalogu),
- jawne błędy: log, powiadomienie, wpis w **Naprawach** (Repairs), status w encjach,
- przeładowanie konfiguracji bez restartu, akcje `ha_mailer.send_now`, `ha_mailer.clear_pending` i `ha_mailer.reload`.

## Wymagania

- Home Assistant 2024.8 lub nowszy,
- konto Gmail z weryfikacją dwuetapową i **hasłem aplikacji**,
- dostęp HA do internetu (`smtp.gmail.com`, port 465) i zapis w wybranym folderze.

Biblioteki `openpyxl` i `croniter` Home Assistant instaluje sam.

## Instalacja

### Przez HACS

1. W HACS wybierz **⋮ → Własne repozytoria**, wklej adres repozytorium, kategoria **Integracja**.
2. Zainstaluj **HA_mailer**.
3. Dodaj blok `ha_mailer:` do `configuration.yaml` (patrz niżej) i **zrestartuj Home Assistanta**.

### Ręcznie

Skopiuj folder `custom_components/ha_mailer` do `/config/custom_components/`, dodaj blok do `configuration.yaml` i zrestartuj HA.

## Hasło aplikacji Gmail

Zwykłe hasło **nie zadziała**. Włącz weryfikację dwuetapową i utwórz hasło aplikacji: https://myaccount.google.com/apppasswords. Spacje w haśle nie są potrzebne. Zapisz je w `secrets.yaml` (`ha_mailer_gmail_password: abcdefghijklmnop`) i odwołuj się przez `!secret`.

## Konfiguracja w configuration.yaml

```yaml
ha_mailer:
  id: glowny                       # opcjonalne, wymagane przy kilku blokach
  gmail_user: moj.adres@gmail.com
  gmail_app_password: !secret ha_mailer_gmail_password
  retry_minutes: 10                # opcjonalne, domyślnie 10
  retention_days: 30               # opcjonalne, domyślnie 30 (0 = nie usuwaj)
  storage_path: HA_mailer_reports  # opcjonalne, domyślnie HA_mailer_reports
  language: pl                     # opcjonalne: pl (domyślnie) lub en, dla wszystkich raportów
  no_value: warning                # opcjonalne: warning (domyślnie), info lub ignore
  reports:
    - name: Raport dzienny
      cron: "0 8 * * *"
      recipient: biuro@example.com
      cc: szef@example.com                       # opcjonalne
      bcc: archiwum@example.com                  # opcjonalne
      subject: "HA_mailer - raport dzienny"      # opcjonalne
      body: "Raport [raport] z [data], encji: [liczba_encji]."   # opcjonalne
      language: en                               # opcjonalne, nadpisuje ustawienie bloku
      no_value: info                             # opcjonalne, nadpisuje ustawienie bloku
      columns: [lp, opis, encja, nazwa, wartosc, jednostka, zmiana]   # opcjonalne
      rows:
        - sheet: Salon                   # nowy arkusz
        - description: Temperatury       # wiersz z samym opisem
        - entity: sensor.salon_temperatura
          description: Temperatura
        - entity: sensor.salon_czujnik
          attribute: battery_level       # wartość atrybutu zamiast stanu
    - name: Raport z szablonu
      cron: "0 18 * * *"
      recipient: biuro@example.com
      template: HA_mailer_templates/dzienny.xlsx   # zamiast columns i rows
    - name: Raport tygodniowy
      id: tygodniowy                # opcjonalne, domyślnie z nazwy
      cron: "0 7 * * 1"
      recipient: "szef@example.com, ksiegowosc@example.com"
      columns: [encja, wartosc]
      rows:
        - entity: sensor.licznik_energii
```

Można podać listę kilku bloków `ha_mailer:` (każdy z własnym `id` i własnym kontem Gmail, folderem i ustawieniami).

### Klucze bloku

| Klucz | Wymagany | Opis |
|---|---|---|
| `id` | nie | identyfikator bloku (domyślnie numer bloku: `0`, `1`...). Ustaw go jawnie, gdy masz kilka bloków lub zmieniasz ich kolejność, bo na nim opierają się wpisy, statystyki i kolejka |
| `gmail_user` | tak | konto Gmail nadawcy |
| `gmail_app_password` | tak | hasło aplikacji (najlepiej `!secret`) |
| `retry_minutes` | nie | co ile minut ponawiać nieudaną wysyłkę (1-1440, domyślnie 10) |
| `retention_days` | nie | po ilu dniach usuwać pliki (0-3650, domyślnie 30, `0` = nie usuwaj) |
| `storage_path` | nie | folder na pliki Excel. Domyślnie `HA_mailer_reports` obok `configuration.yaml`. Ścieżka względna jest liczona od katalogu z `configuration.yaml`, można też podać bezwzględną (np. `/media/raporty`) |
| `language` | nie | `pl` (domyślnie) lub `en`: język tekstów generowanych przez dodatek, domyślny dla raportów bloku |
| `no_value` | nie | `warning` (domyślnie), `info` lub `ignore`: jak traktować encje bez wartości, domyślnie dla raportów bloku |
| `reports` | tak | lista raportów |

### Klucze raportu

| Klucz | Wymagany | Opis |
|---|---|---|
| `name` | tak | nazwa raportu (widoczna w nazwie urządzenia i w mailu) |
| `id` | nie | identyfikator raportu (małe litery, cyfry, `_`), domyślnie z nazwy. Nie zmieniaj go, jeśli chcesz zachować statystyki i kolejkę |
| `cron` | tak | harmonogram, 5 pól, strefa czasowa jak w HA |
| `recipient` | tak | jeden lub kilka adresów oddzielonych przecinkiem |
| `cc`, `bcc` | nie | DW (widoczna kopia) i UDW (ukryta kopia) |
| `subject` | nie | stały tytuł (domyślnie `HA_mailer - raport`), do którego dopisywana jest data i godzina |
| `body` | nie | własna treść maila (znaczniki: `[raport]`, `[data]`, `[liczba_encji]`, `[ostrzezenia]`, `[informacje]`) |
| `template` | nie | ścieżka do szablonu `.xlsx` (zamiast `columns` i `rows`) |
| `language`, `no_value` | nie | nadpisują ustawienia bloku dla tego raportu |
| `columns` | nie | kolumny w kolejności (domyślnie `lp, opis, encja, nazwa, wartosc, jednostka, zmiana`) |
| `rows` | tak, bez `template` | lista wierszy |

### Klucze wiersza

| Klucz | Opis |
|---|---|
| `entity` | encja, której wartość trafia do pliku |
| `attribute` | (z `entity`) atrybut encji zamiast stanu |
| `description` | tekst w kolumnie „Opis”. Bez `entity` powstaje wiersz z samym opisem (nagłówek lub separator) i wymaga kolumny `opis` |
| `sheet` | nazwa nowego arkusza. Wiersze od tego miejsca trafiają do nowego arkusza |

Błędy składni (zły cron, nieistniejąca kolumna, wiersz bez encji i opisu, brak `rows` i `template`, powtórzone `id`) zgłasza Home Assistant przy sprawdzaniu konfiguracji. Błędów logowania do Gmaila dowiesz się przy pierwszej wysyłce, a problemy z folderem i szablonami są zgłaszane w **Naprawach** zaraz po starcie lub przeładowaniu.

## Zastosowanie zmian w YAML

- **Bez restartu:** **Narzędzia deweloperskie → Akcje → HA_mailer: Przeładuj konfigurację YAML** (`ha_mailer.reload`). Zmienione bloki są przeładowywane, nowe tworzone, a usunięte kasowane. Przed zapisaniem zmian możesz sprawdzić składnię w *Narzędzia deweloperskie → YAML → Sprawdź konfigurację*.
- **Restart** Home Assistanta robi to samo.

YAML jest jedynym źródłem konfiguracji. W **Ustawienia → Urządzenia i usługi** widać wpisy `HA_mailer (id)`, ale są one tylko odbiciem YAML: nie mają przycisku „Konfiguruj”, przycisk „Dodaj integrację” pokazuje wyłącznie informację o konfiguracji w YAML, a ręcznie usunięty wpis wróci po przeładowaniu. Aby usunąć raport lub cały blok, usuń go z YAML i przeładuj konfigurację.

Niezmieniony blok nie jest ponownie ładowany, więc przeładowanie nie przerywa trwających wysyłek. Zmiana bloku przeładowuje tylko jego, a trwająca wysyłka jest przed tym dokończona.

## Język i encje bez wartości

### Język (`language`)

`language: pl` (domyślnie) lub `en`, ustawiane w bloku i nadpisywane w raporcie. Obejmuje teksty, które generuje dodatek:

| | PL | EN |
|---|---|---|
| Nagłówki kolumn | Lp., Opis, Encja, Atrybut, Nazwa, Wartość, Jednostka, Ostatnia zmiana | No., Description, Entity, Attribute, Name, Value, Unit, Last changed |
| Wiersz 2 pliku | Czas migawki | Snapshot time |
| Pierwszy arkusz | Migawka | Snapshot |
| Wartości | brak danych, brak encji, brak atrybutu, tak, nie | no data, no entity, no attribute, yes, no |
| Treść maila | „Raport „X” z ... W załączniku wartości 12 encji.” | „Report “X” from ... The attachment contains the values of 12 entities.” |

Przy szablonie własnym zmieniają się tylko wartości wstawiane przez dodatek (np. `brak danych` / `no data`), a Twoje etykiety zostają bez zmian. Język nie dotyczy `body`, `subject`, nazw encji, logów, Napraw ani komunikatów HA.

### Brak wartości (`no_value`)

**Brak wartości** to stan `unavailable`, `unknown`, `none`, pusty lub złożony ze spacji tekst (bez rozróżniania wielkości liter) oraz atrybut `None` lub pusty. Dotyczy to także **brakującej encji i brakującego atrybutu**. W komórce pojawia się wtedy tekst `brak danych` / `brak encji` / `brak atrybutu` (lub angielski odpowiednik), a sposób oznaczenia wybierasz opcją `no_value`:

| Tryb | Wygląd w pliku | Liczenie | Mail |
|---|---|---|---|
| `warning` (domyślnie) | czerwone pole, pogrubiony czerwony tekst | `ostrzezenia` | „Uwaga: encje niedostępne lub brakujące: N” |
| `info` | szare pole, kursywa | `informacje` | „Informacja: encje bez wartości: N” |
| `ignore` | bez oznaczenia | nie liczone | bez wzmianki |

W wierszu 2 pliku pojawia się odpowiednio „Encje niedostępne lub brakujące: N” albo „Encje bez wartości: N” (przy `ignore` nic). Licznik dotyczy tylko wybranego trybu, więc `{{ostrzezenia}}` i `[ostrzezenia]` liczą tylko tryb `warning`, a `{{informacje}}` i `[informacje]` tylko tryb `info`. Atrybuty statystyk: `ostrzezenia_w_ostatnim_pliku` i `informacje_w_ostatnim_pliku`.

**Własny tryb w komórce szablonu.** W szablonie możesz nadpisać tryb raportu dla pojedynczego znacznika, dopisując `#no_value:tryb`:

| Znacznik w komórce | Gdy encja nie ma wartości |
|---|---|
| `{{sensor.sypialnia_temperatura}}` | komórka oznaczona wg `no_value` raportu |
| `{{sensor.sypialnia_temperatura#no_value:ignore}}` | wpisane `brak danych`, bez koloru i bez liczenia |
| `{{sensor.sypialnia_temperatura#no_value:info}}` | szara komórka, liczona jako informacja |
| `{{sensor.sypialnia_temperatura#no_value:warning}}` | czerwona komórka, liczona jako ostrzeżenie |

Zasady liczenia: każda encja (lub para encja i atrybut) bez wartości jest liczona **raz w każdym trybie, w którym wystąpiła**, więc ta sama encja użyta w kilku komórkach nie zawyża wyniku. Encja w trybie `ignore` nie jest liczona wcale. Znacznik z własnym trybem zmienia więc zarówno kolor komórki, jak i wartości `{{ostrzezenia}}`, `{{informacje}}`, `[ostrzezenia]`, `[informacje]` oraz atrybuty statystyk. Znacznik wpisany w środku tekstu (np. `Temp: {{sensor.a#no_value:info}} °C`) nie koloruje komórki, ale jest liczony. Własny tryb działa tylko w szablonach, a w raportach bez szablonu obowiązuje `no_value` raportu lub bloku.

Tryb nie wpływa na status wysyłki, Naprawy ani ponawianie, bo to tylko sposób oznaczenia danych w raporcie. Przy `info` i `ignore` literówka w identyfikatorze encji nie będzie czerwona (zobaczysz tylko `brak encji`), więc po zmianie konfiguracji warto otworzyć pierwszy raport.

## Układ pliku Excel

**Kolumny** (klucz → nagłówek): `lp` → Lp., `opis` → Opis, `encja` → Encja, `atrybut` → Atrybut, `nazwa` → Nazwa, `wartosc` → Wartość, `jednostka` → Jednostka, `zmiana` → Ostatnia zmiana. Kolumny występują w podanej kolejności, a niewybrana kolumna nie trafia do pliku. Kolumny nie mogą się powtarzać. Kolumna `atrybut` pokazuje nazwę atrybutu przy wierszach z `attribute`.

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
- **Nowy arkusz** (`sheet`): numeracja Lp. zaczyna się od 1. Nazwy są dostosowywane do ograniczeń Excela (31 znaków, bez znaków `[ ] : * ? / \`).
- **Brak wartości**: komórka wartości ma wtedy tekst `brak danych` (`brak encji`, `brak atrybutu`) i jest oznaczona wg `no_value` (patrz wyżej), a liczba takich wierszy jest podana w wierszu 2 i w treści maila.
- **Wygląd**: ramki, wyrównanie (liczby po prawej), naprzemienne tło wierszy, zamrożony nagłówek, dopasowana szerokość kolumn, wydruk poziomy na jednej stronie szerokości z powtarzanym nagłówkiem.
- Pozycja wiersza w pliku odpowiada pozycji na liście. Usunięcie wiersza przesuwa kolejne w górę.

## Szablon Excel

Zamiast układu generowanego przez dodatek możesz użyć **własnego pliku `.xlsx`** jako szablonu. W komórkach wpisujesz znaczniki w podwójnych nawiasach klamrowych, a dodatek podmienia je na wartości i zapisuje wynik jako nowy plik (szablon nie jest zmieniany). Całe formatowanie, szerokości kolumn, scalone komórki, ramki, kolory i ustawienia wydruku zostają takie, jak zaprojektowałeś.

**Jak włączyć:** podaj w raporcie `template:` ze ścieżką do pliku. Ścieżka względna jest liczona od katalogu z `configuration.yaml`, np. `HA_mailer_templates/dzienny.xlsx`. Plik wgraj ręcznie (Samba, edytor plików, SSH). Przy szablonie **`columns` i `rows` są ignorowane**, bo układ wyznacza szablon. Cała reszta (harmonogram, adresy, tytuł, treść, kolejka, retencja, przełącznik) działa tak samo.

**Znaczniki:**

| Znacznik | Wartość |
|---|---|
| `{{sensor.temperatura}}` | stan encji (liczba zapisywana jest jako liczba) |
| `{{sensor.czujnik.battery}}` | wartość atrybutu encji |
| `{{sensor.temperatura#nazwa}}` | przyjazna nazwa encji |
| `{{sensor.temperatura#jednostka}}` | jednostka miary |
| `{{sensor.temperatura#zmiana}}` | czas ostatniej zmiany stanu |
| `{{sensor.temperatura#no_value:ignore}}` | stan encji z **własnym trybem braku wartości** tylko dla tej komórki (`warning`, `info` lub `ignore`), patrz niżej |
| `{{raport}}` | nazwa raportu |
| `{{data}}`, `{{czas}}` | data (`2026-11-03`) i godzina (`08:00:00`) utworzenia pliku |
| `{{data_czas}}` | data i godzina jako wartość daty (w komórce z formatem „Ogólne” dostaje format `rrrr-mm-dd gg:mm:ss`) |
| `{{ostrzezenia}}` | liczba encji bez wartości w trybie `warning` (w pozostałych trybach 0) |
| `{{informacje}}` | liczba encji bez wartości w trybie `info` (w pozostałych trybach 0) |

Zasady:

- **Komórka z samym znacznikiem** dostaje wartość z właściwym typem (liczba, data) i zachowuje swój format, np. `0,00`. Dzięki temu w Excelu można na niej liczyć.
- **Znacznik wewnątrz tekstu** jest zamieniany na tekst, np. `Temperatura: {{sensor.a}} {{sensor.a#jednostka}}` daje `Temperatura: 21.5 °C`. W jednej komórce może być kilka znaczników.
- Znaczniki działają **w każdym arkuszu** szablonu, a w komórkach z formułami (zaczynających się od `=`) nie są podmieniane. Formuła może jednak odwoływać się do komórki ze znacznikiem, np. `=B5*2`.
- **Formuły nie są przeliczane przez dodatek.** Wypełniony plik ma włączone przeliczanie przy otwarciu, więc Excel policzy je sam, ale podgląd w poczcie lub na telefonie może pokazać puste wyniki.
- Encje bez wartości (stany `unavailable`, `unknown`, `none`, puste, brak encji lub atrybutu) są wpisywane do komórki jako `brak danych` / `brak encji` / `brak atrybutu` i **oznaczane wg `no_value`**, jeśli komórka zawiera sam znacznik (czerwone w `warning`, szare w `info`, bez oznaczenia w `ignore`). Liczby trafiają do maila oraz do `{{ostrzezenia}}` i `{{informacje}}`.
- **Własny tryb dla komórki:** przyrostek `#no_value:tryb` (`warning`, `info` lub `ignore`) nadpisuje opcję `no_value` raportu tylko dla tego znacznika. Zmienia on kolor komórki oraz sposób liczenia (patrz niżej). Przyrostki można łączyć w kolejności: `{{sensor.a.atrybut#jednostka#no_value:info}}`, a spacja po dwukropku jest dozwolona.
- Znacznik o nieznanej składni (np. `{{cos}}` lub `{{sensor.a#no_value:zly}}`) zostaje w komórce bez zmian. Zgłasza to wpis w Naprawach po starcie lub przeładowaniu oraz ostrzeżenie w logu.
- Znaczniki z encjami piszesz małymi literami (`domena.encja`), nazwy atrybutów rozróżniają wielkość liter. Spacje wewnątrz nawiasów są dozwolone (`{{ sensor.a }}`).
- Wspierane są tylko pliki `.xlsx`, bez makr, wykresów, obrazów i tabel przestawnych, bo biblioteka używana do zapisu może ich nie zachować.
- Brak pliku szablonu, błąd jego wczytania lub nieznane znaczniki są zgłaszane w Naprawach po starcie i po przeładowaniu YAML. Jeśli szablon zniknie później, wysyłka kończy się jawnym błędem (log, powiadomienie, Naprawy, status), a pozostałe raporty działają dalej.

**Gotowy przykład** to plik `examples/szablon_przyklad.xlsx` w repozytorium: sformatowany raport dzienny z temperaturami, wilgotnością, baterią czujników, podsumowaniem (formuły średniej, minimum i maksimum) oraz sekcją energii. Skopiuj go do `/config/HA_mailer_templates/`, zamień przykładowe identyfikatory encji (np. `sensor.salon_temperatura`) na własne i wpisz w raporcie `template: HA_mailer_templates/szablon_przyklad.xlsx`. Komentarz w komórce A1 przypomina o dostosowaniu go, a potem możesz go usunąć.

## Mail

- **Tytuł**: stały tytuł raportu plus data i godzina utworzenia pliku, np. `HA_mailer - raport 2026-11-03 08:00:00`.
- **Adresaci**: odbiorca (kilka adresów po przecinku), DW (widoczna kopia) i UDW (ukryta kopia, nie trafia do nagłówków).
- **Treść**: domyślnie krótka informacja z liczbą encji oraz uwagą lub informacją o encjach bez wartości (zależnie od `no_value`), w wybranym języku. Własną treść ustawisz kluczem `body`, używając znaczników: `[raport]`, `[data]`, `[liczba_encji]`, `[ostrzezenia]`, `[informacje]`. W przypadku opóźnionej wysyłki do treści dopisywana jest adnotacja.

## Jak działa wysyłka

1. **O czasie z harmonogramu** raport jest tworzony w tle, od razu zapisywany w folderze i dodawany do kolejki. Tworzenie nie czeka na wysyłkę, więc migawka zawsze odpowiada terminowi, nawet gdy poczta nie działa.
2. **Wysyłka w tle** obejmuje nowy plik i wszystkie zaległe, w kolejności tworzenia.
3. Jeśli **błąd dotyczy połączenia lub logowania**, pozostałe pliki nie są dodatkowo męczone: kolejna próba całej kolejki następuje po interwale ponawiania. Błąd dotyczący tylko jednego raportu (np. odrzucony adres) nie blokuje pozostałych.
4. Ponowiony mail ma ten sam tytuł i załącznik (stan z chwili harmonogramu), a ponawianie trwa **do skutku**. Kolejka jest zapisywana na dysku, więc przetrwa restart HA.
5. Wyłączenie raportu przełącznikiem zatrzymuje tworzenie nowych plików z harmonogramu, a zaległe pliki nadal są wysyłane. Ręczna wysyłka (`send_now`) działa zawsze.
6. Zmiana bloku w YAML przeładowuje go, ale najpierw czeka na dokończenie trwającej wysyłki, więc nie dochodzi do duplikatów.

Ponawianie nie ma limitu prób, np. zły adres lub hasło będą ponawiane bez końca. Aby przerwać, użyj akcji `ha_mailer.clear_pending`, a jeśli przyczyną były błędne dane, popraw je w YAML i przeładuj konfigurację, bo kolejka wyśle się przy najbliższej próbie.

**Folder niedostępny**: gdy plik nie da się zapisać (np. odłączony dysk), raport nie przepada. Jest wysyłany z pamięci, w Naprawach pojawia się ostrzeżenie, a przy kolejnych próbach dodatek ponawia zapis pliku. Plik utworzony tylko w pamięci nie przetrwa restartu HA.

## Pliki i retencja

Każdy raport tworzy plik `HA_mailer_<id raportu>_<data>_<godzina>.xlsx` we wspólnym folderze bloku.

- Retencja opiera się na **fizycznych plikach** w katalogu i **dacie ich zapisu**, a nie na wewnętrznym indeksie. Co godzinę i przy starcie dodatek skanuje folder (oraz foldery używane wcześniej) i usuwa pliki starsze niż `retention_days`.
- Usuwane są wyłącznie pliki o nazwie w formacie tworzonym przez HA_mailer, więc pozostałe pliki w tym folderze są bezpieczne.
- Pliki czekające w kolejce wysyłki nie są usuwane.
- Po zmianie folderu stare pliki nadal podlegają retencji.
- Jeśli ustawisz folder wewnątrz `/config/www`, pliki będą dostępne publicznie pod `/local/` bez logowania. Nie rób tego dla danych wrażliwych.

## Encje

Dla każdego raportu powstaje osobne urządzenie (o nazwie raportu), a dodatkowo urządzenie **HA_mailer** z licznikiem kolejki. Identyfikatory encji pochodzą z nazw, więc mogą się różnić od przykładów.

| Encja | Stan | Atrybuty |
|---|---|---|
| Wysyłka z harmonogramu (przełącznik) | wł. / wył. | `nastepna_wysylka` |
| Status wysyłki | `OK` / `Błąd` / `Brak danych` | `blad`, `ostatnia_proba`, `oczekujace_pliki`, `aktywny`, `odbiorca` |
| Ostatnia wysyłka | czas ostatniej **udanej** wysyłki | `ostatnia_proba` |
| Statystyki | liczba udanych wysyłek | `udane`, `udane_po_ponowieniu`, `nieudane_proby`, `razem_prob`, `skutecznosc_prob_proc`, `ostatni_blad`, `czas_ostatniego_bledu`, `rozmiar_ostatniego_pliku_kb`, `nastepna_wysylka`, `szablon`, `liczba_encji`, `ostrzezenia_w_ostatnim_pliku`, `informacje_w_ostatnim_pliku`, `liczba_wierszy`, `liczba_arkuszy` (dla szablonu `liczba_encji` pochodzi z ostatniego pliku, a liczba wierszy i arkuszy jest pusta) |
| Oczekujące wysyłki (urządzenie HA_mailer) | liczba plików w kolejce | `kolejka` (raport, plik, zapisany_na_dysku, utworzono, nieudane_proby, nastepna_proba, blad), `katalog_plikow`, `katalog_dostepny`, `pliki_w_katalogu`, `retencja_dni` |

Stan przełącznika jest zapamiętywany i przetrwa restart (nie zależy od YAML). `nieudane_proby` liczy każdą nieudaną próbę, także ponowienia. Dane są zapisywane w `/config/.storage/ha_mailer.<entry_id>`. Encje usuniętych raportów są usuwane automatycznie.

## Obsługa błędów

Każda nieudana próba: trafia do logu (filtr `ha_mailer`), tworzy powiadomienie w panelu HA i wpis w **Ustawienia → Naprawy** dla danego raportu (znikają po wysłaniu wszystkich jego zaległych plików), ustawia status `Błąd` z opisem w atrybucie `blad` i zwiększa licznik `nieudane_proby`. Niedostępny folder i problemy z szablonem mają osobne wpisy w Naprawach.

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
| `ha_mailer.reload` | wczytuje ponownie blok `ha_mailer` z `configuration.yaml` bez restartu |
| `ha_mailer.send_now` | natychmiast tworzy plik i wysyła go (poza harmonogramem), bez ponawiania; błąd wraca do wywołującego |
| `ha_mailer.clear_pending` | porzuca pliki czekające w kolejce (zostają w folderze do końca retencji) |

`send_now` i `clear_pending` mają opcjonalne pola `config` (`id` bloku) i `report` (nazwa lub `id` raportu). Puste pola oznaczają wszystkie. Test: **Narzędzia deweloperskie → Akcje**.

## Przykłady harmonogramu

| Wyrażenie | Kiedy |
|---|---|
| `0 8 * * *` | codziennie o 8:00 |
| `0 7 * * 1` | w poniedziałki o 7:00 |
| `0 18 1 * *` | pierwszego dnia miesiąca o 18:00 |
| `0 6,18 * * 1-5` | pn-pt o 6:00 i 18:00 |
| `*/30 * * * *` | co 30 minut |

Format: `minuta godzina dzień-miesiąca miesiąc dzień-tygodnia`, strefa czasowa jak w HA.

## Aktualizacja z wersji z konfiguracją w interfejsie (do 2.x)

Wersja 3.0 usuwa konfigurację przez interfejs. Przy pierwszym starcie:

- wpisy utworzone w interfejsie są **automatycznie usuwane**, a w powiadomieniu HA dostajesz **gotowy odpowiednik w YAML** (hasło zastąpione odwołaniem `!secret ha_mailer_gmail_password`, które zapisz w `secrets.yaml`),
- statystyki i kolejka wysyłki usuniętych wpisów są wyzerowane, a pliki w folderze raportów zostają nietknięte,
- nowe wpisy z YAML tworzą nowe encje (identyfikatory mogą się różnić),
- wpisy utworzone wcześniej z YAML w formacie 1.2 (płaski blok) i 2.x nadal działają, a starsze wersje są odtwarzane na nowo z YAML.

Stary, płaski format YAML (jeden raport na poziomie bloku, z kluczami `cron`, `recipient`, `rows`) nadal jest akceptowany i traktowany jako jeden raport o nazwie „Raport”.

Wersja 3.1 zmienia tekst w komórce encji bez wartości: zamiast surowego `unavailable` lub `unknown` pojawia się `brak danych` (`no data` po angielsku), a za brak wartości uznawane są też `none`, pusty tekst i puste atrybuty. Domyślny tryb `warning` zachowuje czerwone oznaczenie i uwagę w mailu.

## Rozwiązywanie problemów

| Objaw | Przyczyna / rozwiązanie |
|---|---|
| „Invalid config for [ha_mailer]” przy starcie | Błąd składni w YAML. Komunikat wskazuje klucz. Sprawdź *Narzędzia deweloperskie → YAML → Sprawdź konfigurację*. |
| „Logowanie do Gmaila nie powiodło się” | Użyto zwykłego hasła zamiast hasła aplikacji albo brak weryfikacji dwuetapowej. |
| Naprawa „folder na raporty jest niedostępny” | Dysk lub udział sieciowy nie jest dostępny albo brakuje uprawnień. Sprawdź `storage_path`. Do czasu naprawy raporty nie są archiwizowane. |
| Naprawa „problem z szablonem” | Zła ścieżka `template`, uszkodzony plik lub nieznany znacznik. Popraw i przeładuj YAML. |
| „Nie można połączyć się z smtp.gmail.com” | Brak internetu lub zablokowany port 465. |
| Kolejka oczekujących ciągle rośnie | Wysyłka stale się nie udaje. Sprawdź atrybut `blad` statusu i Naprawy, popraw dane i ewentualnie użyj `clear_pending`. |
| „Wiersze bez encji wymagają kolumny 'opis'” | Wiersz z samym `description` wymaga kolumny `opis` w `columns`. |
| W komórce jest `brak danych` zamiast wartości | Encja ma stan `unavailable`, `unknown`, `none` lub pusty (albo pusty atrybut). Sprawdź jej stan w HA. Sposób oznaczenia zmienisz opcją `no_value`. |
| Wartość atrybutu nie pojawia się w pliku | Wartość jest w kolumnie „Wartość”. Dodaj kolumnę `atrybut`, jeśli chcesz widzieć jego nazwę. |
| Raport nie powstaje o czasie | Sprawdź przełącznik „Wysyłka z harmonogramu” i atrybut `nastepna_wysylka`. |
| Zmiana w YAML nie działa | Wywołaj `ha_mailer.reload` lub zrestartuj HA. |
| Mail nie przychodzi, status `OK` | Sprawdź spam i adres odbiorcy. |

## Ograniczenia

- Wysyłki, które wypadły w czasie wyłączenia HA, nie są tworzone wstecz, ale kolejka wysyłki działa po restarcie.
- Hasło aplikacji jest w `configuration.yaml` (użyj `!secret`) oraz w konfiguracji HA (`/config/.storage`).
- Obowiązują limity Gmaila (załącznik do 25 MB, dzienny limit wysyłek).
- Tekst wyglądający na liczbę (np. wersja `1.10`) jest zapisywany jako liczba.
- Szablony: tylko `.xlsx`, bez makr, wykresów, obrazów i tabel przestawnych; formuły nie są przeliczane przez dodatek.
- Komunikaty i opisy są po polsku.

## Struktura repozytorium

```
hacs.json
README.md
examples/szablon_przyklad.xlsx   # przykładowy szablon Excel
custom_components/ha_mailer/
├── __init__.py        # logika: raporty, harmonogramy, Excel, szablony, kolejka, archiwum, YAML
├── config_flow.py     # tylko import wpisów z YAML (bez formularzy)
├── sensor.py          # encje: status, ostatnia wysyłka, statystyki, kolejka
├── switch.py          # przełącznik harmonogramu raportu
├── services.yaml      # opis akcji
├── manifest.json
└── translations/pl.json
```
