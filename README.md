# HA_mailer

Integracja HACS dla Home Assistanta. W wybranych terminach (wyrażenia **cron**) robi migawki wartości encji, zapisuje je w plikach **Excel (.xlsx)** i wysyła mailem z jednego konta **Gmail** na wskazane adresy. Jedna integracja obsługuje **wiele raportów**, a każdy z nich ma własny plik, harmonogram, odbiorców, tytuł, kolumny i wiersze.

## Funkcje

- **wiele raportów** w jednej integracji (wspólne konto Gmail),
- **konfigurowalny układ pliku** dla każdego raportu: wybór i kolejność kolumn, wiersze z encją lub z samym opisem,
- **ponawianie nieudanych wysyłek** co zadany czas, tym samym plikiem, który powstał zgodnie z harmonogramem,
- **archiwum plików** w wybranym folderze, z automatycznym usuwaniem po zadanym czasie,
- edycja wierszy graficznie lub tekstowo (YAML), konfiguracja także przez `configuration.yaml`,
- jawne błędy: log, powiadomienie w HA, status w encjach,
- encje: status, ostatnia wysyłka i statystyki dla każdego raportu oraz licznik oczekujących wysyłek,
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
| Folder na pliki Excel | wspólny folder dla wszystkich raportów, domyślnie `/config/ha_mailer` (ścieżka względna liczona od `/config`) |

2. **Pierwszy raport**: nazwa, harmonogram (cron), odbiorca, stały tytuł, kolumny i encje.

**Zmiany**: przycisk **Konfiguruj** otwiera listę raportów z menu: *Ustawienia wspólne*, *Dodaj raport*, *Edytuj raport*, *Usuń raporty* oraz *Zapisz i zamknij*. Zmiany zapisują się dopiero po „Zapisz i zamknij”, zamknięcie okna je odrzuca.

Po wybraniu raportu do edycji dostępne jest menu:

| Pozycja | Co robi |
|---|---|
| Ustawienia raportu | nazwa, harmonogram, odbiorca (kilka po przecinku), tytuł, kolumny |
| Wiersze - edytor graficzny | dodaj wiersz z encją, dodaj wiersz z samym opisem, edytuj, usuń, przesuń na wskazaną pozycję |
| Kolumny i wiersze - edycja tekstowa | cały układ raportu jako YAML w jednym polu |
| Wróć do listy raportów | |

W edytorze graficznym zmiany trafiają do raportu po „Zapisz wiersze”, a trwale zapisują się po „Zapisz i zamknij” na liście raportów. Jeśli dodasz wiersz z samym opisem, a układ nie ma kolumny „Opis”, zostanie ona dodana automatycznie.

Przykład edycji tekstowej:

```yaml
columns: [lp, opis, encja, nazwa, wartosc, jednostka, zmiana]
rows:
  - description: Salon            # wiersz z samym opisem
  - entity: sensor.salon_temperatura
    description: Temperatura
  - entity: sensor.salon_wilgotnosc
```

## Układ pliku Excel

**Kolumny** (klucz → nagłówek): `lp` → Lp., `opis` → Opis, `encja` → Encja, `nazwa` → Nazwa, `wartosc` → Wartość, `jednostka` → Jednostka, `zmiana` → Ostatnia zmiana. Kolumny występują w podanej kolejności, a niewybrana kolumna nie trafia do pliku. Kolumny nie mogą się powtarzać, a „Lp.” numeruje tylko wiersze z encjami.

**Wiersze** (arkusz „Migawka”): w wierszu 1 jest czas migawki, w wierszu 3 nagłówki, a od wiersza 4 kolejne wiersze raportu.

- Wiersz z encją ma wartości z jej stanu w chwili tworzenia pliku. Liczby są zapisywane jako liczby, a teksty zawsze jako tekst. Brakująca encja zostaje na miejscu z napisem `brak encji`.
- Wiersz z samym opisem ma tekst w kolumnie „Opis”, pogrubiony na jasnoniebieskim tle (nagłówek lub separator).
- Pozycja wiersza w pliku odpowiada pozycji na liście. Usunięcie wiersza przesuwa kolejne w górę.

Tytuł maila to stały tytuł raportu plus data i godzina utworzenia pliku, np. `HA_mailer - raport 2026-11-03 08:00:00`.

## Zapis plików i retencja

Każdy raport tworzy plik `HA_mailer_<id raportu>_<data>_<godzina>.xlsx` we wspólnym folderze. Plik jest zapisywany **zawsze**, także gdy wysyłka się nie powiedzie.

- Usuwane są tylko pliki utworzone przez HA_mailer (dodatek prowadzi ich indeks), więc inne pliki w tym samym folderze są bezpieczne.
- Pliki starsze niż ustawiona liczba dni są usuwane co godzinę i przy starcie HA.
- Pliki czekające na ponowienie wysyłki **nie są usuwane**, dopóki nie zostaną wysłane lub porzucone.
- Zmiana folderu nie przenosi starych plików, ale będą one nadal usuwane po upływie retencji.

## Ponawianie nieudanych wysyłek

Jeśli wysyłka z harmonogramu się nie uda:

1. plik zostaje zapisany, a błąd jest widoczny (log, powiadomienie, status),
2. ten sam plik trafia do kolejki i jest wysyłany ponownie **co ustawiony interwał, aż do skutku**,
3. ponowiony mail ma ten sam tytuł i załącznik (stan encji z chwili harmonogramu), a w treści adnotację o opóźnieniu,
4. jeśli w międzyczasie nadejdzie kolejny termin, nowy plik **dołącza do kolejki**, a po powrocie łączności zostaną wysłane wszystkie zaległe pliki,
5. kolejka jest zapisywana na dysku, więc przetrwa restart HA.

Ponawianie nie ma limitu prób, np. zły adres lub hasło będą ponawiane bez końca. Gdy chcesz je przerwać, użyj akcji `ha_mailer.clear_pending` (porzuca kolejkę), a jeśli przyczyną były błędne dane, popraw je w ustawieniach, bo kolejka wyśle się przy najbliższej próbie. Wysyłka ręczna (`send_now`) nie trafia do kolejki, a jej błąd widzisz od razu w odpowiedzi akcji.

## Konfiguracja przez configuration.yaml

```yaml
ha_mailer:
  id: glowny                       # opcjonalne, wymagane przy kilku blokach
  gmail_user: moj.adres@gmail.com
  gmail_app_password: !secret gmail_app_password
  retry_minutes: 10                # opcjonalne, domyślnie 10
  retention_days: 30               # opcjonalne, domyślnie 30 (0 = nie usuwaj)
  storage_path: /config/ha_mailer  # opcjonalne, domyślnie /config/ha_mailer
  reports:
    - name: Raport dzienny
      cron: "0 8 * * *"
      recipient: biuro@example.com
      subject: "HA_mailer - raport dzienny"    # opcjonalne
      columns: [lp, opis, encja, nazwa, wartosc, jednostka, zmiana]   # opcjonalne
      rows:
        - description: Salon
        - entity: sensor.salon_temperatura
          description: Temperatura
        - entity: sensor.salon_wilgotnosc
    - name: Raport tygodniowy
      id: tygodniowy                # opcjonalne, domyślnie z nazwy
      cron: "0 7 * * 1"
      recipient: "szef@example.com, ksiegowosc@example.com"
      columns: [encja, wartosc]
      rows:
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

Dla każdego raportu powstaje osobne urządzenie (o nazwie raportu) z trzema sensorami, a dodatkowo urządzenie **HA_mailer** z licznikiem kolejki. Identyfikatory encji pochodzą z nazw, więc mogą się różnić od przykładów.

| Encja | Stan | Atrybuty |
|---|---|---|
| Status wysyłki | `OK` / `Błąd` / `Brak danych` | `blad`, `ostatnia_proba`, `oczekujace_ponowienia`, `odbiorca` |
| Ostatnia wysyłka | czas ostatniej **udanej** wysyłki | `ostatnia_proba` |
| Statystyki | liczba udanych wysyłek | `udane`, `udane_po_ponowieniu`, `nieudane_proby`, `razem_prob`, `skutecznosc_prob_proc`, `ostatni_blad`, `czas_ostatniego_bledu`, `rozmiar_ostatniego_pliku_kb`, `nastepna_wysylka`, `liczba_encji`, `liczba_wierszy` |
| Oczekujące wysyłki (urządzenie HA_mailer) | liczba plików w kolejce | `kolejka` (raport, plik, liczba prób, następna próba, błąd), `katalog_plikow`, `zapisane_pliki`, `retencja_dni` |

`nieudane_proby` liczy każdą nieudaną próbę, także ponowienia. Dane są zapisywane w `/config/.storage/ha_mailer.<entry_id>`. Encje usuniętych raportów są usuwane automatycznie.

## Obsługa błędów

Każda nieudana próba: trafia do logu (filtr `ha_mailer`), tworzy powiadomienie w panelu HA dla danego raportu (znika po wysłaniu wszystkich jego zaległych plików), ustawia status `Błąd` z opisem w atrybucie `blad` i zwiększa licznik `nieudane_proby`.

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
| `ha_mailer.send_now` | natychmiast tworzy plik i wysyła go (poza harmonogramem), bez ponawiania |
| `ha_mailer.clear_pending` | porzuca oczekujące ponowienia (pliki zostają do końca retencji) |

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

Istniejące wpisy są migrowane automatycznie: dotychczasowa konfiguracja staje się raportem „Raport”, a ustawienia wspólne dostają wartości domyślne (ponawianie co 10 min, retencja 30 dni, folder `/config/ha_mailer`). Encje, statystyki i układ pliku zostają zachowane.

## Rozwiązywanie problemów

| Objaw | Przyczyna / rozwiązanie |
|---|---|
| „Logowanie do Gmaila nie powiodło się” | Użyto zwykłego hasła zamiast hasła aplikacji albo brak weryfikacji dwuetapowej. |
| „Nie można utworzyć folderu lub zapisać w nim plików” | Sprawdź ścieżkę i uprawnienia (na HA OS użyj np. `/config/...`, `/media/...` lub `/share/...`). |
| „Nie można połączyć się z smtp.gmail.com” | Brak internetu lub zablokowany port 465. |
| Kolejka oczekujących ciągle rośnie | Wysyłka stale się nie udaje. Sprawdź atrybut `blad` statusu, popraw dane i ewentualnie użyj `clear_pending`. |
| „Masz wiersze bez encji - dodaj kolumnę Opis” | Wiersze z samym opisem wymagają kolumny „Opis” w układzie. |
| Zmiany z interfejsu zniknęły po restarcie | Zmieniono blok `ha_mailer` w YAML, co nadpisuje wpis. |
| Mail nie przychodzi, status `OK` | Sprawdź spam i adres odbiorcy. |

## Ograniczenia

- Wysyłki, które wypadły w czasie wyłączenia HA, nie są tworzone wstecz, ale kolejka ponawiania działa po restarcie.
- Hasło aplikacji jest przechowywane w konfiguracji HA (`/config/.storage`).
- Obowiązują limity Gmaila (załącznik do 25 MB, dzienny limit wysyłek).
- Interfejs konfiguracji jest po polsku.

## Struktura repozytorium

```
hacs.json
README.md
custom_components/ha_mailer/
├── __init__.py        # logika: raporty, harmonogramy, Excel, kolejka, archiwum, import YAML
├── config_flow.py     # konfiguracja, opcje, edytor graficzny i tekstowy
├── sensor.py          # encje: status, ostatnia wysyłka, statystyki, kolejka
├── services.yaml      # opis akcji
├── manifest.json
└── translations/pl.json
```
