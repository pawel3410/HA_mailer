# HA_mailer

Integracja HACS dla Home Assistanta. W wybranym terminie (wyrażenie **cron**) robi migawkę wartości wskazanych encji, zapisuje ją w pliku **Excel (.xlsx)** i wysyła mailem z konta **Gmail** na podany adres.

## Funkcje

- dowolna liczba encji w jednym raporcie i dowolny harmonogram (cron),
- **konfigurowalny układ pliku**: wybierasz, które kolumny się w nim znajdą (opis, encja, nazwa, wartość, jednostka, ostatnia zmiana, lp.) i w jakiej kolejności,
- **wiersze z samym opisem** (bez encji), np. nagłówki sekcji i separatory,
- stała pozycja: każdy wiersz zawsze w tym samym miejscu pliku,
- edycja wierszy **graficznie** (menu w interfejsie) lub **tekstowo** (YAML w interfejsie),
- konfiguracja także przez **configuration.yaml** (import do wpisu w UI),
- tytuł maila: stała wartość + data i godzina wysyłki,
- jawne błędy wysyłki: log, powiadomienie w HA, status w encji,
- trzy encje: status wysyłki, ostatnia wysyłka, statystyki (przetrwają restart),
- akcja `ha_mailer.send_now` do natychmiastowej wysyłki.

## Wymagania

- Home Assistant 2024.8 lub nowszy,
- konto Gmail z włączoną weryfikacją dwuetapową i **hasłem aplikacji**,
- dostęp HA do internetu (`smtp.gmail.com`, port 465).

Biblioteki `openpyxl` i `croniter` Home Assistant instaluje sam.

## Instalacja

### Przez HACS

1. W HACS wybierz **⋮ → Własne repozytoria** (Custom repositories).
2. Wklej adres tego repozytorium, kategoria **Integracja**.
3. Zainstaluj **HA_mailer** i **zrestartuj Home Assistanta**.
4. **Ustawienia → Urządzenia i usługi → Dodaj integrację → HA_mailer** (albo skonfiguruj przez YAML, patrz niżej).

### Ręcznie

Skopiuj folder `custom_components/ha_mailer` do `/config/custom_components/` i zrestartuj HA.

## Hasło aplikacji Gmail

Zwykłe hasło **nie zadziała**. Włącz weryfikację dwuetapową, a potem utwórz hasło aplikacji: https://myaccount.google.com/apppasswords. Spacje w haśle są ignorowane.

## Konfiguracja w interfejsie

**Dodawanie integracji.** Podajesz encje, harmonogram, dane Gmaila, odbiorcę, tytuł i kolumny. Przy zapisie sprawdzana jest poprawność danych oraz logowanie do Gmaila. Z wybranych encji powstają wiersze, a szczegóły (opisy, wiersze bez encji) ustawisz po dodaniu integracji.

**Zmiany.** Przycisk **Konfiguruj** przy integracji otwiera menu:

| Pozycja | Co robi |
|---|---|
| Ustawienia | Gmail (puste hasło = bez zmian), harmonogram, tytuł, kolumny |
| Wiersze - edytor graficzny | dodawanie, edycja, usuwanie i przesuwanie wierszy |
| Kolumny i wiersze - edycja tekstowa | cały układ jako YAML w jednym polu |

### Edytor graficzny

Menu pokazuje aktualną listę wierszy i pozwala: **dodać wiersz z encją** (z opcjonalnym opisem), **dodać wiersz z samym opisem**, **edytować**, **usunąć**, **przesunąć** (na wskazaną pozycję) oraz **zapisać**. Nowy wiersz można wstawić na wybranej pozycji, a bez niej trafia na koniec. Zmiany zapisują się dopiero po wybraniu „Zapisz i zamknij”, zamknięcie okna je odrzuca.

Jeśli dodasz wiersz z samym opisem, a układ nie ma kolumny „Opis”, zostanie ona dodana automatycznie.

### Edycja tekstowa

```yaml
columns:
  - lp
  - opis
  - encja
  - nazwa
  - wartosc
  - jednostka
  - zmiana
rows:
  - description: Salon            # wiersz z samym opisem
  - entity: sensor.salon_temperatura
    description: Temperatura
  - entity: sensor.salon_wilgotnosc
  - description: Kuchnia
  - entity: sensor.kuchnia_temperatura
```

Przy zapisie układ jest walidowany, a błąd pokazuje konkretny powód.

## Układ pliku Excel

**Kolumny** (klucz → nagłówek): `lp` → Lp., `opis` → Opis, `encja` → Encja, `nazwa` → Nazwa, `wartosc` → Wartość, `jednostka` → Jednostka, `zmiana` → Ostatnia zmiana.

- Kolumny występują w pliku dokładnie w podanej kolejności. Kolumny, której nie wybrano, w pliku nie ma, więc encję lub nazwę można pominąć.
- Kolumny nie mogą się powtarzać.
- „Lp.” numeruje tylko wiersze z encjami.

**Wiersze** (arkusz „Migawka”):

| Wiersz arkusza | Zawartość |
|---|---|
| 1 | „Czas migawki:” i data wysyłki |
| 3 | nagłówki wybranych kolumn |
| 4 i dalej | wiersze z konfiguracji, po kolei |

- **Wiersz z encją** ma wartości z jej stanu w momencie wysyłki. Kolumna „Opis” zawiera Twój opis (może być pusty). Liczby zapisywane są jako liczby, teksty zawsze jako tekst. Brakująca encja zostaje na swoim miejscu z napisem `brak encji`.
- **Wiersz z samym opisem** ma tylko tekst w kolumnie „Opis”, pogrubiony na jasnoniebieskim tle (nagłówek lub separator).
- Pozycja wiersza w pliku odpowiada jego pozycji na liście. Usunięcie wiersza przesuwa kolejne w górę.

## Konfiguracja przez configuration.yaml

```yaml
ha_mailer:
  id: glowny                       # opcjonalne, wymagane przy kilku blokach
  cron: "0 8 * * *"
  gmail_user: moj.adres@gmail.com
  gmail_app_password: !secret gmail_app_password
  recipient: odbiorca@example.com  # lub kilka po przecinku
  subject: "HA_mailer - raport"    # opcjonalne
  columns: [lp, opis, encja, nazwa, wartosc, jednostka, zmiana]   # opcjonalne
  rows:
    - description: Salon
    - entity: sensor.salon_temperatura
      description: Temperatura
    - entity: sensor.salon_wilgotnosc
    - description: Kuchnia
    - entity: sensor.kuchnia_temperatura
```

Można podać listę kilku bloków (każdy z własnym `id`). Po restarcie HA dla każdego bloku powstaje wpis widoczny w interfejsie (**HA_mailer (YAML)**).

**Jak działa synchronizacja YAML ↔ interfejs:**

- YAML jest importowany przy starcie Home Assistanta, a wpis można potem edytować w interfejsie.
- Zmiany z interfejsu **zostają**, dopóki nie zmienisz treści bloku w YAML.
- Gdy zmienisz blok w YAML i zrestartujesz HA, wpis jest **w całości nadpisywany** zawartością YAML, także zmiany zrobione wcześniej w interfejsie.
- Usunięcie bloku z YAML nie usuwa wpisu. Usuń go ręcznie w interfejsie.
- Przy importie nie sprawdza się logowania do Gmaila (HA może startować bez internetu). Błąd pojawi się przy pierwszej wysyłce.
- Błędy w YAML (zły cron, nieistniejąca kolumna, wiersz bez encji i bez opisu) zgłasza Home Assistant przy sprawdzaniu konfiguracji.

## Encje

Urządzenie **HA_mailer** ma trzy sensory (identyfikatory przykładowe):

| Encja | Stan | Atrybuty |
|---|---|---|
| `sensor.ha_mailer_status_wysylki` | `OK` / `Błąd` / `Brak danych` | `blad`, `ostatnia_proba`, `odbiorca` |
| `sensor.ha_mailer_ostatnia_wysylka` | czas ostatniej **udanej** wysyłki | `ostatnia_proba` |
| `sensor.ha_mailer_statystyki` | liczba udanych wysyłek | `udane`, `nieudane`, `razem`, `skutecznosc_proc`, `ostatni_blad`, `czas_ostatniego_bledu`, `rozmiar_ostatniego_pliku_kb`, `nastepna_wysylka`, `liczba_encji`, `liczba_wierszy` |

Statystyki są zapisywane w `/config/.storage/ha_mailer.<entry_id>`.

## Obsługa błędów

Gdy wysyłka się nie uda, HA_mailer zapisuje błąd w logu (filtr `ha_mailer`), tworzy powiadomienie w panelu HA (znika po udanej wysyłce), ustawia status `Błąd` z opisem w atrybucie `blad` i zwiększa licznik `nieudane`. Akcja `ha_mailer.send_now` też kończy się komunikatem.

```yaml
automation:
  - alias: "HA_mailer - błąd wysyłki"
    triggers:
      - trigger: state
        entity_id: sensor.ha_mailer_status_wysylki
        to: "Błąd"
    actions:
      - action: notify.notify
        data:
          title: "HA_mailer"
          message: "{{ state_attr('sensor.ha_mailer_status_wysylki', 'blad') }}"
```

## Akcja `ha_mailer.send_now`

Natychmiast tworzy raport i wysyła go, niezależnie od harmonogramu. Pole **Konfiguracja** wskazuje wpis (puste = wszystkie). Test: **Narzędzia deweloperskie → Akcje**.

## Przykłady harmonogramu

| Wyrażenie | Kiedy |
|---|---|
| `0 8 * * *` | codziennie o 8:00 |
| `0 7 * * 1` | w poniedziałki o 7:00 |
| `0 18 1 * *` | pierwszego dnia miesiąca o 18:00 |
| `0 6,18 * * 1-5` | pn-pt o 6:00 i 18:00 |
| `*/30 * * * *` | co 30 minut |

Format: `minuta godzina dzień-miesiąca miesiąc dzień-tygodnia`, strefa czasowa jak w HA.

## Aktualizacja z wersji 1.1

Istniejące wpisy są migrowane automatycznie: lista encji staje się listą wierszy, a układ pliku pozostaje taki jak dotąd (bez kolumny „Opis”).

## Rozwiązywanie problemów

| Objaw | Przyczyna / rozwiązanie |
|---|---|
| „Logowanie do Gmaila nie powiodło się” | Użyto zwykłego hasła zamiast hasła aplikacji albo brak weryfikacji dwuetapowej. |
| „Nie można połączyć się z smtp.gmail.com” | Brak internetu lub zablokowany port 465. |
| „Masz wiersze bez encji - dodaj kolumnę Opis” | Wiersze z samym opisem wymagają kolumny „Opis” w układzie. |
| Mail nie przychodzi, status `OK` | Sprawdź spam i adres odbiorcy. |
| Zmiany z interfejsu zniknęły po restarcie | Zmieniono blok `ha_mailer` w YAML, co nadpisuje wpis. |
| Brak integracji po instalacji | Zrestartuj Home Assistanta. |

## Ograniczenia

- Wysyłki, które wypadły w czasie wyłączenia HA, nie są nadrabiane.
- Hasło aplikacji jest przechowywane w konfiguracji HA (`/config/.storage`).
- Obowiązują limity Gmaila (załącznik do 25 MB, dzienny limit wysyłek).
- Stan encji jest odczytywany w momencie wysyłki, dane nie są logowane w czasie.
- Interfejs konfiguracji jest po polsku.

## Struktura repozytorium

```
hacs.json
README.md
custom_components/ha_mailer/
├── __init__.py        # logika, harmonogram, Excel, wysyłka, statystyki, import YAML
├── config_flow.py     # konfiguracja, opcje, edytor graficzny i tekstowy
├── sensor.py          # encje: status, ostatnia wysyłka, statystyki
├── services.yaml      # opis akcji send_now
├── manifest.json
└── translations/pl.json
```
