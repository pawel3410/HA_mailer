# HA_mailer

Integracja HACS dla Home Assistanta. W wybranym terminie (wyrażenie **cron**) robi migawkę wartości wskazanych encji, zapisuje ją w pliku **Excel (.xlsx)** i wysyła mailem z konta **Gmail** na podany adres.

## Funkcje

- dowolna liczba encji w jednym raporcie,
- dowolny harmonogram wysyłki (wyrażenie cron),
- plik Excel o stałym układzie: każda encja zawsze w tym samym wierszu,
- tytuł maila: stała wartość + data i godzina wysyłki,
- jawne błędy wysyłki: log, powiadomienie w HA, status w encji,
- trzy encje: status wysyłki, ostatnia wysyłka, statystyki (zapisywane na dysku, przetrwają restart),
- akcja `ha_mailer.send_now` do natychmiastowej wysyłki (np. na test),
- konfiguracja w całości z poziomu interfejsu HA.

## Wymagania

- Home Assistant 2024.8 lub nowszy,
- konto Gmail z włączoną weryfikacją dwuetapową i **hasłem aplikacji**,
- dostęp HA do internetu (serwer `smtp.gmail.com`, port 465).

Biblioteki `openpyxl` i `croniter` Home Assistant instaluje sam przy pierwszym uruchomieniu integracji.

## Instalacja

### Przez HACS

1. W HACS wybierz menu **⋮ → Własne repozytoria** (Custom repositories).
2. Wklej adres tego repozytorium, jako kategorię wybierz **Integracja** (Integration).
3. Znajdź na liście **HA_mailer**, zainstaluj i **zrestartuj Home Assistanta**.
4. Wejdź w **Ustawienia → Urządzenia i usługi → Dodaj integrację** i wyszukaj **HA_mailer**.

### Ręcznie

Skopiuj folder `custom_components/ha_mailer` do katalogu `/config/custom_components/`, zrestartuj Home Assistanta i dodaj integrację jak w punkcie 4 powyżej.

## Hasło aplikacji Gmail

Zwykłe hasło do Gmaila **nie zadziała**.

1. Włącz weryfikację dwuetapową na koncie Google.
2. Otwórz https://myaccount.google.com/apppasswords i utwórz hasło aplikacji (np. o nazwie „HA_mailer”).
3. Wklej 16-znakowe hasło do konfiguracji. Spacje możesz zostawić, zostaną usunięte automatycznie.

## Konfiguracja

| Pole | Opis |
|---|---|
| Encje | Encje do raportu. Kolejność wyboru = kolejność wierszy w Excelu. |
| Harmonogram (cron) | Wyrażenie cron z 5 polami, liczone wg strefy czasowej HA. |
| Konto Gmail | Adres nadawcy. |
| Hasło aplikacji Gmail | Hasło aplikacji (patrz wyżej). |
| Adres odbiorcy | Jeden lub kilka adresów oddzielonych przecinkiem. |
| Stały tytuł maila | Do tytułu zawsze dodawana jest data i godzina wysyłki. |

Przy zapisie sprawdzana jest poprawność cron, adresów oraz **logowanie do Gmaila**, więc błędne dane wychodzą od razu. Ustawienia zmienisz później przyciskiem **Konfiguruj** przy integracji (puste pole hasła = bez zmian).

### Przykłady harmonogramu

| Wyrażenie | Kiedy |
|---|---|
| `0 8 * * *` | codziennie o 8:00 |
| `0 7 * * 1` | w poniedziałki o 7:00 |
| `0 18 1 * *` | pierwszego dnia miesiąca o 18:00 |
| `0 6,18 * * 1-5` | od poniedziałku do piątku o 6:00 i 18:00 |
| `*/30 * * * *` | co 30 minut |

Format: `minuta godzina dzień-miesiąca miesiąc dzień-tygodnia`.

## Mail i plik Excel

Przykładowy tytuł: `HA_mailer - raport 2026-11-03 08:00:00`.
Załącznik: `HA_mailer_2026-11-03_08-00-00.xlsx`.

Arkusz **Migawka** ma stały układ:

| Wiersz | Zawartość |
|---|---|
| 1 | „Czas migawki:” i data wysyłki |
| 3 | nagłówki: Lp., Encja, Nazwa, Wartość, Jednostka, Ostatnia zmiana |
| 4 i dalej | po jednej encji na wiersz, w kolejności z konfiguracji |

- Wartości liczbowe są zapisywane jako liczby, więc można na nich liczyć w Excelu.
- Jeśli encja nie istnieje, jej wiersz zostaje, a w kolumnie „Wartość” jest `brak encji`.
- W opcjach istniejące encje zostają na swoich miejscach, a nowe są dopisywane na końcu. Usunięcie encji z konfiguracji przesuwa kolejne wiersze w górę.

## Encje

Integracja tworzy urządzenie **HA_mailer** z trzema sensorami (identyfikatory przykładowe, mogą się różnić):

| Encja | Stan | Atrybuty |
|---|---|---|
| `sensor.ha_mailer_status_wysylki` | `OK` / `Błąd` / `Brak danych` | `blad`, `ostatnia_proba`, `odbiorca` |
| `sensor.ha_mailer_ostatnia_wysylka` | czas ostatniej **udanej** wysyłki | `ostatnia_proba` |
| `sensor.ha_mailer_statystyki` | liczba udanych wysyłek | `udane`, `nieudane`, `razem`, `skutecznosc_proc`, `ostatni_blad`, `czas_ostatniego_bledu`, `rozmiar_ostatniego_pliku_kb`, `nastepna_wysylka`, `liczba_encji` |

Statystyki są zapisywane w `/config/.storage/ha_mailer.<entry_id>` i usuwane razem z integracją.

## Obsługa błędów

Gdy wysyłka się nie uda (np. złe hasło, brak internetu, odrzucony adres), HA_mailer:

1. zapisuje błąd w logu (`Ustawienia → System → Logi`, filtr `ha_mailer`),
2. tworzy powiadomienie w panelu HA, które znika po następnej udanej wysyłce,
3. ustawia status `Błąd` z opisem w atrybucie `blad`,
4. zwiększa licznik `nieudane` w statystykach.

Akcja `ha_mailer.send_now` w razie błędu również kończy się komunikatem.

### Przykład: powiadomienie na telefon o błędzie

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

Natychmiast tworzy raport i wysyła go, niezależnie od harmonogramu. Pole **Konfiguracja** pozwala wskazać wpis HA_mailer (puste = wszystkie). Test wykonasz w **Narzędzia deweloperskie → Akcje**.

```yaml
action: ha_mailer.send_now
```

## Rozwiązywanie problemów

| Objaw | Przyczyna / rozwiązanie |
|---|---|
| „Logowanie do Gmaila nie powiodło się" | Użyto zwykłego hasła zamiast hasła aplikacji albo brak weryfikacji dwuetapowej. |
| „Nie można połączyć się z smtp.gmail.com" | Brak internetu lub zablokowany port 465 w sieci/firewallu. |
| Mail nie przychodzi, status `OK` | Sprawdź folder spam oraz adres odbiorcy. |
| Brak integracji na liście po instalacji | Zrestartuj Home Assistanta. |
| Wysyłka nie nastąpiła o planowanej porze | Sprawdź atrybut `nastepna_wysylka` w statystykach i strefę czasową HA (*Ustawienia → System → Ogólne*). |

## Ograniczenia i uwagi

- Wysyłki, które wypadły w czasie wyłączenia HA, nie są nadrabiane. Po starcie liczony jest kolejny termin z harmonogramu.
- Hasło aplikacji jest przechowywane w konfiguracji Home Assistanta (`/config/.storage`), tak jak dane logowania innych integracji.
- Obowiązują limity Gmaila (m.in. rozmiar załącznika do 25 MB i dzienny limit wysyłek).
- Stan encji jest odczytywany w momencie wysyłki, dane nie są logowane w czasie.
- Interfejs konfiguracji jest dostępny w języku polskim.

## Struktura repozytorium

```
hacs.json
README.md
custom_components/ha_mailer/
├── __init__.py        # logika, harmonogram, Excel, wysyłka, statystyki
├── config_flow.py     # konfiguracja i opcje w interfejsie
├── sensor.py          # encje: status, ostatnia wysyłka, statystyki
├── services.yaml      # opis akcji send_now
├── manifest.json
└── translations/pl.json
```
