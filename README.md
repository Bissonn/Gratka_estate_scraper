# Scraper Gratka.pl → Excel Dashboard
### *Automatyzacja Selekcji Ofert Nieruchomości High-End*

System zaprojektowany w celu wyeliminowania szumu informacyjnego z rynku nieruchomości. Narzędzie automatycznie przeszukuje serwis Gratka.pl, filtruje oferty i generuje estetyczny, w pełni funkcjonalny dashboard analityczny w formacie `.xlsx`.

## Możliwości Systemu
*   **Deep Scraping:** Pobieranie pełnych danych (cena, metraż, rok budowy, lokalizacja) z pominięciem duplikatów i ofert bez cen
*   **Analiza Typów:** Obsługa mieszkań, domów, kawalerek, działek oraz lokali użytkowych
*   **Excel Dashboard:** Generowanie zaawansowanego pliku Excel z automatycznymi wykresami (słupkowymi i kołowymi), kaflami KPI oraz formatowaniem warunkowym (skala kolorów dla cen za $m^2$)
*   **Intuicyjne GUI:** Interfejs oparty na Tkinter z logowaniem zdarzeń w czasie rzeczywistym i paskiem postępu
## Architektura Techniczna
Projekt oparty jest na architekturze modułowej:
1.  `scraper_nieruchomosci.py`: Silnik parsujący oparty na `BeautifulSoup4` i `Requests`
2.  `excel_generator.py`: Zaawansowany moduł tworzący warstwę wizualną raportu przy użyciu `openpyxl`
3.  `gui.py`: Front-end aplikacji umożliwiający sterowanie filtrami bez znajomości kodu

## Instalacja i Uruchomienie

### Wymagania:
Zalecany Python 3.10+ oraz biblioteki wymienione w `requirements.txt`
```bash
pip install -r requirements.txt
```

### Start:
Aby uruchomić aplikację z interfejsem graficznym:
```bash
python gui.py
```

## Co znajdziesz w raporcie?
*   **Arkusz Dashboard:** Średnie ceny za $m^2$ wg miast, udział % typów ofert i kluczowe metryki rynkowe
*   **Arkusz Dane:** Kompletna tabela z aktywnymi linkami do ofert i parametrami technicznymi
*   **Arkusz Analiza:** Rankingi TOP 10 ofert z najkorzystniejszą ceną za $m^2$

---
*Status projektu: Gotowy do wdrożenia biznesowego.*
*Autor: Damian Bisewski*

---
