from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re
import requests
from bs4 import BeautifulSoup

from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

#MODUŁY
# from claude_to_csv import process_urls_to_xlsx
# from claude_to_csv import get_shop_name_from_url
# from claude_to_csv import ctc_get_olx_ads_count
# from claude_to_csv import ctc_get_olx_ads_count_selenium

# === KONFIGURACJA ===
#CATEGORY_URL = "https://www.olx.pl/elektronika/sprzet-audio/"
#CATEGORY_URL = "https://www.olx.pl/dom-ogrod/instalacje/"
#CATEGORY_URL = "https://www.olx.pl/dla-firm/maszyny-i-urzadzenia/"
#CATEGORY_URL = "https://www.olx.pl/muzyka-edukacja/muzyka/"
#CATEGORY_URL = "https://www.olx.pl/muzyka-edukacja/instrumenty/"
CATEGORY_URL = "https://www.olx.pl/dom-ogrod/budowa/"
MAX_PAGES    = 1

# === INICJALIZACJA WEBDRIVERA ===
def get_webdriver():
    service = Service(ChromeDriverManager().install())
    opts = webdriver.ChromeOptions()
    # opts.add_argument("--headless")  # odkomentuj, by uruchomić w tle
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(service=service, options=opts)

def get_shop_info_improved(listing_url):
    """
    Ulepszona wersja - rozróżnia sklepy premium od zwykłych użytkowników
    """
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(options=chrome_options)

    try:
        print(f"\nŁadowanie strony: {listing_url}")
        driver.get(listing_url)
        time.sleep(5)

        #wczesniejsza deklaracja słownika
        #shop_info = {}

        # Inicjalizacja struktury danych z domyślnymi wartościami
        shop_record = {
            'profile_url': None,
            'ads_count': None,
            'name': None,
            'type': None,
            #'source_listing': listing_url  # Dodatkowe pole - z jakiego ogłoszenia pobrano dane
        }

        # Najpierw sprawdź czy to sklep premium (ma parametr w URL)
        is_premium_shop = 'olx_shop_premium' in listing_url

        # Metoda 1: Szukaj linku "Więcej od tego ogłoszeniodawcy"
        try:
            # Ten link prowadzi do profilu sprzedawcy
            more_link = driver.find_element(By.PARTIAL_LINK_TEXT, "Więcej od tego ogłoszeniodawcy")
            profile_url = more_link.get_attribute('href')

            if profile_url:
                shop_record['profile_url'] = profile_url
                print(f"  ✓ Link do profilu: {profile_url}")

                #Zmiana: podstawienie funkcji z claude_to_csv
                #get_olx_ads_count nieaktywne
                ads_count = ctc_get_olx_ads_count(profile_url)  # LICZBA OGLOSZEN
                #mozna tez dać dodawanie ads_count do shop record bez warunku tutaj
                if ads_count is not None:
                    shop_record['ads_count'] = ads_count #brakowało tej linii!
                    print(f"Liczba ogłoszeń: {ads_count}")
                else:
                    print("Nie udało się pobrać liczby ogłoszeń.")

                # Teraz musimy pobrać nazwę sprzedawcy
                # Nazwa powinna być gdzieś obok tego linku
                parent = more_link.find_element(By.XPATH, "../..")

                # Szukaj nazwy w rodzicu
                name_elements = parent.find_elements(By.CSS_SELECTOR, "h2, h3, h4, strong")
                for elem in name_elements:
                    text = elem.text.strip()
                    if text and text != "Więcej od tego ogłoszeniodawcy" and len(text) < 100:
                        shop_record['name'] = text
                        print(f"  ✓ Nazwa: {text}")
                        break
        except:
            pass

        # Metoda 2: JavaScript - bardziej precyzyjne szukanie (POPRAWIONE)
        if 'profile_url' not in shop_record or 'name' not in shop_record:
            try:
                js_result = driver.execute_script("""
                    // Znajdź sekcję ze sprzedawcą
                    const sections = document.querySelectorAll('section, div[role="region"]');
                    let result = null;

                    for (let section of sections) {
                        // Sprawdź czy sekcja zawiera link do profilu
                        const profileLink = section.querySelector('a[href*="/oferty/uzytkownik/"], a[href*="/sklepy/"], a[href*=".olx.pl/home/"]');
                        if (!profileLink) continue;

                        // Znajdź nazwę - zwykle jest w h2, h3 lub strong w tej samej sekcji
                        const nameElements = section.querySelectorAll('h2, h3, h4, strong, [class*="title"]');

                        for (let elem of nameElements) {
                            const text = elem.textContent.trim();
                            // Pomijamy teksty które są linkami lub za długie
                            if (text &&
                                text !== "Więcej od tego ogłoszeniodawcy" &&
                                text.length > 2 &&
                                text.length < 100 &&
                                !text.includes('Zestaw') &&  // Pomijamy tytuły ogłoszeń
                                !text.includes('LEGO')) {     // Pomijamy tytuły ogłoszeń

                                result = {
                                    name: text,
                                    profileUrl: profileLink.href,
                                    isPremium: profileLink.href.includes('/sklepy/') || profileLink.href.includes('.olx.pl/home/')
                                };
                                break;
                            }
                        }

                        if (result) break;
                    }

                    // Jeśli nie znaleziono nazwy, zwróć przynajmniej link
                    if (!result) {
                        const anyProfileLink = document.querySelector('a[href*="/oferty/uzytkownik/"], a[href*="/sklepy/"], a[href*=".olx.pl/home/"]');
                        if (anyProfileLink) {
                            result = {
                                profileUrl: anyProfileLink.href,
                                isPremium: anyProfileLink.href.includes('/sklepy/') || anyProfileLink.href.includes('.olx.pl/home/')
                            };
                        }
                    }

                    return result;
                """)

                if js_result:
                    shop_record.update(js_result)
                    print(f"  ✓ Dane z JS: {js_result}")
            except Exception as e:
                print(f"  ! Błąd JS (kontynuuję): {str(e)[:100]}...")

        # Metoda 3: Jeśli mamy link do profilu ale nie mamy nazwy, możemy go odwiedzić
        if 'profile_url' in shop_record and 'name' not in shop_record:
            print("  → Odwiedzam profil aby pobrać nazwę...")
            driver.get(shop_record['profile_url'])
            time.sleep(3)

            # Na stronie profilu nazwa jest bardziej widoczna
            try:
                # Dla sklepów
                shop_name = driver.find_element(By.CSS_SELECTOR, "h1, [class*='shop-name'], [class*='seller-name']")
                if shop_name:
                    shop_record['name'] = shop_name.text.strip()
                    print(f"  ✓ Nazwa z profilu: {shop_record['name']}")
            except:
                # Dla zwykłych użytkowników - nazwa może być w tytule strony
                title = driver.title
                if " - " in title:
                    potential_name = title.split(" - ")[0].strip()
                    if len(potential_name) > 2 and len(potential_name) < 50:
                        shop_record['name'] = potential_name
                        print(f"  ✓ Nazwa z tytułu: {shop_record['name']}")

        # Określ typ konta
        if 'profile_url' in shop_record:
            if '/sklepy/' in shop_record['profile_url'] or '.olx.pl/home/' in shop_record['profile_url']:
                shop_record['type'] = 'sklep_premium'
            elif '/oferty/uzytkownik/' in shop_record['profile_url']:
                shop_record['type'] = 'uzytkownik'

        return shop_record

    except Exception as e:
        print(f"Błąd główny: {e}")
        import traceback
        traceback.print_exc()
        return {}
    finally:
        driver.quit()

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept-Language": "pl-PL,pl;q=0.9"
}


# === ETAP 1: ZBIERANIE LINKÓW DO OGŁOSZEŃ ===
def extract_ad_links(driver, category_url, max_pages, test_mode=False):
    ad_links = set()
    for page in range(1, max_pages + 1):
        page_url = f"{category_url}?page={page}"
        print(f"🔍 Scraping: {page_url}")
        driver.get(page_url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//a[contains(@href, '/d/oferta/')]")
            )
        )
        elements = driver.find_elements(
            By.XPATH, "//a[contains(@href, '/d/oferta/')]")
        for el in elements:
            href = el.get_attribute('href')
            if href:
                ad_links.add(href.split('?')[0])

                # TRYB TESTOWY: Przerwij po 5 linkach
                if test_mode and len(ad_links) >= 5:
                    print(f"🧪 TRYB TESTOWY: Zatrzymano po {len(ad_links)} linkach")
                    return list(ad_links)
        time.sleep(1)
    print(f"⚡ Found {len(ad_links)} unique ads")
    return list(ad_links)


def extract_store_urls(driver, ad_links):

    store_urls={}

    for ad in ad_links:
        print(f"🔗 Opening ad: {ad}")
        # get_shop_info_improved otworzy swoją własną przeglądarkę
        # podmiana na funkcje z ctc
        shop_record = get_shop_info_improved(ad)

        # Dodaj rekord do listy (nawet jeśli niepełny)
        #all_shop_records.append(shop_record)

        # 2 sposoby - dodaj URL do setu jeśli istnieje
        # Użyj danych które znalazła funkcja!
        if shop_record and 'profile_url' in shop_record:
            #store_urls.add(shop_record['profile_url'])
            key_dict = shop_record['profile_url']
            store_urls[key_dict]=shop_record
            print(f"   Dodano sklep: {key_dict} (typ: {shop_record.get('type', 'nieznany')})")
        else:
            print("   Brak profile_url w shop_info")

        #print(f"   Typ: {shop_record.get('type', 'nieznany')}")

        time.sleep(1) #potrzebne?

    print(f"⚡ Found {len(store_urls)} unique stores")
    #print(f"📊 Collected {len(all_shop_records)} shop records")
    #return list(store_urls), all_shop_records
    return store_urls

def ctc_get_olx_ads_count_selenium(shop_url):
    """
    Używa Selenium do pobrania liczby ogłoszeń
    """
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(options=chrome_options)

    try:
        driver.get(shop_url)
        time.sleep(3)

        page_text = driver.find_element(By.TAG_NAME, "body").text

        # Dla użytkowników
        if "/oferty/uzytkownik/" in shop_url:
            # Sprawdź czy to nie jest przekierowanie do wszystkich ogłoszeń
            if "wszystkie ogłoszenia w" in page_text.lower():
                return 0

            match = re.search(r'Znaleźliśmy (\d+) ogłoszeń', page_text)
            if match:
                count = int(match.group(1))
                if count > 1000000:
                    return 0
                return count

            match = re.search(r'Wszystkie ogłoszenia\s*(\d+)', page_text)
            if match:
                count = int(match.group(1))
                if count > 1000000:
                    return 0
                return count

            if any(text in page_text for text in ["Brak ogłoszeń", "Nie ma ogłoszeń", "0 ogłoszeń"]):
                return 0

        driver.quit()
        return None

    except Exception as e:
        driver.quit()
        return None


def ctc_get_olx_ads_count(shop_url):
    """
    Pobiera liczbę ogłoszeń
    """
    resp = requests.get(shop_url, headers=HEADERS)
    if resp.status_code != 200:
        return None

    html = resp.text
    soup = BeautifulSoup(html, 'html.parser')

    is_user_page = "/oferty/uzytkownik/" in shop_url
    is_shop_page = ".olx.pl/home/" in shop_url

    if is_shop_page:
        # Dla sklepów firmowych
        for element in soup.find_all(['button', 'a', 'span', 'div', 'h1', 'h2', 'h3']):
            text = element.get_text().strip()

            patterns = [
                r'(\d+[\s\d]*)\s*ogłoszeń',
                r'(\d+[\s\d]*)\s*ofert',
            ]

            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    count = int(match.group(1).replace(' ', '').replace('\xa0', ''))
                    if 0 < count <= 10000:
                        return count

    elif is_user_page:
        return ctc_get_olx_ads_count_selenium(shop_url)

    if is_shop_page:
        return ctc_get_olx_ads_count_selenium(shop_url)

    return None

def process_urls_to_xlsx(store_data, output_filename="olx_sellers.xlsx"):
    """
    Przetwarza listę URL-i i zapisuje wyniki do XLSX
    """
    results = []

    print(f"Przetwarzanie {len(store_data)} rekordów sklepów...\n")

    # Iteruj po słowniku - klucz to URL, wartość to shop_record
    for i, (profile_url, shop_record) in enumerate(store_data.items(), 1):
        # Dodaj do wyników
        results.append({
            'Nazwa użytkownika/firmy': shop_record.get('name', ''),
            #'Link do konta': shop_record.get['profile_url', ''],
            'Link do konta': profile_url,  # Możesz użyć klucza lub shop_record.get('profile_url', '')
            'Nr telefonu': '',
            'Login': '',
            'Hasło': '',
            'Ilość ogłoszeń': shop_record.get('ads_count', 0) or 0,
            'Nazwa platformy': 'olx.pl'
        })

    # Utwórz DataFrame
    df = pd.DataFrame(results)

    # Zapisz do XLSX z formatowaniem
    print(f"\n\nZapisywanie wyników do {output_filename}...")

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sprzedawcy OLX', index=False)

        # Pobierz arkusz
        worksheet = writer.sheets['Sprzedawcy OLX']

        # Formatowanie nagłówków
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # Dostosuj szerokość kolumn
        column_widths = {
            'A': 30,  # Nazwa
            'B': 50,  # Link
            'C': 15,  # Nr telefonu
            'D': 15,  # Login
            'E': 15,  # Hasło
            'F': 15,  # Ilość ogłoszeń
            'G': 15  # Platforma
        }

        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width

        # Wyrównaj dane
        for row in worksheet.iter_rows(min_row=2):
            row[5].alignment = Alignment(horizontal="center")  # Ilość ogłoszeń
            row[6].alignment = Alignment(horizontal="center")  # Platforma

    print(f"✅ Zapisano {len(results)} rekordów do {output_filename}")

    # Podsumowanie
    print("\nPodsumowanie:")
    # total_ads = df['Ilość ogłoszeń'].sum()
    # valid_counts = df[df['Ilość ogłoszeń'] > 0]
    # print(f"  - Łączna liczba ogłoszeń: {total_ads}")
    # print(f"  - Średnia liczba ogłoszeń: {df['Ilość ogłoszeń'].mean():.1f}")
    # print(f"  - Rekordy z ogłoszeniami: {len(valid_counts)}/{len(df)}")

    # Zamień None na 0 dla obliczeń
    ads_counts = [record.get('ads_count', 0) or 0 for record in store_data.values()]
    total_ads = sum(ads_counts)
    valid_counts = [count for count in ads_counts if count > 0]

    print(f"  - Łączna liczba ogłoszeń: {total_ads}")
    if ads_counts:
        print(f"  - Średnia liczba ogłoszeń: {sum(ads_counts) / len(ads_counts):.1f}")
    print(f"  - Rekordy z ogłoszeniami: {len(valid_counts)}/{len(store_data)}")

# === GŁÓWNA FUNKCJA ===
def main():
    driver = get_webdriver()
    try:
        ad_links   = extract_ad_links(driver, CATEGORY_URL, MAX_PAGES, test_mode=False)
        store_data = extract_store_urls(driver, ad_links)


    finally:
        driver.quit()

    # Upewnij się, że masz zainstalowane wymagane pakiety
    try:
        import pandas
        import openpyxl
    except ImportError:
        print("Instaluję wymagane pakiety...")
        import subprocess

        subprocess.check_call(["pip", "install", "pandas", "openpyxl"])

# Generuj nazwę pliku z datą i czasem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"olx_sellers_v3{timestamp}.xlsx"

    print(f"DEBUG: Przekazuję {len(store_data)} URL-i do process_urls_to_xlsx")
    if store_data:
        first_url = list(store_data.keys())[0]
        first_record = store_data[first_url]  # Pierwszy rekord
        print(f"DEBUG: Przykładowy URL: {first_url}")
        print(f"DEBUG: Przykładowy rekord: {first_record}")
    else:
        print("DEBUG: Słownik store_data jest pusty!")

    # Przetwórz URL-e i zapisz do XLSX
    process_urls_to_xlsx(store_data, output_file)

if __name__ == '__main__':
    main()
