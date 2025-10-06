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

#Timeout
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# === KONFIGURACJA ===

# 1. Podaj adres URL sklepu
CATEGORY_URL = "https://www.olx.pl/uslugi/budowa-remont/"
#CATEGORY_URL = ""

# 2. Ustaw tryb testowy (przetworzenie pierwszych 5 ogłoszeń/produktów) lub pełny
TEST_MODE = False

# 3. Wybierz zakładki do scrapowania
START_PAGE=1
END_PAGE=1


# === INICJALIZACJA WEBDRIVERA ===
def get_webdriver():
    service = Service(ChromeDriverManager().install())
    opts = webdriver.ChromeOptions()
    # opts.add_argument("--headless")  # odkomentuj, by uruchomić w tle
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(service=service, options=opts)

# Optymalizacja czasu przetwarzania
def quick_get_profile_url(listing_url):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    # Dodane do optymalizacji:
    chrome_options.add_argument('--disable-images')
    chrome_options.add_argument('--disable-css')
    chrome_options.add_argument('--disable-plugins')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--disable-default-apps')

    driver = webdriver.Chrome(options=chrome_options)

    try:
        driver.set_page_load_timeout(5) # Czas ładowania witryny
        print(f"\nŁadowanie strony: {listing_url}")
        driver.get(listing_url)
        #time.sleep(4) # time_optim
        # Dotyczy szukania elementu na juz załadowanej stronie
        wait = WebDriverWait(driver, 2)
        more_link = wait.until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Więcej od tego ogłoszeniodawcy"))
            # powinny być podwójne nawiasy okrągłe!
        )
        profile_url = more_link.get_attribute('href')
        return profile_url, more_link, driver

    except TimeoutException:
        print(f"Timeout - nie znaleziono linku w {listing_url}")
        driver.quit()
        return None, None, None
    except Exception as e:
        print(f"Błąd {e}")
        driver.quit()
        return None, None, None

    # finally:
    #     driver.quit()
    # Bez finally bo nie pobierze nazwy!


def get_shop_info_improved(listing_url, seen:set, treshold=100):
    """
    Ulepszona wersja - rozróżnia sklepy premium od zwykłych użytkowników
    """
    # Przekazanie drivera z quick_get_profile_url

    try:

    # Inicjalizacja struktury danych z domyślnymi wartościami
    # Musi być na poczatku!
        shop_record = {
            'profile_url': None,
            'ads_count': None,
            'name': None,
            'type': None,
        }
        # Sprawdz czy to sklep premium (ma parametr w URL)
        is_premium_shop = 'olx_shop_premium' in listing_url

        # Szukaj linku
        try:

            # Funkcja do skip
            profile_url, more_link, driver = quick_get_profile_url(listing_url)
            # 1) Filtr profilu
            if not profile_url:
                # Early return
                print(f"   ❌ Brak profile_url w szybkim fetchu – skip {listing_url}")
                return {}
            # 2) Filtr duplikatów
            if profile_url in seen:
                print(f"   ⚠ Duplikat {profile_url} – skip")
                return {}
            # Dodanie nowego URL do setu
            seen.add(profile_url)

            if profile_url:
                ads_count = ctc_get_olx_ads_count(profile_url)  # liczba ogłoszeń
                # Ponizej progu funkcja nie przepuszcza sklepu dalej - optymalizacja czasu
                if ads_count < treshold:
                    return {}
                    #return None

                shop_record['profile_url'] = profile_url
                print(f"  ✓ Link do profilu: {profile_url}")

                # Opcjonalnie można zaimplementować dodawanie ads_count do shop record bez warunku tutaj
                if ads_count is not None:
                    shop_record['ads_count'] = ads_count
                    print(f"Liczba ogłoszeń: {ads_count}")
                else:
                    print("Nie udało się pobrać liczby ogłoszeń.")

                # Pobieranie nazwy - zabezpieczenie
                try:
                    if more_link:
                    # Sprawdź czy more_link jest nadal aktywne
                        # Przejście do pobrania nazwy sprzedawcy
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
        except:
            pass

        # # 1) Określ typ konta
        # if 'profile_url' in shop_record:
        #     # Uwaga wrazliwa linia!
        #     if '/sklepy/' in shop_record['profile_url'] or '.olx.pl/home/' in shop_record['profile_url']:
        #         shop_record['type'] = 'sklep_premium'
        #     elif '/oferty/uzytkownik/' in shop_record['profile_url']:
        #         shop_record['type'] = 'uzytkownik'

    # 2) Określ typ konta - poprawione
        profile_rec = shop_record.get('profile_url')
        if profile_rec:
            if '/sklepy/' in profile_rec or '.olx.pl/home/' in profile_rec:
                shop_record['type'] = 'sklep_premium'
            elif '/oferty/uzytkownik/' in profile_rec:
                shop_record['type'] = 'uzytkownik'
            else:
                shop_record['type'] = 'nieznany'
        else:
            shop_record['type'] = 'brak_profilu'

        return shop_record

    except Exception as e:
        print(f"Błąd główny: {e}")
        import traceback
        traceback.print_exc()
        return {}
    finally:
        #if 'driver' in locals() and driver:
        try:
            driver.quit()
        except:
            pass

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept-Language": "pl-PL,pl;q=0.9"
}


# === ZBIERANIE LINKÓW DO OGŁOSZEŃ ===
def extract_ad_links(driver, category_url, start_page, end_page, TEST_MODE):
    ad_links = set()
    for page in range(start_page, end_page+1):
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
                if TEST_MODE and len(ad_links) >= 5:
                    print(f"TRYB TESTOWY: Zatrzymano po {len(ad_links)} linkach")
                    return list(ad_links)
        time.sleep(1)
    print(f"⚡ Found {len(ad_links)} unique ads")
    return list(ad_links)


def extract_store_urls(driver, ad_links):
    seen=set()
    store_urls={}

    for ad in ad_links:
        shop_record = get_shop_info_improved(ad, seen)

        # Dodaj URL do setu jeśli istnieje
        # Użyj danych które znalazła funkcja
        if shop_record and 'profile_url' in shop_record:
            #store_urls.add(shop_record['profile_url'])
            key_dict = shop_record['profile_url']
            store_urls[key_dict]=shop_record
            print(f"   Dodano sklep: {key_dict} (typ: {shop_record.get('type', 'nieznany')})")
        else:
            print("   Brak profile_url w shop_info")

        time.sleep(1) # opcjonalnie

    print(f"⚡ Found {len(store_urls)} unique stores")

    return store_urls

# Funkcja używa selenium do pobrania ads_count
def ctc_get_olx_ads_count_selenium(shop_url):

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
            # Sprawdzenie czy to nie jest przekierowanie do ogólnej kategorii
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

# Pobiera liczbę ogłoszeń
def ctc_get_olx_ads_count(shop_url):
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

# Przetwarza listę URL-i i zapisuje wyniki do XLSX
def process_urls_to_xlsx(store_data, output_filename="olx_sellers.xlsx"):

    results = []

    print(f"Przetwarzanie {len(store_data)} rekordów sklepów...\n")

    # Iteruje po słowniku - klucz to URL, wartość to shop_record
    for i, (profile_url, shop_record) in enumerate(store_data.items(), 1):
        # Dodaj do wyników
        results.append({
            'Nazwa użytkownika/firmy': shop_record.get('name', ''),
            'Link do konta': profile_url,  # Można użyć klucza lub shop_record.get('profile_url', '')
            'Nr telefonu': '',
            'Login': '',
            'Hasło': '',
            'Ilość ogłoszeń': shop_record.get('ads_count', 0) or 0,
            'Nazwa platformy': 'olx.pl'
        })

    # Utwórz DataFrame
    df = pd.DataFrame(results)

    # Zapisz do pliku .xlsx z formatowaniem
    print(f"\n\nZapisywanie wyników do {output_filename}...")

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sprzedawcy OLX', index=False)

        # Pobranie arkusza
        worksheet = writer.sheets['Sprzedawcy OLX']

        # Formatowanie nagłówków
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # Dostosowanie szerokości kolumn
        column_widths = {
            'A': 30,  # Nazwa
            'B': 50,  # Link
            'C': 15,  # Nr telefonu
            'D': 15,  # Login
            'E': 15,  # Hasło
            'F': 15,  # Ilość ogłoszeń
            'G': 15   # Platforma
        }

        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width

        # Wyrównanie danych
        for row in worksheet.iter_rows(min_row=2):
            row[5].alignment = Alignment(horizontal="center")  # Ilość ogłoszeń
            row[6].alignment = Alignment(horizontal="center")  # Platforma

    print(f"✅ Zapisano {len(results)} rekordów do {output_filename}")

    # Podsumowanie
    print("\nPodsumowanie:")
    total_ads = df['Ilość ogłoszeń'].sum()
    valid_counts = df[df['Ilość ogłoszeń'] > 0]
    print(f"  - Łączna liczba ogłoszeń: {total_ads}")
    print(f"  - Średnia liczba ogłoszeń: {df['Ilość ogłoszeń'].mean():.1f}")
    print(f"  - Rekordy z ogłoszeniami: {len(valid_counts)}/{len(df)}")

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
    # Pomiar czasu przetwarzania
    start_time = time.time()

    driver = get_webdriver()
    try:
        ad_links   = extract_ad_links(driver, CATEGORY_URL, START_PAGE, END_PAGE, TEST_MODE)
        store_data = extract_store_urls(driver, ad_links)

    finally:
        driver.quit()

    # Sprawdzenie stanu pakietów
    try:
        import pandas
        import openpyxl
    except ImportError:
        print("Instaluję wymagane pakiety...")
        import subprocess

        subprocess.check_call(["pip", "install", "pandas", "openpyxl"])

    # Generowanie nazwy pliku z datą i czasem
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

    # Zapis URL-i do .xlsx
    process_urls_to_xlsx(store_data, output_file)

    # Podsumowanie optymalizacji czasowej
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"\n⏱️  CZAS WYKONANIA: {execution_time:.1f} sekund ({execution_time / 60:.1f} minut)")

if __name__ == '__main__':
    main()