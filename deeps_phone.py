import requests
import re
import json
from urllib.parse import urlparse
import time

def get_phone_number(offer_url):
    # Wyciągnij ad_id z URL
    ad_id_match = re.search(r'ID([a-zA-Z0-9]+)\.html', offer_url)
    if not ad_id_match:
        return None
    ad_id = ad_id_match.group(1)

    # Wyślij zapytanie do API
    # Ządanie do endpointu (endpoint moze ulegac zmianie)
    api_url = f"https://www.olx.pl/api/v1/offers/{ad_id}/phone/"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        return data.get('phone', {}).get('phone')
    else:
        print(f"Błąd API: {response.status_code}")
        return None

def extract_phone_number(offer_url: str) -> str:
    """
    Pobiera numer telefonu z ogłoszenia OLX na podstawie bezpośredniego URL oferty.

    :param offer_url: Bezpośredni link do ogłoszenia OLX
    :return: Numer telefonu lub None jeśli nie uda się pobrać
    """
    # Weryfikacja poprawności URL
    parsed = urlparse(offer_url)
    if not parsed.scheme or not parsed.netloc or "olx.pl" not in parsed.netloc:
        raise ValueError("Nieprawidłowy URL ogłoszenia OLX")

    # Ekstrakcja ad_id z URL (obsługa różnych formatów)
    ad_id_match = re.search(r'(?:ID|ad_id=)([a-zA-Z0-9]+)', offer_url)
    if not ad_id_match:
        raise ValueError("Nie znaleziono identyfikatora oferty w URL")

    ad_id = ad_id_match.group(1)

    # Konstrukcja URL API
    api_url = f"https://www.olx.pl/api/v1/offers/{ad_id}/phone/"

    # Nagłówki imitujące przeglądarkę
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
        'Accept': 'application/json',
        'Referer': offer_url,
        'X-Requested-With': 'XMLHttpRequest'
    }

    try:
        response = requests.get(api_url, headers=headers, timeout=10)
        response.raise_for_status()  # Sprawdza status HTTP

        # Parsowanie odpowiedzi JSON
        data = response.json()
        return data.get('phone', {}).get('phone', {}).get('number') or data.get('phone', {}).get('number')

    except (requests.RequestException, json.JSONDecodeError, KeyError) as e:
        print(f"Błąd podczas pobierania numeru: {e}")
        return None

#Jeśli API przestanie działać, alternatywą jest użycie Selenium do symulacji kliknięcia przycisku "Pokaż numer"

if __name__ == '__main__':

    #Przykład użycia dla jednego ogłoszenia:
    ad_url="https://www.olx.pl/d/oferta/lego-figurka-piraci-pirates-pi143-pirate-female-piratka-CID88-ID14PhVH.html?reason=seller_listing%7Colx_shop_premium"

    # Najlepiej podać bezpośredni link do ogłoszenia. Wersja z linkiem do sklepu wymaga dodatkowego kroku (pobierania listy ogłoszeń), co zwiększa ryzyko błędów i blokad
    test_urls2 = [
        "https://www.olx.pl/d/oferta/uzywane-czesci-audi-a6-c5-a4-b5-a4-b6-a3-audi-blask-warszawa-goldap-CID5-IDTj4r4.html?reason=seller_listing%7Colx_shop_premium",  #
        "https://www.olx.pl/d/oferta/fabrycznie-nowe-gumy-oslony-lag-zawieszenia-przod-simson-s51-s-50-CID5-IDCxnom.html?reason=seller_listing%7Colx_shop_basic",  #
        "https://autoczesci.olx.pl/home/"  # strona gł
    ]
    # phone = extract_phone_number(ad_url)
    # print(f"Numer telefonu: {phone}")

    for url in test_urls2:
        print(f"Testowanie: {url}")
        print(f"Numer: {extract_phone_number(url)}\n")
        # Nie wysyłaj wiecej niz 1 zapytanie/sek.
        time.sleep(1.5)  #olx moze blokowac nadmiarowe ządania

        #Rotacja User-Agent: Dodaj rotację nagłówków dla wielu żądań
        #Wykrywanie blokad: Sprawdzaj status 429/403 w odpowiedziach