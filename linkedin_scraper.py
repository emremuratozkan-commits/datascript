# -*- coding: utf-8 -*-
"""
LinkedIn Veri Çekme, Derin Profil Taraması ve Excel Raporlama Çekirdek Motoru
Geliştirici: Antigravity
Gereksinimler: pip install selenium beautifulsoup4 pandas openpyxl webdriver-manager streamlit
"""

import os
import re
import io
import time
import random
from urllib.parse import quote_plus, urlsplit, urlunsplit

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

try:
    from webdriver_manager.chrome import ChromeDriverManager
    HAS_WEBDRIVER_MANAGER = True
except ImportError:
    HAS_WEBDRIVER_MANAGER = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def varsayilan_logger(mesaj, seviye="INFO"):
    """Konsol için varsayılan log yazıcı."""
    zaman_damgasi = time.strftime("%H:%M:%S")
    etiket = f"[{seviye.upper()}]"
    print(f"{zaman_damgasi} {etiket} {mesaj}")


def rastgele_bekle(min_sn=1.2, max_sn=2.5):
    """Bot korumasına takılmamak için rastgele aralıklarla bekler."""
    time.sleep(random.uniform(min_sn, max_sn))


def tarayiciyi_baslat(headless=False, logger=varsayilan_logger):
    """
    Hızlı ve kararlı Selenium WebDriver örneği başlatır.
    Hem yerel Windows hem de Streamlit Cloud (Linux) ortamlarıyla %100 uyumludur.
    """
    import sys
    logger("Tarayıcı hazırlanıyor...", "INFO")
    chrome_options = Options()

    # Streamlit Cloud (Linux) üzerinde ekran olmadığı için zorunlu headless ve flags
    if sys.platform.startswith("linux") or os.environ.get("STREAMLIT_SERVER_PORT"):
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        # Linux Chromium binary konumu
        for b_path in ["/usr/bin/chromium", "/usr/bin/chromium-browser"]:
            if os.path.exists(b_path):
                chrome_options.binary_location = b_path
                break
    elif headless:
        chrome_options.add_argument("--headless=new")

    # Temiz ve kilitlenmeyen ayarlar
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        logger(f"Sistem Chrome sürücüsü deneniyor ({e})...", "WARNING")
        if HAS_WEBDRIVER_MANAGER:
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception:
                driver = webdriver.Chrome(options=chrome_options)
        else:
            driver = webdriver.Chrome(options=chrome_options)

    # Sayfa yükleme zaman aşımı (donmayı önler)
    driver.set_page_load_timeout(35)

    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except Exception:
        pass

    logger("Tarayıcı başarıyla başlatıldı.", "SUCCESS")
    return driver


def linkedin_cerezle_giris_yap(driver, li_at_cookie, logger=varsayilan_logger):
    """
    LinkedIn'e li_at çerezi ile şifresiz ve CAPTCHA'sız doğrudan giriş yapar.
    En güvenli ve LinkedIn tarafından engellenmeyen yöntemdir.
    """
    logger("Çerez ile hızlı giriş yapılıyor...", "INFO")
    try:
        driver.get("https://www.linkedin.com")
        rastgele_bekle(1.5, 2.5)

        cookie_dict = {
            "name": "li_at",
            "value": li_at_cookie.strip(),
            "domain": ".linkedin.com",
            "path": "/"
        }
        driver.add_cookie(cookie_dict)
        rastgele_bekle(0.5, 1.0)

        # Sayfayı yenileyerek oturumu aktifleştir
        driver.get("https://www.linkedin.com/feed/")
        rastgele_bekle(2, 3.5)

        if "feed" in driver.current_url or "mynetwork" in driver.current_url:
            logger("Çerez doğrulaması başarılı! Oturum doğrudan açıldı.", "SUCCESS")
            return True
        else:
            logger("Çerez geçersiz veya süresi dolmuş olabilir.", "WARNING")
            return False
    except Exception as e:
        logger(f"Çerez girişinde hata: {e}", "ERROR")
        return False


def linkedin_manuel_giris_bekle(driver, logger=varsayilan_logger, max_bekleme=120):
    """
    Tarayıcıyı doğrudan LinkedIn giriş sayfasına götürür ve kullanıcının
    giriş yapmasını bekler. Oturum açıldığı anda otomatik olarak aramaya geçer.
    """
    logger("Chrome açıldı. LinkedIn giriş sayfasına gidiliyor...", "INFO")
    try:
        driver.get("https://www.linkedin.com/login")
        rastgele_bekle(1.5, 2.5)
        logger("LinkedIn giriş sayfası yüklendi! Lütfen açılan pencereden giriş yapın.", "SUCCESS")
    except Exception as e:
        logger(f"Giriş sayfası açılırken hata: {e}", "ERROR")
        return False
    
    for sn in range(max_bekleme):
        try:
            url = driver.current_url
            if any(k in url for k in ["feed", "check/add-phone", "mynetwork", "messaging"]):
                logger("Tebrikler! Oturum açıldığı tespit edildi. Taramaya geçiliyor...", "SUCCESS")
                rastgele_bekle(1.5, 2.5)
                return True
        except Exception:
            pass
        if sn % 15 == 0 and sn > 0:
            logger(f"Giriş yapmanız bekleniyor... ({max_bekleme - sn} sn kaldı)", "INFO")
        time.sleep(1)

    logger("Zaman aşımı! 120 saniye içinde giriş tamamlanamadı.", "ERROR")
    return False


def linkedin_giris_yap(driver, email, sifre, logger=varsayilan_logger, max_bekleme=30):
    """
    LinkedIn giriş sayfasına gider, kullanıcı bilgilerini doldurur ve oturum açar.
    Hem 'username' hem 'session_key' seçicilerini destekler.
    """
    logger("LinkedIn giriş sayfasına bağlanılıyor...", "INFO")
    try:
        driver.get("https://www.linkedin.com/login")
        rastgele_bekle(2, 3)

        # E-posta alanı (username veya session_key)
        email_alani = None
        for selector in [
            (By.ID, "username"),
            (By.ID, "session_key"),
            (By.NAME, "session_key"),
            (By.CSS_SELECTOR, "input[name='session_key']")
        ]:
            try:
                email_alani = WebDriverWait(driver, 3).until(EC.presence_of_element_located(selector))
                if email_alani:
                    break
            except Exception:
                continue

        if not email_alani:
            logger("E-posta giriş alanı bulunamadı! Sayfa yapısı değişmiş veya güvenlik kontrolünde olabilir.", "ERROR")
            return False

        email_alani.clear()
        for harf in email:
            email_alani.send_keys(harf)
            time.sleep(random.uniform(0.03, 0.08))

        rastgele_bekle(0.4, 0.8)

        # Şifre alanı (password veya session_password)
        sifre_alani = None
        for selector in [
            (By.ID, "password"),
            (By.ID, "session_password"),
            (By.NAME, "session_password"),
            (By.CSS_SELECTOR, "input[name='session_password']")
        ]:
            try:
                sifre_alani = driver.find_element(*selector)
                if sifre_alani:
                    break
            except Exception:
                continue

        if not sifre_alani:
            logger("Şifre alanı bulunamadı!", "ERROR")
            return False

        sifre_alani.clear()
        for harf in sifre:
            sifre_alani.send_keys(harf)
            time.sleep(random.uniform(0.03, 0.08))

        rastgele_bekle(0.6, 1.2)

        # Giriş butonu
        giris_butonu = None
        for selector in [
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "button.btn__primary--large"),
            (By.CSS_SELECTOR, "button[data-litms-control-urn*='login']")
        ]:
            try:
                giris_butonu = driver.find_element(*selector)
                if giris_butonu:
                    break
            except Exception:
                continue

        if giris_butonu:
            giris_butonu.click()
        else:
            sifre_alani.submit()

        logger("Giriş bilgileri iletildi. Doğrulama/Oturum kontrol ediliyor...", "INFO")

    except Exception as e:
        logger(f"Giriş formunu doldururken hata: {e}", "ERROR")
        return False

    # Güvenlik Kontrolü / CAPTCHA / 2FA Bekleme Döngüsü
    logger("Doğrulama kontrol ediliyor (CAPTCHA veya 2FA çıkarsa lütfen açılan Chrome penceresinden tamamlayın)...", "WARNING")

    giris_basarili = False
    for sn in range(max_bekleme):
        try:
            url = driver.current_url
            if any(k in url for k in ["feed", "check/add-phone", "mynetwork", "messaging"]):
                giris_basarili = True
                break
            if "checkpoint" in url or "challenge" in url:
                if sn % 5 == 0:
                    logger(f"LinkedIn güvenlik doğrulaması (CAPTCHA/SMS) bekleniyor... Lütfen tarayıcıdan tamamlayın.", "WARNING")
        except Exception:
            pass
        time.sleep(1)

    if giris_basarili:
        logger("LinkedIn oturumu başarıyla açıldı!", "SUCCESS")
        return True
    else:
        logger("Otomatik yönlendirme tamamlanamadı. Eğer tarayıcıda giriş yaptıysanız devam edebilirsiniz.", "WARNING")
        return True


def sayfayi_asagi_kaydir(driver, kaydirma_sayisi=7, logger=varsayilan_logger):
    """
    Arama sonuçlarının (lazy loading / infinite scroll) tam olarak yüklenmesi için
    sayfayı kademeli ve rastgele gecikmelerle aşağı kaydırır.
    """
    logger("Sayfadaki profiller yükleniyor (sonsuz kaydırma çalıştırılıyor)...", "INFO")
    
    # Sayfanın ilk yüklenmesini bekle
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='/in/'], div.entity-result, div[data-view-name*='search']"))
        )
    except Exception:
        pass

    for _ in range(kaydirma_sayisi):
        adim = random.randint(350, 700)
        driver.execute_script(f"window.scrollBy(0, {adim});")
        rastgele_bekle(0.5, 1.0)

    # Sayfanın en altına kadar son bir kaydırma
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    rastgele_bekle(1.5, 2.5)


def sonraki_sayfaya_gec(driver, sayfa_no, arama_kriteri, logger=varsayilan_logger):
    """
    Sonraki arama sayfasına geçer.
    Önce sayfadaki 'Sonraki / İleri' butonuna basmayı dener, bulunamazsa URL üzerinden ilerler.
    """
    logger(f"{sayfa_no}. sayfaya geçiş yapılıyor...", "INFO")
    sayfa_degisti = False

    try:
        # Sayfanın altına in
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight - 200);")
        rastgele_bekle(1, 1.5)

        # LinkedIn pagination ileri butonu seçicileri
        next_button_selectors = [
            "button[aria-label='İleri']",
            "button[aria-label='Next']",
            "button.artdeco-pagination__button--next",
            "li.artdeco-pagination__indicator--number button",
            "a[aria-label='Next']"
        ]

        for selector in next_button_selectors:
            butonlar = driver.find_elements(By.CSS_SELECTOR, selector)
            for btn in butonlar:
                if btn.is_displayed() and btn.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                    rastgele_bekle(0.5, 0.8)
                    btn.click()
                    logger(f"Sayfalama butonuna tıklandı ({selector}).", "INFO")
                    sayfa_degisti = True
                    break
            if sayfa_degisti:
                break

    except Exception as e:
        logger(f"Sayfalama butonuna tıklanamadı: {e}", "WARNING")

    # Eğer butonla geçilemediyse URL ile geç
    if not sayfa_degisti:
        encoded_query = quote_plus(arama_kriteri)
        yeni_url = (
            f"https://www.linkedin.com/search/results/people/"
            f"?keywords={encoded_query}&origin=GLOBAL_SEARCH_HEADER&page={sayfa_no}"
        )
        logger(f"URL ile yönlendiriliyor: {yeni_url}", "INFO")
        driver.get(yeni_url)

    rastgele_bekle(3, 4.5)
    return True


def ismi_akilli_ayristir(tam_isim):
    """
    İsim ve soyisimi akıllı bir şekilde ayrıştırır.
    Gereksiz unvanlar, zamirler (He/Him, Ph.D. vb.) temizlenir.
    """
    if not tam_isim:
        return "", ""

    temiz_isim = re.sub(r"\(.*?\)", "", tam_isim)
    temiz_isim = temiz_isim.split(",")[0].strip()
    temiz_isim = re.sub(r"\s+", " ", temiz_isim).strip()

    parcalar = temiz_isim.split(" ")
    if len(parcalar) == 1:
        return parcalar[0], ""
    else:
        ad = " ".join(parcalar[:-1])
        soyad = parcalar[-1]
        return ad, soyad


def pozisyon_ve_sirket_ayristir(baslik):
    """Headline metninden Pozisyon ve Şirket/Departman ayrıştırmayı dener."""
    if not baslik:
        return "", ""

    ayraclar = [" at ", " @ ", " - ", " | "]
    for ayrac in ayraclar:
        if ayrac in baslik:
            bolumler = baslik.split(ayrac, 1)
            unvan = bolumler[0].strip()
            sirket = bolumler[1].strip()
            return unvan, sirket

    return baslik.strip(), ""


def okul_ve_egitim_ayristir(snippet_listesi, baslik):
    """Kart altındaki özet metinlerden okul/eğitim bilgisini yakalamaya çalışır."""
    anahtar_kelimeler = ["üniversite", "university", "faculty", "lise", "fakülte", "bölüm", "eğitim", "education", "student", "mezun"]
    
    for snippet in snippet_listesi:
        snippet_lower = snippet.lower()
        if any(k in snippet_lower for k in anahtar_kelimeler):
            temiz = re.sub(r"^(eğitim|education|geçmiş|past|current):\s*", "", snippet, flags=re.IGNORECASE)
            return temiz.strip()

    if baslik:
        baslik_lower = baslik.lower()
        if any(k in baslik_lower for k in ["üniversitesi", "university", "college", "institute"]):
            return baslik.strip()

    return "Belirtilmemiş"


def derin_profil_taramasi(driver, profil_url, logger=varsayilan_logger):
    """
    Belirtilen profilin detay sayfasına giderek derinlemesine veri toplar:
    - Hakkında (About) metni
    - Detaylı Deneyim (Experience)
    - Detaylı Eğitim (Education)
    - Varsa İletişim Bilgileri (Web sitesi, e-posta, telefon vb.)
    """
    detaylar = {
        "Hakkında": "",
        "Detaylı Deneyim": "",
        "Detaylı Eğitim": "",
        "İletişim Detayları": ""
    }

    if not profil_url or "/in/" not in profil_url:
        return detaylar

    try:
        logger(f"Derin tarama: {profil_url} ziyaret ediliyor...", "INFO")
        driver.get(profil_url)
        rastgele_bekle(2.5, 4.0)

        # Sayfayı yavaşça biraz kaydır
        driver.execute_script("window.scrollBy(0, 500);")
        rastgele_bekle(0.8, 1.4)

        soup = BeautifulSoup(driver.page_source, "html.parser")

        # 1. Hakkında (About) Alanı
        about_section = soup.select_one("section#about, div#about, div.display-flex.ph5.pv3")
        if about_section:
            about_text = about_section.get_text(" ", strip=True)
            # Başlık kelimelerini temizle
            about_text = re.sub(r"^(Hakkında|About)\s*", "", about_text, flags=re.IGNORECASE)
            detaylar["Hakkında"] = about_text[:500]  # Çok uzunsa makul boyutta sınırla
        else:
            # Alternatif inline-show-more-text
            about_elem = soup.select_one("div.inline-show-more-text--is-collapsed")
            if about_elem:
                detaylar["Hakkında"] = about_elem.get_text(strip=True)[:500]

        # 2. Deneyim (Experience) Alanı
        exp_elements = soup.select("section#experience li.artdeco-list__item, div#experience ~ div li")
        deneyimler = []
        for exp in exp_elements[:3]:  # Son 3 deneyim
            metin = exp.get_text(" | ", strip=True)
            metin = re.sub(r"\s+", " ", metin)
            if metin and len(metin) > 5:
                deneyimler.append(metin)
        if deneyimler:
            detaylar["Detaylı Deneyim"] = " ;; ".join(deneyimler)

        # 3. Eğitim (Education) Alanı
        edu_elements = soup.select("section#education li.artdeco-list__item, div#education ~ div li")
        egitimler = []
        for edu in edu_elements[:2]:  # Son 2 eğitim
            metin = edu.get_text(" | ", strip=True)
            metin = re.sub(r"\s+", " ", metin)
            if metin and len(metin) > 5:
                egitimler.append(metin)
        if egitimler:
            detaylar["Detaylı Eğitim"] = " ;; ".join(egitimler)

        # 4. İletişim Bilgileri (Overlay sayfasına gitmeyi dene)
        try:
            contact_url = profil_url.rstrip("/") + "/overlay/contact-info/"
            driver.get(contact_url)
            rastgele_bekle(1.5, 2.5)
            c_soup = BeautifulSoup(driver.page_source, "html.parser")
            contact_section = c_soup.select_one("div.pv-profile-section__section-info, section.ci-email, div.artdeco-modal__content")
            if contact_section:
                c_links = [a.get("href") for a in contact_section.select("a") if a.get("href")]
                c_text = contact_section.get_text(" | ", strip=True)
                detaylar["İletişim Detayları"] = f"{c_text[:300]} (Linkler: {', '.join(c_links[:2])})"
        except Exception:
            pass

    except Exception as e:
        logger(f"Derin tarama sırasında hata ({profil_url}): {e}", "WARNING")

    return detaylar


def profil_kartlarini_ayristir(driver, derin_tarama=False, logger=varsayilan_logger):
    """
    Mevcut sayfadaki profil kartlarını çok katmanlı ve yedekli seçicilerle çeker.
    LinkedIn'in yeni veya eski tüm arama sonucu DOM yapılarıyla uyumludur.
    """
    sayfa_kaynagi = driver.page_source
    soup = BeautifulSoup(sayfa_kaynagi, "html.parser")
    profil_listesi = []
    görülen_urller = set()

    # 1. Aşama: Bilinen tüm kart konteyner seçicilerini dene
    kartlar = soup.select("li.reusable-search__result-container")
    if not kartlar:
        kartlar = soup.select("div.entity-result")
    if not kartlar:
        kartlar = soup.select("div[data-view-name*='search-entity-result']")
    if not kartlar:
        kartlar = soup.select("ul.reusable-search__entity-result-list > li")
    if not kartlar:
        kartlar = soup.select("li[data-chameleon-result-urn]")

    for kart in kartlar:
        try:
            link_etiketi = kart.select_one("span.entity-result__title-text a") or \
                           kart.select_one("a.app-aware-link[href*='/in/']") or \
                           kart.select_one("a[data-test-app-aware-link][href*='/in/']") or \
                           kart.select_one("a[href*='/in/']")

            if not link_etiketi:
                continue

            ham_url = link_etiketi.get("href", "").strip()
            if not ham_url or "/in/" not in ham_url:
                continue

            parcali_url = urlsplit(ham_url)
            temiz_url = urlunsplit((parcali_url.scheme, parcali_url.netloc, parcali_url.path.rstrip("/"), "", ""))

            if temiz_url in görülen_urller or "/in/ACoAA" in temiz_url:
                continue

            isim_span = link_etiketi.select_one("span[aria-hidden='true']") or \
                        link_etiketi.select_one("span.t-bold") or link_etiketi
            tam_isim = isim_span.get_text(strip=True) if isim_span else ""

            if not tam_isim or ("LinkedIn" in tam_isim and ("Member" in tam_isim or "Üyesi" in tam_isim)):
                continue

            ad, soyad = ismi_akilli_ayristir(tam_isim)

            baslik_etiketi = kart.select_one("div.entity-result__primary-subtitle") or \
                             kart.select_one("div[class*='primary-subtitle']") or \
                             kart.select_one(".entity-result__summary")
            baslik_metni = baslik_etiketi.get_text(strip=True) if baslik_etiketi else ""
            unvan, sirket = pozisyon_ve_sirket_ayristir(baslik_metni)

            lokasyon_etiketi = kart.select_one("div.entity-result__secondary-subtitle") or \
                               kart.select_one("div[class*='secondary-subtitle']")
            lokasyon_metni = lokasyon_etiketi.get_text(strip=True) if lokasyon_etiketi else ""

            snippet_etiketleri = kart.select("p.entity-result__custom-summary, div.entity-result__insights, div[class*='insights']")
            snippet_metinleri = [s.get_text(strip=True) for s in snippet_etiketleri if s]
            okul_bilgisi = okul_ve_egitim_ayristir(snippet_metinleri, baslik_metni)

            if unvan and sirket:
                calistigi_yer = f"{unvan} - {sirket}"
            elif baslik_metni:
                calistigi_yer = baslik_metni
            else:
                calistigi_yer = "Belirtilmemiş"

            görülen_urller.add(temiz_url)
            profil_listesi.append({
                "İsim": ad,
                "Soyisim": soyad,
                "Mezun Olduğu Yer / Okul": okul_bilgisi,
                "Çalıştığı Şirket / Departman / Unvan": calistigi_yer,
                "İletişim Bilgisi / Profil URL": temiz_url,
                "Lokasyon": lokasyon_metni if lokasyon_metni else "Belirtilmemiş",
                "Ham Başlık (Headline)": baslik_metni,
                "Hakkında": "",
                "Detaylı Deneyim": "",
                "Detaylı Eğitim": "",
                "Ek İletişim": ""
            })
        except Exception:
            continue

    # 2. Aşama: Eğer kart konteynerleri bulunamadıysa doğrudan sayfadaki tüm /in/ linklerinden çek (Evrensel Yedek Katman)
    if not profil_listesi:
        logger("Konteyner seçicisi eşleşmedi, evrensel profil bağlantı tarayıcısı devreye girdi...", "INFO")
        tum_in_linkleri = soup.find_all("a", href=re.compile(r"/in/[^/?#]+"))

        for a in tum_in_linkleri:
            ham_url = a.get("href", "")
            if not ham_url or "/in/" not in ham_url:
                continue

            parcali_url = urlsplit(ham_url)
            temiz_url = urlunsplit((parcali_url.scheme, parcali_url.netloc, parcali_url.path.rstrip("/"), "", ""))

            if temiz_url in görülen_urller:
                continue

            # İsim genelde linkin içindeki span'da veya linkin kendi metnindedir
            isim_span = a.select_one("span[aria-hidden='true']") or a
            tam_isim = isim_span.get_text(strip=True) if isim_span else ""

            # Resim veya ikon linklerini (boş veya çok kısa olanları) atla
            if not tam_isim or len(tam_isim) < 2 or "LinkedIn" in tam_isim:
                continue

            ad, soyad = ismi_akilli_ayristir(tam_isim)

            # Bu linkin ait olduğu ebeveyn bloğunu bul
            parent_kart = a.find_parent("li") or a.find_parent("div", class_=re.compile(r"result|entity|card"))
            baslik_metni = ""
            lokasyon_metni = ""
            okul_bilgisi = "Belirtilmemiş"

            if parent_kart:
                b_elem = parent_kart.select_one("div[class*='primary-subtitle'], div[class*='subtitle'], p")
                if b_elem:
                    baslik_metni = b_elem.get_text(strip=True)
                l_elem = parent_kart.select_one("div[class*='secondary-subtitle'], div[class*='caption']")
                if l_elem:
                    lokasyon_metni = l_elem.get_text(strip=True)
                snip_elems = parent_kart.select("p, div[class*='insights']")
                snip_texts = [s.get_text(strip=True) for s in snip_elems]
                okul_bilgisi = okul_ve_egitim_ayristir(snip_texts, baslik_metni)

            unvan, sirket = pozisyon_ve_sirket_ayristir(baslik_metni)
            if unvan and sirket:
                calistigi_yer = f"{unvan} - {sirket}"
            elif baslik_metni:
                calistigi_yer = baslik_metni
            else:
                calistigi_yer = "Belirtilmemiş"

            görülen_urller.add(temiz_url)
            profil_listesi.append({
                "İsim": ad,
                "Soyisim": soyad,
                "Mezun Olduğu Yer / Okul": okul_bilgisi,
                "Çalıştığı Şirket / Departman / Unvan": calistigi_yer,
                "İletişim Bilgisi / Profil URL": temiz_url,
                "Lokasyon": lokasyon_metni if lokasyon_metni else "Belirtilmemiş",
                "Ham Başlık (Headline)": baslik_metni,
                "Hakkında": "",
                "Detaylı Deneyim": "",
                "Detaylı Eğitim": "",
                "Ek İletişim": ""
            })

    logger(f"Sayfadan başarıyla {len(profil_listesi)} adet profil ayrıştırıldı.", "INFO")

    # Derin tarama aktifse ve profiller varsa detay sayfalarını ziyaret et
    if derin_tarama and profil_listesi:
        logger(f"Derin profil taraması başlatılıyor ({len(profil_listesi)} profil taranacak)...", "INFO")
        for i, p in enumerate(profil_listesi, start=1):
            logger(f"[{i}/{len(profil_listesi)}] Derin analiz: {p['İsim']} {p['Soyisim']}", "INFO")
            detaylar = derin_profil_taramasi(driver, p["İletişim Bilgisi / Profil URL"], logger=logger)
            p["Hakkında"] = detaylar.get("Hakkında", "")
            if detaylar.get("Detaylı Deneyim"):
                p["Detaylı Deneyim"] = detaylar["Detaylı Deneyim"]
            if detaylar.get("Detaylı Eğitim"):
                p["Detaylı Eğitim"] = detaylar["Detaylı Eğitim"]
                if p["Mezun Olduğu Yer / Okul"] == "Belirtilmemiş":
                    p["Mezun Olduğu Yer / Okul"] = detaylar["Detaylı Eğitim"][:100]
            if detaylar.get("İletişim Detayları"):
                p["Ek İletişim"] = detaylar["İletişim Detayları"]

            rastgele_bekle(1.5, 3.0)

    return profil_listesi


def verileri_filtrele_ve_temizle(veri_listesi, filtre_anahtar="", haric_anahtar="", filtre_sehir=""):
    """
    Verileri Excel'e kaydetmeden önce filtrelere göre süzer ve mükerrer kayıtları temizler.
    """
    if not veri_listesi:
        return pd.DataFrame(), 0, 0

    df = pd.DataFrame(veri_listesi)
    toplam_ham = len(df)

    # 1. Mükerrer (Duplicate) Kayıt Temizliği
    df.drop_duplicates(subset=["İletişim Bilgisi / Profil URL"], inplace=True)
    mukerrer_sayisi = toplam_ham - len(df)

    # 2. Dahil Edilecek Anahtar Kelime Filtresi (Başlık, Unvan, Okul, vb.)
    if filtre_anahtar and filtre_anahtar.strip():
        kelimeler = [k.strip().lower() for k in filtre_anahtar.split(",") if k.strip()]
        if kelimeler:
            filtre_maskesi = df.apply(
                lambda row: any(
                    any(k in str(row[col]).lower() for k in kelimeler)
                    for col in ["Çalıştığı Şirket / Departman / Unvan", "Mezun Olduğu Yer / Okul", "Ham Başlık (Headline)", "Lokasyon"]
                ),
                axis=1
            )
            df = df[filtre_maskesi]

    # 3. Hariç Tutulacak Anahtar Kelime Filtresi
    if haric_anahtar and haric_anahtar.strip():
        kelimeler_haric = [k.strip().lower() for k in haric_anahtar.split(",") if k.strip()]
        if kelimeler_haric:
            haric_maskesi = df.apply(
                lambda row: not any(
                    any(k in str(row[col]).lower() for k in kelimeler_haric)
                    for col in ["Çalıştığı Şirket / Departman / Unvan", "Mezun Olduğu Yer / Okul", "Ham Başlık (Headline)"]
                ),
                axis=1
            )
            df = df[haric_maskesi]

    # 4. Lokasyon / Şehir Filtresi
    if filtre_sehir and filtre_sehir.strip():
        sehirler = [s.strip().lower() for s in filtre_sehir.split(",") if s.strip()]
        if sehirler:
            sehir_maskesi = df["Lokasyon"].astype(str).str.lower().apply(
                lambda loc: any(s in loc for s in sehirler)
            )
            df = df[sehir_maskesi]

    return df, toplam_ham, mukerrer_sayisi


def dataframe_to_styled_excel(df):
    """DataFrame'i profesyonel stillendirilmiş Excel binary verisi (bytes) olarak döndürür."""
    output = io.BytesIO()

    # İstenen zorunlu sütun sırası
    sutun_sirasi = [
        "İsim",
        "Soyisim",
        "Mezun Olduğu Yer / Okul",
        "Çalıştığı Şirket / Departman / Unvan",
        "İletişim Bilgisi / Profil URL",
        "Lokasyon",
        "Ham Başlık (Headline)",
        "Hakkında",
        "Detaylı Deneyim",
        "Detaylı Eğitim",
        "Ek İletişim"
    ]
    sirali_sutunlar = [col for col in sutun_sirasi if col in df.columns]
    kalan_sutunlar = [col for col in df.columns if col not in sirali_sutunlar]
    final_df = df[sirali_sutunlar + kalan_sutunlar]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="LinkedIn Profilleri", index=False)

        if HAS_OPENPYXL:
            workbook = writer.book
            worksheet = writer.sheets["LinkedIn Profilleri"]

            baslik_fontu = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            baslik_dolgusu = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            govde_fontu = Font(name="Segoe UI", size=10)
            hizalama = Alignment(vertical="center", wrap_text=True)
            ince_kenarlik = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9")
            )

            for col_idx, col in enumerate(worksheet.columns, start=1):
                header_cell = col[0]
                header_cell.font = baslik_fontu
                header_cell.fill = baslik_dolgusu
                header_cell.alignment = Alignment(horizontal="center", vertical="center")

                max_len = len(str(header_cell.value or ""))
                for cell in col[1:]:
                    cell.font = govde_fontu
                    cell.alignment = hizalama
                    cell.border = ince_kenarlik
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)

                col_letter = get_column_letter(col_idx)
                worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 55)

            worksheet.row_dimensions[1].height = 28
            for row in range(2, worksheet.max_row + 1):
                worksheet.row_dimensions[row].height = 24

    output.seek(0)
    return output.getvalue()


def sonuclari_excele_kaydet(veriler, dosya_adi="linkedin_detayli_sonuclar.xlsx", logger=varsayilan_logger):
    """Toplanan verileri diskteki Excel dosyasına kaydeder."""
    if not veriler:
        logger("Kaydedilecek veri bulunamadı.", "WARNING")
        return False

    if isinstance(veriler, list):
        df = pd.DataFrame(veriler)
    else:
        df = veriler

    excel_bytes = dataframe_to_styled_excel(df)
    with open(dosya_adi, "wb") as f:
        f.write(excel_bytes)

    logger(f"Başarılı! Toplam {len(df)} adet profil verisi '{dosya_adi}' dosyasına kaydedildi.", "SUCCESS")
    return True
