# -*- coding: utf-8 -*-
import time
from linkedin_scraper import tarayiciyi_baslat

def main():
    print("=" * 60)
    print("   LINKEDIN GIRIS PENCERESI")
    print("=" * 60)
    print("\n[+] Chrome tarayicisi aciliyor...")

    driver = tarayiciyi_baslat(headless=False)
    driver.get("https://www.linkedin.com/login")

    print("\n[!] Chrome acildi ve LinkedIn giris sayfasi yuklendi.")
    print("[*] Lutfen tarayicidan hesabiniza giris yapin.")
    print("[*] Giris yaptiginizda program bunu otomatik algilayacaktir.")

    while True:
        try:
            url = driver.current_url
            if any(k in url for k in ["feed", "check/add-phone", "mynetwork", "messaging"]):
                print("\n[+] TEBRIKLER! LinkedIn oturumu basariyla acildi.")
                print("[*] Bu pencere 5 saniye icinde kapanacaktir.")
                time.sleep(5)
                break
            time.sleep(1.5)
        except Exception:
            break

    try:
        driver.quit()
    except Exception:
        pass

if __name__ == "__main__":
    main()
