# -*- coding: utf-8 -*-
"""
LinkedIn Veri Çekme ve Excel Raporlama - Streamlit Web Arayüzü
Mobil ve Masaüstü Uyumlu Modern UI
Geliştirici: Antigravity
"""

import time
import queue
from urllib.parse import quote_plus
import pandas as pd
import streamlit as st

# Çekirdek kazıma modülü
from linkedin_scraper import (
    tarayiciyi_baslat,
    linkedin_giris_yap,
    linkedin_cerezle_giris_yap,
    linkedin_manuel_giris_bekle,
    sayfayi_asagi_kaydir,
    sonraki_sayfaya_gec,
    profil_kartlarini_ayristir,
    verileri_filtrele_ve_temizle,
    dataframe_to_styled_excel,
    sonuclari_excele_kaydet,
    rastgele_bekle
)

# Sayfa Yapılandırması (Mobil ve Masaüstü Uyumlu)
st.set_page_config(
    page_title="LinkedIn Veri Çekme Otomasyonu",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Özel CSS: Mobil duyarlılık, modern kartlar ve log alanı
st.markdown("""
<style>
    /* Ana sayfa ve mobil ayarları */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 1.5rem;
        padding-right: 1.5rem;
        max-width: 100% !important;
    }
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 0.8rem;
            padding-right: 0.8rem;
        }
    }
    
    /* Başlık ve Kart Tasarımı */
    .app-header {
        background: linear-gradient(135deg, #0A66C2 0%, #004182 100%);
        color: white;
        padding: 1.2rem 1.6rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(10, 102, 194, 0.15);
    }
    .app-header h1 {
        color: white !important;
        font-size: 1.8rem;
        margin: 0;
        font-weight: 700;
    }
    .app-header p {
        color: #E1E9EE;
        margin-top: 0.3rem;
        margin-bottom: 0;
        font-size: 0.95rem;
    }
    
    /* Metrik Kartları */
    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        margin-bottom: 1rem;
    }
    .metric-val {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0A66C2;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 500;
    }

    /* Canlı Log Konsolu */
    .log-container {
        background-color: #0f172a;
        color: #f1f5f9;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 0.84rem;
        padding: 1rem;
        border-radius: 8px;
        height: 280px;
        overflow-y: auto;
        white-space: pre-wrap;
        border: 1px solid #334155;
    }
    .log-INFO { color: #38bdf8; }
    .log-SUCCESS { color: #4ade80; font-weight: bold; }
    .log-WARNING { color: #facc15; }
    .log-ERROR { color: #f87171; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# Oturum Durumu (Session State) Başlatma
if "raw_profiles" not in st.session_state:
    st.session_state.raw_profiles = []
if "filtered_df" not in st.session_state:
    st.session_state.filtered_df = pd.DataFrame()
if "logs" not in st.session_state:
    st.session_state.logs = []
if "scraping_active" not in st.session_state:
    st.session_state.scraping_active = False


def log_ekle(mesaj, seviye="INFO"):
    """Canlı log listesine yeni mesaj ekler."""
    zaman = time.strftime("%H:%M:%S")
    st.session_state.logs.append({
        "zaman": zaman,
        "seviye": seviye,
        "mesaj": mesaj
    })


# ---------------- SIDEBAR: GİRİŞ VE AYARLAR ----------------
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3536/3536505.png", width=50)
    st.title("Giriş & Ayarlar")

    giris_yontemi = st.radio(
        "Giriş Yöntemini Seçin:",
        [
            "🔑 Manuel Giriş (Önerilen & En Güvenli)",
            "🍪 Çerez ile Giriş (li_at Cookie)",
            "🤖 Otomatik Giriş (E-posta + Şifre)"
        ],
        index=0,
        help="LinkedIn bot korumasına takılmamak ve hesabın kilitlenmesini önlemek için 'Manuel Giriş' en sorunsuz yöntemdir."
    )

    email = ""
    sifre = ""
    li_at_cookie = ""

    if giris_yontemi == "🔑 Manuel Giriş (Önerilen & En Güvenli)":
        st.info("💡 **Nasıl Çalışır?**\n1. 'Arama & Otomasyon' sekmesinden kriterinizi girin.\n2. **'Taramayı Başlat'** butonuna basın.\n3. Chrome anında açılıp LinkedIn giriş sayfasını yükleyecektir.\n4. Kendi elinizle giriş yaptığınız anda bot otomatik olarak aramaya geçecektir!")
    elif giris_yontemi == "🍪 Çerez ile Giriş (li_at Cookie)":
        li_at_cookie = st.text_input(
            "LinkedIn 'li_at' Çerezi:",
            type="password",
            placeholder="AQEDATXXXX...",
            help="Tarayıcınızda F12 -> Application -> Cookies -> li_at değerini yapıştırın."
        )
        st.caption("🔒 Şifre, SMS kodu veya CAPTCHA olmadan doğrudan oturum açar.")
    else:
        email = st.text_input("E-posta veya Telefon", placeholder="ornek@domain.com")
        sifre = st.text_input("LinkedIn Şifresi", type="password", placeholder="••••••••")

    st.markdown("---")
    st.subheader("⚙️ Tarayıcı Ayarları")
    headless_secim = st.checkbox(
        "Arka Planda Çalıştır (Headless)", 
        value=False,
        disabled=(giris_yontemi == "🔑 Manuel Giriş (Önerilen & En Güvenli)"),
        help="Manuel giriş modunda tarayıcıyı görebilmeniz için açık kalmalıdır."
    )
    if giris_yontemi == "🔑 Manuel Giriş (Önerilen & En Güvenli)":
        headless_secim = False


# ---------------- ANA EKRAN BAŞLIK ----------------
st.markdown("""
<div class="app-header">
    <h1>💼 LinkedIn Veri Çekme & Otomasyon Platformu</h1>
    <p>Dinamik arama kriterleri, çoklu sayfalama, derin profil taraması ve Excel raporlama.</p>
</div>
""", unsafe_allow_html=True)

# Sekmeler
tab_arama, tab_filtre, tab_sonuc, tab_log = st.tabs([
    "🚀 1. Arama & Otomasyon",
    "🧹 2. Filtreleme & Temizleme",
    "📊 3. Sonuçlar & Excel İndirme",
    "📜 4. Canlı Log Paneli"
])


# ---------------- TAB 1: ARAMA & OTOMASYON ----------------
with tab_arama:
    st.subheader("🔍 Arama Kriterleri ve Tarama Parametreleri")

    arama_modu = st.radio(
        "Arama Yöntemi:",
        ["🎛️ Kategorilere Göre Arama (Önerilen)", "✍️ Serbest Metin ile Arama"],
        horizontal=True
    )

    arama_kriteri = ""

    if arama_modu == "🎛️ Kategorilere Göre Arama (Önerilen)":
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            kat_sirket = st.text_input("🏢 Şirket / Kurum Adı", placeholder="Örn: Google, Trendyol, Aselsan")
            kat_unvan = st.text_input("💼 Pozisyon / Meslek Unvanı", placeholder="Örn: Yazılım Mühendisi, Veri Analisti, İK")
            kat_kidem = st.selectbox(
                "⭐ Kıdem / Deneyim Seviyesi",
                ["Tümü / Farketmez", "Stajyer / Intern", "Junior", "Mid-Level", "Senior", "Lead / Takım Lideri", "Manager / Yönetici", "Director / C-Level"]
            )
        with col_k2:
            kat_okul = st.text_input("🎓 Üniversite / Okul", placeholder="Örn: ODTÜ, İTÜ, Boğaziçi, Yıldız Teknik")
            kat_bolum = st.text_input("📚 Bölüm / Alan", placeholder="Örn: Bilgisayar Mühendisliği, İşletme")

        # Kategorileri doğal ve esnek arama sorgusu haline getirelim
        sorgu_parcalari = []
        if kat_sirket.strip():
            sorgu_parcalari.append(kat_sirket.strip())
        if kat_kidem != "Tümü / Farketmez":
            temiz_kidem = kat_kidem.split(" / ")[0].strip()
            sorgu_parcalari.append(temiz_kidem)
        if kat_unvan.strip():
            sorgu_parcalari.append(kat_unvan.strip())
        if kat_okul.strip():
            sorgu_parcalari.append(kat_okul.strip())
        if kat_bolum.strip():
            sorgu_parcalari.append(kat_bolum.strip())

        arama_kriteri = " ".join(sorgu_parcalari) if sorgu_parcalari else "Yazılım Mühendisi"
        
        st.markdown(f"🎯 **Oluşturulan Arama Sorgusu:** `{arama_kriteri}`")

    else:
        arama_kriteri = st.text_input(
            "Aramak İstediğiniz Kriterler (Serbest Arama)",
            value="Google Yazılım Mühendisi",
            placeholder="Örn: Trendyol Backend Developer veya ODTÜ Bilgisayar"
        )

    st.markdown("---")
    col_p1, col_p2 = st.columns([1, 1])

    with col_p1:
        toplam_sayfa = st.slider(
            "Taranacak Sayfa Sayısı (Pagination)",
            min_value=1,
            max_value=10,
            value=2,
            help="LinkedIn arama sonuçlarında kaç sayfa ilerleneceğini belirler (Her sayfa ~10 profil içerir)."
        )

    with col_p2:
        st.write("")
        derin_tarama_aktif = st.checkbox(
            "🔎 Derin Profil Taraması (Deep Scraping)",
            value=False,
            help="İşaretlenirse bot sadece arama listesiyle yetinmez; her bir profilin sayfasına giderek 'Hakkında', detaylı deneyimler ve iletişim bilgilerini toplar (Süre uzayabilir)."
        )

    st.markdown("---")

    baslat_col, durdur_col = st.columns([1, 4])
    with baslat_col:
        baslat_butonu = st.button("🚀 Taramayı Başlat", type="primary", use_container_width=True)

    # Otomasyon Yürütme Bloğu
    if baslat_butonu:
        if giris_yontemi == "🤖 Otomatik Giriş (E-posta + Şifre)" and (not email or not sifre):
            st.error("Lütfen sol taraftaki menüden LinkedIn e-posta ve şifrenizi girin!")
        elif giris_yontemi == "🍪 Çerez ile Giriş (li_at Cookie)" and not li_at_cookie.strip():
            st.error("Lütfen sol taraftaki menüden 'li_at' çerez değerini yapıştırın!")
        elif not arama_kriteri.strip():
            st.error("Lütfen bir arama kriteri girin!")
        else:
            st.session_state.scraping_active = True
            st.session_state.raw_profiles = []
            st.session_state.logs = []

            ilerleme_cubugu = st.progress(0, text="Otomasyon başlatılıyor...")
            log_alani = st.empty()

            def canli_logger(mesaj, seviye="INFO"):
                log_ekle(mesaj, seviye)
                son_loglar = st.session_state.logs[-8:]
                html = "<div class='log-container'>"
                for l in son_loglar:
                    html += f"<span class='log-{l['seviye']}'>[{l['zaman']}] [{l['seviye']}] {l['mesaj']}</span>\n"
                html += "</div>"
                log_alani.markdown(html, unsafe_allow_html=True)

            driver = None
            try:
                canli_logger("Tarayıcı hazırlanıyor...", "INFO")
                driver = tarayiciyi_baslat(headless=headless_secim, logger=canli_logger)

                giris_ok = False
                if giris_yontemi == "🔑 Manuel Giriş (Önerilen & En Güvenli)":
                    canli_logger("Lütfen açılan Chrome penceresinde LinkedIn hesabınıza giriş yapın...", "WARNING")
                    giris_ok = linkedin_manuel_giris_bekle(driver, logger=canli_logger, max_bekleme=90)
                elif giris_yontemi == "🍪 Çerez ile Giriş (li_at Cookie)":
                    canli_logger("Çerez doğrulaması yapılıyor...", "INFO")
                    giris_ok = linkedin_cerezle_giris_yap(driver, li_at_cookie, logger=canli_logger)
                else:
                    canli_logger("Otomatik form dolduruluyor...", "INFO")
                    giris_ok = linkedin_giris_yap(driver, email, sifre, logger=canli_logger, max_bekleme=30)

                if not giris_ok:
                    canli_logger("Giriş adımı tamamlanamadı!", "ERROR")
                    st.error("LinkedIn'e giriş yapılamadı. Lütfen 'Manuel Giriş' yöntemini deneyin.")
                else:
                    canli_logger(f"Arama başlatılıyor: '{arama_kriteri}'", "SUCCESS")
                    encoded_query = quote_plus(arama_kriteri.strip())

                    # İlk sayfaya git
                    ilk_url = (
                        f"https://www.linkedin.com/search/results/people/"
                        f"?keywords={encoded_query}&origin=GLOBAL_SEARCH_HEADER&page=1"
                    )
                    driver.get(ilk_url)
                    rastgele_bekle(3, 4.5)

                    for sayfa_idx in range(1, toplam_sayfa + 1):
                        yuzde = int((sayfa_idx - 1) / toplam_sayfa * 100)
                        ilerleme_cubugu.progress(yuzde, text=f"Sayfa {sayfa_idx}/{toplam_sayfa} taranıyor...")

                        canli_logger(f"--- {sayfa_idx}. Sayfa Taranıyor ---", "INFO")

                        # Sonsuz kaydırma
                        sayfayi_asagi_kaydir(driver, kaydirma_sayisi=6, logger=canli_logger)

                        # Kartları çek
                        yeni_profiller = profil_kartlarini_ayristir(
                            driver, 
                            derin_tarama=derin_tarama_aktif, 
                            logger=canli_logger
                        )
                        st.session_state.raw_profiles.extend(yeni_profiller)
                        canli_logger(f"Bu sayfadan {len(yeni_profiller)} profil eklendi. Toplam: {len(st.session_state.raw_profiles)}", "SUCCESS")

                        # Sonraki sayfaya geçiş (son sayfa değilse)
                        if sayfa_idx < toplam_sayfa:
                            canli_logger(f"Sonraki sayfaya ({sayfa_idx + 1}) geçiliyor...", "INFO")
                            sonraki_sayfaya_gec(driver, sayfa_idx + 1, arama_kriteri, logger=canli_logger)
                            rastgele_bekle(3, 5)

                    ilerleme_cubugu.progress(100, text="Tarama tamamlandı!")
                    canli_logger(f"İşlem tamamlandı! Toplam {len(st.session_state.raw_profiles)} profil toplandı.", "SUCCESS")
                    
                    # Otomatik varsayılan filtreleme uygula
                    df_temiz, _, _ = verileri_filtrele_ve_temizle(st.session_state.raw_profiles)
                    st.session_state.filtered_df = df_temiz

                    st.success("Tüm sayfalar başarıyla tarandı! 'Filtreleme' veya 'Sonuçlar' sekmesinden verileri inceleyebilirsiniz.")

            except Exception as genel_hata:
                canli_logger(f"Beklenmeyen bir hata oluştu: {genel_hata}", "ERROR")
                st.error(f"Hata detayı: {genel_hata}")
            finally:
                if driver:
                    canli_logger("Tarayıcı kapatılıyor...", "INFO")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                st.session_state.scraping_active = False


# ---------------- TAB 2: FİLTRELEME & TEMİZLEME ----------------
with tab_filtre:
    st.subheader("🧹 Akıllı Veri Filtreleme ve Mükerrer Kayıt Temizliği")
    st.markdown("Excel çıktısını oluşturmadan önce kriterlerinize uymayan verileri eleyebilir ve mükerrerleri temizleyebilirsiniz.")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        dahil_kelimeler = st.text_input(
            "İçermesi Gereken Anahtar Kelimeler (Virgülle ayırın)",
            value="",
            placeholder="Örn: Senior, Lead, Mühendis, Engineer"
        )
        sehir_kelimeleri = st.text_input(
            "Şehir / Lokasyon Filtresi (Virgülle ayırın)",
            value="",
            placeholder="Örn: İstanbul, Ankara, İzmir, Remote, Turkey"
        )

    with col_f2:
        haric_kelimeler = st.text_input(
            "Hariç Tutulacak Anahtar Kelimeler (Virgülle ayırın)",
            value="",
            placeholder="Örn: Stajyer, Intern, Assistant"
        )

    filtrele_butonu = st.button("🔍 Filtreleri ve Temizliği Uygula", type="primary")

    if filtrele_butonu or not st.session_state.filtered_df.empty:
        if st.session_state.raw_profiles:
            df_filtreli, toplam_ham, mukerrer = verileri_filtrele_ve_temizle(
                st.session_state.raw_profiles,
                filtre_anahtar=dahil_kelimeler,
                haric_anahtar=haric_kelimeler,
                filtre_sehir=sehir_kelimeleri
            )
            st.session_state.filtered_df = df_filtreli

            st.markdown("### 📈 Filtreleme İstatistikleri")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-val">{toplam_ham}</div>
                    <div class="metric-label">Toplam Toplanan Ham Profil</div>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-val" style="color: #ef4444;">-{mukerrer}</div>
                    <div class="metric-label">Temizlenen Mükerrer Profil</div>
                </div>
                """, unsafe_allow_html=True)
            with c3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-val" style="color: #22c55e;">{len(df_filtreli)}</div>
                    <div class="metric-label">Kalan Filtrelenmiş Profil</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.warning("Henüz çekilmiş bir veri bulunmuyor. Lütfen önce 1. sekmeden taramayı başlatın.")


# ---------------- TAB 3: SONUÇLAR & EXCEL İNDİRME ----------------
with tab_sonuc:
    st.subheader("📊 Yapılandırılmış Profil Tablosu ve Excel Aktarımı")

    aktif_df = st.session_state.filtered_df

    if not aktif_df.empty:
        st.write(f"**Görüntülenen Profil Sayısı:** `{len(aktif_df)}`")

        # Excel İndirme Butonu (Doğrudan Tarayıcıdan ve Mobilden)
        excel_bytes = dataframe_to_styled_excel(aktif_df)
        
        col_down1, col_down2 = st.columns([1, 2])
        with col_down1:
            st.download_button(
                label="📥 Excel Dosyasını İndir (.xlsx)",
                data=excel_bytes,
                file_name="linkedin_detayli_sonuclar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with col_down2:
            if st.button("💾 Sunucuya Kaydet (Çalışma Klasörüne)", use_container_width=True):
                sonuclari_excele_kaydet(aktif_df, dosya_adi="linkedin_detayli_sonuclar.xlsx")
                st.success("Dosya 'linkedin_detayli_sonuclar.xlsx' olarak kaydedildi!")

        st.markdown("---")
        st.dataframe(aktif_df, use_container_width=True, height=450)

    else:
        st.info("Henüz görüntülenecek veri yok. Lütfen 'Arama & Otomasyon' sekmesinden işlemi başlatın.")


# ---------------- TAB 4: CANLI LOG PANELİ ----------------
with tab_log:
    st.subheader("📜 Canlı İşlem Günlüğü ve Sistem Durumu")
    
    col_l1, col_l2 = st.columns([1, 4])
    with col_l1:
        if st.button("Logları Temizle"):
            st.session_state.logs = []
            st.rerun()

    if st.session_state.logs:
        html = "<div class='log-container' style='height: 480px;'>"
        for l in st.session_state.logs:
            html += f"<span class='log-{l['seviye']}'>[{l['zaman']}] [{l['seviye']}] {l['mesaj']}</span>\n"
        html += "</div>"
        st.markdown(html, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="log-container" style="height: 200px;">
        Sistem hazır. Tarama başladığında canlı loglar burada anlık olarak akacaktır.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    ---
    ### 🛡️ Olası Sorunlar ve Çözümler
    - **CAPTCHA / Güvenlik Doğrulaması:** LinkedIn bazen sıradışı hareket algılayıp doğrulama isteyebilir. Arayüzden *Headless* seçeneğini kapatıp açılan Chrome ekranından doğrulamayı tamamlayabilirsiniz.
    - **Element Bulunamadı Hatası:** LinkedIn periyodik olarak DOM yapısını günceller. Scriptimiz çoklu yedekli (fallback) seçiciler ve `try-except` blokları kullanarak çökmeden çalışmaya devam eder.
    - **Rate Limit Koruması:** Çok fazla sayfa taraması yaparken oturumunuzun kısıtlanmaması için her istek arasında insan benzeri rastgele gecikmeler uygulanmaktadır.
    """)
