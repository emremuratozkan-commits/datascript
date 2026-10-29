# LinkedIn Akıllı Veri Çekme, Derin Profil Taraması ve Excel Platformu

Python, Selenium, BeautifulSoup4, Pandas ve **Streamlit** kullanılarak geliştirilmiş; hem **Masaüstü** hem de **Mobil** cihazlarla uyumlu, çoklu sayfalama, derin profil analizi, veri filtreleme ve profesyonel Excel raporlama sunan uçtan uca LinkedIn otomasyon platformu.

---

## 🌟 Öne Çıkan Özellikler

1. **📱 Mobil ve Masaüstü Uyumlu Modern Streamlit Arayüzü:**
   - Web tarayıcısından veya mobil cihazlardan erişilebilen kullanıcı dostu arayüz.
   - LinkedIn giriş bilgilerinizi, arama kriterlerinizi, taranacak sayfa sayısını ve filtreleri zahmetsizce yönetebileceğiniz sekmeli tasarım.

2. **📄 Çoklu Sayfa Desteği (Pagination):**
   - Belirlenen sayfa sayısı kadar (1-10 sayfa) LinkedIn sonuçları arasında otomatik ilerler.
   - Sayfalama butonlarına (`Sonraki / Next`) basarak ve URL yönlendirme desteğiyle kesintisiz tarama yapar.

3. **🔎 Derin Profil Taraması (Deep Scraping):**
   - Arama listesiyle sınırlı kalmaz; profillerin detay sayfalarına giderek:
     - **Hakkında (About)** metnini,
     - **Detaylı Deneyimler** (Şirket, unvan, tarihler),
     - **Detaylı Eğitim** (Üniversite, bölüm, derece),
     - **İletişim Detayları** (E-posta, web sitesi bağlantıları vb.) çeker.

4. **🧹 Akıllı Veri Filtreleme ve Mükerrer (Duplicate) Temizliği:**
   - Excel'e aktarılmadan önce arayüzden belirlenebilen kriterler:
     - **Dahil edilecek anahtar kelimeler** (Örn: `Senior, Lead, Mühendis`)
     - **Hariç tutulacak anahtar kelimeler** (Örn: `Stajyer, Intern`)
     - **Şehir / Lokasyon Filtresi** (Örn: `İstanbul, Ankara, Remote`)
   - Mükerrer profilleri temizler ve kaç adet profilin elendiğini canlı metriklerle raporlar.

5. **🛡️ Hata Yönetimi ve Canlı Log Paneli:**
   - Olası CAPTCHA, internet kopması veya element bulunamaması durumunda çökmez; yedek seçiciler ve sağlam hata bloklarıyla çalışmaya devam eder.
   - Canlı log paneli ile işlemler anlık olarak (INFO, SUCCESS, WARNING, ERROR) takip edilir.

6. **📊 Profesyonel Excel Çıktısı & Tek Tıkla İndirme:**
   - `openpyxl` ile stillendirilmiş (lacivert başlıklar, otomatik sütun genişlikleri, kenarlıklar) `.xlsx` dosyası.
   - Doğrudan web arayüzünden tarayıcıya/telefona tek tıkla indirme veya sunucu klasörüne kaydetme seçeneği.

---

## 📦 Kurulum

Terminalde proje klasörüne gidin ve gerekli bağımlılıkları yükleyin:

```bash
pip install -r requirements.txt
```

*(Gereken paketler: `selenium`, `beautifulsoup4`, `pandas`, `openpyxl`, `webdriver-manager`, `streamlit`)*

---

## 🚀 Çalıştırma

### 1. Web Arayüzü ile Başlatma (Tavsiye Edilen)
Modern Streamlit arayüzünü açmak için terminalde şu komutu verin:

```bash
streamlit run app.py
```

Komutu çalıştırdıktan sonra tarayıcınızda otomatik olarak açılacaktır (Genellikle `http://localhost:8501`). Aynı ağdaki cep telefonunuzdan erişmek için konsolda gösterilen `Network URL` adresini kullanabilirsiniz.

### 2. Konsol (CLI) Üzerinden Başlatma
İsterseniz arayüzsüz klasik komut satırı modunda da çalıştırabilirsiniz:

```bash
python linkedin_scraper.py
```

---

## 💡 Güvenlik & CAPTCHA İpuçları
- LinkedIn ilk giriş anında CAPTCHA veya 2FA SMS/E-posta doğrulama kodu sorabilir. Bu nedenle arayüzdeki **"Arka Planda Çalıştır (Headless)"** seçeneğini ilk etapta kapalı tutarak açılan Chrome penceresinden doğrulamayı kolayca tamamlayabilirsiniz.
- Hesap güvenliğiniz için arama istekleri arasında otomatik rastgele insansı gecikmeler uygulanır.
