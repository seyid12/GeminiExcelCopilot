# Gemini Excel Copilot (Veri Asistanı)

**Veri Asistanı**, Microsoft Excel'i Google'ın en yeni **Gemini 2.5 Flash** yapay zeka modeli ile güçlendiren, C# ve VSTO (Visual Studio Tools for Office) mimarisi üzerine inşa edilmiş kapsamlı bir eklentidir.

Bu proje, Excel'e sadece bir "Yan Panel" eklemekle kalmaz; aynı zamanda C# ve VBA arasında kurduğu özel **COM Automation Köprüsü** sayesinde, yapay zekayı doğrudan hücre içinde bir formül gibi kullanmanıza (`=GEMINI()`) olanak tanır.

![Proje Durumu](https://img.shields.io/badge/Durum-v1.1%20Yayında-success)
![Lisans](https://img.shields.io/badge/Lisans-MIT-blue)
![Yapay Zeka](https://img.shields.io/badge/Model-Gemini%202.5%20Flash-orange)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Excel-lightgrey)

## 🚀 Öne Çıkan Özellikler

### 1. 📊 Görsel ve Zengin Veri Analizi
* Veri setlerini analiz ederken sıkıcı düz metinler yerine **HTML ve Markdown** destekli raporlar sunar.
* Yan panelde kalın başlıklar, düzenli listeler ve tablolar içeren şık bir görünüm sağlar (`WebBrowser` entegrasyonu).

### 2. 🔗 Hücre İçi Yapay Zeka Fonksiyonu (UDF)
* Excel'e yerleşik olmayan bir yetenek kazandırır: **Doğrudan hücre içinde AI kullanımı.**
* **Kullanım:** `=GEMINI("Bu metni İngilizceye çevir"; A1)`
* **Teknoloji:** C# tarafındaki metodları VBA üzerinden çağırarak (COM Interop) çalışır.

### 3. 📈 Akıllı Grafik Motoru
* Veriyi rastgele çizmez; içeriğini analiz ederek en uygun grafik türüne karar verir.
* **Zaman Serisi** algılarsa -> Çizgi Grafik (Line Chart) 📈
* **Kategorik Veri** algılarsa -> Sütun Grafik (Column Chart) 📊
* **Parça/Bütün** algılarsa -> Pasta Grafik (Pie Chart) 🍰

### 4. 🧮 Doğal Dil ile Formül Üretimi
* *"A sütunundaki değerler 100'den büyükse topla"* gibi bir isteği Excel formülüne dönüştürür.
* Excel'in dilini (Türkçe/İngilizce) otomatik algılar ve formülü ona göre yazar (`=ETOPLA` veya `=SUMIF`).

---

## 🛠️ Kurulum ve Geliştirme

Projeyi kendi bilgisayarınızda çalıştırmak veya geliştirmek için aşağıdaki adımları izleyin.

### Gereksinimler
* **IDE:** Visual Studio 2022 (Workload: *Office/SharePoint Development*)
* **Framework:** .NET Framework 4.7.2 veya 4.8
* **API:** [Google AI Studio](https://aistudio.google.com/)'dan alınmış ücretsiz bir API Anahtarı.

### Adım Adım Kurulum

1.  **Repoyu Klonlayın:**
    ```bash
    git clone [https://github.com/seyid12/GeminiExcelCopilot.git](https://github.com/seyid12/GeminiExcelCopilot.git)
    ```

2.  **Paketleri Yükleyin:**
    Visual Studio'da projeyi açın (`.sln`). Solution Explorer'da projeye sağ tıklayıp **"Manage NuGet Packages"** diyerek şunları yükleyin/güncelleyin:
    * `Google.Ai.GenerativeLanguage`
    * `Markdig` (HTML Dönüşümü için)

3.  **API Anahtarını Girin:**
    Projeyi çalıştırdıktan sonra (`F5`), yan paneldeki "Ayarlar" kutusuna API anahtarınızı girip "Kaydet"e basın.

---

## ⚠️ Kritik: Hücre İçi Fonksiyon Kurulumu (VBA)

`=GEMINI()` fonksiyonunun çalışması için Excel dosyanızın içinde C# eklentisiyle konuşacak bir makro bulunmalıdır.

1.  Excel'de `Alt + F11` ile VBA editörünü açın.
2.  **Insert > Module** diyerek yeni bir modül ekleyin.
3.  Modülün adını `modGemini` olarak değiştirin (Properties penceresinden).
4.  Şu kodu yapıştırın:

```vb
Function GEMINI(talimat As String, Optional hucre As Range) As String
    On Error GoTo HataYakala
    Dim eklenti As COMAddIn, otomasyonNesnesi As Object
    
    ' Eklenti bağlantısını kur
    Set eklenti = Application.COMAddIns("GeminiExcelCopilot")
    If eklenti Is Nothing Then
        GEMINI = "HATA: Eklenti bulunamadı."
        Exit Function
    End If
    
    ' C# nesnesini al
    Set otomasyonNesnesi = eklenti.Object
    
    ' Veriyi hazırla
    Dim veri As String: veri = ""
    If Not hucre Is Nothing Then veri = CStr(hucre.Value2)
    
    ' Fonksiyonu çağır
    GEMINI = otomasyonNesnesi.AskGeminiSync(talimat, veri)
    Exit Function
    
HataYakala:
    GEMINI = "HATA: " & Err.Description
End Function
```
## 🏗️ Teknoloji Yığını

* **Dil:** C# (.NET Framework 4.8)
* **Platform:** VSTO (Visual Studio Tools for Office) Excel Add-in
* **Yapay Zeka:** Google Gemini 2.5 Flash (`GenerativeAI` SDK)
* **Arayüz:** Windows Forms (WinForms) & WebBrowser Control
* **Kütüphane:** Markdig (HTML Dönüşümü)

---

## 🤝 Katkıda Bulunma

Katkılarınızı bekliyoruz! Lütfen önce bir "Issue" açarak yapmak istediğiniz değişikliği tartışın.

1.  Bu repoyu Fork'layın.
2.  Kendi branch'inizi oluşturun (`git checkout -b feature/YeniOzellik`).
3.  Değişikliklerinizi commit yapın (`git commit -m 'Yeni özellik eklendi'`).
4.  Branch'inizi Push yapın (`git push origin feature/YeniOzellik`).
5.  Bir Pull Request oluşturun.

## 📄 Lisans

Bu proje [MIT Lisansı](LICENSE) altında lisanslanmıştır.

---
**Geliştirici Notu:** Bu proje, Excel'in yerel dil ayarlarını (Localization) otomatik algılayarak formülleri dönüştüren özel bir yapıya sahiptir.
