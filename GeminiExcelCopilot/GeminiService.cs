using GenerativeAI;
using System;
using System.Threading.Tasks;

namespace GeminiExcelCopilot
{
    public class GeminiService
    {
        private GenerativeModel model;

        public GeminiService()
        {
            // 1. KODDAN DEĞİL, KAYITLI AYARLARDAN OKU
            string apiKey = Properties.Settings.Default.GeminiApiKey;

            // 2. ANAHTARIN BOŞ OLUP OLMADIĞINI KONTROL ET
            if (string.IsNullOrEmpty(apiKey))
            {
                // Anahtar yoksa, Task Pane'in yakalaması için bir hata fırlat
                throw new InvalidOperationException("API Anahtarı bulunamadı. Lütfen aşağıdaki 'API Anahtarı' bölümünden anahtarınızı girip 'Kaydet' butonuna basın.");
            }

            // 3. MODELİ BU ANAHTARLA BAŞLAT
            // "gemini-1.5-flash" en hızlı olanıdır.
            model = new GenerativeModel(apiKey, "gemini-2.5-flash");
        }

        public async Task<string> GenerateContentAsync(string userPrompt)
        {
            try
            {
                var response = await model.GenerateContentAsync(userPrompt);
                return response.Text;
            }
            catch (Exception ex)
            {
                // API'den gelen hataları (örn: geçersiz anahtar) yakala
                return $"API Hatası: {ex.Message}";
            }
        }
    }
}