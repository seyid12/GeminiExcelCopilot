using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace GeminiExcelCopilot
{
    // 1. Önce bir "Arayüz" (Interface) tanımlıyoruz.
    // Bu, dış dünyaya "Benim hangi metodlarım kullanılabilir" listesini verir.
    [ComVisible(true)]
    [Guid("D8660990-E593-417F-9D45-C639D9770932")] // Rastgele bir kimlik
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGeminiInterface
    {
        string AskGeminiSync(string prompt, string context);
    }

    // 2. Şimdi bu arayüzü uygulayan gerçek sınıfı yazıyoruz.
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class GeminiFunctions : IGeminiInterface
    {
        public string AskGeminiSync(string prompt, string context)
        {
            // API Anahtarı kontrolü
            string apiKey = Properties.Settings.Default.GeminiApiKey;
            if (string.IsNullOrEmpty(apiKey)) return "HATA: API Anahtarı Yok";

            try
            {
                // Servisi başlat
                GeminiService service = new GeminiService();

                // Kullanıcıdan gelen prompt ile hücredeki veriyi birleştir
                string finalPrompt = $"Kullanıcı İsteği: {prompt}\nVeri: {context}\nCevap sadece sonuç olsun.";

                // ÖNEMLİ: Excel hücreleri "Asenkron" (await) çalışmayı sevmez.
                // Bu yüzden cevabı bekletip (Synchronous) alıyoruz. 
                // Bu işlem sırasında Excel 1-2 saniye donabilir, bu normaldir.
                var task = Task.Run(async () => await service.GenerateContentAsync(finalPrompt));
                return task.GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                return "HATA: " + ex.Message;
            }
        }
    }
}