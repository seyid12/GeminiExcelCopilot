using GenerativeAI;
using System;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GeminiExcelCopilot
{
    public partial class TaskPaneControl : UserControl
    {
        private GeminiService geminiService;

        public TaskPaneControl()
        {
            InitializeComponent();
            InitializeCustomUI();

            // API Anahtarını Yükle
            txtApiKey.Text = Properties.Settings.Default.GeminiApiKey;
            InitializeGemini();
        }

        private void InitializeCustomUI()
        {
            if (cmbActionType.Items.Count == 0)
            {
                cmbActionType.Items.AddRange(new object[] {
                    "Soru Sor (Genel)",
                    "Formül Üret",
                    "Makro (VBA) Üret",
                    "Seçili Alanı Analiz Et",
                    "Otomatik Grafik Oluştur"
                });
            }
            if (cmbActionType.SelectedIndex == -1) cmbActionType.SelectedIndex = 0;

            // Başlangıçta WebBrowser gizli
            if (webBrowserResult != null) webBrowserResult.Visible = false;
        }

        private void InitializeGemini()
        {
            try
            {
                geminiService = new GeminiService();
                txtResult.ForeColor = System.Drawing.Color.Green;
                txtResult.Text = "Veri Asistanı hazır.";
                SetControlsEnabled(true);
            }
            catch (Exception ex)
            {
                txtResult.ForeColor = System.Drawing.Color.Red;
                txtResult.Text = ex.Message;
                SetControlsEnabled(false);
            }
        }

        private void SetControlsEnabled(bool isEnabled)
        {
            button1.Enabled = isEnabled;
            btnApply.Enabled = isEnabled;
            cmbActionType.Enabled = isEnabled;
            textBox1.Enabled = isEnabled;
        }

        private void btnSaveApiKey_Click(object sender, EventArgs e)
        {
            string newKey = txtApiKey.Text.Trim();
            Properties.Settings.Default.GeminiApiKey = newKey;
            Properties.Settings.Default.Save();
            InitializeGemini();
            MessageBox.Show("API Anahtarı kaydedildi!", "Bilgi");
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // UI Durumu: Yükleniyor
            button1.Enabled = false;
            button1.Text = "İnceleniyor...";
            txtResult.Text = "";
            if (webBrowserResult != null) webBrowserResult.DocumentText = "";

            string userPrompt = textBox1.Text;
            string action = cmbActionType.SelectedItem.ToString();
            string language = Globals.ThisAddIn.ExcelLanguageName;
            string finalPrompt = "";
            string csvData = "";
            bool useWebBrowser = false;

            // --- PROMPT HAZIRLAMA MANTIĞI ---
            switch (action)
            {
                case "Formül Üret":
                    finalPrompt = $"Talimat: Kullanıcının '{language}' dilindeki şu isteğini anla: \"{userPrompt}\". Bu isteği yerine getiren **İngilizce** Excel formülünü yaz. Sadece formülü döndür. Örn: =SUM(A1:A10)";
                    useWebBrowser = false;
                    break;

                case "Makro (VBA) Üret":
                    finalPrompt = $"Talimat: Kullanıcının şu isteğini yerine getiren tam bir Excel VBA Sub...End Sub kodu yaz. Sadece kodu ver: \"{userPrompt}\"";
                    useWebBrowser = false;
                    break;

                case "Soru Sor (Genel)":
                    // Genel sohbette veri okumuyoruz, sadece Markdown istiyoruz.
                    finalPrompt = $"Sen uzman bir asistansın. Soru: \"{userPrompt}\". Cevabı Markdown formatında (kalın başlıklar, listeler) ver.";
                    useWebBrowser = true;
                    break;

                case "Seçili Alanı Analiz Et":
                    csvData = GetSelectedRangeAsCsv();
                    if (string.IsNullOrEmpty(csvData)) { ShowError("Lütfen veri seçin."); return; }

                    finalPrompt = $"Sen kıdemli bir veri analistisin. Veri Seti:\n---\n{csvData}\n---\n" +
                                  $"Soru: \"{userPrompt}\"\n\n" +
                                  $"Analizi Markdown formatında yap. Başlıklar, kalın vurgular ve tablolar kullan. Türkçe ve profesyonel bir dil kullan.";
                    useWebBrowser = true;
                    break;

                case "Otomatik Grafik Oluştur":
                    csvData = GetSelectedRangeAsCsv();
                    if (string.IsNullOrEmpty(csvData)) { ShowError("Lütfen grafik için veri seçin."); return; }

                    // Gelişmiş Grafik Karar Prompt'u
                    finalPrompt = $"Sen uzman bir Excel Veri Analistisin. Veri Seti:\n---\n{csvData}\n---\n" +
                                  $"GÖREV: Bu veriyi en iyi temsil edecek Excel Grafik Türünü (XlChartType) belirle.\n" +
                                  $"KURALLAR:\n" +
                                  $"1. Zaman serisi (Yıllar, Aylar) -> 'xlLine'\n" +
                                  $"2. Kategorik karşılaştırma -> 'xlColumnClustered'\n" +
                                  $"3. Parça/Bütün (Yüzde) -> 'xlPie'\n" +
                                  $"4. İki sayısal değişken ilişkisi -> 'xlXYScatter'\n" +
                                  $"CEVAP: SADECE tek kelime (Örn: xlLine). Açıklama yapma.";
                    useWebBrowser = false;
                    break;

                default:
                    finalPrompt = userPrompt;
                    break;
            }

            try
            {
                // API ÇAĞRISI (ARKA PLAN)
                string response = await geminiService.GenerateContentAsync(finalPrompt);

                // --- UI GÜNCELLEME (ANA THREAD) ---
                this.Invoke(new Action(() =>
                {
                    txtResult.Visible = !useWebBrowser;
                    if (webBrowserResult != null) webBrowserResult.Visible = useWebBrowser;
                    btnApply.Enabled = !useWebBrowser;

                    // 1. GRAFİK İŞLEMİ
                    if (action == "Otomatik Grafik Oluştur")
                    {
                        string chartTypeString = response.Trim().Replace("`", "").Replace("'", "");
                        bool success = CreateChartFromGemini(chartTypeString, userPrompt);

                        txtResult.Visible = true;
                        if (webBrowserResult != null) webBrowserResult.Visible = false;

                        if (success) txtResult.Text = $"Grafik çizildi: {chartTypeString}";
                        else txtResult.Text = $"Grafik çizilemedi. Yanıt: {chartTypeString}";
                    }
                    // 2. WEB BROWSER (ANALİZ / SOHBET)
                    else if (useWebBrowser)
                    {
                        string htmlContent = MarkdownHelper.ConvertToHtml(response);
                        webBrowserResult.DocumentText = htmlContent;
                    }
                    // 3. TEXTBOX (FORMÜL / VBA)
                    else
                    {
                        string cleanResponse = response.Replace("```excel", "").Replace("```vba", "").Replace("```", "").Trim();
                        txtResult.Text = cleanResponse;
                        txtResult.ForeColor = System.Drawing.Color.Black;
                    }

                    ResetButton();
                }));
            }
            catch (Exception ex)
            {
                // Hata durumunda da Invoke kullanmalıyız
                this.Invoke(new Action(() => { ShowError($"Hata: {ex.Message}"); }));
            }
        }

        private void ShowError(string msg)
        {
            txtResult.Visible = true;
            if (webBrowserResult != null) webBrowserResult.Visible = false;
            txtResult.ForeColor = System.Drawing.Color.Red;
            txtResult.Text = msg;
            ResetButton();
        }

        private void ResetButton()
        {
            button1.Enabled = true;
            button1.Text = "Gönder";
        }

        // --- YARDIMCI METOTLAR ---

        private string GetSelectedRangeAsCsv()
        {
            Excel.Range selection = null;
            try
            {
                selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (selection == null || selection.Cells.Count < 2) return null;
            }
            catch { return null; }

            StringBuilder csvBuilder = new StringBuilder();
            object[,] values = null;

            try
            {
                // Performans için 500 satır / 20 sütun sınırı
                int rowCount = Math.Min(selection.Rows.Count, 500);
                int colCount = Math.Min(selection.Columns.Count, 20);

                // Value2 ile toplu okuma (Hızlı Yöntem)
                Excel.Range resizedRange = selection.Resize[rowCount, colCount];

                // Tek hücre değilse dizi döner
                if (rowCount > 1 || colCount > 1)
                    values = resizedRange.Value2 as object[,];
                else
                    return null;

                if (values == null) return null;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        string val = values[i, j]?.ToString() ?? "";
                        val = val.Replace(";", " ").Replace("\n", " ");
                        csvBuilder.Append(val + ";");
                    }
                    csvBuilder.AppendLine();
                }
            }
            catch { return null; }

            return csvBuilder.ToString();
        }

        private bool CreateChartFromGemini(string chartTypeString, string title)
        {
            // Temizlik
            chartTypeString = chartTypeString.Replace("```", "").Trim();

            // Bilinen tipleri yakala
            if (chartTypeString.Contains("xlLine")) chartTypeString = "xlLine";
            else if (chartTypeString.Contains("xlPie")) chartTypeString = "xlPie";
            else if (chartTypeString.Contains("xlColumn")) chartTypeString = "xlColumnClustered";
            else if (chartTypeString.Contains("xlBar")) chartTypeString = "xlBarClustered";
            else if (chartTypeString.Contains("xlScatter")) chartTypeString = "xlXYScatter";

            Excel.XlChartType chartType;
            if (!Enum.TryParse(chartTypeString, true, out chartType))
            {
                chartType = Excel.XlChartType.xlColumnClustered; // Varsayılan
            }

            try
            {
                Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (selection == null) return false;

                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

                // Grafiği verinin sağına yerleştir
                double leftPos = selection.Left + selection.Width + 10;

                Excel.ChartObject chartObj = sheet.ChartObjects().Add(leftPos, selection.Top, 450, 300);
                chartObj.Chart.SetSourceData(selection);
                chartObj.Chart.ChartType = chartType;
                chartObj.Chart.HasTitle = true;

                // Başlık boşsa otomatik ver
                if (!string.IsNullOrEmpty(title) && title.Length > 2)
                    chartObj.Chart.ChartTitle.Text = title;
                else
                    chartObj.Chart.ChartTitle.Text = "Analiz Grafiği";

                return true;
            }
            catch
            {
                return false;
            }
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            if (!txtResult.Visible)
            {
                MessageBox.Show("Analiz modu etkinken formül uygulanamaz.");
                return;
            }
            string formula = txtResult.Text.Trim();
            if (string.IsNullOrWhiteSpace(formula)) return;

            try
            {
                Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
                if (activeCell == null) return;

                if (formula.StartsWith("=")) activeCell.Formula = formula;
                else activeCell.Value2 = formula;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
    }
}