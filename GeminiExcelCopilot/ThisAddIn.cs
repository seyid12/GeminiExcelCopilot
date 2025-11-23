using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace GeminiExcelCopilot
{
    public partial class ThisAddIn
    {
        // 1. Dil adını saklamak için genel (public) bir değişken.
        public string ExcelLanguageName = "English"; // Varsayılan

        // 2. Görev bölmesi için değişkenler
        private TaskPaneControl taskPaneControl;
        // GÖREV BÖLMESİNİ GENEL (PUBLIC) YAPIYORUZ
        public Microsoft.Office.Tools.CustomTaskPane geminiTaskPane;

        // =================================================================
        // YENİ EKLENEN KISIM: VBA KÖPRÜSÜ İÇİN DEĞİŞKEN
        // =================================================================
        private GeminiFunctions utilities;

        // =================================================================
        // YENİ EKLENEN KISIM: EXCEL BU METODU OTOMATİK ÇAĞIRIR
        // "VBA'nın erişebileceği bir nesnen var mı?" diye sorar.
        // =================================================================
        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
            {
                utilities = new GeminiFunctions();
            }
            return utilities;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 3. Eklenti başlarken Excel'in arayüz dilini (LCID) al
            int lcid = this.Application.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI];

            // LCID'yi "Turkish", "English" gibi bir metne çevir
            ExcelLanguageName = GetExcelLanguageName(lcid);
            // ------------------------------------

            // 4. Görev bölmesini oluştur
            taskPaneControl = new TaskPaneControl();
            // AD DEĞİŞİKLİĞİ: Başlık "Gemini Copilot" -> "Veri Asistanı" olarak güncellendi.
            geminiTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "Veri Asistanı");
            geminiTaskPane.Visible = true;

            // 5. ŞERİT (RIBBON) İÇİN EVENT EKLİYORUZ
            // Kullanıcı bölmeyi "X" ile kapattığında, şerit butonunun
            // durumunu güncellemek için bu olayı dinleriz.
            geminiTaskPane.VisibleChanged += GeminiTaskPane_VisibleChanged;
        }

        // 6. ŞERİT İÇİN YENİ METOD
        /// <summary>
        /// Görev bölmesinin görünürlüğü değiştiğinde tetiklenir.
        /// </summary>
        private void GeminiTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            // Şeritteki 'toggleButtonShowPane' butonunun durumunu
            // görev bölmesinin görünürlüğüne eşitleriz.
            // Bu, 'GeminiRibbon' sınıfında oluşturduğumuz butondur.
            // NOT: Eğer Ribbon henüz yüklenmediyse hata vermemesi için kontrol ekleyebilirsiniz
            if (Globals.Ribbons.GeminiRibbon != null)
            {
                Globals.Ribbons.GeminiRibbon.toggleButtonShowPane.Checked = geminiTaskPane.Visible;
            }
        }

        // 7. LCID kodunu metne çeviren yardımcı metod
        /// <summary>
        /// Excel Dil Kodunu (LCID) okunabilir bir metne çevirir.
        /// </summary>
        private string GetExcelLanguageName(int lcid)
        {
            switch (lcid)
            {
                case 1055:
                    return "Turkish"; // Türkçe
                case 1033:
                    return "English"; // İngilizce (ABD)
                case 1031:
                    return "German";  // Almanca
                case 1036:
                    return "French";  // Fransızca
                case 1034:
                    return "Spanish"; // İspanyolca
                default:
                    return "English"; // Bilinmiyorsa İngilizce varsay
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}