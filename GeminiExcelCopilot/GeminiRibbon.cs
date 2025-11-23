using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GeminiExcelCopilot
{
    public partial class GeminiRibbon
    {
        // Butona tıklandığında (Hem kullanıcı tıkladığında, hem de kodla değiştiğinde)
        private void toggleButtonShowPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.geminiTaskPane != null)
            {
                // Butonun durumunu (Checked) görev bölmesinin görünürlüğüne eşitle
                Globals.ThisAddIn.geminiTaskPane.Visible = toggleButtonShowPane.Checked;
            }
        }
    }
}