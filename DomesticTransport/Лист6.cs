using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DomesticTransport
{
    public partial class Лист6
    {
        private void Лист6_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист6_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Лист6_Startup);
            this.Shutdown += new System.EventHandler(Лист6_Shutdown);
        }

        #endregion

    }
}
