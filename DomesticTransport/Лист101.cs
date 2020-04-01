using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelTools = Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Button = Microsoft.Office.Interop.Excel.Button;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DomesticTransport
{
    public partial class Лист10
    {
        private void Лист10_Startup(object sender, System.EventArgs e)
        {
          //  Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Лист1"];
          // OLEObject btnSave = deliverySheet.OLEObjects("CommandButton1");
          ////  object CommandButtonStart = this.GetType().InvokeMember("BtnSave_GotFocus", System.Reflection.BindingFlags.GetProperty, null, this, null);
          //  btnSave.GotFocus += BtnSave_GotFocus ; 
        }

       
        private void Лист10_Shutdown(object sender, System.EventArgs e)
        {

        }
        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Лист10_Startup);
            this.Shutdown += new System.EventHandler(Лист10_Shutdown);
        }

        #endregion

    }
}
