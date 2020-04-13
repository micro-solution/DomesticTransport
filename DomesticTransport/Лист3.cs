using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DomesticTransport
{
    public partial class Лист3
    {
        private void Лист3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.TableRoutes.Change += new Microsoft.Office.Tools.Excel.ListObjectChangeHandler(this.TableRoutes_Change);
            this.Startup += new System.EventHandler(this.Лист3_Startup);
            this.Shutdown += new System.EventHandler(this.Лист3_Shutdown);

        }

        #endregion

        private void TableRoutes_Change(Excel.Range targetRange, Microsoft.Office.Tools.Excel.ListRanges changedRanges)
        {
            ShefflerWB.RoutesList = null;             
        }
    }
}
