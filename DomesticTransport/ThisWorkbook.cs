namespace DomesticTransport
{
    public partial class ThisWorkbook
    {

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Shutdown += new System.EventHandler(this.ThisWorkbook_Shutdown);
            this.Open += new Microsoft.Office.Interop.Excel.WorkbookEvents_OpenEventHandler(this.ThisWorkbook_Open);

        }

        #endregion

        private void ThisWorkbook_Open()
        {
            Properties.Settings.Default.AllOrders = "";
            Globals.Лист3.Change += SheetChange; 
            Globals.Лист4.Change += SheetChange;
            Globals.Лист6.Change += SheetChange;
        }

        private void SheetChange(Microsoft.Office.Interop.Excel.Range Target)
        {
            ShefflerWB.SetToZero();
        }
    }
}
