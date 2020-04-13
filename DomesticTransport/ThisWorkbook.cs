namespace DomesticTransport
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            // try
            // {
            // Excel.Worksheet MsgSheet = Globals.ThisWorkbook.Sheets["Email"];
            // Excel.Shape btnSaveReestr = MsgSheet.Shapes.Item("btnSave");
            //     Debug.WriteLine(btnSaveReestr.TextFrame2.TextRange.Text);
            // }
            //catch (Exception ex)
            // {
            //     Debug.WriteLine(ex.Message);
            // }
            //btnSaveReestr.
        }

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
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
