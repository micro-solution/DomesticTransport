using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DomesticTransport
{
    public partial class Лист2
    {
        


        private void Лист2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.TableCarrier.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.TableCarrier_SelectionChange);
            this.TableCarrier.SelectedIndexChanged += new System.EventHandler(this.TableCarrier_SelectedIndexChanged);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Лист2_SelectionChange);
            this.Startup += new System.EventHandler(this.Лист2_Startup);
            this.Shutdown += new System.EventHandler(this.Лист2_Shutdown);

        }


        #endregion



        private void TableCarrier_SelectedIndexChanged(object sender, EventArgs e)
        {
         
          ///ListObject listCarrier = (Microsoft.Office.Tools.Excel.ListObject)sender;
          // // Excel.ListObject listCarrier = (Excel.ListObject )sender;
          ////  Excel.ListRow listRowCarrier = (Excel.ListRow)sender;
          //  //Excel.ListObject listCarrier = listRowCarrier.Parent;
          //  listCarrier.ListRows[1].Range.Select();
          //  Worksheet worksheet = (Worksheet)listCarrier.Parent ;
         // Excel.ListObject listOrders = worksheet.ListObjects["TableOrders"];
          //  listOrders.ShowAutoFilter = false;

            //listCarrier. ();
            //listOrders.
        }

        private void TableCarrier_SelectionChange(Range Target)
        {
            int row = Target.Row;
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];

            OrdersTable.Range.AutoFilter(Field: 1) ;
            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, carrierTable.DataBodyRange);
            if (commonRng !=null)
            {                  
                string numberDelivery = deliverySheet.Cells[Target.Row, carrierTable.ListColumns[1].Range.Column].Text;
                 OrdersTable.Range.AutoFilter(Field: 1, Criteria1: numberDelivery);
            }

        }

        private void Лист2_SelectionChange(Range Target)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"]; 
            ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];
            OrdersTable.Range.AutoFilter(Field: 1);
        }
    }
}
