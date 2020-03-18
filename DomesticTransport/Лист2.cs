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
            this.TableCarrier.SelectionChange += new DocEvents_SelectionChangeEventHandler(this.TableCarrier_SelectionChange);
            this.TableOrders1.Change += new Microsoft.Office.Tools.Excel.ListObjectChangeHandler(this.TableOrders1_Change);
            this.SelectionChange += new DocEvents_SelectionChangeEventHandler(this.Лист2_SelectionChange);
            this.Startup += new EventHandler(this.Лист2_Startup);
            this.Shutdown += new EventHandler(this.Лист2_Shutdown);
            this.Change += new DocEvents_ChangeEventHandler(this.Лист2_Change);

        }


        #endregion





        private void TableCarrier_SelectionChange(Range Target)
        {

            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];

            OrdersTable.Range.AutoFilter(Field: 1);
            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, carrierTable.DataBodyRange);

            if (commonRng != null)
            {
                string numberDelivery = deliverySheet.Cells[Target.Row, carrierTable.ListColumns[1].Range.Column].Text;
                OrdersTable.Range.AutoFilter(Field: 1, Criteria1: numberDelivery);
            }

        }

        private void Лист2_SelectionChange(Range Target)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet?.ListObjects["TableCarrier"];
            ListObject OrdersTable = deliverySheet?.ListObjects["TableOrders"];
            if (carrierTable == null || OrdersTable == null) return;

            Range commonOrdrrRng = Globals.ThisWorkbook.Application.Intersect(Target, OrdersTable.Range);
            if (carrierTable?.DataBodyRange == null) return;

            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, carrierTable.DataBodyRange);
            if (commonRng == null && commonOrdrrRng == null)
            {
                OrdersTable.Range.AutoFilter(Field: 1);
            }
        }

        private void Лист2_Change(Range Target)
        {

        }

        private void TableOrders1_Change(Range targetRange, Microsoft.Office.Tools.Excel.ListRanges changedRanges)
        {
            //Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            //ListObject OrdersTable = deliverySheet?.ListObjects["TableOrders"];
            //if ( OrdersTable == null) return;

            //         //При заполнении программой выключено обновление экрана
                     
            //if (Globals.ThisWorkbook.Application.ScreenUpdating && targetRange.Text !="")
            //{
            //    if (targetRange.Column == OrdersTable.ListColumns["№ Доставки"].Range.Column)
            //    {
            //        if (int.TryParse(targetRange.Text, out int num))
            //        {                    
            //            Functions functions = new Functions();
            //            functions.СhangeDelivery();
            //        }
            //        else
            //        {
            //           targetRange.Value = 0;
            //        }

                    //numberDelivery 
                    //MessageBox.Show("dkj " + targetRange.Value);
                    // Пересчитать все доставки
                //}
           /// }
        }
    }
}
