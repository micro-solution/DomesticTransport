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

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.TableCarrier.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.TableCarrier_SelectionChange);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Лист2_SelectionChange);
            this.Startup += new System.EventHandler(this.Лист2_Startup);

        }
        #endregion

        private void Лист2_Startup(object sender, System.EventArgs e)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            deliverySheet.Cells[2, 3].Formula = "=TODAY()+1";
            deliverySheet.Calculate();
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


        // Фильтр заказов по активной доставке   
        private void TableCarrier_SelectionChange(Range Target)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];
            Range TargetCell = Globals.ThisWorkbook.Application.ActiveCell;
            OrdersTable.Range.AutoFilter(Field: 1);
            if (TargetCell == null) return;
            try
            {      
                Range commonRng = Globals.ThisWorkbook.Application.Intersect(TargetCell, carrierTable.DataBodyRange);

                if (commonRng != null)
                {
                    string numberDelivery = deliverySheet.Cells[TargetCell.Row, carrierTable.ListColumns[1].Range.Column].Text;
                    OrdersTable.Range.AutoFilter(Field: 1, Criteria1: numberDelivery);
                }
            }
            catch (Exception ex){ MessageBox.Show(ex.Message);
            }
        }

        //private void TableOrders1_Change(Range targetRange, Microsoft.Office.Tools.Excel.ListRanges changedRanges)
        //{
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
        //}
    }
}
