using DomesticTransport.Forms;
using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

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
            this.TableCarrier.BeforeDoubleClick += new Microsoft.Office.Interop.Excel.DocEvents_BeforeDoubleClickEventHandler(this.TableCarrier_BeforeDoubleClick);
            this.TableCarrier.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.TableCarrier_SelectionChange);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Лист2_SelectionChange);
            this.Startup += new System.EventHandler(this.Лист2_Startup);

        }
        #endregion

        private void Лист2_Startup(object sender, System.EventArgs e)
        {
            ShefflerWorkBook.ExcelOptimizateOff();
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
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
            if (carrierTable.DataBodyRange == null) return;
            try
            {
                Range commonRng = Globals.ThisWorkbook.Application.Intersect(TargetCell, carrierTable.DataBodyRange);

                if (commonRng != null)
                {
                    string numberDelivery = deliverySheet.Cells[TargetCell.Row, carrierTable.ListColumns[1].Range.Column].Text;
                    OrdersTable.Range.AutoFilter(Field: 1, Criteria1: numberDelivery);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void TableCarrier_BeforeDoubleClick(Range Target, ref bool Cancel)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];

            if (Target.Column == carrierTable.ListColumns["Компания"].Range.Column &&
                Target.Row > carrierTable.HeaderRowRange.Row &&
                Target.Text != "")
            {
                ProviderEditor providerFrm = new ProviderEditor();
                string wt = deliverySheet.Cells[Target.Row, carrierTable.ListColumns["Вес доставки"].Range.Column].Text;
                Functions functions = new Functions();
                List<Order> orders = functions.GetOrdersFromTable(ordersTable);
                Delivery delivery = new Delivery();
                string numStr = deliverySheet.Cells[Target.Row, carrierTable.ListColumns["№ Доставки"].Range.Column].Text;
                int number = int.TryParse(numStr, out int n) ? n : 0;
                if (number == 0) return;
                delivery.Orders = orders.FindAll(o => o.DeliveryNumber == number);

                if (orders.Count == 0) {return;  }
                    providerFrm.Weight = double.TryParse(wt, out double weight) ? weight : 0;
                    providerFrm.ProviderName = Target.Text;
                    providerFrm.MapDpelivery = delivery.MapDelivery;
                    providerFrm.ShowDialog();
                if (providerFrm.DialogResult == DialogResult.OK)
                {
                    Target.Value = providerFrm.ProviderName;
                }

            }

        }
    }
}
