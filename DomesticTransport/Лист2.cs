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
            this.DateDelivery.BeforeDoubleClick += new Microsoft.Office.Interop.Excel.DocEvents_BeforeDoubleClickEventHandler(this.DateDelivery_BeforeDoubleClick);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.Лист2_SelectionChange);
            this.Startup += new System.EventHandler(this.Лист2_Startup);

        }
        #endregion

        private void Лист2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист2_SelectionChange(Range Target)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

      


        /// <summary>
        /// Фильтр заказов по активной доставке 
        /// </summary>
        /// <param name="Target"></param>
        private void TableCarrier_SelectionChange(Range Target)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
                ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
                ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];
                Range TargetCell = Globals.ThisWorkbook.Application.ActiveCell;
                OrdersTable.Range.AutoFilter(Field: 1);
                if (TargetCell == null) return;
                if (carrierTable.DataBodyRange == null) return;

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
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Смена провайдера по двойному клику 
        /// </summary>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        private void TableCarrier_BeforeDoubleClick(Range Target, ref bool Cancel)
        {
            Worksheet deliverySheet = ShefflerWB.DeliverySheet;
            ListObject deliveryTable = ShefflerWB.DeliveryTable;
            ListObject ordersTable = ShefflerWB.OrdersTable;

            if (Target.Column == deliveryTable.ListColumns["Компания"].Range.Column &&
                Target.Row > deliveryTable.HeaderRowRange.Row &&
                Target.Text != "")
            {
                Cancel = true;
                ProviderEditor providerFrm = new ProviderEditor();
                string wt = deliverySheet.Cells[Target.Row, deliveryTable.ListColumns["Вес доставки"].Range.Column].Text;
                Functions functions = new Functions();
                List<Order> orders = functions.GetOrdersFromTable();
                Delivery delivery = new Delivery();
                string numStr = deliverySheet.Cells[Target.Row, deliveryTable.ListColumns["№ Доставки"].Range.Column].Text;
                int number = int.TryParse(numStr, out int n) ? n : 0;
                if (number == 0) return;
                delivery.Orders = orders.FindAll(o => o.DeliveryNumber == number);
                               
                if (orders.Count == 0) return;
                providerFrm.Weight = double.TryParse(wt, out double weight) ? weight : 0;
                providerFrm.ProviderName = Target.Text;
                providerFrm.DeliveryTarget = delivery;
                providerFrm.ShowDialog();
                if (providerFrm.DialogResult == DialogResult.OK)
                {
                    Target.Value = providerFrm.ProviderName;

                    if (providerFrm.ProviderName== "Деловые линии")
                    {
                        delivery.MapDelivery.ForEach(p => p.RouteName = "Сборный груз");
                    Target.Offset[0, 5].Value ="Сборный груз";
                    Target.Offset[0, 2].Value = "0";
                    Target.Offset[0, 1].Value = "0";
                    }

                    Target.Offset[0, 4].Value = providerFrm.CostDelivery;
                    //На лист отгрузки 
                    string idOrder = delivery.Orders[0].Id;
                    Range row = null;
                    row = new ShefflerWB().GetRowOrderTotal(idOrder);
                    if (row != null)
                    {
                        row.Cells[1, ShefflerWB.TotalTable.ListColumns["Перевозчик"].Index].Value = providerFrm.ProviderName;
                        row.Cells[1, ShefflerWB.TotalTable.ListColumns["Стоимость доставки"].Index].Value = providerFrm.CostDelivery;
                        row.Cells[1, ShefflerWB.TotalTable.ListColumns["Тип ТС, тонн"].Index].Value ="";
                      // row.Cells[0, ShefflerWB.TotalTable.ListColumns["Направление"].Index].Value = "Сборный груз";

                    }
                }
            }

        }

        private void DateDelivery_BeforeDoubleClick(Range Target, ref bool Cancel)
        {
            if (Target.Column != ShefflerWB.DateCell.Column ||
            Target.Row != ShefflerWB.DateCell.Row) return;
            Cancel = true;
            new ShefflerWB().SetDateCell();
        }

       
    }
}
