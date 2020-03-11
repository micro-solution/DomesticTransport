using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Config = DomesticTransport.Properties.Settings;
using Excel = Microsoft.Office.Interop.Excel;

namespace DomesticTransport
{
    class Functions
    {
        // private DomesticTransport.Properties.Settings 

        /// <summary>
        /// Вывести на рабочий лист доставки 
        /// </summary>
        public void SetDelivery()
        {

            SapFiles sapFiles = new SapFiles();
            sapFiles.ShowDialog();
            if (sapFiles.DialogResult == DialogResult.OK)
            {
                string sap = "";
                string orders = "";
                try
                {
                    sap = sapFiles.ExportFile;
                    orders = sapFiles.OrderFile;
                }
                catch (Exception ex)
                {
                    return;
                }
                finally
                {
                    sapFiles.Close();
                }

                List<Delivery> deliveries = GetDeliveries(sap, orders);

                Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
                ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
                ListObject invoiciesTable = deliverySheet.ListObjects["TableInvoicies"];

                ClearListObj(carrierTable);
                if (invoiciesTable.DataBodyRange.Rows.Count > 0)
                { invoiciesTable.DataBodyRange.Rows.Delete(); }

                if (deliveries != null && deliveries.Count > 0)
                {
                    PrintDelivery(deliveries, deliverySheet, carrierTable, invoiciesTable);
                }
            }
        }

        private void ClearListObj(ListObject listObject)
        {

            Worksheet worksheet = listObject.Parent;
            for (int i = listObject.ListRows.Count; i > 0; i--)
            {
                ListRow listRow = listObject.ListRows[i];
                worksheet.Rows[listRow.Range.Row].Delete();
            }

        }
        private void AddListRow(ListObject listObject)
        {
            Worksheet worksheet = listObject.Parent;
            if (listObject.ListRows.Count > 0)
            {
                worksheet.Rows[listObject.ListRows[listObject.ListRows.Count].Range.Row + 1].Insert();
            }
            else
            {
                worksheet.Rows[listObject.HeaderRowRange.Row + 2].Insert();
            }
            listObject.ListRows.Add();
        }

        private void PrintDelivery(List<Delivery> deliveries, Worksheet deliverySheet, ListObject CarrierTable, ListObject OrderTable)
        {
            for (int i = 0; i < deliveries.Count; i++)
            {
                Delivery delivery = deliveries[i];
                System.Windows.Forms.Application.DoEvents();


                if (CarrierTable == null || OrderTable == null)
                {
                    MessageBox.Show("Отсутствует таблица");
                    return;
                }
                ListRow rowCarrier = null;
                if (CarrierTable.ListRows.Count == 0)
                {
                    AddListRow(CarrierTable);
                    rowCarrier = CarrierTable.ListRows[1];//  }
                }
                else
                {
                    AddListRow(CarrierTable);
                    rowCarrier = CarrierTable.ListRows[CarrierTable.ListRows.Count - 1];
                }
                rowCarrier.Range[1, 1].Value = i + 1;
                rowCarrier.Range[1, 2].Value = delivery.Truck?.ShippingCompany?.Name ?? "";
                rowCarrier.Range[1, 3].Value = delivery.Truck?.Mark ?? "";
                rowCarrier.Range[1, 4].Value = delivery.Truck?.Tonnage ?? 0;
                Debug.WriteLine($"{delivery.TotalWeight} " + delivery.TotalWeight.ToString().Replace(".", ","));
                rowCarrier.Range[1, 5].Value = delivery.TotalWeight;
                rowCarrier.Range[1, 6].Value = delivery.CostProducts;
                rowCarrier.Range[1, 7].Value = delivery.CostDelivery;

                ListRow rowOrder;

                foreach (Order order in delivery.Orders)
                {
                    // if (CarrierTable.ListRows.Count == 0) CarrierTable.ListRows.AddEx()

                    if (OrderTable.ListRows.Count == 0)
                    {
                        OrderTable.ListRows.Add();
                        rowOrder = OrderTable.ListRows[1];
                    }
                    else
                    {
                        OrderTable.ListRows.Add();
                        rowOrder = OrderTable.ListRows[OrderTable.ListRows.Count - 1];
                    }
                    int column = 0;
                    rowOrder.Range[1, ++column].Value = rowCarrier.Index;
                    rowOrder.Range[1, ++column].Value = order.TransportationUnit;
                    rowOrder.Range[1, ++column].Value = order.Customer?.Id ?? "";
                    rowOrder.Range[1, ++column].Value = "";
                    rowOrder.Range[1, ++column].Value = order.Route;
                    rowOrder.Range[1, ++column].Value = order.PalletsCount;
                    rowOrder.Range[1, ++column].Value = order.WeightNetto;
                    rowOrder.Range[1, ++column].Value = order.Cost.ToString();
                }
            }
        }

        /// <summary>
        /// Поиск 
        /// </summary>
        /// <param name="sap"></param>
        /// <returns></returns>
        public List<Delivery> GetDeliveries(string sap, string orders)
        {
            List<Order> rourerOrders = new List<Order>();
            Delivery delivery = null;
            Workbook sapBook = null;
            Workbook orderBook = null;
            try
            {
                orderBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: orders);
                sapBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: sap);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть книгу Excel");
            }
            Worksheet sheet = sapBook.Sheets[1];
            if (sheet != null)
            {
                int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;
                int lastColumn = sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1;
                Range range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[lastRow, lastColumn]];

                foreach (Range row in range.Rows)
                {
                    Order order = ReadSapRow(row);
                    if (order != null)
                    {

                        if (!string.IsNullOrWhiteSpace(order.TransportationUnit))
                        {
                            List<string> orderInfo = GetOrderInfo(orderBook.Sheets[1], order.TransportationUnit);
                            if (orderInfo != null)
                            {
                                string costStr = orderInfo[1];
                                costStr = costStr.Replace("Стоимость товаров без НДС:", "");
                                costStr = costStr.Replace("RUB", "");
                                costStr = costStr.Replace(".", "");
                                costStr = costStr.Trim();
                                order.Cost = double.TryParse(costStr, out double cost) ? cost : 0;

                                string pallets = orderInfo.Find(x => x.Contains("грузовых мест:")) ?? "";
                                pallets = string.Join("", pallets.Where(c => char.IsDigit(c)));
                                order.PalletsCount = int.TryParse(pallets, out int p) ? p : 0;

                                string weightBrutto = orderInfo.Find(x => x.Contains("вес")) ?? "";
                                weightBrutto = weightBrutto.Replace(".", "");
                                Regex regex = new Regex(@"\d+(\.\d+)?");
                                weightBrutto = regex.Match(weightBrutto).Value;
                                order.WeightBrutto = double.TryParse(weightBrutto, out double wb) ? wb : 0;


                                // order.Customer.   //Улица , Город
                            }
                        }
                        rourerOrders.Add(order);
                    }
                }
            }
            return CompleteAuto(rourerOrders);
        }

        /// <summary>
        /// Распределить заказы по автомобилям
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        public List<Delivery> CompleteAuto(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orderList = orders.OrderByDescending(x => x.WeightNetto).ToList();
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable.OrderBy(x => x.PriorityRoute).ThenBy(y => y.PriorityPoint).ToList();

            #region Проверка если клиента (точки) нет в таблице маршрутов
                Delivery emptyDelivery = null;
                bool emptyPoint;
            for (int k = orderList.Count-1; k >=0 ; k--)
            {
                emptyPoint = true;
                foreach (DeliveryPoint point in pointMap)
                {
                    emptyPoint = true;
                    if (orderList[k].Customer.Id == point.IdCustomer)
                    {
                        emptyPoint = false ;
                        break;
                    }
                    
                }
                if (emptyPoint)
                {
                    if (emptyDelivery == null) emptyDelivery = new Delivery(orderList[k]);
                    orderList.RemoveAt(k);
                }
            }
            deliveries.Add(emptyDelivery);
            #endregion 

            Delivery delivery = null;
            int pointNumber = 0;
            while (orderList.Count > 0)
            {

                for (int orderNumber = orderList.Count - 1; orderNumber >= 0; orderNumber--)
                {
                    
                    if (  orderList[orderNumber].Customer.Id != pointMap[pointNumber].IdCustomer) continue;

                    if (delivery == null)
                    {
                        orderList[orderNumber].DeliveryPoint = pointMap[pointNumber];
                        delivery = new Delivery(orderList[orderNumber]);
                        orderList.RemoveAt(orderNumber);
                    }
                    else
                    {
                        List<DeliveryPoint> points = delivery.MapDelivery;
                        DeliveryPoint point = points.First();
                        Debug.WriteLine($"pointDelivery={point.IdRoute} pointMap={pointMap[pointNumber].IdRoute}");
                        //  новый маршрут в таблице
                        if (delivery.MapDelivery.First().IdRoute != pointMap[pointNumber].IdRoute)
                        {
                            deliveries.Add(delivery);
                            orderList[orderNumber].DeliveryPoint = pointMap[pointNumber];
                            delivery = new Delivery(orderList[orderNumber]);                            
                            orderList.RemoveAt(orderNumber);
                        }
                        else
                        {
                            //По весу 
                            if (!delivery.CheckDeliveryWeght(orderList[orderNumber]))
                            {
                                continue;
                            }
                            else
                            {
                                orderList[orderNumber].DeliveryPoint = pointMap[pointNumber];
                                delivery.Orders.Add(orderList[orderNumber]);
                                orderList.RemoveAt(orderNumber);
                            }

                        }
                    }
                    
                  
                    Debug.WriteLine($"pointNumber={pointNumber} , orderNumber={orderNumber}   idCustomer {orderList[orderNumber].Customer.Id}");
                  
                }
                

                Debug.WriteLine($"pointNumber = {pointNumber}");
                pointNumber++;
                if (pointNumber >= pointMap.Count)  pointNumber = 0;                 
            }
            if (delivery != null && delivery.Orders.Count > 0) deliveries.Add(delivery);
            return deliveries;
        }

        /// <summary>
        /// Собираем из строки выгруза данные для формирования доставки
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Order ReadSapRow(Range row)
        {
            /// ТТН
            Order order = new Order();
            Debug.WriteLine(row.Row);

            order.TransportationUnit = row.Cells[1, GetColumn(row.Parent, "Transportation Unit", 1)].Text;
            if (string.IsNullOrWhiteSpace(order.TransportationUnit)) return null;


            string idCusomer = row.Cells[1, GetColumn(row.Parent, "Получатель материала", 1)].Text;
            order.Customer = string.IsNullOrWhiteSpace(idCusomer) ? null : new Customer(idCusomer);

            order.Id = row.Cells[1, GetColumn(row.Parent, "Delivery", 1)].Text;

            if (string.IsNullOrWhiteSpace(idCusomer) || string.IsNullOrWhiteSpace(order.Id))
            {
                return null;
            }
            string weight = row.Cells[1, GetColumn(row.Parent, "Вес брутто", 1)].Text;
            order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
            order.Route = row.Cells[1, GetColumn(row.Parent, "Маршрут", 1)].Text;
            return order;
        }



        /// <summary>
        /// Записать накладную по доставке из файла выгруза в Лист
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="transportationUnit"></param>
        /// <returns></returns>
        private List<string> GetOrderInfo(Worksheet sheet, string transportationUnit)
        {
            Range findRange = sheet.Columns[1];
            //string search = "№ ТТН:" + new string('0', 18 - transportationUnit.Length) + transportationUnit;
            string search = new string('0', 18 - transportationUnit.Length) + transportationUnit;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);

            if (fcell == null && fcell.Value.Trim().Contains("ТТН:")) return null;

            int rowStart = fcell.Row;
            int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            int rowEnd = rowStart;
            List<string> info = new List<string>();
            do
            {
                fcell = findRange.Cells[rowEnd++, 1];
                string cellText = fcell.Value;
                cellText.Trim();
                cellText = cellText.Replace("\t", "");
                cellText = cellText.Replace(";;;", "");
                if (string.IsNullOrEmpty(cellText.Replace(";", ""))) break;
                info.Add(cellText);
            }
            while (rowEnd <= lastRow);
            return info; //findRange[findRange.Cells[rowStart, 1], findRange.Cells[rowEnd, 1]];
        }

        /// <summary>
        /// Ищет в строке или на листе ячейку с заголовком и возвращает столбец
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="header"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private int GetColumn(Worksheet sh, string header, int row = 0)
        {
            Range findRange = row == 0 ? sh.UsedRange : sh.Rows[row];
            Range fcell = findRange.Find(What: header, LookIn: XlFindLookIn.xlValues);
            return fcell == null ? 0 : fcell.Column;
        }


    }
}
