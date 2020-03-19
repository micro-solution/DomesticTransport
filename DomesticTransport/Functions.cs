using DomesticTransport.Forms;
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
            ExcelOptimizateOn();
            SapFiles sapFiles = new SapFiles();
            sapFiles.ShowDialog();
            if (sapFiles.DialogResult == DialogResult.OK)
            {
                string sapPath = "";
                string ordersPath = "";
                try
                {
                    sapPath = sapFiles.ExportFile;
                    ordersPath = sapFiles.OrderFile;
                }
                catch (Exception ex)
                {
                    return;
                }
                finally
                {
                    sapFiles.Close();
                }

                List<Order> ordersSap = GetSapOrders(sapPath);  // Export from SAP
                if (ordersPath != "" && File.Exists(ordersPath))
                    ordersSap = GetOrdersInfo(ordersPath, ordersSap);  // Поиск свойств в файле All orders
                List<Delivery> deliveries = CompleteAuto(ordersSap);


                Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
                ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
                ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
                Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
                ListObject TotalTable = TotalSheet.ListObjects["TableTotal"];

                ClearListObj(carrierTable);
                if (ordersTable.DataBodyRange.Rows.Count > 0)
                { ordersTable.DataBodyRange.Rows.Delete(); }

                if (deliveries != null && deliveries.Count > 0)
                {
                    PrintDelivery(deliveries, carrierTable);
                    PrintOrders(deliveries, ordersTable);
                    PrintShipping(TotalTable, deliveries);
                }
            }
            ExcelOptimizateOff();
        }



         /// <summary>
         /// очистка таблицы удалением строк листа
         /// </summary>
         /// <param name="listObject"></param>
        private void ClearListObj(ListObject listObject)
        {
            Globals.ThisWorkbook.Application.DisplayAlerts = false;
            Worksheet worksheet = listObject.Parent;
            for (int i = listObject.ListRows.Count; i > 0; i--)
            {
                ListRow listRow = listObject.ListRows[i];
                worksheet.Rows[listRow.Range.Row].Delete();
            }
            Globals.ThisWorkbook.Application.DisplayAlerts = true;
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
            // worksheet.Rows[listObject.ListRows[listObject.ListRows.Count].Range.Row + 1].Insert();
            listObject.ListRows.Add();
        }

        /// <summary>
        ///кнопка  Добавить строку авто
        /// </summary>
        internal void AddAuto()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];

            ListRow rowDelivery;
            if (carrierTable.ListRows.Count == 0)
            {
                AddListRow(carrierTable);
                rowDelivery = carrierTable.ListRows[1];//  }
            }
            else
            {
                AddListRow(carrierTable);
                rowDelivery = carrierTable.ListRows[carrierTable.ListRows.Count - 1];
            }
            int number = 0;
            foreach (Range rng in carrierTable.ListColumns["№ Доставки"].DataBodyRange)
            {
                if (int.TryParse(rng.Text, out int valueCell))
                {
                    if (number < valueCell) number = valueCell;
                }
            }

            rowDelivery.Range[1, carrierTable.ListColumns["№ Доставки"].Index].Value = number + 1;
        }

        /// <summary>
        ///кнопка Добавить авто
        /// </summary>
        internal void DeleteAuto()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject deliveryTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject TotalTable = totalSheet.ListObjects["TableTotal"];

            if (deliveryTable == null || ordersTable == null) return;
            Range Target = Globals.ThisWorkbook.Application.Selection;
            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, deliveryTable.DataBodyRange);
            int numberDelivery = 0;
            if (commonRng != null)
            {
                int row = commonRng.Row;
                // commonRng = Globals.ThisWorkbook.Application.Intersect(
                commonRng = deliveryTable.ListColumns["№ Доставки"].Range.Columns[row, 1];
                numberDelivery = int.TryParse(commonRng.Text, out int nmDelivery) ? nmDelivery : 0;
            }

            foreach (ListRow listDeliveryRow in deliveryTable.ListRows)
            {
                string str = listDeliveryRow.Range[1, deliveryTable.ListColumns["№ Доставки"].Index].value;
                if (int.TryParse(str, out int number))
                {
                    if (number == numberDelivery) listDeliveryRow.Range.Rows.Delete();
                }

            }
            foreach (ListRow listDeliveryRow in ordersTable.ListRows)
            {
                string str = listDeliveryRow.Range[1, ordersTable.ListColumns["№ Доставки"].Index].value;
                if (int.TryParse(str, out int number))
                {
                    if (number == numberDelivery) listDeliveryRow.Range.Rows.Delete();
                }
            }
        }


        /// <summary>
        /// Запись доставок в таблицы  лист Delivery
        /// </summary>
        /// <param name="deliveries"></param>
        /// <param name="DeliveryTable"></param>
        /// <param name="OrderTable"></param>
        private void PrintDelivery(List<Delivery> deliveries, ListObject DeliveryTable)
        {

            ProcessBar pb = ProcessBar.Init("Вывод данных", deliveries.Count, 1, "Формирование доставок");
            if (pb == null) return;
            pb.Show();

            for (int i = 0; i < deliveries.Count; i++)
            {
                if (pb.Cancel) break;
                pb.Action($"Доставка {i + 1} из {pb.Count}");

                Delivery delivery = deliveries[i];

                ListRow rowDelivery;
                if (DeliveryTable.ListRows.Count == 0)
                {
                    AddListRow(DeliveryTable);
                    rowDelivery = DeliveryTable.ListRows[1];//  }
                }
                else
                {
                    AddListRow(DeliveryTable);
                    rowDelivery = DeliveryTable.ListRows[DeliveryTable.ListRows.Count - 1];
                }

                //int numberDelivery = 0;
                //if (delivery.hasRoute )
                //{
                //    numberDelivery = i + 1;
                //}
                rowDelivery.Range[1, DeliveryTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Компания"].Index].Value =
                                                                delivery.Truck?.ShippingCompany?.Name ?? "";
                if (delivery?.MapDelivery.Count > 0)
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["ID Route"].Index].Value =
                                                                        delivery?.MapDelivery[0].IdRoute;
                }
                rowDelivery.Range[1, DeliveryTable.ListColumns["Тоннаж"].Index].Value = delivery.Truck?.Tonnage ?? 0;

                //rowCarrier.Range[1, CarrierTable.ListColumns["Вес доставки"].Index].Value = delivery.TotalWeight;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Вес доставки"].Index].FormulaR1C1 =
                                                "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес нетто])";

                //rowCarrier.Range[1, CarrierTable.ListColumns["Стоимость товаров"].Index].Value = delivery.CostProducts;
              //  rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость товаров"].Index].Value =
               //                             "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Стоимость товаров])";
                rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.CostDelivery;

                int columnMap = 0;
                foreach (DeliveryPoint point in delivery.MapDelivery)
                {
                    ++columnMap;
                    rowDelivery.Range[1, DeliveryTable.ListColumns.Count].Offset[0, 2 + columnMap].Value
                                    = $"{point.IdCustomer} - {point.City} ";
                }
            }
            pb.Close();
        }

        //Вывод заказов
        private void PrintOrders(List<Delivery> deliveries, ListObject OrderTable)
        {
            ProcessBar pb = ProcessBar.Init("Вывод данных", deliveries.Count, 1, "Печать заказов");
            if (pb == null) return;
            pb.Show();
            for (int i = 0; i < deliveries.Count; i++)
            {
                if (pb == null) return;

                if (pb.Cancel) break;
                pb.Action($"Доставка {i + 1} из {pb.Count}");

                Delivery delivery = deliveries[i];

                ListRow rowOrder;
                foreach (Order order in delivery.Orders)
                {

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
                    rowOrder.Range[1, OrderTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                    rowOrder.Range[1, OrderTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                    rowOrder.Range[1, OrderTable.ListColumns["Доставка"].Index].Value = order.Id;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Получателя"].Index].Value = order.Customer?.Id ?? "";
                    rowOrder.Range[1, OrderTable.ListColumns["Получатель"].Index].Value = order.Customer.Name;
                    rowOrder.Range[1, OrderTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Маршрута"].Index].Value = order.DeliveryPoint.IdRoute;
                    rowOrder.Range[1, OrderTable.ListColumns["Вес нетто"].Index].Value = order.WeightNetto;
                   
                    
                }
            }
            pb.Close();
        }

        /// <summary>
        /// Распределить заказы по автомобилям
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        public List<Delivery> CompleteAuto(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orderList = orders.OrderBy(x => x.WeightNetto).ToList();
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;

            #region Проверка если клиента (точки) нет в таблице маршрутов
            Delivery emptyDelivery = null;
            bool emptyPoint;
            for (int k = orderList.Count - 1; k >= 0; k--)
            {
                emptyPoint = true;
                foreach (DeliveryPoint point in pointMap)
                {
                    emptyPoint = true;
                    if (orderList[k].Customer.Id == point.IdCustomer)
                    {
                        emptyPoint = false;
                        break;
                    }

                }
                if (emptyPoint)
                {
                    if (emptyDelivery == null)
                    {
                        emptyDelivery = new Delivery(orderList[k]);
                    }
                    else
                    {
                        emptyDelivery.Orders.Add(orderList[k]);
                    }
                    orderList.RemoveAt(k);
                }
            }
            if (emptyDelivery != null)
            {
                emptyDelivery.hasRoute = false;
                deliveries.Add(emptyDelivery);
            }
            #endregion 

            Delivery delivery = null;
            int pointNumber = 0;
            while (orderList.Count > 0)
            {

                for (int orderNumber = orderList.Count - 1; orderNumber >= 0; orderNumber--)
                {

                    if (orderList[orderNumber].Customer.Id != pointMap[pointNumber].IdCustomer) continue;

                    if (delivery == null)
                    {
                        orderList[orderNumber].DeliveryPoint = pointMap[pointNumber];
                        orderList[orderNumber].PointNumber = 1;

                        delivery = new Delivery(orderList[orderNumber]);
                        delivery.Number = 1;
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
                            delivery.Number = deliveries.Count + 1;
                            Order orderLastAdd = delivery.Orders.Last();
                            orderLastAdd.PointNumber = delivery.MapDelivery.Count;
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
                                Order orderLastAdd = delivery.Orders.Last();
                                orderLastAdd.PointNumber = delivery.MapDelivery.Count;
                                orderList.RemoveAt(orderNumber);
                            }
                        }
                    }
                }

                Debug.WriteLine($"pointNumber = {pointNumber}");
                pointNumber++;
                if (pointNumber >= pointMap.Count) pointNumber = 0;
            }
            if (delivery != null && delivery.Orders.Count > 0)
            {
                deliveries.Add(delivery);
            }
            return deliveries;
        }

        #region  Сбор данных sap

        /// <summary>
        /// Поиск 
        /// </summary>
        /// <param name="sapPath"></param>
        /// <returns></returns>
        public List<Order> GetSapOrders(string sapPath)
        {
            List<Order> rourerOrders = new List<Order>();
           
            Workbook sapBook = null;
           
            try
            {
                sapBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: sapPath);                  
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
                ProcessBar pb = ProcessBar.Init("Сбор данных", range.Rows.Count, 1, "Формирование доставок");

                if (pb == null) return null;
                pb.Show();
                foreach (Range row in range.Rows)
                {
                    if (pb.Cancel) break;
                    pb.Action("Заказ " + (row.Row - range.Row + 1) + " из " + pb.Count);
                    Order order = ReadSapRow(row);
                    if (order !=null)
                    rourerOrders.Add(order);
                }
                pb.Close();
            }
            sapBook.Close();
           
            return rourerOrders;
        }



         /// <summary>
         /// Искать Стоимость посылки кол-во паллетов и ТТН
         /// </summary>
         /// <param name="ordersPath"></param>
         /// <param name="ordersSap"></param>
         /// <returns></returns>
        public List<Order> GetOrdersInfo(string ordersPath, List<Order> ordersSap)
        {
            List<Order> OrdersInfo = new List<Order>();
            Workbook orderBook = null;
            try
            {
                if (ordersPath != "")
                   orderBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: ordersPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть книгу Excel");
            }
            foreach (Order order in ordersSap)
            {
                
                    // read  All orders
                    if (!string.IsNullOrWhiteSpace(order.Id))
                    {
                        List<string> orderInfo = GetOrderInfo(orderBook.Sheets[1], order.Id);
                        if (orderInfo != null)
                        {
                            string costStr = orderInfo.Find(x => x.Contains("Стоимость")) ?? "";
                            Regex regexCost = new Regex(@"\d+(\,\d+)?");
                            costStr = costStr.Replace(".", "");
                            costStr = regexCost.Match(costStr).Value;
                            order.Cost = double.TryParse(costStr, out double cost) ? cost : 0;

                            int ix = orderInfo.FindIndex(x => x.Contains("грузовых мест:"));
                            if (ix > 0)
                            {
                                string pallets = orderInfo[ix] ?? "";
                                pallets = string.Join("", pallets.Where(c => char.IsDigit(c)));
                                order.PalletsCount = int.TryParse(pallets, out int p) ? p : 0;
                                //  order.Customer.Name = orderInfo[ix + 1]; 
                                order.Customer.AddresStreet = orderInfo[ix + 2];
                                order.Customer.AddresCity = orderInfo[ix + 3];
                            }

                            string weightBrutto = orderInfo.Find(x => x.Contains("вес")) ?? "";
                            weightBrutto = weightBrutto.Replace(".", "");
                            Regex regex = new Regex(@"\d+(,\d+)?");
                            weightBrutto = regex.Match(weightBrutto).Value;
                            order.WeightBrutto = double.TryParse(weightBrutto, out double wb) ? wb : 0;
                        }
                    }
                    OrdersInfo.Add(order);                  
            }

            orderBook.Close();
            return OrdersInfo;
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

            order.Id = row.Cells[1, GetColumn(row.Parent, "Delivery", 1)].Text;
            if (string.IsNullOrWhiteSpace(order.Id)) return null;

            order.TransportationUnit = row.Cells[1, GetColumn(row.Parent, "Transportation Unit", 1)].Text;
            string idCusomer = row.Cells[1, GetColumn(row.Parent, "Получатель материала", 1)].Text;
            order.Customer.Id = idCusomer;
            order.Customer.Name = row.Cells[1, GetColumn(row.Parent, "Описание получателя материала", 1)].Text;


            if (string.IsNullOrWhiteSpace(idCusomer) || string.IsNullOrWhiteSpace(order.Id))
            {
                return null;
            }
            // Вес брутто для товара будет весом нетто для доставки 
            string weight = row.Cells[1, GetColumn(row.Parent, "Вес брутто", 1)].Text;
            order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
            order.Route = row.Cells[1, GetColumn(row.Parent, "Маршрут", 1)].Text;
            return order;
        }

        #endregion Сбор данных sap

        /// <summary>
        /// Записать накладную по доставке из файла выгруза в Лист
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="transportationUnit"></param>
        /// <returns></returns>
        private List<string> GetOrderInfo(Worksheet sheet, string delivery)
        {
            Range findRange = sheet.Columns[1];
            //string search = "№ ТТН:" + new string('0', 18 - transportationUnit.Length) + transportationUnit;
            //  string search = new string('0', 18 - transportationUnit.Length) + transportationUnit;


            string search = new string('0', 10 - delivery.Length) + delivery;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);

            if (fcell == null && fcell.Value.Trim().Contains("Доставка")) return null;

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
        /// Заполнить таблицу отгрузки
        /// </summary>
        /// <param name="totalTable"></param>
        /// <param name="deliveries"></param>
        private void PrintShipping(ListObject totalTable, List<Delivery> deliveries)
        {
            ShefflerWorkBook shefflerBook = new ShefflerWorkBook();
            ListRow row;
            if (totalTable.ListRows.Count > 0)
            {
                totalTable.DataBodyRange.Rows.Delete();
            }

            if (totalTable.ListRows.Count == 0)
            {
                totalTable.ListRows.Add();
                row = totalTable.ListRows[1];//  }
            }
            row = totalTable.ListRows[totalTable.ListRows.Count - 1];

            foreach (Delivery delivery in deliveries)
            {
                row.Range[1, totalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.CostDelivery;


                foreach (Order order in delivery.Orders)
                {
                    string date = shefflerBook.GetDateDelivery();
                    row.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value = date ;

                    row.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value =
                            delivery.MapDelivery.FindIndex(x => x.IdCustomer == order.Customer.Id) + 1;

                    row.Range[1, totalTable.ListColumns["Номер накладной"].Index].Value = order.TransportationUnit;
                    row.Range[1, totalTable.ListColumns["Номер поставки"].Index].Value = order.Id;
                    row.Range[1, totalTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    row.Range[1, totalTable.ListColumns["Направление"].Index].Value = order.Route;
                    row.Range[1, totalTable.ListColumns["Номер грузополучателя"].Index].Value = order.Customer.Id;
                    row.Range[1, totalTable.ListColumns["Брутто вес"].Index].Value = order.WeightBrutto;
                    row.Range[1, totalTable.ListColumns["Нетто вес"].Index].Value = order.WeightNetto;

                    row.Range[1, totalTable.ListColumns["Грузополучатель"].Index].Value = $"{order.Customer.Name}";
                    //                   $"{order.Customer.Name}  {order.Customer.Name} {order.Customer.AddresStreet}";
                    row.Range[1, totalTable.ListColumns["Стоимость поставки"].Index].Value = order.Cost;
                    row.Range[1, totalTable.ListColumns["Кол-во паллет"].Index].Value = order.PalletsCount;            
                   

                    totalTable.ListRows.Add();
                    row = totalTable.ListRows[totalTable.ListRows.Count - 1];
                }
            }
        }
        /// <summary>
        /// Изменить
        /// </summary>
        internal void СhangeDelivery()
        {
            ExcelOptimizateOn();
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];

            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];

            List<Order> orders = GetOrdersFromTable(ordersTable);
            List<Delivery> deliveries = EditDeliveres(orders);
            ClearListObj(carrierTable);
            PrintDelivery(deliveries, carrierTable);

            foreach (ListRow row in ordersTable.ListRows)
            {
                string strNum = row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : 0;
                if (deliveryNumber == 0) continue;
                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                Delivery delivery = deliveries.Find(d => d.Number == deliveryNumber);
                if (delivery == null) continue;

                Order order = delivery.Orders.Find(r => r.Customer.Id == customerId);
                if (order != null)
                {
                    row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                }
            }

            ExcelOptimizateOff();
        }

        private List<Order> GetOrdersFromTable(ListObject ordersTable)
        {
            List<Order> orders = new List<Order>();
            //  List<int> deliveryNumbers = new List<int>();


            foreach (ListRow row in ordersTable.ListRows)
            {
                Order order = new Order();
                string strNum = row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : 0;
                if (deliveryNumber == 0) continue;
                order.NumberDelivery = deliveryNumber;
                order.TransportationUnit = row.Range[1, ordersTable.ListColumns["Накладная"].Index].Text;

                strNum = row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Text;
                order.PointNumber = int.TryParse(strNum, out int pointnum) ? pointnum : 0;

                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                Customer customer = new Customer(customerId);
                order.Customer = customer;
                string CityStr = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;
                order.DeliveryPoint = new DeliveryPoint() { City = CityStr };

                string weight = row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
                orders.Add(order);
            }
            return orders;
        }


        internal void AcceptDelivery()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];

            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orders = GetOrdersFromTable(ordersTable);


            foreach (ListRow delveryRow in carrierTable.ListRows)
            {
                string str = delveryRow.Range[1, carrierTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(str, out int num) ? num : 0;

                if (deliveryNumber > 0)
                {
                    Delivery delivery = new Delivery();
                    delivery.Orders = orders.FindAll(x => x.NumberDelivery == deliveryNumber).ToList();
                    deliveries.Add(delivery);
                }
            }

            foreach (ListRow totalRow in totalTable.ListRows)
            {
                string transportationUnit = totalRow.Range[1,
                                 totalTable.ListColumns["Номер накладной"].Index].Text;

                foreach (Delivery delivery in deliveries)
                {
                    transportationUnit = new string('0', 18 - transportationUnit.Length) + transportationUnit;
                    Order orderf = delivery.Orders.Find(x => x.TransportationUnit == transportationUnit);
                    if (orderf != null)
                    {

                        totalRow.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value = orderf.PointNumber;
                    }
                }
            }

            totalSheet.Activate();
        }




        /// <summary>
        /// Прменять список доставок для списка заказов
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        private List<Delivery> EditDeliveres(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();

            /// Список номеров доставок
            List<int> deliveryNumbers = (from o in orders
                                         select o.NumberDelivery).Distinct().ToList();
            for (int i = 0; i < deliveryNumbers.Count; i++)
            {
                int deliveryNumber = deliveryNumbers[i];
                if (deliveryNumber > 0)
                {

                    List<Order> orderList = orders.FindAll(
                                o => o.NumberDelivery == deliveryNumber).ToList().OrderBy(
                                                                x => x.PointNumber).ToList();
                    if (orderList.Count > 0)
                    {
                        Delivery delivery = EditDelivery(orderList);
                        delivery.Number = deliveries.Count + 1;
                        deliveries.Add(delivery);
                    }
                }
            }
            // По каждой доставке создать список заказов 
            // найти подходящий маршрут
            //


            #region Добавление нового маршрута
            #endregion
            return deliveries;
        }

        /// <summary>
        /// Изменить доставку
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        private Delivery EditDelivery(List<Order> orders)
        {
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
            Delivery delivery = new Delivery();
            int idRoute = FindRoute(orders, functionsBook);
            if (idRoute == 0)
            {

            }
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;

            foreach (Order order in orders)
            {
                order.DeliveryPoint = pointMap.Find(p => p.IdRoute == idRoute &&
                                                 p.IdCustomer == order.Customer.Id);

            }
            orders = orders.OrderBy(b => b.DeliveryPoint.PriorityPoint).ToList();
            int number = 1;
            for (int i = 0; i < orders.Count; i++)
            {
                if (i > 0 && orders[i].DeliveryPoint.IdCustomer != orders[i - 1].DeliveryPoint.IdCustomer)
                {
                    ++number;
                }
                orders[i].PointNumber = number;
            }
            delivery.Orders = orders;
            return delivery;
        }

        private int AddRoute(List<Order> orders)
        {
            Worksheet sheetRoute = Globals.ThisWorkbook.Sheets["Routes"];
            ListObject TableRoutes = sheetRoute?.ListObjects["TableRoutes"];
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
            DeliveryPoint deliveryPoint = functionsBook.RoutesTable.Last();
            int idRoute = deliveryPoint.IdRoute + 1;
            int PriorityRoute = deliveryPoint.PriorityRoute + 1;
            int point = 0;
            foreach (Order order in orders)
            {
                ListRow row = TableRoutes.ListRows[TableRoutes.ListRows.Count];
                TableRoutes.ListRows.Add();
                row.Range[1, 1].Value = idRoute;
                row.Range[1, 2].Value = PriorityRoute;
                row.Range[1, 3].Value = ++point;
                row.Range[1, 5].Value = order.Customer.Id;

            }
            return idRoute;
        }





        /// <summary>
        /// Поиск маршрута где есть все клиенты из списка заказов
        /// </summary>
        /// <param name="orders"></param>
        /// <param name="functionsBook"></param>
        /// <returns></returns>
        private int FindRoute(List<Order> orders, ShefflerWorkBook functionsBook)
        {
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;
            //список id маршрутов
            List<int> uRoutes = (from p in pointMap
                                 select p.IdRoute).Distinct().ToList();

            for (int i = 0; i < uRoutes.Count; i++)
            {
                int idRoute = uRoutes[i];
                bool hasRoute = true;
                foreach (Order order in orders)
                {
                    List<DeliveryPoint> routes = pointMap.FindAll(
                                 x => x.IdRoute == idRoute &&
                                 x.IdCustomer == order.Customer.Id).ToList();
                    if (routes.Count == 0)
                    {
                        hasRoute = false;
                        break;
                    }
                }
                if (hasRoute)
                {
                    return idRoute;
                }
            }
            return 0;
        }

        #region Вспомогательные

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

        /// <summary>
        /// Оптимизация Excel
        /// </summary>
        public static void ExcelOptimizateOn()
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Globals.ThisWorkbook.Application.Calculation = XlCalculation.xlCalculationManual;
        }

        /// <summary>
        /// Возврат Excel в исходное состояние
        /// </summary>
        public static void ExcelOptimizateOff()
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
            Globals.ThisWorkbook.Application.Calculation = XlCalculation.xlCalculationAutomatic;
        }
        #endregion Вспомогательные




    }
}
