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
                ListObject OrdersTable = deliverySheet.ListObjects["TableOrders"];
                Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
                ListObject TotalTable = TotalSheet.ListObjects["TableTotal"];

                ClearListObj(carrierTable);
                if (OrdersTable.DataBodyRange.Rows.Count > 0)
                { OrdersTable.DataBodyRange.Rows.Delete(); }

                if (deliveries != null && deliveries.Count > 0)
                {
                    PrintDelivery(deliveries, carrierTable, OrdersTable);

                    PrintShipping(TotalTable, deliveries);
                }
            }
            ExcelOptimizateOff();
        }



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
        /// Запись доставок в таблицы  лист Delivery
        /// </summary>
        /// <param name="deliveries"></param>
        /// <param name="CarrierTable"></param>
        /// <param name="OrderTable"></param>
        private void PrintDelivery(List<Delivery> deliveries, ListObject CarrierTable, ListObject OrderTable)
        {

            ProcessBar pb = ProcessBar.Init("Вывод данных", deliveries.Count, 1, "Формирование доставок");
            if (pb == null) return ;
            pb.Show();

            for (int i = 0; i < deliveries.Count; i++)
            {
                if (pb == null) return;

                if (pb.Cancel) break;
                pb.Action($"Доставка {i + 1} из {pb.Count}");

                Delivery delivery = deliveries[i];
                System.Windows.Forms.Application.DoEvents();
                // Worksheet deliverySheet

                if (CarrierTable == null || OrderTable == null)
                {
                    MessageBox.Show("Отсутствует таблица");
                    return;
                }
                ListRow rowCarrier;
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

                int numberDelivery = 0;
                if (delivery.hasRoute )
                {
                    numberDelivery = i + 1;
                }
                rowCarrier.Range[1, CarrierTable.ListColumns["№ Доставки"].Index].Value = numberDelivery;
                rowCarrier.Range[1, CarrierTable.ListColumns["Компания"].Index].Value = delivery.Truck?.ShippingCompany?.Name ?? "";
                rowCarrier.Range[1, CarrierTable.ListColumns["ID Route"].Index].Value = delivery.MapDelivery[0].IdRoute;
                rowCarrier.Range[1, CarrierTable.ListColumns["Тоннаж"].Index].Value = delivery.Truck?.Tonnage ?? 0;

                //rowCarrier.Range[1, CarrierTable.ListColumns["Вес доставки"].Index].Value = delivery.TotalWeight;
                rowCarrier.Range[1, CarrierTable.ListColumns["Вес доставки"].Index].FormulaR1C1 =
                                                "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес нетто])";

                //rowCarrier.Range[1, CarrierTable.ListColumns["Стоимость товаров"].Index].Value = delivery.CostProducts;
                rowCarrier.Range[1, CarrierTable.ListColumns["Стоимость товаров"].Index].Value =
                                            "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Стоимость товаров])";
                rowCarrier.Range[1, CarrierTable.ListColumns["Стоимость доставки"].Index].Value = delivery.CostDelivery;

                int columnMap = 0;
                foreach (DeliveryPoint point in delivery.MapDelivery)
                {
                    ++columnMap;
                    rowCarrier.Range[1, CarrierTable.ListColumns.Count].Offset[0, 2 + columnMap].Value
                                    = $"{point.IdCustomer} - {point.City} ";
                }
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
                    rowOrder.Range[1, OrderTable.ListColumns["№ Доставки"].Index].Value = numberDelivery;
                    rowOrder.Range[1, OrderTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;

                    rowOrder.Range[1, OrderTable.ListColumns["Накладная"].Index].Value = order.TransportationUnit;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Получателя"].Index].Value = order.Customer?.Id ?? "";
                    rowOrder.Range[1, OrderTable.ListColumns["Получатель"].Index].Value = order.Customer.Name;
                    rowOrder.Range[1, OrderTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Маршрута"].Index].Value = order.DeliveryPoint.IdRoute ;
                    rowOrder.Range[1, OrderTable.ListColumns["Колличество паллет"].Index].Value = order.PalletsCount;
                    rowOrder.Range[1, OrderTable.ListColumns["Вес нетто"].Index].Value = order.WeightNetto;
                    rowOrder.Range[1, OrderTable.ListColumns["Стоимость товаров"].Index].Value = order.Cost;
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
            if (delivery != null && delivery.Orders.Count > 0) deliveries.Add(delivery);
            return deliveries;
        }

        #region  Сбор данных sap

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
                ProcessBar pb = ProcessBar.Init("Сбор данных", range.Rows.Count, 1, "Формирование доставок");

                if (pb == null) return null;
                pb.Show();
                foreach (Range row in range.Rows)
                {
                    if (pb.Cancel) break;
                    pb.Action("Заказ " + (row.Row - range.Row + 1) + " из " + pb.Count);
                    Order order = ReadSapRow(row);
                    if (order != null)
                    {

                        if (!string.IsNullOrWhiteSpace(order.TransportationUnit))
                        {
                            List<string> orderInfo = GetOrderInfo(orderBook.Sheets[1], order.TransportationUnit);
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
                        rourerOrders.Add(order);
                    }
                }
                pb.Close();
            }
            sapBook.Close();
            orderBook.Close();
            return CompleteAuto(rourerOrders);
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
            order.Customer.Id = idCusomer;
            order.Customer.Name = row.Cells[1, GetColumn(row.Parent, "Описание получателя материала", 1)].Text;

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

        #endregion Сбор данных sap

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
        /// Заполнить таблицу отгрузки
        /// </summary>
        /// <param name="totalTable"></param>
        /// <param name="deliveries"></param>
        private void PrintShipping(ListObject totalTable, List<Delivery> deliveries)
        {
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
                    //row.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value =
                    //        delivery.MapDelivery.FindIndex(x => x.IdCustomer == order.Customer.Id) + 1;

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

        internal void СhangeDelivery()
        {
            ExcelOptimizateOn();
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];


            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];

            List<Order> orders = GetOrdersFromTable(ordersTable);
            List<Delivery> deliveries = ChangeDeliveres(orders);
            PrintDelivery(deliveries, carrierTable, ordersTable);
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
            Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject TotalTable = TotalSheet.ListObjects["TableTotal"];

            List<Order> orders = GetOrdersFromTable(ordersTable);
            List<int> deliveryNumbers = new List<int>();

            foreach (ListRow delveryRow in carrierTable.ListRows)
            {
                string str = delveryRow.Range[1, carrierTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(str, out int num) ? num : 0;
                if (deliveryNumbers.Find(x => x == deliveryNumber) == 0)
                {
                    deliveryNumbers.Add(deliveryNumber);
                }
            }

            foreach (ListRow deliveryRow in carrierTable.ListRows)
            {


            }
            //  for (int i = OrdersTable.ListRows.Count; i >= 0; --i)
            foreach (ListRow row in ordersTable.ListRows)
            {
                //ListRow row = OrdersTable.ListRows[i];
                Order order = new Order();
                //string str = row.Range[1, OrdersTable.ListColumns["№ Доставки"].Index].Text;
                //int deliveryNumber = int.TryParse(str, out int num) ? num : 0;
                //if (deliveryNumber > 0 && deliveryNumbers.Find(x => x == deliveryNumber) == 0)
                //{
                //    //  row.Range.Rows.Delete();
                //    //добавить
                //}


                order.TransportationUnit = row.Range[1, ordersTable.ListColumns["Накладная"].Index].Text;
                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                Customer customer = new Customer(customerId);
                order.Customer = customer;
                string weight = row.Range[1, ordersTable.ListColumns["Накладная"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
                 
                //orders.Add
            }
            List<Delivery> deliveries = ChangeDeliveres(orders);
        }

       


         /// <summary>
         /// Прменять список доставок для списка заказов
         /// </summary>
         /// <param name="orders"></param>
         /// <returns></returns>
        private List<Delivery> ChangeDeliveres(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            
            /// Список номеров доставок
            List<int> deliveryNumbers = (from o in orders
                                         select o.NumberDelivery).Distinct().ToList();
            for ( int i =0;i< deliveryNumbers.Count; i++)
            {
              int deliveryNumber =  deliveryNumbers[i];
                if (deliveryNumber > 0)
                {

                List<Order> orderList = orders.FindAll(
                            o=>o.NumberDelivery == deliveryNumber).ToList().OrderBy(
                                                            x => x.PointNumber).ToList();
                    if (orderList.Count > 0)
                    {
                      Delivery delivery= EditDelivery(orderList);
                        deliveries.Add(delivery);
                    }
                }
            }
            // По каждой доставке создать список заказов 
            // найти подходящий маршрут
            //


            #region Добавление нового маршрута
            # endregion
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
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;

            foreach (Order order in orders)
            {
                order.DeliveryPoint = pointMap.Find(p => p.IdRoute == idRoute &&
                                                 p.IdCustomer == order.Customer.Id);
            }             
            delivery.Orders = orders;
            return delivery;
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

           for (int i = 0; i< uRoutes.Count; i++)
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
