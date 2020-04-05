using DomesticTransport.Forms;
using DomesticTransport.Model;
using DomesticTransport.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Provider = DomesticTransport.Model.Provider;

namespace DomesticTransport
{
    /// <summary>
    /// Основной алгоритм 
    /// </summary>
    class Functions
    {
        /// <summary>
        /// Вывести на рабочий лист доставки 
        /// </summary>
        public void SetDelivery()
        {
            SapFiles sapFiles = new SapFiles();
            sapFiles.ShowDialog();
            if (sapFiles.DialogResult != DialogResult.OK) return;

            string sapPath = "";
            string ordersPath = "";
            try
            {
                sapPath = sapFiles.ExportFile;
                ordersPath = sapFiles.OrderFile;
            }
            catch
            {
                return;
            }
            finally
            {
                sapFiles.Close();
            }

            ShefflerWB.ExcelOptimizateOn();
            List<Order> orders = GetOrdersFromSap(sapPath);

            if (ordersPath != "" && File.Exists(ordersPath))
            {
                orders = GetOrdersInfo(ordersPath, orders);  // Поиск свойств в файле All orders
            }

            List<Delivery> deliveries = CompleteAuto(orders);

            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject TotalTable = TotalSheet.ListObjects["TableTotal"];

            ClearListObj(carrierTable);
            if (ordersTable.DataBodyRange.Rows.Count > 0)
            {
                Globals.ThisWorkbook.Application.DisplayAlerts = false;
                ordersTable.DataBodyRange.Rows.Delete();
                Globals.ThisWorkbook.Application.DisplayAlerts = true;
            }

            if (deliveries != null && deliveries.Count > 0)
            {
                PrintDelivery(deliveries, carrierTable);
                PrintOrders(deliveries, ordersTable);
                PrintShipping(TotalTable, deliveries);
            }
            deliverySheet.Columns.AutoFit();

            ShefflerWB.ExcelOptimizateOff();
        }



        /// <summary>
        /// Загрузка All Orders 
        /// </summary>
        public void LoadAllOrders()
        {
            ShefflerWB functionsBook = new ShefflerWB();
            Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = TotalSheet.ListObjects["TableTotal"];
            Range range = totalTable.DataBodyRange;
            if (range == null || totalTable == null) return;
            string file = SapFiles.SelectFile();
            if (!File.Exists(file)) return;
            List<Order> orders = GetOrdersFromTotalTable(range);
            orders = GetOrdersInfo(file, orders);
            if (orders == null || orders.Count == 0) return;
            int columnId = totalTable.ListColumns["Номер поставки"].Index;
            foreach (Range row in range.Rows)
            {
                string idOrder = row.Cells[1, columnId].Text;
                if (string.IsNullOrWhiteSpace(idOrder)) continue;
                idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                Order order = orders.Find(o => o.Id == idOrder);
                if (order == null) continue;

                row.Cells[1, totalTable.ListColumns["Брутто вес"].Index].Value = order.WeightBrutto;
                row.Cells[1, totalTable.ListColumns["Стоимость поставки"].Index].Value = order.Cost;
                row.Cells[1, totalTable.ListColumns["Кол-во паллет"].Index].Value = order.PalletsCount;
            }
        }

        private List<Order> GetOrdersFromTotalTable(Range range)
        {
            List<Order> orders = new List<Order>();
            Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = TotalSheet.ListObjects["TableTotal"];

            int column = totalTable.ListColumns["Номер поставки"].Index;

            foreach (Range row in range.Rows)
            {
                string idOrder = row.Cells[1, column].Text;
                if (string.IsNullOrWhiteSpace(idOrder)) continue;
                Order order = new Order();
                idOrder = new string('0', 10 - idOrder.Length) + idOrder;
                order.Id = idOrder;
                orders.Add(order);
            }

            return orders;
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
        public void AddAuto()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject deliveryTable = deliverySheet.ListObjects["TableCarrier"];
            int idRoute = 0;
            int number = 0;
            if (deliveryTable.ListColumns["№ Доставки"].DataBodyRange != null)
            {
                foreach (Range rng in deliveryTable.ListColumns["№ Доставки"].DataBodyRange)
                {
                    if (int.TryParse(rng.Text, out int valueCell))
                    {
                        if (number < valueCell) number = valueCell;
                    }
                }
            }
            number++;

            // Выделенный диапазон
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Range selection = Globals.ThisWorkbook.Application.Selection;
            Range orfderRng = Globals.ThisWorkbook.Application.Intersect(selection, ordersTable.DataBodyRange);
            Delivery delivery = null;
            if (orfderRng != null)
            {
                Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
                ListObject totalTable = totalSheet.ListObjects["TableTotal"];
                ShefflerWB workBook = new ShefflerWB();

                string orderId = "";
                List<Order> orders = new List<Order>();

                foreach (Range orderLine in orfderRng.Rows)
                {
                    Range cl = deliverySheet.Cells[orderLine.Row, 2];
                    orderId = cl.Offset[0, 1].Text; //  "Номер поставки"
                    cl.Value = number;
                    double weight = double.TryParse(cl.Offset[0, 4].Text, out double wgt) ? wgt : 0;
                    string idCustomer = cl.Offset[0, 5].Text;
                    Customer customer = new Customer(idCustomer);
                    orders.Add(new Order()
                    {
                        Id = orderId,
                        WeightNetto = weight,
                        Customer = customer
                    });
                }
                List<Delivery> deliveries = CompleteAuto(orders);
                Range totalRng = workBook.GetCurrentShippingRange();
                if (deliveries != null && deliveries.Count > 0 && totalRng != null)
                {
                    delivery = deliveries[0];
                    idRoute = delivery.MapDelivery[0].Id;

                    foreach (Range row in totalRng.Rows)
                    {
                        string idOrderTotal = row.Cells[0, totalTable.ListColumns["Номер поставки"].Index].Text;
                        idOrderTotal = idOrderTotal.Length < 10 ? new string('0', 10 - idOrderTotal.Length) + idOrderTotal : idOrderTotal;
                        Order findOrder = orders.Find(x => x.Id == idOrderTotal);
                        if (findOrder != null)
                        {
                            row.Cells[0, totalTable.ListColumns["№ Доставки"].Index].Value = number.ToString();
                            row.Cells[0, totalTable.ListColumns["Порядок выгрузки"].Index].Value = findOrder.PointNumber;
                            row.Cells[0, totalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                        }
                    }
                    foreach (Range orderLine in orfderRng.Rows)
                    {
                        deliverySheet.Cells[orderLine.Row, 4].Value = delivery.MapDelivery[0].Id;   //ID Route
                    }
                }


            }
            ListRow rowDelivery;
            if (deliveryTable.ListRows.Count == 0)
            {
                AddListRow(deliveryTable);
                rowDelivery = deliveryTable.ListRows[1];//  }
            }
            else
            {
                AddListRow(deliveryTable);
                rowDelivery = deliveryTable.ListRows[deliveryTable.ListRows.Count - 1];
            }
            rowDelivery.Range[1, deliveryTable.ListColumns["№ Доставки"].Index].Value = number;
            if (delivery != null)
            {
                rowDelivery.Range[1, deliveryTable.ListColumns["ID Route"].Index].Value = delivery.MapDelivery[0].Id;
                rowDelivery.Range[1, deliveryTable.ListColumns["Компания"].Index].Value = delivery.Truck?.ShippingCompany?.Name ?? "";
                rowDelivery.Range[1, deliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                rowDelivery.Range[1, deliveryTable.ListColumns["Тоннаж"].Index].Value = delivery.Truck.Tonnage;
            }
        }


        /// <summary>
        ///кнопка Добавить авто
        /// </summary>
        public void DeleteAuto()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject deliveryTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject TotalTable = totalSheet.ListObjects["TableTotal"];

            if (deliveryTable == null || ordersTable == null) return;
            Range Target = Globals.ThisWorkbook.Application.Selection;

            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, deliveryTable.DataBodyRange);
            if (commonRng == null) return;

            DialogResult msg = MessageBox.Show("Удалить авто с заказами", "Удалить", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (DialogResult.No == msg) return;
            ShefflerWB workBook = new ShefflerWB();

            int numberDelivery = 0;
            int row = commonRng.Row;
            int column = deliveryTable.ListColumns["№ Доставки"].Range.Column;
            // commonRng = Globals.ThisWorkbook.Application.Intersect(
            commonRng = deliverySheet.Cells[row, column];
            numberDelivery = int.TryParse(commonRng.Text, out int nmDelivery) ? nmDelivery : 0;

            //foreach (ListRow listDeliveryRow in deliveryTable.ListRows)
            for (int i = deliveryTable.ListRows.Count; i > 0; --i)
            {
                ListRow listDeliveryRow = deliveryTable.ListRows[i];
                Range deliveryCell = listDeliveryRow.Range[1, deliveryTable.ListColumns["№ Доставки"].Index];
                string str = deliveryCell != null ? deliveryCell.Text : "";
                if (int.TryParse(str, out int number))
                {
                    if (number == numberDelivery)
                        deliverySheet.Rows[listDeliveryRow.Range.Row].Delete();
                }
            }

            for (int j = ordersTable.ListRows.Count; j > 0; --j)
            {
                ListRow listOrderRow = ordersTable.ListRows[j];
                Range orderCell = listOrderRow.Range[1, ordersTable.ListColumns["№ Доставки"].Index];
                string strDeliveryNum = orderCell.Offset[0, 1].Text;
                strDeliveryNum = orderCell != null ? orderCell.Text : "";
                if (int.TryParse(strDeliveryNum, out int DeliveryNum))
                {
                    if (DeliveryNum == numberDelivery)
                        deliverySheet.Rows[listOrderRow.Range.Row].Delete();

                }

            }
            Range rng = workBook.GetCurrentShippingRange();
            if (rng == null) return;
            for (int k = rng.Rows.Count; k > 0; k--)
            {
                string idDelivery = rng.Rows[k].Cells[0,
                         TotalTable.ListColumns["№ Доставки"].Index].Text;
                if (int.TryParse(idDelivery, out int num))
                {
                    if (num == numberDelivery)
                    {
                        totalSheet.Rows[rng.Rows[k].Row - 1].Delete();
                    }
                }
            }

        }

        //TODO УДАЛИТЬ из таблицы Total

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
                // if (delivery.Truck == null) continue;
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
                rowDelivery.Range[1, DeliveryTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                
                if (!string.IsNullOrWhiteSpace(delivery.Truck?.ShippingCompany?.Name)) //delivery?.MapDelivery.Find(m => m.RouteName == "") != null)
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["Компания"].Index].Value =
                                                                        delivery.Truck?.ShippingCompany?.Name ?? "";
                    rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                }
                else
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["Компания"].Index].Value = "Деловые линии";
                    rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = "";
                }

                if (delivery?.MapDelivery.Count > 0)
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["ID Route"].Index].Value =
                                                                        delivery?.MapDelivery[0].Id;
                }
                rowDelivery.Range[1, DeliveryTable.ListColumns["Тоннаж"].Index].Value = delivery.Truck?.Tonnage ?? 0;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Вес доставки"].Index].FormulaR1C1 =
                                                "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес нетто])";

            }
            pb.Close();
        }

        /// <summary>
        /// Вывод заказов
        /// </summary>
        /// <param name="deliveries"></param>
        /// <param name="OrderTable"></param>
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
                    PrintOrder(rowOrder, order, delivery.Number);
                }
            }
            pb.Close();
        }
          public void PrintOrder(ListRow row ,Order order,  int deliveryNumber)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];              
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];

            row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Value = deliveryNumber;
            row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
            row.Range[1, ordersTable.ListColumns["Доставка"].Index].Value = order.Id;
            row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Value = order.Customer?.Id ?? "";
            row.Range[1, ordersTable.ListColumns["Получатель"].Index].Value = order.Customer.Name;
            row.Range[1, ordersTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
            row.Range[1, ordersTable.ListColumns["ID Route"].Index].Value = order.DeliveryPoint.Id;
            row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Value = order.WeightNetto;
            row.Range[1, ordersTable.ListColumns["Маршрут"].Index].Value = order.Route;
        }


        /// <summary>
        /// Распределить заказы по автомобилям
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        public List<Delivery> CompleteAuto(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            orders = orders.OrderBy(x => x.WeightNetto).ToList();
            ShefflerWB functionsBook = new ShefflerWB();
            List<DeliveryPoint> points = functionsBook.RoutesTable;
            Delivery deliveryNoRoute = new Delivery();
            deliveryNoRoute.HasRoute = false;

            while (orders.Count > 0)
            {
                bool findDelivery = false;
                // Проходим по возможным маршрутам
                foreach (DeliveryPoint point in points)
                {
                    // Ищем товар, который можно отправить указанным маршрутом                    
                    for (int iOrder = orders.Count - 1; iOrder >= 0; iOrder--)
                    {
                        if (! orders[iOrder].Customer.Id.Contains(point.IdCustomer)) continue;
                        findDelivery = true;
                        orders[iOrder].DeliveryPoint = point;
                        // Пытаемся добавить к имеющимся машинам
                        Delivery delivery = null;
                        foreach (Delivery iDelivery in deliveries)
                        {
                            string city = iDelivery.Orders[0].DeliveryPoint.City;
                            city = city.Trim();
                            if (iDelivery.Orders[0].DeliveryPoint.Id != point.Id) continue;
                            if ((city.Contains("MSK") ||
                                city.Contains("MO")) && iDelivery.MapDelivery.Count > 3) { continue; }

                            if ((city.Contains("Yerevan") ||
                                city.Contains("Nur-Sultan")) )
                            {
                               if (iDelivery.TotalWeight + orders[iOrder].WeightNetto <= 3300)
                                {
                                delivery = iDelivery;
                                break;
                                }
                            }
                            else if (iDelivery.CheckDeliveryWeght(orders[iOrder]))
                            {
                                delivery = iDelivery;
                                break;
                            }
                        }
                        if (delivery == null)
                        {
                            delivery = new Delivery();
                            deliveries.Add(delivery);
                        }
                        orders[iOrder].DeliveryPoint = point;
                        Order orderCurrentCustomer = delivery.Orders.Find(x => x.Customer.Id == orders[iOrder].Customer.Id);
                        //Порядок выгруза / Если уже есть груз для заказчика 
                        int number = orderCurrentCustomer == null ?
                              delivery.Orders.Count + 1
                            : orderCurrentCustomer.PointNumber;
                        orders[iOrder].PointNumber = number;
                        delivery.Orders.Add(orders[iOrder]);
                        delivery.Number = deliveries.Count;
                        orders.RemoveAt(iOrder);
                    }
                    if (findDelivery) break;
                }
                // не нашли маршрут
                if (!findDelivery)
                {
                    deliveryNoRoute.Orders.Add(orders[0]);
                    deliveryNoRoute.Number = deliveries.Count;
                    orders.RemoveAt(0);
                }
            }
            if (deliveryNoRoute.Orders.Count > 0) deliveries.Add(deliveryNoRoute);
            return deliveries;
        }

        #region  Сбор данных sap

        /// <summary>
        /// Получение списка заказов из файла с выгрузкой из SAP 
        /// </summary>
        /// <param name="sapPath">Путь к файлу</param>
        /// <returns></returns>
        public List<Order> GetOrdersFromSap(string sapPath)
        {
            List<Order> orders = new List<Order>();
            Workbook sapBook;
            try
            {
                sapBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: sapPath);
            }
            catch
            {
                MessageBox.Show("Не удалось открыть книгу Excel");
                return null;
            }

            Worksheet sheet = sapBook.Sheets[1];
            int lastRow = sheet.Cells[sheet.Rows.Count, 5].End(XlDirection.xlUp).Row;
            int lastColumn = sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1;
            Range range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[lastRow, lastColumn]];
            ProcessBar pb = ProcessBar.Init("Сбор данных", range.Rows.Count, 1, "Формирование доставок");

            if (pb == null) return null;
            pb.Show();

            foreach (Range row in range.Rows)
            {
                if (pb.Cancel) break;
                pb.Action("Заказ " + (row.Row - range.Row + 1) + " из " + pb.Count);
                Order order = GetOrder(row);
                if (order != null)
                {
                    //orderClient.Id = orderClient.Id + " \\ " + order.Id;
                    //orderClient.TransportationUnit = orderClient.TransportationUnit +
                    //                                " \\ " + order.TransportationUnit;
                    //orderClient.PalletsCount += order.PalletsCount;
                    //orderClient.WeightNetto += order.WeightNetto;
                    //orderClient.WeightBrutto += order.WeightBrutto;
                    //orderClient.Cost += order.Cost;
                    orders.Add(order);
                }
            }
            pb.Close();
            sapBook.Close();

            return orders;
        }

        /// <summary>
        /// Получение данных заказа из строки
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Order GetOrder(Range row)
        {
            Order order = new Order();
            Debug.WriteLine("Загрузить заказ строка -" + row.Row);

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

        /// <summary>
        /// Получение дополнительной информации о заказах из файла All orders
        /// </summary>
        /// <param name="ordersPath">Путь к файлу All orders</param>
        /// <param name="ordersSap">Список заказов</param>
        /// <returns></returns>
        public List<Order> GetOrdersInfo(string ordersPath, List<Order> ordersSap)
        {
            if (ordersPath == "") return null;
            Workbook orderBook = null;
            try
            {
                orderBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: ordersPath);
            }
            catch
            {
                MessageBox.Show("Не удалось открыть книгу Excel");
                return null;
            }

            List<Order> ordersInfo = new List<Order>();
            foreach (Order order in ordersSap)
            {
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
                ordersInfo.Add(order);
            }

            orderBook.Close();
            return ordersInfo;
        }

        /// <summary>
        /// Получение дополнительной информации по заказу
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="delivery"></param>
        /// <returns></returns>
        private List<string> GetOrderInfo(Worksheet sheet, string delivery)
        {
            Range findRange = sheet.Columns[1];

            string search = new string('0', 10 - delivery.Length) + delivery;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);
            if (fcell == null) return null;

            string strCell = fcell.Text.Trim();
            if (!strCell.Contains("Доставка")) return null;

            int rowStart = 0;
            for (int i = fcell.Row; i > 1; --i)
            {
                strCell = findRange.Cells[i, 1].Text.Trim();
                if (strCell.Contains("ТТН:") || string.IsNullOrWhiteSpace(strCell))
                {
                    rowStart = i;
                    break;
                }
            }

            int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            int rowEnd = rowStart;
            List<string> info = new List<string>();
            do
            {
                fcell = findRange.Cells[rowEnd++, 1];
                string cellText = fcell.Text;
                cellText.Trim();
                cellText = cellText.Replace("\t", "");
                cellText = cellText.Replace(";;;", "");
                if (string.IsNullOrEmpty(cellText.Replace(";", ""))) break;
                info.Add(cellText);
            }
            while (rowEnd <= lastRow);
            return info;
        }

        #endregion Сбор данных sap



        /// <summary>
        /// Заполнить таблицу отгрузки
        /// </summary>
        /// <param name="totalTable"></param>
        /// <param name="deliveries"></param>
        private void PrintShipping(ListObject totalTable, List<Delivery> deliveries)
        {
            ShefflerWB shefflerBook = new ShefflerWB();
            ListRow row;
            Range CurrentDateRng = shefflerBook.GetCurrentShippingRange();

            if (CurrentDateRng != null)
            {
                Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
                for (int k = CurrentDateRng.Rows.Count; k > 0; k--)
                {
                    totalSheet.Rows[CurrentDateRng.Rows[k].Row].Delete();
                }
            }

            if (totalTable.ListRows.Count == 0)
            {
                totalTable.ListRows.Add();
                row = totalTable.ListRows[1];
            }
            else
            {
                row = totalTable.ListRows[totalTable.ListRows.Count - 1];
            }

            foreach (Delivery delivery in deliveries)
            {
                row.Range[1, totalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;

                foreach (Order order in delivery.Orders)
                {
                    string date = shefflerBook.DateDelivery;

                    row.Range[1, totalTable.ListColumns["Дата доставки"].Index].Value = date;
                    row.Range[1, totalTable.ListColumns["Перевозчик"].Index].Value = delivery.Truck?.ShippingCompany?.Name;
                    row.Range[1, totalTable.ListColumns["Тип ТС, тонн"].Index].Value = delivery.Truck?.Tonnage ?? 0;
                    row.Range[1, totalTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;

                    row.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value =
                            delivery.MapDelivery.FindIndex(x => x.IdCustomer == order.Customer.Id) + 1;

                    row.Range[1, totalTable.ListColumns["Номер накладной"].Index].Value = order.TransportationUnit;
                    row.Range[1, totalTable.ListColumns["Номер поставки"].Index].Value = order.Id;
                    row.Range[1, totalTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    row.Range[1, totalTable.ListColumns["Направление"].Index].Value = order.Route;
                    row.Range[1, totalTable.ListColumns["Номер грузополучателя"].Index].Value = order.Customer?.Id ?? "";
                    row.Range[1, totalTable.ListColumns["Брутто вес"].Index].Value = order.WeightBrutto;
                    row.Range[1, totalTable.ListColumns["Нетто вес"].Index].Value = order.WeightNetto;
                    row.Range[1, totalTable.ListColumns["Грузополучатель"].Index].Value = $"{order.Customer?.Name ?? ""}";
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
        public void СhangeDelivery()
        {
            ShefflerWB.ExcelOptimizateOn();
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];

            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];

            List<Order> orders = GetOrdersFromTable(ordersTable);
            List<Delivery> deliveries = EditDeliveres(orders);
            ClearListObj(carrierTable);
            PrintDelivery(deliveries, carrierTable);
            // EditPrintOrders()

            ShefflerWB.ExcelOptimizateOff();
            foreach (ListRow row in ordersTable.ListRows)
            {
                string strNum = row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : 0;
                if (deliveryNumber == 0) continue;
                string orderId = row.Range[1, ordersTable.ListColumns["Доставка"].Index].Text;
                orderId = new string('0', 10 - orderId.Length) + orderId;
                Delivery delivery = deliveries.Find(d => d.Number == deliveryNumber);
                if (delivery == null) continue;

                Order order = delivery.Orders.Find(r => r.Id == orderId);
                if (order != null)
                {
                    row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                    row.Range[1, ordersTable.ListColumns["ID Route"].Index].Value = order.DeliveryPoint.Id;
                    row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                }
            }
        }

        /// <summary>
        ///  Считать заказы с листа
        /// </summary>
        /// <param name="ordersTable"></param>
        /// <returns></returns>
        public List<Order> GetOrdersFromTable(ListObject ordersTable)
        {
            List<Order> orders = new List<Order>();
            //  List<int> deliveryNumbers = new List<int>();


            foreach (ListRow row in ordersTable.ListRows)
            {
                Order order = new Order();
                string strNum = row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : 0;
                if (deliveryNumber == 0) continue;
                order.DeliveryNumber = deliveryNumber;
                order.Id = row.Range[1, ordersTable.ListColumns["Доставка"].Index].Text;

                string city = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;

                //strNum = row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Text;
                // order.PointNumber = int.TryParse(strNum, out int pointnum) ? pointnum : 0;

                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                Customer customer = new Customer(customerId);
                DeliveryPoint point = new DeliveryPoint()
                {
                    City = city,
                    Customer = customerId
                };
                order.DeliveryPoint = point;

                order.Customer = customer;
                //string CityStr = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;
                //order.DeliveryPoint = new DeliveryPoint() { City = CityStr };

                string weight = row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
                orders.Add(order);
            }
            return orders;
        }

        /// <summary>
        ///    Пересчитать маршруты
        /// </summary>
        public void AcceptDelivery()
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
                    delivery.Orders = orders.FindAll(x => x.DeliveryNumber == deliveryNumber).ToList();
                    deliveries.Add(delivery);
                }
            }

            foreach (ListRow totalRow in totalTable.ListRows)
            {
                string transportationUnit = totalRow.Range[1,
                                 totalTable.ListColumns["Номер накладной"].Index].Text;

                foreach (Delivery delivery in deliveries)
                {
                    //transportationUnit = new string('0', 18 - transportationUnit.Length) + transportationUnit;

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
        /// перенести с деливери на лист Отгрузка
        /// </summary>
        internal void CopyDelivery()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];
            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orders = GetOrdersFromTable(ordersTable);

            string dateDelivery = deliverySheet.Range["DateDelivery"].Text;

            foreach (ListRow delveryRow in carrierTable.ListRows)
            {
                string str = delveryRow.Range[1, carrierTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(str, out int num) ? num : 0;

                if (deliveryNumber > 0)
                {
                    Delivery delivery = new Delivery();
                    delivery.Orders = orders.FindAll(x => x.DeliveryNumber == deliveryNumber).ToList();
                    deliveries.Add(delivery);
                }
            }

            foreach (ListRow totalRow in totalTable.ListRows)
            {
                string idOrder = totalRow.Range[1,
                                 totalTable.ListColumns["Номер поставки"].Index].Text;

                foreach (Delivery delivery in deliveries)
                {
                    //transportationUnit = new string('0', 18 - transportationUnit.Length) + transportationUnit;

                    Order orderf = delivery.Orders.Find(x => x.Id.Contains(idOrder));
                    if (orderf != null)
                    {
                        totalRow.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value = orderf.PointNumber;
                        totalRow.Range[1, totalTable.ListColumns["Дата доставки"].Index].Value = dateDelivery;

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

            /// Список новых номеров доставок
            List<int> deliveryNumbers = (from o in orders
                                         select o.DeliveryNumber).Distinct().ToList();
            // По каждой доставке создать список заказов 
            for (int i = 0; i < deliveryNumbers.Count; i++)
            {
                int deliveryNumber = deliveryNumbers[i];
                if (deliveryNumber > 0)
                {

                    List<Order> orderList = orders.FindAll(
                                o => o.DeliveryNumber == deliveryNumber).ToList().OrderBy(
                                                                x => x.PointNumber).ToList();
                    if (orderList.Count > 0)
                    {
                        Delivery delivery = EditDelivery(orderList);

                        delivery.Number = deliveryNumber;
                        //delivery.Number = deliveries.Count + 1;
                        deliveries.Add(delivery);
                    }
                }
            }
            // найти подходящий маршрут
            //
            #region Добавление нового маршрута
            #endregion
            return deliveries;
        }

        /// <summary>
        /// Изменить доставку
        /// </summary>
        /// <param name="ordersCurrentDelivery"></param>
        /// <returns></returns>
        private Delivery EditDelivery(List<Order> ordersCurrentDelivery)
        {
            ShefflerWB functionsBook = new ShefflerWB();
            Delivery delivery = new Delivery();
            int idRoute = FindRoute(ordersCurrentDelivery, functionsBook);
            if (idRoute == 0)
            {
                // Добавить маршрут 
                idRoute = functionsBook.CreateRoute(ordersCurrentDelivery);
                functionsBook = new ShefflerWB();
            }
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;

            foreach (Order order in ordersCurrentDelivery)
            {
                order.DeliveryPoint = pointMap.Find(p => p.Id == idRoute &&
                                                 p.IdCustomer == order.Customer.Id);
            }
            ordersCurrentDelivery = ordersCurrentDelivery.OrderBy(b => b.DeliveryPoint.PriorityPoint).ToList();
            int number = 1;
            for (int i = 0; i < ordersCurrentDelivery.Count; i++)
            {
                if (i > 0 && ordersCurrentDelivery[i].DeliveryPoint.IdCustomer != ordersCurrentDelivery[i - 1].DeliveryPoint.IdCustomer)
                {
                    ++number;
                }
                ordersCurrentDelivery[i].PointNumber = number;
            }
            delivery.Orders = ordersCurrentDelivery;

            return delivery;
        }



        /// <summary>
        /// Поиск маршрута где есть все клиенты из списка заказов
        /// </summary>
        /// <param name="orders"></param>
        /// <param name="functionsBook"></param>
        /// <returns></returns>
        private int FindRoute(List<Order> orders, ShefflerWB functionsBook)
        {
            //Таблица routes
            List<DeliveryPoint> pointMap = functionsBook.RoutesTable;
            //список id маршрутов
            List<int> uRoutes = (from p in pointMap
                                 select p.Id).Distinct().ToList();

            for (int i = 0; i < uRoutes.Count; i++)
            {
                int idRoute = uRoutes[i];
                bool hasRoute = true;
                foreach (Order order in orders)
                {
                    List<DeliveryPoint> routes = pointMap.FindAll(
                                 x => x.Id == idRoute &&
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

        /// <summary>
        /// Загрузка заказов, формирование доставок вывод на лист
        /// </summary>
        public void GetOrdersFromFiles()
        {
            //string path = OpenFileDialog();
            string path = "";
            //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //dialog.InitialDirectory = Settings.Default.SapUnloadPath; //Directory.GetCurrentDirectory() ;
            //dialog.IsFolderPicker = true;
            //if (dialog.ShowDialog() != CommonFileDialogResult.Ok) { return; }
            //path = dialog.FileName;
            //string[] files = Directory.GetFiles(path);


            string file = SapFiles.SelectFile();
            if (!File.Exists(file)) return;

            List<Order> orders = new List<Order>();

            Order order = GetFromFile(file);
            if (order != null) orders.Add(order);

            List<Delivery> deliveries = CompleteAuto(orders);
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];
            PrintDelivery(deliveries, carrierTable);
            PrintOrders(deliveries, ordersTable);
            PrintShipping(totalTable, deliveries);

            return;
        }

        /// <summary>
        /// Вывод заказа в таблицу
        /// </summary>
        /// <param name="table"></param>
        /// <param name="order"></param>
        public void PrintRow(ListObject table, Order order)
        {
            ListRow rowDelivery;
            if (table.ListRows.Count == 0)
            {
                AddListRow(table);
                rowDelivery = table.ListRows[1];
            }
            else
            {
                AddListRow(table);
                rowDelivery = table.ListRows[table.ListRows.Count - 1];
            }

            // rowDelivery.Range[1, table.ListColumns["№ Доставки"].Index].Value = delivery.Number;
        }

        /// <summary>
        /// Получить инфо из выгруза  
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public Order GetFromFile(string file)
        {
            Order order = new Order();
            Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: file);
            Worksheet sh = wb.Sheets[1];
            Range rng = sh.UsedRange;
            string str = ShefflerWB.FindValue("Заявка на перевозку", rng, 0, 0);
            if (str == "") return null;

            str = ShefflerWB.FindValue("Номер грузополучателя", rng, 0, 1);
            // str = str.Remove(0, str.IndexOf("ИНН") + 3).Trim();
            Regex regexId = new Regex(@"\d+");
            string idcustomer = regexId.Match(str).Value;
            order.Customer.Id = idcustomer;

            str = ShefflerWB.FindValue("Грузополучатель", rng, 0, 1);
            order.Customer.Name = str.Trim();

            str = ShefflerWB.FindValue("Номер накладной", rng, 0, 1);
            order.Id = str.Replace(", ", " / ");

            str = ShefflerWB.FindValue("Стоимость", rng, 0, 0);
            Regex regexCost = new Regex(@"(\d+\s?)+(\,\d+)?");
            str = regexCost.Match(str).Value;
            order.Cost = double.TryParse(str, out double ct) ? ct : 0;

            str = ShefflerWB.FindValue("брутто", rng, 0, 0);
            str = regexCost.Match(str).Value;
            double weight = double.TryParse(str, out double wt) ? wt : 0;
            order.WeightNetto = weight;

            str = ShefflerWB.FindValue("грузовых", rng, 0, 0);
            str = regexId.Match(str).Value;
            int countPallets = int.TryParse(str, out int count) ? count : 0;
            order.PalletsCount = countPallets;
            wb.Close();
            return order;
        }

        private string GetVal(ListObject table, int row, string header)
        {
            int col = table.ListColumns[header].Index;
            string value = table.ListRows[row].Range[1, col].Text;
            return value;
        }
        private string GetVal(ListObject table, Range rng, int row, string header)
        {
            int col = table.ListColumns[header].Index;
            string value = table.Range[row, col].Text;
            return value;
        }

        /// <summary>
        /// Собрать доставки из актуального диапазона таблицы Отгрузка
        /// </summary>
        /// <returns></returns>
        public List<Delivery> GetDeliveriesFromTotalSheet()
        {

            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carTable = deliverySheet.ListObjects["TableCarrier"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];
            List<Delivery> deliveries = new List<Delivery>();
            Range total = new ShefflerWB().GetCurrentShippingRange();
            ShefflerWB functionsBook = new ShefflerWB();
            List<DeliveryPoint> points = functionsBook.RoutesTable;

            if (total == null) return deliveries;

            for (int i = 0; i < total.Rows.Count; i++)
            {
                string numDelivery = total.Cells[i, totalTable.ListColumns["№ Доставки"].Index].Text;
                int numD = int.TryParse(numDelivery, out int numDel) ? numDel : 0;
                if (numD == 0) continue;
                Delivery delivery = deliveries.Find(x => x.Number == numD);
                if (delivery == null)
                {
                    delivery = new Delivery();
                    delivery.Number = numD;
                    delivery.Truck = new Truck();
                    delivery.Truck.ShippingCompany = new Provider();
                    string providerName = total.Cells[i, totalTable.ListColumns["Перевозчик"].Index].Text;
                    delivery.Truck.ShippingCompany.Name = providerName;
                    string tonn = total.Cells[i, totalTable.ListColumns["Тип ТС, тонн"].Index].Text;
                    delivery.Truck.Tonnage = double.TryParse(tonn, out double ton) ? ton : 0;
                    string costDelivery = total.Cells[i, totalTable.ListColumns["Стоимость доставки"].Index].Text;
                    delivery.Cost = double.TryParse(costDelivery, out double cd) ? cd : 0;
                    delivery.Carrier.Id = ShefflerWB.GetProviderId(providerName);
                    total.Cells[i, totalTable.ListColumns["ID перевозчика"].Index].Value = delivery.Carrier.Id;
                    deliveries.Add(delivery);
                }

                string ID = total.Cells[i, totalTable.ListColumns["Номер поставки"].Index].Text;

                if (ID != "")
                {
                    Order order = new Order();
                    order.Id = ID;
                    string cost = total.Cells[i, totalTable.ListColumns["Стоимость поставки"].Index].text;
                    order.Cost = double.TryParse(cost, out double ct) ? ct : 0;

                    string customerId = total.Cells[i, totalTable.ListColumns["Номер грузополучателя"].Index].Text;
                    order.Customer = new Customer(customerId);
                    order.Customer.Name = total.Cells[i, totalTable.ListColumns["Грузополучатель"].Index].text;

                    order.TransportationUnit = total.Cells[i, totalTable.ListColumns["Номер накладной"].Index].Text;

                    DeliveryPoint point = new DeliveryPoint();
                    point.City = total.Cells[i, totalTable.ListColumns["Город"].Index].Text; ;
                    order.DeliveryPoint = point;

                    string palletCount = total.Cells[i, totalTable.ListColumns["Кол-во паллет"].Index].Text;
                    order.PalletsCount = int.TryParse(palletCount, out int countPalets) ? countPalets : 0;

                    string nom = total.Cells[i, totalTable.ListColumns["Порядок выгрузки"].Index].Text;
                    order.PointNumber = int.TryParse(nom, out int nd) ? nd : 0;

                    string weightBr = total.Cells[i, totalTable.ListColumns["Брутто вес"].Index].Text;
                    order.WeightBrutto = double.TryParse(weightBr, out double wb) ? wb : 0;

                    string weightNt = total.Cells[i, totalTable.ListColumns["Нетто вес"].Index].Text;
                    order.WeightNetto = double.TryParse(weightNt, out double wn) ? wn : 0;

                    order.Route = total.Cells[i, totalTable.ListColumns["Направление"].Index].Text;

                    delivery.Orders.Add(order);
                }
            }
            return deliveries;
        }


        /// <summary>
        /// Подготовка сообщений перевозчикам
        /// </summary>
        public void CreateMasseges()
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carTable = deliverySheet.ListObjects["TableCarrier"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            List<Delivery> deliveries = GetDeliveriesFromTotalSheet();
            if (deliveries?.Count == 0) return;

            //Уникальны провайдеры в списке доставок
            string[] shippingComp = (from d in deliveries
                                     select d.Truck.ShippingCompany.Name).Distinct().ToArray();
            ClearFolder();
            //ДДля каждого провайдера
            for (int i = 0; i < shippingComp.Length; i++)
            {
                string сompanyShipping = shippingComp[i];

                List<Delivery> deliverShipping = deliveries.FindAll(x =>
                                   x.Truck.ShippingCompany.Name == сompanyShipping).ToList();
                string date = deliverySheet.Range["DateDelivery"].Text;
                string subject = messageSheet.Cells[8, 2].Text;
                subject = subject.Replace("[date]", date).Replace("[provider]", shippingComp[i]);
                //Прикрепленный файл
                string attachment = GenerateFile(deliverShipping, subject);
                //   Range findCell = tableEmail.ListColumns["Компания"]?.Range.Find(What: Company);                       

                /// Найти Email
                Email messenger = new Email();
                messenger.CreateMessage(сompany: сompanyShipping,
                                        date: date,
                                        attachment: attachment,
                                        subject: subject);
            }
        }

        /// <summary>
        /// Создать файл отгрузки для провайдера
        /// </summary>
        /// <param name="delivery"></param>
        /// <returns></returns>
        private string GenerateFile(List<Delivery> deliveries, string name)
        {
            if (deliveries.Count == 0) return "";

            string folder = GenerateFolder();
            string filename = $"{folder}\\{name}.xlsx";

            Workbook workbook = Globals.ThisWorkbook.Application.Workbooks.Add();

            Worksheet sh = workbook.Sheets[1];
            string[] headers = {
                "ID перевозчика",
                "Перевозчик",
                "Тип ТС, тонн" ,
                "Водитель (ФИО)",
                "Номер, марка",
                "Телефон водителя",
                "Город"            ,
                "Направление"   ,
                "Порядок выгрузки",
                "Номер грузополучателя",
                "Номер накладной",
                "Номер поставки",
                "Грузополучатель",
                "Брутто вес",
                "Нетто вес",
                "Кол-во паллет" ,
                "Стоимость поставки" ,
                "Стоимость доставки"
                    };
            for (int i = 1; i <= headers.Length; i++)
            {
                sh.Cells[1, i].Value = headers[i - 1];
            }
            int row = 2;
            foreach (Delivery delivery in deliveries)
            {
                string providerName = delivery.Truck.ShippingCompany.Name;
                if (string.IsNullOrWhiteSpace(providerName)) continue;
                sh.Cells[row, 1].Value = delivery.Carrier.Id;
                sh.Cells[row, 18].Value = delivery.Cost;

                foreach (Order order in delivery.Orders)
                {
                    sh.Cells[row, 2].Value = providerName;
                    sh.Cells[row, 3].Value = delivery.Truck.Tonnage;

                    sh.Cells[row, 7].Value = order.DeliveryPoint.City;
                    sh.Cells[row, 8].Value = order.Route;

                    sh.Cells[row, 9].Value = order.PointNumber;
                    sh.Cells[row, 10].Value = order.Customer.Id;
                    sh.Cells[row, 11].Value = order.TransportationUnit;
                    sh.Cells[row, 12].Value = order.Id;
                    sh.Cells[row, 13].Value = order.Customer.Name ?? "";
                    sh.Cells[row, 14].Value = order.WeightBrutto;
                    sh.Cells[row, 15].Value = order.WeightNetto;
                    sh.Cells[row, 16].Value = order.PalletsCount;
                    sh.Cells[row, 17].Value = order.Cost;
                    row++;
                }
                Range spaceRow = sh.Range[sh.Cells[row, 1], sh.Cells[row, 18]];
                spaceRow.Interior.Color = System.Drawing.Color.FromArgb(81, 135, 245);
                sh.Rows[spaceRow.Row].RowHeight = 3;
                row++;
            }
            sh.Rows[row].Delete();
            Range rng = sh.Range[sh.Cells[1, 1], sh.Cells[row - 1, headers.Length]];
            ListObject list =
                sh.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rng,
                XlListObjectHasHeaders: XlYesNoGuess.xlYes);

            workbook.SaveAs(filename);
            workbook.Close();
            return filename;
        }

        /// <summary>
        /// Создать папку для заявок прикрепления к письмам
        /// </summary>
        /// <returns></returns>
        private string GenerateFolder()
        {
            string folder = "";
            folder = Globals.ThisWorkbook.Path + "\\ShippingOrders";

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }


        /// <summary>
        /// Очистить папку вложений
        /// </summary>
        private void ClearFolder()
        {
            string folder = "";
            folder = Globals.ThisWorkbook.Path + "\\ShippingOrders";

            if (Directory.Exists(folder))
            {
                string[] fls = Directory.GetFiles(folder);
                foreach (string f in fls)
                {

                    try
                    {
                        File.Delete(f);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(folder);
                    }
                }
            }

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

    }
    #endregion Вспомогательные

}