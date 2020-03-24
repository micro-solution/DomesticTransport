using DomesticTransport.Forms;
using DomesticTransport.Model;
using DomesticTransport.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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

            ExcelOptimizateOn();
            List<Order> orders = GetOrdersFromSap(sapPath);

            if (ordersPath != "" && File.Exists(ordersPath))
            {
                orders = GetOrdersInfo(ordersPath, orders);  // Поиск свойств в файле All orders
            }

            List<Delivery> deliveries = CompleteAuto2(orders);

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

            ExcelOptimizateOff();
        }

        /// <summary>
        /// Загрузка All Orders
        /// </summary>
        public void LoadAllOrders()
        {
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
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
            foreach (Range rng in deliveryTable.ListColumns["№ Доставки"].DataBodyRange)
            {
                if (int.TryParse(rng.Text, out int valueCell))
                {
                    if (number < valueCell) number = valueCell;
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
                ShefflerWorkBook workBook = new ShefflerWorkBook();

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
                List<Delivery> deliveries = CompleteAuto2(orders);
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

            DialogResult msg = MessageBox.Show("Удалить авто с заказами", "Удалить", MessageBoxButtons.YesNo);
            if (DialogResult.No == msg) return;
            ShefflerWorkBook workBook = new ShefflerWorkBook();

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
                rowDelivery.Range[1, DeliveryTable.ListColumns["Компания"].Index].Value =
                                                                delivery.Truck?.ShippingCompany?.Name ?? "";
                if (delivery?.MapDelivery.Count > 0)
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["ID Route"].Index].Value =
                                                                        delivery?.MapDelivery[0].Id;
                }
                rowDelivery.Range[1, DeliveryTable.ListColumns["Тоннаж"].Index].Value = delivery.Truck?.Tonnage ?? 0;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Вес доставки"].Index].FormulaR1C1 =
                                                "=SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес нетто])";

                rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;

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
                    rowOrder.Range[1, OrderTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                    rowOrder.Range[1, OrderTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                    rowOrder.Range[1, OrderTable.ListColumns["Доставка"].Index].Value = order.Id;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Получателя"].Index].Value = order.Customer?.Id ?? "";
                    rowOrder.Range[1, OrderTable.ListColumns["Получатель"].Index].Value = order.Customer.Name;
                    rowOrder.Range[1, OrderTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    rowOrder.Range[1, OrderTable.ListColumns["ID Route"].Index].Value = order.DeliveryPoint.Id;
                    rowOrder.Range[1, OrderTable.ListColumns["Вес нетто"].Index].Value = order.WeightNetto;
                    rowOrder.Range[1, OrderTable.ListColumns["Маршрут"].Index].Value = order.Route;

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
                emptyDelivery.HasRoute = false;
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
                        Debug.WriteLine($"pointDelivery={point.Id} pointMap={pointMap[pointNumber].Id}");
                        //  новый маршрут в таблице
                        if (delivery.MapDelivery.First().Id != pointMap[pointNumber].Id)
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




        /// <summary>
        /// Распределить заказы по автомобилям
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        public List<Delivery> CompleteAuto2(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            orders = orders.OrderBy(x => x.WeightNetto).ToList();
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
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
                        if (orders[iOrder].Customer.Id != point.IdCustomer) continue;
                        findDelivery = true;
                        orders[iOrder].DeliveryPoint = point;
                        // Пытаемся добавить к имеющимся машинам
                        Delivery delivery = null;
                        foreach (Delivery iDelivery in deliveries)
                        {
                            if (iDelivery.Orders[0].DeliveryPoint.Id != point.Id) continue;
                            if (iDelivery.CheckDeliveryWeght(orders[iOrder]))
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
                        orders[iOrder].PointNumber = delivery.Orders.Count + 1;
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
                Order order = GetOrder(row);
                if (order != null) orders.Add(order);
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
            ShefflerWorkBook shefflerBook = new ShefflerWorkBook();
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
                    row.Range[1, totalTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;

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
        public void СhangeDelivery()
        {
            ExcelOptimizateOn();
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];

            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];

            List<Order> orders = GetOrdersFromTable(ordersTable);
            List<Delivery> deliveries = EditDeliveres(orders);
            ClearListObj(carrierTable);
            PrintDelivery(deliveries, carrierTable);
            // EditPrintOrders()

            ExcelOptimizateOff();
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
        /// 
        /// </summary>
        /// <param name="ordersTable"></param>
        /// <returns></returns>
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
                order.DeliveryNumber = deliveryNumber;
                order.Id = row.Range[1, ordersTable.ListColumns["Доставка"].Index].Text;

                string city = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;
                DeliveryPoint point = new DeliveryPoint() { City = city };
                order.DeliveryPoint = point;

                //strNum = row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Text;
                //order.PointNumber = int.TryParse(strNum, out int pointnum) ? pointnum : 0;

                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                Customer customer = new Customer(customerId);
                order.Customer = customer;
                //string CityStr = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;
                //order.DeliveryPoint = new DeliveryPoint() { City = CityStr };

                string weight = row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;
                orders.Add(order);
            }
            return orders;
        }


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
            ShefflerWorkBook functionsBook = new ShefflerWorkBook();
            Delivery delivery = new Delivery();
            int idRoute = FindRoute(ordersCurrentDelivery, functionsBook);
            if (idRoute == 0)
            {
                // Добавить маршрут 
                idRoute = functionsBook.CreateRoute(ordersCurrentDelivery);
                functionsBook = new ShefflerWorkBook();
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
        private int FindRoute(List<Order> orders, ShefflerWorkBook functionsBook)
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
        public void GetOrdersFromFiles()
        {
            //string path = OpenFileDialog();
            string path = "";
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = Settings.Default.SapUnloadPath; //Directory.GetCurrentDirectory() ;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() != CommonFileDialogResult.Ok)  { return; }             
            path = dialog.FileName;
            string[] files= Directory.GetFiles(path);
            List<Order> orders = new List<Order>();
            
            foreach(string file in files )
            {
                Order order = GetFromFile(file);
                if (order != null) orders.Add(order);
            
            }
            List<Delivery> deliveries = CompleteAuto2(orders);
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject carrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];
            Worksheet totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = totalSheet.ListObjects["TableTotal"];
            PrintDelivery(deliveries, carrierTable);
            PrintOrders( deliveries, ordersTable);
            PrintShipping(totalTable, deliveries);

            return ;
        }
        public Order GetFromFile(string file)
        {
            Order order = new Order();
            Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: file);
            Worksheet sh = wb.Sheets[1];
            Range rng = sh.UsedRange;
            string str = FindValue("Заявка на перевозку", rng, 0, 0);
            if (str == "") return null;

            str = FindValue("Номер грузополучателя", rng, 0, 1);
           // str = str.Remove(0, str.IndexOf("ИНН") + 3).Trim();
            Regex regexId = new Regex(@"\d+");
            string idcustomer = regexId.Match(str).Value;
            order.Customer.Id = idcustomer;

            str = FindValue("Грузополучатель", rng, 0, 1);
            order.Customer.Name = str.Trim();

            str = FindValue("Номер накладной", rng, 0, 1);
            order.Id = str.Replace(", ", " / ");

            str = FindValue("Стоимость", rng, 0, 0);
            Regex regexCost = new Regex(@"(\d+\s?)+(\,\d+)?");
            str = regexCost.Match(str).Value;
            order.Cost = double.TryParse(str, out double ct) ? ct : 0;

            str = FindValue("брутто", rng, 0, 0);
            str = regexCost.Match(str).Value;
            double weight = double.TryParse(str, out double wt) ? wt : 0;
            order.WeightNetto = weight;

            str = FindValue("грузовых", rng, 0, 0);
            str = regexId.Match(str).Value;
            int countPallets = int.TryParse(str, out int count) ? count : 0;
            order.PalletsCount = countPallets;
            wb.Close();
            return order;
        }

        #region Вспомогательные
        /// <summary>
        /// Ищет в диапазоне текст возвращает значение ячейки по указанному смещению
        /// </summary>
        /// <param name="header"></param>
        /// <param name="rng"></param>
        /// <param name="offsetRow"></param>
        /// <param name="offsetCol"></param>
        /// <returns></returns>
        public string FindValue(string header, Range rng, int offsetRow = 0, int offsetCol = 0)
        {
            Range findCell = rng.Find(What: header, LookIn: XlFindLookIn.xlValues);
            if (findCell == null) return "";
            findCell = findCell.Offset[offsetRow, offsetCol];
            string valueCell = findCell.Text;
            return valueCell;
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

        public string OpenFileDialog()
        {
            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = Properties.Settings.Default.SapUnloadPath,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })

                return (ofd.ShowDialog() == DialogResult.OK) ? ofd.FileName : "";
        }
    }
    #endregion Вспомогательные

}