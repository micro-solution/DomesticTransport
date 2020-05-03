using DomesticTransport.Forms;
using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
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
        /// Загрузка from SAP
        /// </summary>
        public void ExportFromSAP()
        {
            SapFiles sapFiles = new SapFiles();
            sapFiles.ShowDialog();
            if (sapFiles.DialogResult != DialogResult.OK) return;

            ShefflerWB.DateCell.Value = sapFiles.calendarControl.SelectionStart;

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

            List<Order> orders = GetOrdersFromSap(sapPath);

            if (ordersPath != "" && File.Exists(ordersPath))
            {
                orders = GetOrdersInfo(ordersPath, orders);  // Поиск свойств в файле All orders
            }

            List<Delivery> deliveries = CompleteAuto(orders);

            ClearListObj(ShefflerWB.DeliveryTable);
            ClearListObj(ShefflerWB.OrdersTable);
            ClearListObj(ShefflerWB.TotalTable);

            if (deliveries != null && deliveries.Count > 0)
            {
                PrintDelivery(deliveries);
                PrintOrders(deliveries);
                PrintTotal(deliveries);
            }
            ShefflerWB.DeliverySheet.Columns.AutoFit();
        }

        /// <summary>
        /// Загрузка All Orders 
        /// </summary>
        public void LoadAllOrders()
        {
            ShefflerWB functionsBook = new ShefflerWB();

            Range range = ShefflerWB.TotalTable.DataBodyRange;
            if (range == null || ShefflerWB.TotalTable == null) return;
            string file = SapFiles.SelectFile();
            if (!File.Exists(file)) return;
         

            List<Order> orders = GetOrdersFromTotalTable(range);
            orders = GetOrdersInfo(file, orders);
            if (orders == null || orders.Count == 0) return;
            int columnId = ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index;
            foreach (Range row in range.Rows)
            {
                string idOrder = row.Cells[1, columnId].Text;
                if (string.IsNullOrWhiteSpace(idOrder)) continue;
                idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                Order order = orders.Find(o => o.Id == idOrder);
                if (order == null) continue;

                row.Cells[1, ShefflerWB.TotalTable.ListColumns["Номер накладной"].Index].Value = order.TransportationUnit;
                row.Cells[1, ShefflerWB.TotalTable.ListColumns["Брутто вес"].Index].Value = order.WeightBrutto;
                row.Cells[1, ShefflerWB.TotalTable.ListColumns["Стоимость поставки"].Index].Value = order.Cost;
                row.Cells[1, ShefflerWB.TotalTable.ListColumns["Кол-во паллет"].Index].Value = order.PalletsCount;
            }
            UpdateOrderFromTotal(); 
        }

        /// <summary>
        /// Загрузка заказов, формирование доставок вывод на лист
        /// </summary>
        public void ExportFromCS()
        {
            string file = SapFiles.SelectFile();
            if (!File.Exists(file)) return;
            Order order = GetFromFile(file);
            if (order == null) return;

            CheckAndAddNewRoute(order);
            Range range = ShefflerWB.TotalTable.DataBodyRange;
            List<Order> orders = GetOrdersFromTotalTable(range);
            int countDouble = 0;
            foreach (Order item in orders)
            {              
                if (item.Id.Contains(order.Id))
                {
                    countDouble++;                 
                }
            }
            if (countDouble > 0) order.Id += $"_{countDouble}";
            orders.Add(order);
            List<Delivery> deliveries = CompleteAuto(orders);

            ClearListObj(ShefflerWB.DeliveryTable);
            ClearListObj(ShefflerWB.OrdersTable);
            ClearListObj(ShefflerWB.TotalTable);

            if (deliveries != null && deliveries.Count > 0)
            {
                PrintDelivery(deliveries);
                PrintOrders(deliveries);
                PrintTotal(deliveries);
            }
            ShefflerWB.DeliverySheet.Columns.AutoFit();

            return;
        }

        /// <summary>
        ///кнопка  Добавить строку авто
        /// </summary>
        public void AddAuto()
        {
            int idRoute = 0;
            int number = 0;
            if (Globals.ThisWorkbook.Application.ActiveSheet.Name != "Delivery") return;
            if (ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].DataBodyRange != null)
            {
                foreach (Range rng in ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].DataBodyRange)
                {
                    if (int.TryParse(rng.Text, out int valueCell))
                    {
                        if (number < valueCell) number = valueCell;
                    }
                }
            }
            number++;
            Range selection;
            Range orfderRng;
            // Выделенный диапазон

            selection = Globals.ThisWorkbook.Application.Selection;
            orfderRng = ((Microsoft.Office.Interop.Excel.Application)(Globals.ThisWorkbook.Application)).Intersect(selection, ShefflerWB.OrdersTable.DataBodyRange);

            Delivery delivery = null;
            int colNumberDelivery = orfderRng.Column;
            int colWtNetto = colNumberDelivery + ShefflerWB.OrdersTable.ListColumns["Вес нетто"].Index - 1;
            int colWtBrutto = colNumberDelivery + ShefflerWB.OrdersTable.ListColumns["Вес брутто"].Index - 1;
            int colID = colNumberDelivery + ShefflerWB.OrdersTable.ListColumns["Поставка"].Index - 1;
            int colCustomerID = colNumberDelivery + ShefflerWB.OrdersTable.ListColumns["ID Получателя"].Index - 1;
            if (orfderRng != null)
            {
                ShefflerWB workBook = new ShefflerWB();

                string orderId = "";
                List<Order> orders = new List<Order>();
                int dpPriority = 1;
                foreach (Range orderLine in orfderRng.Rows)
                {


                    ShefflerWB.DeliverySheet.Cells[orderLine.Row, colNumberDelivery].Value = number;
                    orderId = ShefflerWB.DeliverySheet.Cells[orderLine.Row, colID].Text; //  "Номер поставки"
                    string idCustomer = ShefflerWB.DeliverySheet.Cells[orderLine.Row, colCustomerID].Text;

                    double weightNetto = double.TryParse(ShefflerWB.DeliverySheet.Cells[orderLine.Row, colWtNetto].Text, out double wgt) ? wgt : 0;
                    double weightBrutto = double.TryParse(ShefflerWB.DeliverySheet.Cells[orderLine.Row, colWtBrutto].Text, out double wgtb) ? wgtb : 0;

                    Customer customer = new Customer(idCustomer);
                    DeliveryPoint dp = new DeliveryPoint();
                    dp.IdCustomer = orderId;
                    dp.PriorityPoint = dpPriority;
                    Order ord = new Order();
                    ord.Id = orderId;
                    ord.PointNumber = dpPriority;
                    ord.DeliveryPoint = dp;
                    ord.WeightNetto = weightNetto;
                    ord.WeightBrutto = weightBrutto;
                    ord.Customer = customer;
                    orders.Add(ord);
                    dpPriority++;
                }

                List<Delivery> deliveries = CompleteAuto(orders);    //  EditDelivery(orders);
                                                                     // List<Delivery> deliveries = EditDeliveres(orders); //new List<Delivery>
              //  Range totalRng = workBook.GetCurrentTotalRange();
                if (deliveries != null && deliveries.Count > 0 )
                {
                    delivery = deliveries[0];
                    idRoute = delivery.MapDelivery[0].Id;

                    foreach (Range row in ShefflerWB.TotalTable.DataBodyRange.Rows)
                    {
                        string idOrderTotal = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index].Text;
                        idOrderTotal = idOrderTotal.Length < 10 ? new string('0', 10 - idOrderTotal.Length) + idOrderTotal : idOrderTotal;
                        Order findOrder = orders.Find(x => x.Id == idOrderTotal);
                        if (findOrder != null)
                        {
                            row.Cells[0, ShefflerWB.TotalTable.ListColumns["№ Доставки"].Index].Value = number.ToString();
                            row.Cells[0, ShefflerWB.TotalTable.ListColumns["Порядок выгрузки"].Index].Value = findOrder.PointNumber;
                            row.Cells[0, ShefflerWB.TotalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                        }
                    }
                    foreach (Range orderLine in orfderRng.Rows)
                    {
                        ShefflerWB.DeliverySheet.Cells[orderLine.Row, 4].Value = delivery.MapDelivery[0].Id;   //ID Route
                        string idOrder = ShefflerWB.DeliverySheet.Cells[orderLine.Row, colID].Text;
                        idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                        Order orderFnd = delivery.Orders.Find(x => x.Id==idOrder);
                        if (orderFnd != null)
                        {
                            ShefflerWB.DeliverySheet.Cells[orderLine.Row, 5].Value = orderFnd.PointNumber;
                        }
                    }
                }

            }
            ListRow rowDelivery;
            if (ShefflerWB.DeliveryTable.ListRows.Count == 0)
            {
                AddListRow(ShefflerWB.DeliveryTable);
                rowDelivery = ShefflerWB.DeliveryTable.ListRows[1];
            }
            else
            {
                AddListRow(ShefflerWB.DeliveryTable);
                rowDelivery = ShefflerWB.DeliveryTable.ListRows[ShefflerWB.DeliveryTable.ListRows.Count - 1];
            }
            rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Index].Value = number;
            if (delivery != null)
            {
                rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["ID Route"].Index].Value = delivery.MapDelivery[0].Id;
                rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["Компания"].Index].Value = delivery.Truck?.ProviderCompany?.Name ?? "";
                rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["Тип ТС, тонн"].Index].Value = delivery.Truck.Tonnage;
                rowDelivery.Range[1, ShefflerWB.DeliveryTable.ListColumns["Направление"].Index].Value =
                                                                                 delivery.MapDelivery[0].RouteName;
            }
            СhangeDelivery();
        }

        public static void SetDate()
        {
         foreach(ListRow row in ShefflerWB.TotalTable.ListRows)
            {
                //string dateTable = row.Range[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Text;
                string date = ShefflerWB.DateDelivery;
                string  orderId = row.Range[1, ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index].Text;
                if (!string.IsNullOrWhiteSpace(orderId) && !string.IsNullOrWhiteSpace(date))
                {
                    row.Range[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Value = date;
                }
            }   
        }

        /// <summary>
        /// Пересобрать авто используя основные маршруты
        /// </summary>
        public void UpdateAutoMain()
        {
            Range range = ShefflerWB.TotalTable.DataBodyRange;
            List<Order> orders = GetOrdersFromTotalTable(range);
            List<Delivery> deliveries = CompleteAuto(orders);

            ClearListObj(ShefflerWB.DeliveryTable);
            ClearListObj(ShefflerWB.OrdersTable);
            ClearListObj(ShefflerWB.TotalTable);

            if (deliveries != null && deliveries.Count > 0)
            {
                PrintDelivery(deliveries);
                PrintOrders(deliveries);
                PrintTotal(deliveries);
            }
            ShefflerWB.DeliverySheet.Columns.AutoFit();

            return;
        }

        /// <summary>
        /// Пересобрать авто пытаясь доукомплектовать авто по второстепенным маршрутам
        /// </summary>
        public void UpdateAutoSecond()
        {
            Range range = ShefflerWB.TotalTable.DataBodyRange;
            List<Order> orders = GetOrdersFromTotalTable(range);
            List<Delivery> deliveries = CompleteAuto(orders);
            deliveries = CompileAutoSecond(deliveries);

            ClearListObj(ShefflerWB.DeliveryTable);
            ClearListObj(ShefflerWB.OrdersTable);
            ClearListObj(ShefflerWB.TotalTable);

            if (deliveries != null && deliveries.Count > 0)
            {
                PrintDelivery(deliveries);
                PrintOrders(deliveries);
                PrintTotal(deliveries);
            }
            ShefflerWB.DeliverySheet.Columns.AutoFit();

            return;
        }

        /// <summary>
        ///кнопка Добавить авто
        /// </summary>
        public void DeleteAuto()
        {
            Worksheet sh = (Worksheet)Globals.ThisWorkbook.ActiveSheet;
            if (ShefflerWB.DeliveryTable == null ||
                ShefflerWB.OrdersTable == null ||
                sh.Name != "Delivery") return;
            Range Target = Globals.ThisWorkbook.Application.Selection;

            Range commonRng = Globals.ThisWorkbook.Application.Intersect(Target, ShefflerWB.DeliveryTable.DataBodyRange);
            if (commonRng == null) return;

            DialogResult msg = MessageBox.Show("Удалить авто с заказами", "Удалить", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (DialogResult.No == msg) return;
            ShefflerWB workBook = new ShefflerWB();

            int row = commonRng.Row;
            int column = ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Range.Column;

            commonRng = ShefflerWB.DeliverySheet.Cells[row, column];
            int numberDelivery = int.TryParse(commonRng.Text, out int nmDelivery) ? nmDelivery : 0;


            for (int i = ShefflerWB.DeliveryTable.ListRows.Count; i > 0; --i)
            {
                ListRow listDeliveryRow = ShefflerWB.DeliveryTable.ListRows[i];
                Range deliveryCell = listDeliveryRow.Range[1, ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Index];
                string str = deliveryCell != null ? deliveryCell.Text : "";
                if (int.TryParse(str, out int number))
                {
                    if (number == numberDelivery)
                        ShefflerWB.DeliverySheet.Rows[listDeliveryRow.Range.Row].Delete();
                }
            }

            for (int j = ShefflerWB.OrdersTable.ListRows.Count; j > 0; --j)
            {
                ListRow listOrderRow = ShefflerWB.OrdersTable.ListRows[j];
                Range orderCell = listOrderRow.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index];

                string strDeliveryNum = orderCell != null ? orderCell.Text : "";
                if (int.TryParse(strDeliveryNum, out int DeliveryNum))
                {
                    if (DeliveryNum == numberDelivery)
                        ShefflerWB.DeliverySheet.Rows[listOrderRow.Range.Row].Delete();
                }
            }
            //  Range rng = workBook. //GetCurrentTotalRange();
            Range rng = ShefflerWB.TotalTable.DataBodyRange;
            if (rng == null) return;
            for (int k = rng.Rows.Count; k > 0; k--)
            {
                string idDelivery = rng.Rows[k].Cells[0,
                         ShefflerWB.TotalTable.ListColumns["№ Доставки"].Index].Text;
                if (int.TryParse(idDelivery, out int num))
                {
                    if (num == numberDelivery)
                    {
                        ShefflerWB.TotalSheet.Rows[rng.Rows[k].Row - 1].Delete();
                    }
                }
            }
        }


        /// <summary>
        ///Кнопка Пересчитать маршруты
        /// </summary>
        public void СhangeDelivery()
        {
            List<Order> orders = GetOrdersFromTable();
            List<Delivery> deliveries = EditDeliveres(orders);
            ClearListObj(ShefflerWB.TotalTable);
            PrintTotal(deliveries); 
            PrintChanges(deliveries);
        }

        public void PrintChanges(List<Delivery> deliveries)
        {
            ClearListObj(ShefflerWB.DeliveryTable);
            PrintDelivery(deliveries);

            foreach (ListRow row in ShefflerWB.OrdersTable.ListRows)
            {
                string strNum = row.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : 0;
                if (deliveryNumber == 0) continue;
                string orderId = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Поставка"].Index].Text;
                if (orderId.Length < 10 && !orderId.Contains(" "))
                {
                    orderId = orderId.Length < 10 ? new string('0', 10 - orderId.Length) + orderId : orderId ;
                }

                Delivery delivery = deliveries.Find(d => d.Number == deliveryNumber);

                if (delivery == null) continue;

                Order order = delivery.Orders.Find(r => r.Id == orderId);
                if (order != null)
                {
                    row.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                    row.Range[1, ShefflerWB.OrdersTable.ListColumns["ID Route"].Index].Value = order.DeliveryPoint.Id;
                    row.Range[1, ShefflerWB.OrdersTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                }
            }
            CopyDeliveryToTotal(deliveries);
        }


        /// <summary>
        /// Кнопка изменить нумерацию
        /// </summary>
        public void RenumberDeliveries()
        {
            List<Delivery> deliveries = ReadFromDelivery();
            RenumerateDeliveries(deliveries);
            CopyDeliveryToTotal(ReadFromDelivery());
        }

        /// <summary>
        /// Нумерация доставок по порядку
        /// </summary>
        /// <param name="deliveries"></param>
        public void RenumerateDeliveries(List<Delivery> deliveries)
        {
            Dictionary<int, int> numbers = new Dictionary<int, int>();
            deliveries = (from d in deliveries
                          orderby d.SortPriority
                          select d).ToList();

            foreach (var item in deliveries)
            {
                numbers.Add(item.Number, numbers.Count + 1);
            }

            foreach (ListRow row in ShefflerWB.DeliveryTable.ListRows)
            {
                string oldDeliveryNumber = row.Range[1, ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Index].Text;
                if (!int.TryParse(oldDeliveryNumber, out int num)) continue;
                row.Range[1, ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Index].Value = numbers[num];
            }

            ShefflerWB.DeliveryTableSort();

            foreach (ListRow rowOrder in ShefflerWB.OrdersTable.ListRows)
            {
                string orderDeliveryNumber = rowOrder.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Text;
                if (!int.TryParse(orderDeliveryNumber, out int num)) continue;
                rowOrder.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Value = numbers[num];
            }
            ShefflerWB.OrderTableSort();

            foreach (ListRow rowOrder in ShefflerWB.TotalTable.ListRows)
            {
                string orderDeliveryNumber = rowOrder.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Text;
                if (!int.TryParse(orderDeliveryNumber, out int num)) continue;
                rowOrder.Range[1, ShefflerWB.TotalTable.ListColumns["№ Доставки"].Index].Value = numbers[num];
            }
            ShefflerWB.TotalTableSort();

            return;
        }

        /// <summary>
        /// Перенос данных в таблицу Отгрузки
        /// </summary>
        public void UpdateTotal()
        {
            List<Delivery> deliveries = ReadFromDelivery();
            CopyDeliveryToTotal(deliveries);
        }


        /// <summary>
        /// Подготовка сообщений перевозчикам
        /// </summary>
        public void SendEmailToProviderAdoutOrders()
        {
            List<Delivery> deliveries = GetDeliveriesFromTotalSheet();
            if (deliveries?.Count == 0) return;

            //Уникальны провайдеры в списке доставок
            string[] shippingComp = (from d in deliveries
                                     select d.Truck.ProviderCompany.Name).Distinct().ToArray();
            ClearFolder();
            if (shippingComp.Length == 0) return;

            ProcessBar pb = ProcessBar.Init("Сообщения", shippingComp.Length, 1, "Подготовка писем");
            pb.Show();

            //Для каждого провайдера
            for (int i = 0; i < shippingComp.Length; i++)
            {
                string сompanyShipping = shippingComp[i];           
                if (pb.Cancel) break;
                pb.Action($"Письмо {i + 1} из {pb.Count} {shippingComp[i]} ");

                if (сompanyShipping == "" || сompanyShipping == "Деловые линии") continue;
                List<Delivery> deliverShipping = deliveries.FindAll(x =>
                               x.Truck.ProviderCompany.Name == сompanyShipping).ToList();

                string date = ShefflerWB.DeliverySheet.Range["DateDelivery"].Text;
                string subject = Properties.Settings.Default.ProviderSubjectOrder;
                string message = Properties.Settings.Default.ProviderMessageOrder;

                subject = subject.Replace("[date]", date).Replace("[provider]", shippingComp[i]);
                message = message.Replace("[date]", date).Replace("[provider]", shippingComp[i]);
                string attachment = GenerateAttachmentFile(deliverShipping, subject);
                List<string> attachments = new List<string> { attachment };

                Email messenger = new Email();
                messenger.MailToProvider(сompanyShipping, subject, message, attachments, Email.TypeSend.Display);
            }
            pb.Close();
        }

        /// <summary>
        /// Отправка сообщений провайдерам со списком уточнений
        /// </summary>
        public void SendEmailToProviderAdoutAdding()
        {
            List<Delivery> deliveries = GetDeliveriesFromTotalSheet();
            if (deliveries?.Count == 0) return;

            //Уникальны провайдеры в списке доставок
            string[] shippingComp = (from d in deliveries select d.Truck.ProviderCompany.Name).Distinct().ToArray();
            ClearFolder();
            if (shippingComp.Length == 0) return;
            ProcessBar pb = ProcessBar.Init("Сообщения", shippingComp.Length, 1, "Подготовка писем");
            pb.Show();

            //Для каждого провайдера
            for (int i = 0; i < shippingComp.Length; i++)
            {
                string сompanyShipping = shippingComp[i];
        
                if (pb.Cancel) break;
                pb.Action($"Письмо {i + 1} из {pb.Count} {shippingComp[i]} ");

                if (сompanyShipping == "" || сompanyShipping == "Деловые линии") continue;
                List<Delivery> deliverShipping = deliveries.FindAll(x => x.Truck.ProviderCompany.Name == сompanyShipping).ToList();

                string date = ShefflerWB.DeliverySheet.Range["DateDelivery"].Text;
                string subject = Properties.Settings.Default.ProviderSubjectAdd;
                string message = Properties.Settings.Default.ProviderMessageAdd;

                subject = subject.Replace("[date]", date).Replace("[provider]", shippingComp[i]);
                message = message.Replace("[date]", date).Replace("[provider]", shippingComp[i]);
                string attachment = GenerateAttachmentFile(deliverShipping, subject);
                List<string> attachments = new List<string> { attachment };

                Email messenger = new Email();
                messenger.MailToProvider(сompanyShipping, subject, message, attachments, Email.TypeSend.Display);
            }
            pb.Close();
        }

        /// <summary>
        /// Отправка письма в кастом сервис
        /// </summary>
        public void CreateLetterToCS()
        {
            string path = Globals.ThisWorkbook.Path + "\\MailToCS\\";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string date = ShefflerWB.DeliverySheet.Range["DateDelivery"].Text;
            //      string attachment = path + DateTime.Today.ToString("dd.MM.yyyy") + ".xlsx";
            List<string> attachments = new List<string>();
            string attachment = path + date + ".xlsx";
            attachments.Add(attachment);
            string attachmentAllOrders = Properties.Settings.Default.AllOrders ; // path + date + ".xlsx";
            if (!string.IsNullOrWhiteSpace(attachmentAllOrders)) attachments.Add(attachmentAllOrders);

            ShefflerWB.TotalSheet.Copy();
            Globals.ThisWorkbook.Application.ActiveWorkbook.ActiveSheet.Columns[22].Delete();
            Globals.ThisWorkbook.Application.ActiveWorkbook.SaveAs(attachment, XlFileFormat.xlWorkbookDefault);
            Globals.ThisWorkbook.Application.ActiveWorkbook.Close();

            string to = Properties.Settings.Default.SettingCSLetterTo;
            string copy = Properties.Settings.Default.SettingCSLetterCopy;
            string subject = Properties.Settings.Default.SettingCSLetterSubject;
            subject = subject.Replace("[date]", date);

            string message = Properties.Settings.Default.SettingCSLetterMessage;
            message = message.Replace("[date]", date);

            Email email = new Email();
            email.CreateMail(to, copy, subject, message, attachments);
        }

        /// <summary>
        /// Импорт данных из писем провайдеров
        /// </summary>
        /// <param name="file"></param>
        public void ReadMessageFile(string file)
        {
            Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: file);
            Worksheet sh = wb.Sheets[1];
            try
            {
                ListObject list = sh.ListObjects["Таблица1"];
                for (int i = 1; i <= list.ListRows.Count; i++)
                {
                    ListRow row = list.ListRows[i];
                    string idProvider = row.Range[1, list.ListColumns["ID перевозчика"].Index].Text;
                    if (string.IsNullOrWhiteSpace(idProvider)) continue;
                    string NameProvider = row.Range[1, list.ListColumns["Водитель (ФИО)"].Index].Text;
                    string NumberProvider = row.Range[1, list.ListColumns["Номер, марка"].Index].Text;
                    string PhoneProvider = row.Range[1, list.ListColumns["Телефон водителя"].Index].Text;

                    Driver driver = new Driver()
                    {
                        Id = idProvider,
                        Name = NameProvider,
                        Phone = PhoneProvider,
                        CarNumber = NumberProvider
                    };
                    WriteProviderInfo(driver);
                }

            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Не удалось прочитать таблицу в файле " + file);
                //  MessageBox.Show(ex.Message);
            }
            finally
            {
                Globals.ThisWorkbook.Application.DisplayAlerts = false;
                wb.Close();
                Globals.ThisWorkbook.Application.DisplayAlerts = true;
            }
        }

        /// <summary>
        ///  Считать заказы с листа
        /// </summary>
        /// <param name="ordersTable"></param>
        /// <returns></returns>
        public List<Order> GetOrdersFromTable()
        {
            ListObject ordersTable = ShefflerWB.OrdersTable;
            List<Order> orders = new List<Order>();

            foreach (ListRow row in ordersTable.ListRows)
            {
                Order order = new Order();
                string strNum = row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : -1;
                if (deliveryNumber == -1) continue;
                order.DeliveryNumber = deliveryNumber;
                order.Id = row.Range[1, ordersTable.ListColumns["Поставка"].Index].Text;

                string city = row.Range[1, ordersTable.ListColumns["Город"].Index].Text;

                strNum = row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Text;
                order.PointNumber = int.TryParse(strNum, out int pointnum) ? pointnum : 0;

                string customerId = row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Text;
                customerId = customerId.Length < 10 ? new string('0', 10 - customerId.Length) + customerId : customerId;
                string customerName = row.Range[1, ordersTable.ListColumns["Получатель"].Index].Text;
                Customer customer = new Customer(customerId)
                {
                    Name = customerName
                };
                order.Customer = customer;
                string idRouteStr = row.Range[1, ordersTable.ListColumns["ID Route"].Index].Text;
                int idRoute = int.TryParse(idRouteStr, out int idr) ? idr : 0;
                string Route = row.Range[1, ordersTable.ListColumns["Направление"].Index].Text;

                DeliveryPoint point = new DeliveryPoint
                {
                    Id = idRoute,
                    IdCustomer = customerId,
                    Customer = customerName,
                    City = city,
                    PriorityPoint = pointnum,
                    Route = Route
                };
                order.DeliveryPoint = point;
                order.RouteCity = row.Range[1, ordersTable.ListColumns["Направление"].Index].Text;
                string weight = row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;

                weight = row.Range[1, ordersTable.ListColumns["Вес брутто"].Index].Text;
                order.WeightBrutto = double.TryParse(weight, out wgt) ? wgt : 0;

                order = GetOrdersInfoFromTotal(order);
                orders.Add(order);
            }
            return orders;
        }

        private Order GetOrdersInfoFromTotal(Order order)
        {
            int column = ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index;
            foreach (ListRow row in ShefflerWB.TotalTable.ListRows)
            {
                string idOrder = row.Range[1, column].Text;
                 idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                if (order.Id == idOrder)
                {
                    string dateTable = row.Range[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Text;
                    order.DateDelivery = dateTable;

                    string pc = row.Range[1, ShefflerWB.TotalTable.ListColumns["Кол-во паллет"].Index].text;
                    if (int.TryParse(pc, out int pallets)) order.PalletsCount = pallets;

                    string wb = row.Range[1, ShefflerWB.TotalTable.ListColumns["Брутто вес"].Index].Value.ToString();
                    if (double.TryParse(wb, out double wbrutto)) order.WeightBrutto = wbrutto;

                    string cost = row.Range[1, ShefflerWB.TotalTable.ListColumns["Стоимость поставки"].Index].Value.ToString();
                    if (double.TryParse(cost, out double costProd)) order.Cost = costProd;

                    string ttn = row.Range[1, ShefflerWB.TotalTable.ListColumns["Номер накладной"].Index].Text;
                    order.TransportationUnit = ttn;
                    break;
                }
            }

            return order;
        }
      

        private void UpdateOrderFromTotal()
        {
            List<Order> orders = new List<Order>();

            foreach (ListRow row in ShefflerWB.OrdersTable.ListRows)
            {
                Order order = new Order();

                string dateTable = row.Range[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Text;
                order.DateDelivery = dateTable;
                string strNum = row.Range[1, ShefflerWB.OrdersTable.ListColumns["№ Доставки"].Index].Text;
                int deliveryNumber = int.TryParse(strNum, out int n) ? n : -1;
                if (deliveryNumber == -1) continue;
                order.DeliveryNumber = deliveryNumber;
                order.Id = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Поставка"].Index].Text;

                string city = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Город"].Index].Text;

                strNum = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Порядок выгрузки"].Index].Text;
                order.PointNumber = int.TryParse(strNum, out int pointnum) ? pointnum : 0;

                string customerId = row.Range[1, ShefflerWB.OrdersTable.ListColumns["ID Получателя"].Index].Text;
                customerId = customerId.Length < 10 ? new string('0', 10 - customerId.Length) + customerId : customerId;
                string customerName = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Получатель"].Index].Text;

                Customer customer = new Customer(customerId) { Name = customerName };
                order.Customer = customer;

                DeliveryPoint point = ShefflerWB.RoutesList.Find(r => r.IdCustomer == customerId);
                order.DeliveryPoint = point;
                order.RouteCity = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Направление"].Index].Text;
                string weight = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Вес нетто"].Index].Text;
                order.WeightNetto = double.TryParse(weight, out double wgt) ? wgt : 0;

                weight = row.Range[1, ShefflerWB.OrdersTable.ListColumns["Вес брутто"].Index].Text;
                order.WeightBrutto = double.TryParse(weight, out wgt) ? wgt : 0;

                order = GetOrdersInfoFromTotal(order);
                PrintOrder(row, order);
            }
        }


        /// <summary>
        /// Получение списка заказов из таблицы Отгрузки
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private List<Order> GetOrdersFromTotalTable(Range range)
        {
            List<Order> orders = new List<Order>();
            int column = ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index;

            foreach (Range row in range.Rows)
            {
                // string dateDelivery = ShefflerWB.DateDelivery;
                // if (dateTable != dateDelivery) continue;

                string idOrder = row.Cells[1, column].Text;
                if (string.IsNullOrWhiteSpace(idOrder)) continue;
                Order order = new Order();
                string dateTable = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Text;
                order.DateDelivery = dateTable;

                idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                order.Id = idOrder;
                order.TransportationUnit = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Номер накладной"].Index].Text;

                double.TryParse(row.Cells[1, ShefflerWB.TotalTable.ListColumns["Брутто вес"].Index].Value.ToString(), out double wt);
                order.WeightBrutto = wt;

                double.TryParse(row.Cells[1, ShefflerWB.TotalTable.ListColumns["Нетто вес"].Index].Value.ToString(), out wt);
                order.WeightNetto = wt;

                double.TryParse(row.Cells[1, ShefflerWB.TotalTable.ListColumns["Стоимость поставки"].Index].Value.ToString(), out wt);
                order.Cost = wt;

                int.TryParse(row.Cells[1, ShefflerWB.TotalTable.ListColumns["Кол-во паллет"].Index].Text, out int countpallet);
                order.PalletsCount = countpallet;

                order.Customer.Id = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Номер грузополучателя"].Index].Text;
                order.Customer.Name = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Грузополучатель"].Index].Text;
                order.RouteCity = row.Cells[1, ShefflerWB.TotalTable.ListColumns["Направление"].Index].Text;

                orders.Add(order);
            }

            return orders;
        }

       

       
            //'=============================================================================
            /// <summary>
            /// Получить инфо из выгруза  
            /// </summary>
            /// <param name="file"></param>
            /// <returns></returns>
            private Order GetFromFile(string file)
        {
            Order order = new Order();
            Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: file);
            Worksheet sh = wb.Sheets[1];
            Range rng = sh.UsedRange;
            string strTitle = ShefflerWB.FindValue("Заявка на перевозку", rng, 0, 0);
            if (strTitle == "") return null;

            string strCustomerId = ShefflerWB.FindValue("Номер грузополучателя", rng, 0, 1);
            order.Customer.Id = strCustomerId;

            string strName = ShefflerWB.FindValue("Грузополучатель", rng, 0, 1);
            order.Customer.Name = strName.Trim();

            string strTU = ShefflerWB.FindValue("Номер накладной", rng, 0, 1);
            order.TransportationUnit = strTU.Replace(",", "/");

            string strCost = ShefflerWB.FindValue("Стоимость", rng, 0, 0);
            Regex regexCost = new Regex(@"(\d+\s?)+(\,\d+)?");
            strCost = regexCost.Match(strCost).Value;
            order.Cost = double.TryParse(strCost, out double ct) ? ct : 0;

            string strWeightBrutto = ShefflerWB.FindValue("брутто", rng, 0, 0);
            strWeightBrutto = regexCost.Match(strWeightBrutto).Value;
            double weight = double.TryParse(strWeightBrutto, out double wt) ? wt : 0;
            order.WeightBrutto = weight;

            string strPalletsCount = ShefflerWB.FindValue("грузовых мест", rng, 0, 0);
            Regex regexId = new Regex(@"\d+");
            strPalletsCount = regexId.Match(strPalletsCount).Value;
            int countPallets = int.TryParse(strPalletsCount, out int count) ? count : 0;
            order.PalletsCount = countPallets;

            string strID = ShefflerWB.FindValue("Номер поставки", rng, 0, 1);
            order.Id = strID;
            Globals.ThisWorkbook.Application.DisplayAlerts = false;
            wb.Close();
            Globals.ThisWorkbook.Application.DisplayAlerts = true;
            return order;
        }

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
          //  order.Route = row.Cells[1, GetColumn(row.Parent, "Направление", 1)].Text;
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
            if (ordersPath == "" || ordersSap.Count == 0) return null;
            Workbook orderBook = null;
            try
            {
                orderBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: ordersPath);

                // Хранить файл для отправки в CS
                Properties.Settings.Default.AllOrders = ordersPath;
                Properties.Settings.Default.Save();
            }
            catch
            {
                throw new System.Exception("Не удалось открыть книгу Excel: " + ordersPath);
            }

            List<Order> ordersInfo = new List<Order>();
            foreach (Order order in ordersSap)
            {
                if (!string.IsNullOrWhiteSpace(order.Id))
                {
                    List<string> orderInfo = GetOrderInfo(orderBook.Sheets[1], order.Id);

                    string TTN = GetOrderTTN(orderBook.Sheets[1], order.Id);
                    if (!string.IsNullOrWhiteSpace(TTN) && 
                         string.IsNullOrWhiteSpace(order.TransportationUnit))
                    {
                        order.TransportationUnit = TTN;
                    }
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

        private string GetOrderTTN(Worksheet sheet, string delivery)
        {
            Range findRange = sheet.Columns[1];
            string search = delivery.Length < 10 ? new string('0', 10 - delivery.Length) + delivery : delivery;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);
            if (fcell == null) return null;

            //Начало накладной 
            int rowStart = fcell.Row;
            string TTN = "";
            for (int i = fcell.Row; i > 1; --i)
            {
                // Ограничения вверх
                string strCell = findRange.Cells[i, 1].Text.Trim();
                if (strCell.Contains("ТТН:") )
                {                     
                    Regex regexId = new Regex(@"\d+");
                    TTN = regexId.Match(strCell).Value;
                    int NumberTTN = int.Parse(TTN);
                    TTN =NumberTTN.ToString();
                    break;
                }
            }
            return TTN;
        }

        /// <summary>
        /// Получение дополнительной информации по заказу  AllOrders
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="delivery"></param>
        /// <returns></returns>
        private List<string> GetOrderInfo(Worksheet sheet, string delivery)
        {
            Range findRange = sheet.Columns[1];
            string search = delivery.Length < 10 ? new string('0', 10 - delivery.Length) + delivery : delivery;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);
            if (fcell == null) return null;

            //Начало накладной 
            int rowStart = fcell.Row;

            for (int i = fcell.Row; i > 1; --i)
            {
                // Ограничения вверх
                string strCell = findRange.Cells[i, 1].Text.Trim();
                if (strCell.Contains("KG/") ||
                    strCell.Contains("ТТН:") ||
                    strCell.Contains("№") ||
                    string.IsNullOrWhiteSpace(strCell))
                {
                    rowStart = i;
                    break;
                }
            }

            int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;
            // конец диапазона
            int rowEnd = rowStart + 1;
            List<string> info = new List<string>();
            do
            {
                fcell = findRange.Cells[rowEnd++, 1];
                string cellText = fcell.Text;
                cellText.Trim();
                cellText = cellText.Replace("\t", "");
                cellText = cellText.Replace(";;;", "");
                if (string.IsNullOrEmpty(cellText.Replace(";", "")) ||
                     cellText.Contains("№ грузового места") ||
                        cellText.Contains("KG/"))
                    break;
                info.Add(cellText);
            }
            while (rowEnd <= lastRow);
            return info;
        }

        /// <summary>
        /// Запись информации о провайдере
        /// </summary>
        /// <param name="carrier"></param>
        private void WriteProviderInfo(Driver carrier)
        {
            foreach (ListRow row in ShefflerWB.TotalTable.ListRows)
            {
                string id = row.Range[1, ShefflerWB.TotalTable.ListColumns["ID перевозчика"].Index].Text;
                if (id == carrier.Id)
                {
                    row.Range[1, ShefflerWB.TotalTable.ListColumns["Водитель (ФИО)"].Index].Value = carrier.Name;
                    row.Range[1, ShefflerWB.TotalTable.ListColumns["Телефон водителя"].Index].Value = carrier.Phone;
                    row.Range[1, ShefflerWB.TotalTable.ListColumns["Номер,марка"].Index].Value = carrier.CarNumber;
                }
            }
        }

        /// <summary>
        /// Формирование таблицы Доставки
        /// </summary>
        /// <param name="deliveries"></param>
        /// <param name="DeliveryTable"></param>
        /// <param name="OrderTable"></param>
        private void PrintDelivery(List<Delivery> deliveries)
        {
            ListObject DeliveryTable = ShefflerWB.DeliveryTable;
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
                    rowDelivery = DeliveryTable.ListRows[1];
                }
                else
                {
                    AddListRow(DeliveryTable);
                    rowDelivery = DeliveryTable.ListRows[DeliveryTable.ListRows.Count - 1];
                }
                rowDelivery.Range[1, DeliveryTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Компания"].Index].Value =
                                                                    delivery.Truck?.ProviderCompany?.Name ?? "";
                rowDelivery.Range[1, DeliveryTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;

                if (delivery?.MapDelivery.Count > 0)
                {
                    rowDelivery.Range[1, DeliveryTable.ListColumns["Направление"].Index].Value =
                                                            delivery.MapDelivery[0].RouteName;
                    rowDelivery.Range[1, DeliveryTable.ListColumns["ID Route"].Index].Value =
                                                                        delivery?.MapDelivery[0].Id;
                }
                rowDelivery.Range[1, DeliveryTable.ListColumns["Тип ТС, тонн"].Index].Value 
                                                                                = delivery.Truck?.Tonnage ?? 0;
                rowDelivery.Range[1, DeliveryTable.ListColumns["Вес доставки"].Index].Value 
                                                                               = delivery.TotalWeight;
                                                //FormulaR1C1 =
                                               // "=IF(SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес брутто])=0, SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес нетто]), SUMIF(TableOrders[№ Доставки],[@[№ Доставки]],TableOrders[Вес брутто]))";
            }
            pb.Close();
        }

        /// <summary>
        /// Формирование таблицы Товары (заказы)
        /// </summary>
        /// <param name="deliveries"></param>
        /// <param name="OrderTable"></param>
        private void PrintOrders(List<Delivery> deliveries)
        {
            ListObject OrderTable = ShefflerWB.OrdersTable;
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

        /// <summary>
        /// Запись одной строки в таблицу Товары
        /// </summary>
        /// <param name="row"></param>
        /// <param name="order"></param>
        /// <param name="deliveryNumber"></param>
        private void PrintOrder(ListRow row, Order order, int? deliveryNumber = null)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject ordersTable = deliverySheet.ListObjects["TableOrders"];

            if (deliveryNumber != null)
            {
                row.Range[1, ordersTable.ListColumns["№ Доставки"].Index].Value = deliveryNumber;
            }

            row.Range[1, ordersTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
            row.Range[1, ordersTable.ListColumns["Поставка"].Index].Value = order.Id;
            row.Range[1, ordersTable.ListColumns["ID Получателя"].Index].Value = order.Customer?.Id ?? "";
            row.Range[1, ordersTable.ListColumns["Получатель"].Index].Value = order.Customer.Name;
            row.Range[1, ordersTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
            row.Range[1, ordersTable.ListColumns["ID Route"].Index].Value = order.DeliveryPoint.Id;
            row.Range[1, ordersTable.ListColumns["Вес нетто"].Index].Value = order.WeightNetto;
            row.Range[1, ordersTable.ListColumns["Вес брутто"].Index].Value = order.WeightBrutto;
            row.Range[1, ordersTable.ListColumns["Направление"].Index].Value = order.RouteCity;
        }

        //private void ClearTotal()
        //{

        //    int column = ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index;
        //    foreach (Range row in ShefflerWB.TotalTable.DataBodyRange.Rows)
        //    {
        //        string dataRow = row.Cells[1, column].Text;
        //        if (string.IsNullOrWhiteSpace(dataRow))
        //        {
        //            ShefflerWB.TotalSheet.Cells.Rows[row.Row].Delete();
        //        }
        //    }
        //}

        /// <summary>
        /// Заполнить таблицу отгрузки
        /// </summary>
        /// <param name="totalTable"></param>
        /// <param name="deliveries"></param>
        private void PrintTotal(List<Delivery> deliveries)
        {
            ListObject totalTable = ShefflerWB.TotalTable;
            if (deliveries.Count < 1) return;
           
            if (totalTable.ListRows.Count > 0 &&
                totalTable.ListRows[totalTable.ListRows.Count].Range[1, 1].Text != "")
            {
                ShefflerWB.TotalTable.ListRows.Add();
            }

            foreach (Delivery delivery in deliveries)
            {
                ListRow row;
                if (totalTable.ListRows.Count == 0)
                {
                    ShefflerWB.TotalTable.ListRows.Add();
                    row = totalTable.ListRows[1];
                }
                else
                {
                    ShefflerWB.TotalTable.ListRows.Add();
                    row = totalTable.ListRows[totalTable.ListRows.Count - 1];
                }

                bool mainRow = true;
            
                row.Range[1, totalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                row.Range[1, totalTable.ListColumns["Перевозчик"].Index].Value = delivery.Truck?.ProviderCompany?.Name;
                row.Range[1, totalTable.ListColumns["Тип ТС, тонн"].Index].Value = delivery.Truck?.Tonnage ?? 0;
                row.Range[1, totalTable.ListColumns["Время подачи ТС"].Index].Value = delivery.Time;
                foreach (Order order in delivery.Orders)
                {
                    if (!mainRow)
                    {
                        totalTable.ListRows.Add();
                        row = totalTable.ListRows[totalTable.ListRows.Count - 1];
                    }
                    row.Range[1, totalTable.ListColumns["Дата отгрузки"].Index].Value =  ShefflerWB.DateDelivery;   //   string.IsNullOrWhiteSpace( order.DateDelivery) ?
                    //ShefflerWB.DateDelivery :  string date = ShefflerWB.DateDelivery;
                    row.Range[1, totalTable.ListColumns["Порядок выгрузки"].Index].Value =
                            delivery.MapDelivery.FindIndex(x => x.IdCustomer == order.Customer.Id) + 1;


                    row.Range[1, totalTable.ListColumns["№ Доставки"].Index].Value = delivery.Number;
                    row.Range[1, totalTable.ListColumns["Номер накладной"].Index].Value = order.TransportationUnit;
                    row.Range[1, totalTable.ListColumns["Номер поставки"].Index].Value = order.Id;
                    row.Range[1, totalTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    row.Range[1, totalTable.ListColumns["Направление"].Index].Value = order.RouteCity;
                    row.Range[1, totalTable.ListColumns["Номер грузополучателя"].Index].Value = order.Customer?.Id ?? "";
                    row.Range[1, totalTable.ListColumns["Брутто вес"].Index].Value = order.WeightBrutto;
                    row.Range[1, totalTable.ListColumns["Нетто вес"].Index].Value = order.WeightNetto;
                    row.Range[1, totalTable.ListColumns["Грузополучатель"].Index].Value = $"{order.Customer?.Name ?? ""}";
                    row.Range[1, totalTable.ListColumns["Стоимость поставки"].Index].Value = order.Cost;
                    row.Range[1, totalTable.ListColumns["Кол-во паллет"].Index].Value = order.PalletsCount;

                    mainRow = false;
                }

            }
        }

        /// <summary>
        /// очистка таблицы удалением строк листа
        /// </summary>
        /// <param name="listObject"></param>
        private void ClearListObj(ListObject listObject)
        {
            Globals.ThisWorkbook.Application.DisplayAlerts = false;
            if (listObject.DataBodyRange == null) return;
            listObject.DataBodyRange.EntireRow.Delete();
            Globals.ThisWorkbook.Application.DisplayAlerts = true;
        }

        private void AddListRow(ListObject listObject)
        {
            Worksheet worksheet = listObject.Parent;
            if (listObject.ListRows.Count > 0)
            {
                worksheet.Rows[listObject.DataBodyRange.Row + listObject.DataBodyRange.Rows.Count - 1].Insert();
            }
            else
            {
                worksheet.Rows[listObject.HeaderRowRange.Row + 1].Insert();
            }
        }

        /// <summary>
        /// Распределить заказы по автомобилям
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        private List<Delivery> CompleteAuto(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
            orders = orders.OrderBy(x => x.WeightNetto).ToList();

            List<DeliveryPoint> points = ShefflerWB.RoutesList;
            Delivery deliveryNoRoute = new Delivery
            {
                HasRoute = false
            };

            while (orders.Count > 0)
            {
                bool hasDelivery = false;
                // Проходим по возможным маршрутам
                foreach (DeliveryPoint point in points)
                {
                    // Ищем товар, который можно отправить указанным маршрутом                    
                    for (int iOrder = orders.Count - 1; iOrder >= 0; iOrder--)
                    {
                        if (orders[iOrder].Customer.Id != point.IdCustomer) continue;
                        hasDelivery = true;
                        orders[iOrder].DeliveryPoint = point;
                        // Пытаемся добавить к имеющимся машинам
                        Delivery delivery = null;
                        foreach (Delivery iDelivery in deliveries)
                        {
                            string city = iDelivery.MapDelivery[0].City;
                            // У машины другой маршрут
                            if (iDelivery.Orders[0].DeliveryPoint.Id != point.Id) continue;
                            // Для мск допустимо 3 точки 
                            if ((city.Contains("MSK") || city.Contains("MO")) && iDelivery.MapDelivery.Count == 3)
                            {
                                Order orderFind = iDelivery.Orders.Find(x => x.Customer.Id == orders[iOrder].Customer.Id);
                                if (orderFind == null) continue;
                            }

                            if (ShefflerWB.InternationalCityList.Any(x => x == city) &&
                                     orders[iOrder].DeliveryPoint.City == city) //Nur - Sultan //Yerevan
                            {
                                if (iDelivery.CheckDeliveryWeightLTL(orders[iOrder]))
                                { delivery = iDelivery; break; }
                            }
                            else if (iDelivery.CheckDeliveryWeight(orders[iOrder]))
                            { delivery = iDelivery; break; }
                        }
                        if (delivery == null)
                        {
                            delivery = new Delivery();
                            delivery.Number = deliveries.Count+1;
                            deliveries.Add(delivery);
                        }
                        orders[iOrder].DeliveryPoint = point;
                        Order orderCurrentCustomer = delivery.Orders.Find(x => x.Customer.Id == orders[iOrder].Customer.Id);
                        //Порядок выгруза / Если уже есть груз для заказчика 
                        int number = orderCurrentCustomer == null ?
                              delivery.MapDelivery.Count + 1
                            : orderCurrentCustomer.PointNumber;
                        orders[iOrder].PointNumber = number;
                        delivery.Orders.Add(orders[iOrder]);
                        
                        orders.RemoveAt(iOrder);
                    }
                    if (hasDelivery) break;
                }
                // не нашли маршрут
                if (!hasDelivery)
                {
                    deliveryNoRoute.Orders.Add(orders[0]);
                    deliveryNoRoute.Number = deliveries.Count;
                    orders.RemoveAt(0);
                }
            }
            if (deliveryNoRoute.Orders.Count > 0) deliveries.Add(deliveryNoRoute);
            return deliveries;
        }

        /// <summary>
        /// Получаем список товаров, которые доставляются неукомплектованными авто
        /// </summary>
        /// <param name="firstDeliveries"></param>
        /// <param name="mapDeliveryCount"></param>
        /// <returns></returns>
        private List<Order> GetOrdersFromIncompleteDelivery(List<Delivery> firstDeliveries)
        {
            List<Order> orders = new List<Order>();
            foreach (Delivery delivery in firstDeliveries)
            {
                string city = delivery.MapDelivery[0].City;

                // обрабатываем только МСК и МО и машины с 1-2 точками
                if (!(city.Contains("MSK") || city.Contains("MO")) || delivery.MapDelivery.Count > 2) { continue; }

                foreach (Order order in delivery.Orders)
                {
                    orders.Add(order);
                }
            }
            return orders;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="orders"></param>
        /// <param name="routeId"></param>
        /// <returns></returns>
        private bool IsComplete(List<Order> orders, int routeId)
        {
            List<DeliveryPoint> points = ShefflerWB.RoutesList;
            var pointsInOrder = (from i in orders
                                 select i.Customer.Id).ToList();

            var pointsIdInRoute = (from i in points
                                   where i.Id == routeId
                                   select i.IdCustomer).ToList();
            var result = pointsIdInRoute.Intersect(pointsInOrder);

            return result.Count() == pointsIdInRoute.Count();
        }

        private List<Delivery> CompileAutoSecond(List<Delivery> firstDeliveries)
        {
            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orders = GetOrdersFromIncompleteDelivery(firstDeliveries);

            for (int iDelyvery = firstDeliveries.Count - 1; iDelyvery >= 0; iDelyvery--)
            {
                bool findOrder = false;
                foreach (Order item in firstDeliveries[iDelyvery].Orders)
                {

                    foreach (Order iorder in orders)
                    {
                        if (item.Id == iorder.Id)
                        {
                            firstDeliveries.RemoveAt(iDelyvery);
                            findOrder = true;
                            break;
                        }
                    }
                    if (findOrder) break;
                }
            }

            List<DeliveryPoint> points = ShefflerWB.RoutesList;

            var uniqueRoutesId = from route in points
                                 where route.PriorityRoute > 1
                                 group route by route.Id into g
                                 select new { Id = g.Key, Count = g.Count() };

            /// Проверряем каждый маршрут ищем по какомму можно отправить грузы
            foreach (var routeId in uniqueRoutesId)
            {
                if (!IsComplete(orders, routeId.Id)) continue;

                var pointsRoute = (from i in points
                                   where i.Id == routeId.Id
                                   select i).ToList();

                foreach (DeliveryPoint point in pointsRoute)
                {
                    for (int iOrder = orders.Count - 1; iOrder >= 0; iOrder--)
                    {
                        if (orders[iOrder].Customer.Id != point.IdCustomer) continue;

                        orders[iOrder].DeliveryPoint = point;
                        // Пытаемся добавить к имеющимся машинам
                        Delivery delivery = null;
                        foreach (Delivery iDelivery in deliveries)
                        {
                            string city = iDelivery.MapDelivery[0].City;
                            // У машины другой маршрут
                            if (iDelivery.Orders[0].DeliveryPoint.Id != point.Id) continue;

                            if (iDelivery.CheckDeliveryWeight(orders[iOrder]))
                            {
                                delivery = iDelivery;
                                break;
                            }
                        }

                        if (delivery == null)
                        {
                            delivery = new Delivery();
                            delivery.Number = deliveries.Count+1;
                            deliveries.Add(delivery);
                        }

                        orders[iOrder].DeliveryPoint = point;
                        Order orderCurrentCustomer = delivery.Orders.Find(x => x.Customer.Id == orders[iOrder].Customer.Id);

                        //Порядок выгруза / Если уже есть груз для заказчика 
                        int number = orderCurrentCustomer == null ?
                              delivery.MapDelivery.Count + 1
                            : orderCurrentCustomer.PointNumber;
                        orders[iOrder].PointNumber = number;
                        delivery.Orders.Add(orders[iOrder]);
                        
                        orders.RemoveAt(iOrder);
                    }
                }
            }

            foreach (Delivery item in firstDeliveries)
            {
                item.Number = deliveries.Count + 1;
                deliveries.Add(item);
            }
            if (orders.Count > 0)
            {
                List<Delivery> AutoLost = CompleteAuto(orders);

                foreach (Delivery item in AutoLost)
                {
                    item.Number = deliveries.Count + 1;
                    deliveries.Add(item);
                }
            }
            return deliveries;
        }


        /// <summary>
        /// перенести с деливери на лист Отгрузка
        /// </summary>
        private void CopyDeliveryToTotal(List<Delivery> deliveries)
        {
            ProcessBar pb = ProcessBar.Init("Вывод данных", deliveries.Count, 1, "Обновление доставок");
            if (deliveries == null || deliveries.Count < 1 | pb == null) return;

            for (int ixDelivery = 0; ixDelivery < deliveries.Count; ixDelivery++)
            {
                Delivery delivery = deliveries[ixDelivery];
                if (pb.Cancel) break;
                pb.Action($"Доставка {ixDelivery + 1} из {pb.Count}");
                pb.Show();

                for (int ixOrder = 0; ixOrder < delivery.Orders.Count; ixOrder++)
                {
                    Order order = delivery.Orders[ixOrder];
                    ListRow totalRow = null;
                    for (int i = 1; i <= ShefflerWB.TotalTable.ListRows.Count; i++)
                    {
                        totalRow = ShefflerWB.TotalTable.ListRows[i];
                        string idOrder = totalRow.Range[1,
                                    ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index].Text;
                        idOrder = idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                        if ((!string.IsNullOrWhiteSpace(idOrder)) && (order.Id==idOrder))
                        { break; }
                        totalRow = null;
                    }
                    if (totalRow == null)
                    {
                        ShefflerWB.TotalTable.ListRows.Add();
                        totalRow = ShefflerWB.TotalTable.ListRows[ShefflerWB.TotalTable.ListRows.Count - 1];
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Дата отгрузки"].Index].Value = ShefflerWB.DateDelivery;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index].value = order.Id;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Номер накладной"].Index].Value = order.TransportationUnit;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Грузополучатель"].Index].Value = order.Customer.Name;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Номер грузополучателя"].Index].Value = order.Customer.Id;
                    }
                    string wtBrutto = totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Брутто вес"].Index].Text;
                    if (double.TryParse(wtBrutto, out double wb))
                    {
                        order.WeightBrutto = wb;
                    }

                    totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["№ Доставки"].Index].Value = order.DeliveryNumber;
                    totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Порядок выгрузки"].Index].Value = order.PointNumber;
                    totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Стоимость поставки"].Index].Value = order.Cost;
                    totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Направление"].Index].Value = order.RouteCity;
                    totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Город"].Index].Value = order.DeliveryPoint.City;
                    if (ixOrder == 0)
                    {
                        //Если Заказ 1й в списке доставки выводим инфо о заказе
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Стоимость доставки"].Index].Value = delivery.Cost;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Перевозчик"].Index].Value =
                                                                                delivery.Truck?.ProviderCompany?.Name ?? "";
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Тип ТС, тонн"].Index].Value = delivery.Truck?.Tonnage ?? 0;
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Время подачи ТС"].Index].Value = delivery.Time;
                    }
                    else
                    {
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Стоимость доставки"].Index].Value = "";
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Перевозчик"].Index].Value = "";
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Тип ТС, тонн"].Index].Value = "";
                        totalRow.Range[1, ShefflerWB.TotalTable.ListColumns["Время подачи ТС"].Index].Value = "";

                    }
                }
            }
            // Сортировка
            ShefflerWB.TotalTableSort();
            pb.Close();
            // ShefflerWB.TotalSheet.Activate();
        }

        /// <summary>
        ///  Прочитать доставки 
        /// </summary>
        /// <returns></returns>
        private List<Delivery> ReadFromDelivery()
        {
            List<Delivery> deliveries = new List<Delivery>();
            List<Order> orders = GetOrdersFromTable();

            foreach (ListRow deliveryRow in ShefflerWB.DeliveryTable.ListRows)
            {
                string strNumber = deliveryRow.Range[1, ShefflerWB.DeliveryTable.ListColumns["№ Доставки"].Index].Text;
                if (!string.IsNullOrWhiteSpace(strNumber))
                {
                    int deliveryNumber = int.TryParse(strNumber, out int num) ? num : 0;
                    List<Order> ordersDelivery = orders.FindAll(x => x.DeliveryNumber == deliveryNumber).ToList();
                    if (ordersDelivery.Count == 0) { continue; }
                    Delivery delivery = new Delivery
                    {
                        Orders = ordersDelivery,
                        Number = deliveryNumber
                    };
                    string providerName = deliveryRow.Range[1, ShefflerWB.DeliveryTable.ListColumns["Компания"].Index].Text;
                    Provider shippingCompany = new Provider() { Name = providerName };
                    string carTonnage = deliveryRow.Range[1, ShefflerWB.DeliveryTable.ListColumns["Тип ТС, тонн"].Index].Text;
                    double tonnage = double.TryParse(carTonnage, out double ton) ? ton : 0;
                    delivery.Truck = new Truck() { ProviderCompany = shippingCompany, Tonnage = tonnage };

                    string costStr = deliveryRow.Range[1, ShefflerWB.DeliveryTable.ListColumns["Стоимость доставки"].Index].Value.ToString();
                    delivery.Cost = decimal.TryParse(costStr, out decimal cost) ? cost : 0;
                    deliveries.Add(delivery);
                    //Компания
                    //Деловые линии
                }
            }
            return deliveries;
        }


        /// <summary>
        /// Кнопка сохранить маршруты
        /// </summary>
        public void SaveRoute()
        {
            List<Order> orders = GetOrdersFromTable();
            List<Delivery> deliveries = CreateDeliveries(orders);
            deliveries.ForEach(a => a.SaveRoute());
            PrintChanges(deliveries);
        }



        /// <summary>
        /// Прменять список доставок для списка заказов
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        private List<Delivery> EditDeliveres(List<Order> orders)
        {
            List<Delivery> deliveries = new List<Delivery>();
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
                        deliveries.Add(delivery);
                    }
                }
            }
            return deliveries;
        }
        private List<Delivery> CreateDeliveries(List<Order> orders)
        {

            List<Delivery> deliveries = new List<Delivery>();
            List<int> deliveryNumbers = (from o in orders
                                         select o.DeliveryNumber).Distinct().ToList();
            // По каждой доставке разделить список заказов 
            for (int i = 0; i < deliveryNumbers.Count; i++)
            {
                List<Order> orderList = orders.FindAll(
                            o => o.DeliveryNumber == deliveryNumbers[i]).ToList().OrderBy(
                                                            x => x.PointNumber).ToList();

                if (orderList.Count > 0)
                {

                    Delivery delivery = new Delivery();
                    int deliveryNumber = i + 1;
                    delivery.Number = deliveryNumber;
                    orderList.ForEach(x => x.DeliveryNumber = deliveryNumber);

                    string[] idCustomers = (from o in orderList
                                            orderby o.PointNumber
                                            select o.DeliveryPoint.IdCustomer
                                             ).Distinct().ToArray();
                    delivery.Orders = orderList;
                    //Расставить приоритеты точек 
                    for (int ixMap = 0; ixMap < idCustomers.Length; ixMap++)
                    {
                        orderList = delivery.Orders.FindAll(x => x.DeliveryPoint.IdCustomer == idCustomers[ixMap]);
                        int pnum = ixMap + 1;
                        foreach (Order orderID in orderList)
                        {
                            orderID.PointNumber = pnum;
                            DeliveryPoint dp = orderID.DeliveryPoint;
                            dp.PriorityPoint = pnum;
                            dp.PriorityRoute = deliveryNumber;

                            orderID.DeliveryPoint = dp;
                        }
                        orderList.Clear();
                    }
                    deliveries.Add(delivery);
                }


            }
            return deliveries;
        }
        /// <summary>
        /// Изменить доставку
        /// </summary>
        /// <param name="ordersCurrentDelivery"></param>
        /// <returns></returns>
        private Delivery EditDelivery(List<Order> ordersCurrentDelivery)
        {
            Delivery delivery = new Delivery();
            delivery.Orders = ordersCurrentDelivery;

            int idRoute = FindRoute(delivery.MapDelivery);
            if (idRoute == 0)
            {
                // Добавить маршрут 
                idRoute = new ShefflerWB().CreateRoute(delivery.MapDelivery);
            }
            List<DeliveryPoint> pointMap = ShefflerWB.RoutesList;
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
        private int FindRoute(List<DeliveryPoint> map)
        {
            //список id маршрутов
            List<int> uRoutes = (from p in ShefflerWB.RoutesList
                                 select p.Id).Distinct().ToList();

            for (int i = 0; i < uRoutes.Count; i++)
            {
                int idRoute = uRoutes[i];
                bool hasRoute = true;
                foreach (DeliveryPoint point in map)
                {
                    List<DeliveryPoint> routesVariants = ShefflerWB.RoutesList.FindAll(
                                 x => x.Id == idRoute &&
                                 x.IdCustomer == point.IdCustomer).ToList();
                    if (routesVariants.Count == 0)
                    {
                        hasRoute = false;
                        break;
                    }
                }
                if (hasRoute)
                { return idRoute; }
            }
            return 0;
        }

        /// <summary>
        /// Проверка нового маршрута и добавление по запросу пользователя
        /// </summary>
        /// <param name="order"></param>
        private void CheckAndAddNewRoute(Order order)
        {
            if (string.IsNullOrWhiteSpace(order?.Customer?.Id) || Delivery.CheckCustomerRoute(order.Customer.Id))
            {
                if (MessageBox.Show("Добвить маршрут?",
                                    "Маршрут с клиетном не найден!",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Warning)
                    == DialogResult.Yes)
                {
                    ShefflerWB.RoutesTable.ListRows.Add();
                    ListRow RouteRow = ShefflerWB.RoutesTable.ListRows[ShefflerWB.RoutesTable.ListRows.Count - 1];
                    ShefflerWB.RoutesList = null;  // Чтобы свойство обновилось;
                    RouteRow.Range[1, ShefflerWB.RoutesTable.ListColumns["Получатель материала"].Index].Value = order.Customer.Id;
                    try
                    {
                        ShefflerWB.RoutesSheet.Activate();
                        RouteRow.Range.Select();
                    }
                    catch (System.Exception ex)
                    { Debug.WriteLine(ex.Message); }
                }
                return;
            }
        }

        /// <summary>
        /// Собрать доставки из актуального диапазона таблицы Отгрузка
        /// </summary>
        /// <returns></returns>
        public List<Delivery> GetDeliveriesFromTotalSheet()
        {
            List<Delivery> deliveries = new List<Delivery>();
            Range total =  ShefflerWB.TotalTable.DataBodyRange ; //GetCurrentTotalRange();
            List<DeliveryPoint> points = ShefflerWB.RoutesList;
            if (total == null) return deliveries;

            for (int i = 0; i < total.Rows.Count; i++)
            {
                string numDelivery = total.Cells[i, ShefflerWB.TotalTable.ListColumns["№ Доставки"].Index].Text;
                int numD = int.TryParse(numDelivery, out int numDel) ? numDel : 0;
                if (numD == 0) continue;
                Delivery delivery = deliveries.Find(x => x.Number == numD);
                if (delivery == null)
                {
                    delivery = new Delivery
                    {
                        Number = numD,
                        Truck = new Truck
                        {
                            ProviderCompany = new Provider()
                        }
                    };
                    string providerName = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Перевозчик"].Index].Text;
                    delivery.Truck.ProviderCompany.Name = providerName;
                    delivery.Time = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Время подачи ТС"].Index].Text;
                    string tonn = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Тип ТС, тонн"].Index].Text;
                    delivery.Truck.Tonnage = double.TryParse(tonn, out double ton) ? ton : 0;
                    string costDelivery = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Стоимость доставки"].Index].Value.ToString();
                    delivery.Cost = decimal.TryParse(costDelivery, out decimal cd) ? cd : 0;

                    string id = total.Cells[i, ShefflerWB.TotalTable.ListColumns["ID перевозчика"].Index].Text;
                    string curNumber = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Номер,марка"].Index].Text;
                    string phone = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Телефон водителя"].Index].Text;
                    string fio = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Водитель (ФИО)"].Index].Text;

                    if (string.IsNullOrEmpty(id)) id = ShefflerWB.GetProviderId(providerName);

                    Driver driver = new Driver() 
                    { 
                        Id = id,
                        CarNumber = curNumber,
                        Name = fio,
                        Phone = phone

                    };
                    delivery.Driver = driver;
                    total.Cells[i, ShefflerWB.TotalTable.ListColumns["ID перевозчика"].Index].Value = driver.Id;

                    deliveries.Add(delivery);
                }
                string ID = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Номер поставки"].Index].Text;
                if (ID != "")
                {
                    Order order = new Order
                    {
                        Id = ID
                    };
                    string cost = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Стоимость поставки"].Index].text;
                    order.Cost = double.TryParse(cost, out double ct) ? ct : 0;

                    string customerId = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Номер грузополучателя"].Index].Text;
                    order.Customer = new Customer(customerId)
                    {
                        Name = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Грузополучатель"].Index].text
                    };
                    order.TransportationUnit = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Номер накладной"].Index].Text;
                    DeliveryPoint point = new DeliveryPoint
                    {
                        City = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Город"].Index].Text
                    };
                    ;
                    order.DeliveryPoint = point;
                    string palletCount = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Кол-во паллет"].Index].Text;
                    order.PalletsCount = int.TryParse(palletCount, out int countPalets) ? countPalets : 0;
                    string nom = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Порядок выгрузки"].Index].Text;
                    order.PointNumber = int.TryParse(nom, out int nd) ? nd : 0;

                    string weightBr = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Брутто вес"].Index].Text;
                    order.WeightBrutto = double.TryParse(weightBr, out double wb) ? wb : 0;

                    string weightNt = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Нетто вес"].Index].Text;
                    order.WeightNetto = double.TryParse(weightNt, out double wn) ? wn : 0;
                    order.RouteCity = total.Cells[i, ShefflerWB.TotalTable.ListColumns["Направление"].Index].Text;
                    delivery.Orders.Add(order);
                }
            }
            return deliveries;
        }

        /// <summary>
        /// Создать файл отгрузки для провайдера
        /// </summary>
        /// <param name="delivery"></param>
        /// <returns></returns>
        private string GenerateAttachmentFile(List<Delivery> deliveries, string name)
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
                                "Время подачи ТС",
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
                                "Стоимость доставки" };
            for (int i = 1; i <= headers.Length; i++)
            {
                sh.Cells[1, i].Value = headers[i - 1];
            }
            int row = 2;
            for (int ixDelivery = 0; ixDelivery < deliveries.Count; ixDelivery++)
            {
                Delivery delivery = deliveries[ixDelivery];
                string providerName = delivery.Truck.ProviderCompany.Name;
                if (string.IsNullOrWhiteSpace(providerName)) continue;
                sh.Cells[row, 1].Value = delivery.Driver.Id;
                sh.Cells[row, 7].Value = delivery.Time;
                sh.Cells[row, 19].Value = delivery.Cost;
                sh.Cells[row, 4].Value = delivery.Driver.Name;
                sh.Cells[row, 5].Value = delivery.Driver.CarNumber;
                sh.Cells[row, 6].Value = delivery.Driver.Phone;

                for (int i = 0; i < delivery.Orders.Count; i++)
                {
                    Range rowColor = sh.Range[sh.Cells[row, 1], sh.Cells[row, headers.Length]];
                    Order order = delivery.Orders[i];
                    if (ixDelivery % 2 == 0)
                    {
                        rowColor.Interior.Color = System.Drawing.Color.FromArgb(228, 234, 245);
                    }
                    else
                    {
                        rowColor.Interior.Color = System.Drawing.Color.FromArgb(252, 253, 255);
                    }
                    sh.Cells[row, 2].Value = providerName;
                    sh.Cells[row, 3].Value = delivery.Truck.Tonnage;

                    

                    sh.Cells[row, 8].Value = order.DeliveryPoint.City;
                    sh.Cells[row, 9].Value = order.RouteCity;

                    sh.Cells[row, 10].Value = order.PointNumber;
                    sh.Cells[row, 11].Value = order.Customer.Id;
                    sh.Cells[row, 12].Value = order.TransportationUnit;
                    sh.Cells[row, 13].Value = order.Id;
                    sh.Cells[row, 14].Value = order.Customer.Name ?? "";
                    sh.Cells[row, 15].Value = order.WeightBrutto;
                    sh.Cells[row, 16].Value = order.WeightNetto;
                    sh.Cells[row, 17].Value = order.PalletsCount;
                    sh.Cells[row, 18].Value = order.Cost;
                    row++;
                }
            }
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
            string folder = Globals.ThisWorkbook.Path + "\\ShippingOrders";

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
            string folder = Globals.ThisWorkbook.Path + "\\ShippingOrders";

            if (Directory.Exists(folder))
            {
                string[] files = Directory.GetFiles(folder);
                foreach (string file in files)
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        Debug.WriteLine(folder);
                    }
                }
            }
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