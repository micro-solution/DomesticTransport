using DomesticTransport.Forms;
using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;

namespace DomesticTransport
{
    internal class Archive
    {
        /// <summary>
        /// Список id товаров в Архиве
        /// </summary>
        private static List<string> OrdersId
        {
            get
            {
                if (_ordersId == null || _ordersId.Count == 0)
                {
                    _ordersId = new List<string>();
                    foreach (ListRow archiveRow in ShefflerWB.ArchiveTable.ListRows)
                    {
                        string idOrder = archiveRow.Range[1,
                          ShefflerWB.ArchiveTable.ListColumns["Номер поставки"].Index].Text;
                        if (string.IsNullOrWhiteSpace(idOrder)) continue;

                        idOrder = idOrder.Length < 10 ?
                                    new string('0', 10 - idOrder.Length) + idOrder : idOrder;

                        _ordersId.Add(idOrder);
                    }
                }
                return _ordersId;
            }
            set => _ordersId = value;
        }

        private static List<string> _ordersId;

        public Archive() { }

        /// <summary>
        /// Перенести на Лист Архив
        /// </summary>
        public static void LoadToArhive()
        {
            XLTable table = new XLTable() { ListTable = ShefflerWB.TotalTable };
            List<Delivery> deliveries = GetAllDeliveries(table);
            if (deliveries.Count == 0) return;
            if (!CheckArchive(deliveries)) //Проверить повторение заказов по Id
            {
                CpopyTotalPastArchive();   //Копипастить
            }
            else
            {
                PrintArchive(deliveries);  // Удалять старые если совпадают, печатать по строке
            }
            SortArchive();   //Сортировка
        }

        /// <summary>
        /// Проверить наличие в архиве
        /// </summary>
        /// <param name="deliveries"></param>
        /// <returns></returns>
        private static bool CheckArchive(List<Delivery> deliveries)
        {
            bool chk = false;
            OrdersId = null;
            foreach (Delivery delivery in deliveries)
            {
                chk = CheckDelivery(delivery);
                if (chk) break;
            }
            return chk;
        }

        /// <summary>
        /// Проверить все заказы доставки
        /// </summary>
        /// <param name="delivery"></param>
        /// <returns></returns>
        private static bool CheckDelivery(Delivery delivery)
        {
            bool chk = false;
            ListObject archiveTable = ShefflerWB.ArchiveTable;
            foreach (string idOrder in OrdersId)
            {
                chk = delivery.Orders.Find(a => a.Id == idOrder) != null;
                if (chk) break;
            }
            return chk;
        }

        /// <summary>
        /// Сортировка архива
        /// </summary>
        private static void SortArchive()
        {
            Range table = ShefflerWB.ArchiveTable.Range;
            Range col1 = table.Columns[ShefflerWB.ArchiveTable.ListColumns["Дата отгрузки"].Index];
            Range col2 = table.Columns[ShefflerWB.ArchiveTable.ListColumns["№ Доставки"].Index];
            table.Sort(
                Key1: col1,
                Order1: XlSortOrder.xlAscending,
                Key2: col2,
                Order2: XlSortOrder.xlAscending,
                OrderCustom: Type.Missing, MatchCase: Type.Missing,
                Header: XlYesNoGuess.xlYes, Orientation: XlSortOrientation.xlSortColumns);
        }

        //Скопировать все вставить в архив
        private static void CpopyTotalPastArchive()
        {
            ShefflerWB.TotalTable.DataBodyRange.Copy();
            XLTable archive = new XLTable() { ListTable = ShefflerWB.ArchiveTable };
            Range rng = archive.GetLastRow();
            rng.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
        }


        public static void ClearArchive()
        {
            ShefflerWB.ArchiveTable.DataBodyRange.Clear();
        }

        /// <summary>
        /// Вывод доставок в архив
        /// </summary>
        /// <param name="deliveries"></param>
        private static void PrintArchive(List<Delivery> deliveries)
        {
            XLTable tableArchive = new XLTable { ListTable = ShefflerWB.ArchiveTable };

            foreach (Delivery delivery in deliveries)
            {
                if (CheckDelivery(delivery))
                {
                    DeleteChangedDelivery(delivery, tableArchive);
                }
                for (int i = 0; i < delivery.Orders.Count; i++)
                {
                    tableArchive.SetCurrentRow();
                    if (i == 0) PrintArchiveDelivery(delivery, tableArchive);

                    Order order = delivery.Orders[i];
                    PrintArchiveOrder(order, tableArchive);
                }
            }
        }

        /// <summary>
        /// Удалять строки заказов доставки
        /// </summary>
        /// <param name="date"></param>
        /// <param name="number"></param>
        /// <param name="table"></param>
        private static void DeleteChangedDelivery(Delivery delivery, XLTable table)
        {
            ListObject archive = table.ListTable;
            for (int i = archive.ListRows.Count; i > 0; i--)
            {
                table.CurrentRowRange = archive.ListRows[i].Range;
                string IdOrderRow = table.GetValueString("Номер поставки");
                IdOrderRow = IdOrderRow.Length < 10 ?
                                   new string('0', 10 - IdOrderRow.Length) + IdOrderRow : IdOrderRow;
                ListRow row = table.ListTable.ListRows[i];
                table.CurrentRowRange = row.Range;
                Order order = delivery.Orders.Find(o => o.Id == IdOrderRow);
                if (order != null) row.Range.EntireRow.Delete();
            }
        }

        /// <summary>
        /// Удалять строки всех заказов, отправленных на указанную дату и ранее 
        /// </summary>
        /// <param name="date"></param>
        /// <param name="table"></param>
        private static void DeleteBefore(DateTime date, XLTable table)
        {
            ListObject archive = table.ListTable;
            for (int i = archive.ListRows.Count; i > 0; i--)
            {
                ListRow row = archive.ListRows[i];
                table.CurrentRowRange = row.Range;
                string currentOrderDate = table.GetValueString("Дата отгрузки");
                DateTime orderDate = DateTime.TryParse(currentOrderDate, out DateTime currentDate) ? currentDate : DateTime.MaxValue;
                if (orderDate <= date)
                {
                    row.Range.EntireRow.Delete();
                }
            }
        }


        /// <summary>
        /// Вывести  доставок в строку таблицы
        /// </summary>
        /// <param name="delivery"></param>
        /// <param name="tableArchive"></param>
        private static void PrintArchiveDelivery(Delivery delivery, XLTable tableArchive)
        {
            //delivery. 
            tableArchive.SetValue("№ Доставки", delivery.Number);
            tableArchive.SetValue("Время подачи ТС", delivery.Time);
            tableArchive.SetValue("ID перевозчика", delivery.Driver?.Id);
            tableArchive.SetValue("Дата отгрузки", delivery.DateDelivery);
            tableArchive.SetValue("Перевозчик", delivery.Truck.ProviderCompany.Name);
            tableArchive.SetValue("Тип ТС, тонн", delivery.Truck.Tonnage);
            tableArchive.SetValue("Стоимость доставки", delivery.Cost);
            if (!string.IsNullOrEmpty(delivery.Driver?.Id))
            {
                tableArchive.SetValue("ID перевозчика", delivery.Driver.Id);
                tableArchive.SetValue("Водитель (ФИО)", delivery.Driver.Name);
                tableArchive.SetValue("Телефон водителя", delivery.Driver.Phone);
                tableArchive.SetValue("Номер,марка", delivery.Driver.CarNumber);
            }
        }

        /// <summary>
        ///    Вывести заказ в строку таблицы
        /// </summary>
        /// <param name="order"></param>
        /// <param name="xlTable"></param>
        private static void PrintArchiveOrder(Order order, XLTable xlTable)
        {
            xlTable.SetValue("№ Доставки", order.DeliveryNumber);
            xlTable.SetValue("Номер поставки", order.Id);
            xlTable.SetValue("Дата отгрузки", order.DateDelivery);
            xlTable.SetValue("Порядок выгрузки", order.PointNumber);
            xlTable.SetValue("Грузополучатель", order.Customer.Name);
            xlTable.SetValue("Номер грузополучателя", order.Customer.Id);
            xlTable.SetValue("Номер накладной", order.TransportationUnit);
            xlTable.SetValue("Брутто вес", order.WeightBrutto);
            xlTable.SetValue("Нетто вес", order.WeightNetto);
            xlTable.SetValue("Стоимость поставки", order.Cost);
            xlTable.SetValue("Кол-во паллет", order.PalletsCount);
            xlTable.SetValue("Направление", order.RouteCity);
            xlTable.SetValue("Город", order.DeliveryPoint.City);
        }

        /// <summary>
        /// Перенос текущего архива в таблицы Shepments and TransportTable
        /// </summary>
        public static void ToTransportTableAndShepments()
        {

            XLTable tableArchive = new XLTable
            {
                ListTable = ShefflerWB.ArchiveTable
            };

            List<Delivery> deliveries = GetAllDeliveries(tableArchive);

            TransportTable transportTable = new TransportTable();
            transportTable.ImportDeliveryes(deliveries);
            transportTable.SaveAndClose();

            ShipmentsTable shipmentsTable = new ShipmentsTable();
            shipmentsTable.ImportDeliveryes(deliveries);
            shipmentsTable.SaveAndClose();

            DateTime dateMax = DateTime.Today;
            dateMax = dateMax.AddDays(-(double)dateMax.DayOfWeek);
            DeleteBefore(dateMax, tableArchive);

            System.Windows.Forms.MessageBox.Show("Архив перенесен", "Операция выполнена", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

            return;
        }


        public static void UnoadFromArhive()
        {
            new UnloadArchive().ShowDialog();

        }

        ///=======================================================


        /// <summary>
        /// Собрать доставки из таблицы
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static List<Delivery> GetAllDeliveries(XLTable table)
        {
            List<Order> orders = new List<Order>();
            List<Delivery> deliveries = new List<Delivery>();
            Delivery delivery = new Delivery();
            foreach (ListRow row in table.ListTable.ListRows)
            {
                table.CurrentRowRange = row.Range;
                Delivery deliveryRow = GetDeliveryFromTotalRow(table);
                
                if (deliveryRow != null) 
                {
                    delivery = deliveryRow;
                    deliveries.Add(delivery);
                }
                Order order = GetOrdersFromTotalRow(table);
                if (order != null) delivery.Orders.Add(order);
            }
            return deliveries;
        }

        private static Order GetOrdersFromTotalRow(XLTable xlTable)
        {
            Order order = new Order();
            string idOrder = xlTable.GetValueString("Номер поставки");
            if (string.IsNullOrWhiteSpace(idOrder)) return null;
            order.Id = idOrder;
            order.DeliveryNumber = xlTable.GetValueInt("№ Доставки");
            order.DateDelivery = xlTable.GetValueString("Дата отгрузки");
            order.PointNumber = xlTable.GetValueInt("Порядок выгрузки");
            string customerId = xlTable.GetValueString("Номер грузополучателя");
            string nameCustomer = xlTable.GetValueString("Грузополучатель");
            Customer customer = new Customer(customerId)
            {
                Name = nameCustomer
            };
            order.Customer = customer;

            order.TransportationUnit = xlTable.GetValueString("Номер накладной");
            order.WeightBrutto = xlTable.GetValueDouble("Брутто вес");
            order.WeightNetto = xlTable.GetValueDouble("Нетто вес");
            order.Cost = xlTable.GetValueDouble("Стоимость поставки");
            order.PalletsCount = xlTable.GetValueInt("Кол-во паллет");
            order.RouteCity = xlTable.GetValueString("Направление");
            string city = xlTable.GetValueString("Город");
            DeliveryPoint point = new DeliveryPoint() { City = city };
            order.DeliveryPoint = point;
            return order;
        }
        private static Delivery GetDeliveryFromTotalRow(XLTable xlTable)
        {
            Delivery delivery = new Delivery
            {
                DateDelivery = xlTable.GetValueString("Дата отгрузки"),
                Number = xlTable.GetValueInt("№ Доставки"),
                Time = xlTable.GetValueString("Время подачи ТС"),
                Cost = xlTable.GetValueDecimal("Стоимость доставки")
            };
            string providerName = xlTable.GetValueString("Перевозчик");
            if (string.IsNullOrWhiteSpace(delivery.DateDelivery) ||
                                            delivery.Number == 0 ||
                                            string.IsNullOrWhiteSpace(providerName)
                                            ) return null;
            Truck truck = new Truck
            {
                Tonnage = xlTable.GetValueDouble("Тип ТС, тонн")
            };
            truck.ProviderCompany.Name = providerName;
            delivery.Truck = truck;

            string id = xlTable.GetValueString("ID перевозчика");
            string curNumber = xlTable.GetValueString("Номер,марка");
            string phone = xlTable.GetValueString("Телефон водителя");
            string fio = xlTable.GetValueString("Водитель (ФИО)");
            if (!string.IsNullOrWhiteSpace(id))
            {
                Driver driver = new Driver()
                {
                    Id = id,
                    CarNumber = curNumber,
                    Name = fio,
                    Phone = phone
                };
                delivery.Driver = driver;
            }

            return delivery;
        }
    }
}
