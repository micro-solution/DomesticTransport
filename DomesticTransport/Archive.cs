using DomesticTransport.Forms;
using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DomesticTransport
{
    class Archive
    {
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
        static List<string> _ordersId;

        public Archive() { }

        /// <summary>
        /// Перенести на Лист Архив
        /// </summary>
        public static void LoadToArhive()
        {
            XLTable table = new XLTable() { ListTable = ShefflerWB.TotalTable };
            List<Delivery> deliveries = GetAllDeliveries(table);
            if (!CheckArchive(deliveries))
            {//Проверить повторение заказов
                CpopyTotalPastArchive();
            }
            else
            {
                PrintArchive(deliveries);
            }
        }

        /// <summary>
        /// Проверить наличие в архиве
        /// </summary>
        /// <param name="deliveries"></param>
        /// <returns></returns>
        static bool CheckArchive(List<Delivery> deliveries)
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
        static bool CheckDelivery(Delivery delivery)
        {
            bool chk = false;
            ListObject archiveTable = ShefflerWB.ArchiveTable;
            foreach (string idOrder in OrdersId)
            {
                // string idOrder = archiveRow.Range[1, archiveTable.ListColumns["Номер поставки"].Index].Text;
                //idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                chk = delivery.Orders.Find(a => a.Id == idOrder) != null;
                if (chk) break;
            }
            return chk;
        }

        //Скопировать все вставить в архив
        static void CpopyTotalPastArchive()
        {
            ShefflerWB.TotalTable.DataBodyRange.Copy();
            ListObject arh = ShefflerWB.ArchiveTable;
            Range rng = arh.ListRows[arh.ListRows.Count].Range[1, 1];
            rng.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
        }
        static void ClearArchive()
        {

        }

        private static void PrintArchive(List<Delivery> deliveries)
        {
            XLTable tableArchive = new XLTable();
            tableArchive.ListTable = ShefflerWB.ArchiveTable;

            bool chk = false;
            foreach (Delivery delivery in deliveries)
            {

                chk = CheckDelivery(delivery);
                if (chk)
                {
                    //delivery 
                    for (int i = 0; i < delivery.Orders.Count; i++)
                    {
                        ShefflerWB.ArchiveTable.ListRows.Add();
                        tableArchive.CurrentRowRange = ShefflerWB.ArchiveTable.ListRows[
                                            ShefflerWB.ArchiveTable.ListRows.Count - 1].Range;
                        if (i == 0)  PrintDeliveryArchiveRow(delivery, tableArchive);

                        Order order = delivery.Orders[i];
                        PrintArchiveRow(order, tableArchive);
                    }
                }
            }
        }

        private static void PrintDeliveryArchiveRow(Delivery delivery, XLTable tableArchive)
        {
            //delivery. 
            tableArchive.SetValue("№ Доставки", delivery.Number);
            tableArchive.SetValue("Время подачи ТС", delivery.Time);
            tableArchive.SetValue("ID перевозчика", delivery.Driver.Id);
            //  tableArchive.SetValue("Дата отгрузки", delivery.DateDelivery);
            //tableArchive.SetValue("Грузополучатель", delivery.Driver.Name);
            tableArchive.SetValue("Стоимость доставки", delivery.Cost);
            if (!string.IsNullOrEmpty(delivery.Driver.Id))
            {
                tableArchive.SetValue("ID перевозчика", delivery.Driver.Id);
                tableArchive.SetValue("Водитель (ФИО)", delivery.Driver.Name);
                tableArchive.SetValue("Телефон водителя", delivery.Driver.Phone);
                tableArchive.SetValue("Номер,марка", delivery.Driver.CarNumber);
            }
        }
        private static void PrintArchiveRow(Order order, XLTable xlTable)
        {
            xlTable.SetValue("№ Доставки", order.DeliveryNumber);
            xlTable.SetValue("Номер поставки", order.Id);
            xlTable.SetValue("Дата отгрузки", order.DeliveryNumber);
            xlTable.SetValue("Порядок выгрузки", order.PointNumber);
            xlTable.SetValue("Грузополучатель", order.Customer.Name);
            xlTable.SetValue("Номер грузополучателя", order.Customer.Id);
            xlTable.SetValue("Номер накладной", order.TransportationUnit);
            xlTable.SetValue("Брутто вес", order.WeightBrutto);
            xlTable.SetValue("Нетто вес", order.WeightNetto);
            xlTable.SetValue("Стоимость поставки", order.Cost);
            xlTable.SetValue("Кол - во паллет", order.PalletsCount);
            xlTable.SetValue("Направление", order.RouteCity);
        }




        public static void UnoadFromArhive()
        {
            new UnloadArchive().ShowDialog();
        }

        ///=======================================================
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static List<Delivery> GetAllDeliveries(XLTable table)
        {
            List<Order> orders = new List<Order>();
            List<Delivery> deliveries = new List<Delivery>();

            foreach (ListRow row in ShefflerWB.TotalTable.ListRows)
            {
                table.CurrentRowRange = row.Range;
                Order order = GetOrdersFromTotalRow(table);
                if (order != null) orders.Add(order);
                Delivery delivery = GetDeliveryFromTotalRow(table);
                if (delivery != null) deliveries.Add(delivery);
            }
            SortingOrders(orders, deliveries);
            return deliveries;
        }

        private static List<Delivery> SortingOrders(List<Order> orders, List<Delivery> deliveries)
        {
            foreach (Delivery delivery in deliveries)
            {
                List<Order> ordersDelivery = orders.FindAll(a =>
                                             a.DeliveryNumber == delivery.Number &&
                                             a.DateDelivery == delivery.DateDelivery);
                if (ordersDelivery != null)
                {
                    delivery.Orders = ordersDelivery;
                    //ordersDelivery.ForEach(x => orders.Remove(x));
                }
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
            Customer customer = new Customer(customerId);
            customer.Name = nameCustomer;
            order.Customer = customer;

            order.TransportationUnit = xlTable.GetValueString("Номер накладной");
            order.WeightBrutto = xlTable.GetValueDouble("Брутто вес");
            order.WeightNetto = xlTable.GetValueDouble("Нетто вес");
            order.Cost = xlTable.GetValueDouble("Стоимость поставки");
            order.PalletsCount = xlTable.GetValueInt("Кол-во паллет");
            order.RouteCity = xlTable.GetValueString("Направление");

            return order;
        }
        private static Delivery GetDeliveryFromTotalRow(XLTable xlTable)
        {
            Delivery delivery = new Delivery();
            delivery.DateDelivery = xlTable.GetValueString("Дата отгрузки");
            delivery.Number = xlTable.GetValueInt("№ Доставки");
            delivery.Time = xlTable.GetValueString("Время подачи ТС");
            delivery.Cost = xlTable.GetValueDecimal("Стоимость доставки");
            if (string.IsNullOrWhiteSpace(delivery.DateDelivery) ||
                                            delivery.Number == 0 ||
                                             delivery.Cost == 0
                                            ) return null;

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
