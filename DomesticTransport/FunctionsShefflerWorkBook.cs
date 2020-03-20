using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DomesticTransport
{

    /// <summary>
    /// Действия с текущей книгой
    /// </summary>
    class ShefflerWorkBook
    {

        private List<TruckRate> RateList
        {
            get
            {
                if (_rateList == null)
                {
                    _rateList = GetTruckRateList();
                }

                return _rateList;
            }
        }
        private List<TruckRate> _rateList;

        public string DateDelivery
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_dateDelivery))
                {
                    Worksheet sheetDelidery = GetSheet("Delivery");
                    Range dateCell = sheetDelidery.Range["DateDelivery"];
                    if (dateCell != null)
                    {
                        _dateDelivery = dateCell.Text;

                        DateTime date = DateTime.Parse(_dateDelivery);
                        _dateDelivery = date > DateTime.MinValue ? date.ToShortDateString() : "";
                    }
                }
                return _dateDelivery;
            }
        }
        string _dateDelivery;

        public List<DeliveryPoint> RoutesTable
        {
            get
            {
                if (_routes == null)
                {
                    _routes = new List<DeliveryPoint>();
                    Worksheet sheetRoute = GetSheet("Routes");
                    ListObject TableRoutes = sheetRoute?.ListObjects["TableRoutes"];
                    if (TableRoutes != null)
                    {
                        foreach (ListRow row in TableRoutes.ListRows)
                        {
                            Debug.WriteLine(row.Range.Row.ToString());
                            if (row.Range[1, 1].Value == null ||
                                row.Range[1, 2].Value == null ||
                                row.Range[1, 3].Value == null ||
                                row.Range[1, 5].Value == null ||
                                row.Range[1, 9].Value == null) continue;
                            DeliveryPoint route = new DeliveryPoint()
                            {
                                IdRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Id route"].Index].Text, out int id) ? id : 0,
                                PriorityRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority route"].Index].Text.ToString(), out int prioritRoute) ? prioritRoute : 0,
                                PriorityPoint = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority point"].Index].Text.ToString(), out int prioritPoint) ? prioritPoint : 0,
                                IdCustomer = row.Range[1, TableRoutes.ListColumns["Получатель материала"].Index].Text,
                                City = row.Range[1, TableRoutes.ListColumns["City"].Index].Text,
                                CityLongName = row.Range[1, TableRoutes.ListColumns["Город"].Index].Text,
                                Customer = row.Range[1, TableRoutes.ListColumns["Клиент"].Index].Text,
                                CustomerNumber = row.Range[1, TableRoutes.ListColumns["Номер клиента"].Index].Text,
                                Route = row.Range[1, TableRoutes.ListColumns["Маршрут"].Index].Text,
                                RouteName = row.Range[1, TableRoutes.ListColumns["Направление"].Index].Text

                            };
                            _routes.Add(route);
                        }
                    }
                }
                _routes = _routes.OrderBy(x => x.IdRoute).ThenBy(
                                      y => y.PriorityRoute).ThenBy(y => y.PriorityPoint).ToList();
                return _routes;
            }
            set
            {
                _routes = value;
            }
        }
        List<DeliveryPoint> _routes;

        public object DataTime { get; private set; }





        internal Truck GetTruck(double totalWeight, List<DeliveryPoint> mapDelivery)
        {
            if (mapDelivery.Count <= 0 || totalWeight <= 0) return null;

            List<TruckRate> rates = RateList; //Вся таблица
            Truck truck = null;
            List<TruckRate> rateVariants = new List<TruckRate>();


            rateVariants = rates.FindAll(r =>
                                        r.City == mapDelivery[0].City &&
                                        (r.Tonnage + r.Tonnage * 0.1) * 1000 > totalWeight
                                        ).ToList();
            //  Выборка по Городу в первой точке                                                                              
            // double MaxTottage = truckSelect.Tonnage + truckSelect.Tonnage * 0.1;    Допустимый перегруз       

            if (rateVariants.Count > 0)
            {
                for (int rateIx = 0; rateIx < rateVariants.Count; rateIx++)
                {

                    TruckRate variantRate = rateVariants[rateIx];
                    variantRate.TotalDeliveryCost = rateVariants[rateIx].PriceFirstPoint;
                    for (int pointNumber = 1; pointNumber < mapDelivery.Count; pointNumber++)
                    {
                        TruckRate addPointRate =
                            rates.Where(x => x.Company == variantRate.Company &&
                                                x.Tonnage == variantRate.Tonnage &&
                                                x.City == mapDelivery[pointNumber].City).First();
                        if (addPointRate.PriceAddPoint > 0)
                            variantRate.TotalDeliveryCost += addPointRate.PriceAddPoint;
                    }
                    rateVariants[rateIx] = variantRate;
                }


                rateVariants = rateVariants.OrderBy(r => r.TotalDeliveryCost).ToList();
                truck = new Truck(rateVariants.First());
            }
            return truck;
        }



        /// <summary>
        /// Вернуть лист по имени
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet GetSheet(string sheetName)
        {
            try
            {
                Worksheet sh = Globals.ThisWorkbook.Sheets[sheetName];
                return sh;
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось получить лист \"{sheetName}\"");
            }

        }

        /// <summary>
        /// Получить таблицу цен перевозчиков
        /// </summary>
        /// <returns></returns>
        internal ListObject GetRateList()
        {
            Worksheet sheetRoute = GetSheet("Rate");
            return sheetRoute?.ListObjects["PriceDelivery"];
        }


        /// <summary>
        /// Получить вес список цен перевозчиков в формате списка         
        /// </summary>
        /// <returns></returns>
        internal List<TruckRate> GetTruckRateList()
        {
            List<TruckRate> ListRate = new List<TruckRate>();
            ListObject rateTable = GetRateList();
            foreach (ListRow row in rateTable.ListRows)
            {
                // double tonnage = row.Range[1, rateTable.ListColumns["tonnage, t"].Index].Value ?? 0;
                string valTonnage = row.Range[1, rateTable.ListColumns["tonnage, t"].Index].Text;
                double tonnage = double.TryParse(valTonnage, out double t) ? t : 0;

                string valCity = row.Range[1, rateTable.ListColumns["City"].Index].Text;
                valCity = valCity.Trim();

                string valCompany = row.Range[1, rateTable.ListColumns["Company"].Index].Text;
                valCompany = valCompany.Trim();


                if (tonnage > 0 && !string.IsNullOrWhiteSpace(valCity))
                {

                    string strPrice = row.Range[1, rateTable.ListColumns["vehicle"].Index].Text;
                    int priceFirst = int.TryParse(strPrice, out int pf) ? pf : 0;
                    strPrice = row.Range[1, rateTable.ListColumns["add.point"].Index].Text;
                    int priceAdd = int.TryParse(strPrice, out int pa) ? pa : 0;
                    TruckRate rate = new TruckRate()
                    {
                        City = valCity,
                        Company = valCompany,
                        PriceFirstPoint = priceFirst,
                        PriceAddPoint = priceAdd,
                        PlaceShipment = row.Range[1, 1].Text,
                        PlaceDelivery = row.Range[1, 2].Text,
                        Tonnage = tonnage

                    };

                    ListRate.Add(rate);
                }
            }

            return ListRate;
        }

        internal int CreateRoute(List<Order> ordersCurrentDelivery)
        {
            Worksheet sheetRoutes = GetSheet("Routes");
            ListObject TableRoutes = sheetRoutes?.ListObjects["TableRoutes"];
            List<DeliveryPoint> pointMap = RoutesTable;

            DeliveryPoint LastPoint = RoutesTable.Last();
            int idRoute = LastPoint.IdRoute + 1;
            int priorityRoute = LastPoint.PriorityRoute + 1;
            //Поиск подходящего максимального приоритета
            foreach (Order ord in ordersCurrentDelivery)
            {
                string customerId = ord.Customer.Id;
                List<int> routes = (from p in pointMap
                                    where p.IdCustomer == customerId
                                    select p.PriorityRoute
                                     ).Distinct().ToList();
                int maxPriority = routes.Max();
                priorityRoute = maxPriority > priorityRoute ? maxPriority : priorityRoute;
            }

            int point = 0;

            foreach (Order order in ordersCurrentDelivery)
            {
                ListRow row = TableRoutes.ListRows[TableRoutes.ListRows.Count];
                TableRoutes.ListRows.Add();
                row.Range[1, TableRoutes.ListColumns["Id route"].Index].Value = idRoute;
                row.Range[1, TableRoutes.ListColumns["Priority route"].Index].Value = priorityRoute;
                row.Range[1, TableRoutes.ListColumns["Priority point"].Index].Value = ++point;
                row.Range[1, TableRoutes.ListColumns["Получатель материала"].Index].Value = order.Customer.Id;
                row.Range[1, TableRoutes.ListColumns["City"].Index].Value = order.DeliveryPoint.City;

                //поиск этого же Получателя в другой строке
                DeliveryPoint findPoint = pointMap.Find(x => x.IdCustomer == order.Customer.Id && x.CityLongName != "");
                if (!string.IsNullOrWhiteSpace(findPoint.CustomerNumber))
                {
                    row.Range[1, TableRoutes.ListColumns["Город"].Index].Value = findPoint.CityLongName;
                    row.Range[1, TableRoutes.ListColumns["Маршрут"].Index].Value = findPoint.Route;
                    row.Range[1, TableRoutes.ListColumns["Направление"].Index].Value = findPoint.RouteName;
                    row.Range[1, TableRoutes.ListColumns["Клиент"].Index].Value = findPoint.Customer;
                    row.Range[1, TableRoutes.ListColumns["Номер клиента"].Index].Value = findPoint.CustomerNumber;
                    row.Range[1, TableRoutes.ListColumns["Add"].Index].Value = "Auto";
                }
            }
            RoutesTable = null;
            return idRoute;
        }

       


        public Range GetCurrentShippingRange()
        {
            Worksheet TotalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
            ListObject totalTable = TotalSheet.ListObjects["TableTotal"];
            Range currentRng = null;
            string dateDelivery = DateDelivery;
            int columnDeliveryId = totalTable.ListColumns["Дата доставки"].Index;
            foreach (ListRow row in totalTable.ListRows)
            {
                string dateTable = row.Range[1, columnDeliveryId].Text;
                if (dateTable == dateDelivery)
                {
                    if (currentRng == null)
                    {
                        currentRng = row.Range;
                    }
                    else
                    {
                        currentRng = Globals.ThisWorkbook.Application.Union(currentRng, row.Range);
                    }
                }
            }
            return currentRng;
        }

    }
}
