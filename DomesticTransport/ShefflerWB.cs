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
    class ShefflerWB
    {
        public static Worksheet DeliverySheet
        {
            get
            {
                if (_deliverySheet == null)
                {
                    _deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
                }
                return _deliverySheet;
            }
        }
        private static Worksheet _deliverySheet;
        public static Worksheet TotalSheet
        {
            get
            {
                if (_totalSheet == null)
                {
                    _totalSheet = Globals.ThisWorkbook.Sheets["Отгрузка"];
                }
                return _totalSheet;
            }
        }
        private static Worksheet _totalSheet;

        public static Worksheet RateSheet
        {
            get
            {
                if (_rateSheet == null)
                {
                    _rateSheet = Globals.ThisWorkbook.Sheets["Rate"];
                }
                return _rateSheet;
            }
        }
        private static Worksheet _rateSheet;

        public static Worksheet RoutesSheet
        {
            get
            {
                if (_routesSheet == null)
                {
                    _routesSheet = Globals.ThisWorkbook.Sheets["Routes"];
                }
                return _routesSheet;
            }
        }
        private static Worksheet _routesSheet;

        public static ListObject OrdersTable
        {
            get
            {
                if (_ordersTable == null)
                {
                    _ordersTable = DeliverySheet.ListObjects["TableOrders"];
                }
                return _ordersTable;
            }
        }
        private static ListObject _ordersTable;
        public static ListObject DeliveryTable
        {
            get
            {
                if (_deliveryTable == null)
                {
                    _deliveryTable = DeliverySheet.ListObjects["TableCarrier"];
                }
                return _deliveryTable;
            }
        }
        private static ListObject _deliveryTable;

        public static ListObject TotalTable
        {
            get
            {
                if (_totalTable == null)
                {
                    _totalTable = TotalSheet.ListObjects["TableTotal"];
                }
                return _totalTable;
            }
        }
        private static ListObject _totalTable;

        public static ListObject RateTable
        {
            get
            {
                if (_rateTable == null)
                {
                    _rateTable = RateSheet.ListObjects["PriceDelivery"];
                }
                return _rateTable;
            }
        }
        private static ListObject _rateTable;

        public static ListObject RoutesTable
        {
            get
            {
                if (_routesTable == null)
                {
                    _routesTable = RoutesSheet.ListObjects["TableRoutes"];
                }
                return _routesTable;
            }
        }
        private static ListObject _routesTable;
        public static ListObject ProviderTable
        {
            get
            {
                if (_providerTable == null)
                {
                    _providerTable = RateSheet.ListObjects["ProviderTable"];
                }
                return _providerTable;
            }
        }
        private static ListObject _providerTable;



        /// <summary>
        /// Прайс
        /// </summary>
         private static List<TruckRate> RateList
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
        private static List<TruckRate> _rateList;

        /// <summary>
        /// Дата Доставки
        /// </summary>
        public static string DateDelivery
        {
            get
            {
                Range dateCell = DeliverySheet.Range["DateDelivery"];
                _dateDelivery = dateCell?.Text;
                if (string.IsNullOrWhiteSpace(_dateDelivery))
                {
                    if (dateCell != null)
                        if (string.IsNullOrWhiteSpace(_dateDelivery))
                        {
                            dateCell.Formula = "=Today()+1";
                            _dateDelivery = dateCell?.Text;
                        }
                        else
                        {
                            _dateDelivery = DateTime.Today.AddDays(1).ToShortDateString();
                        }
                }
                return _dateDelivery;
            }
        }
        static string _dateDelivery;

        internal Range GetRowOrderTotal(string idOrder)
        {
            Range range = null;
            foreach (ListRow row in TotalTable.ListRows)
            {
              string cell = row.Range[0, TotalTable.ListColumns["Номер поставки"].Index].Text;
                if ((!string.IsNullOrWhiteSpace(cell)) &&  idOrder.Contains(cell))
                {
                    range = row.Range;
                    break;
                }
            }
            return range;
        }






        /// <summary>
        /// Получить таблицу Маршрутов 
        /// </summary>
        public static List<DeliveryPoint> RoutesList
        {
            get
            {
                if (_routes == null)
                {
                    _routes = new List<DeliveryPoint>();
                    foreach (ListRow row in RoutesTable.ListRows)
                    {
                        Debug.WriteLine(row.Range.Row.ToString());
                        if (row.Range[1, 1].Value == null ||
                            row.Range[1, 2].Value == null ||
                            row.Range[1, 3].Value == null ||
                            row.Range[1, 5].Value == null ||
                            row.Range[1, 9].Value == null) continue;
                        DeliveryPoint route = new DeliveryPoint()
                        {
                            Id = int.TryParse(row.Range[1, RoutesTable.ListColumns["Id route"].Index].Text, out int id) ? id : 0,
                            PriorityRoute = int.TryParse(row.Range[1, RoutesTable.ListColumns["Priority route"].Index].Text.ToString(), out int prioritRoute) ? prioritRoute : 0,
                            PriorityPoint = int.TryParse(row.Range[1, RoutesTable.ListColumns["Priority point"].Index].Text.ToString(), out int prioritPoint) ? prioritPoint : 0,
                            IdCustomer = row.Range[1, RoutesTable.ListColumns["Получатель материала"].Index].Text,
                            City = row.Range[1, RoutesTable.ListColumns["City"].Index].Text,
                            CityLongName = row.Range[1, RoutesTable.ListColumns["Город"].Index].Text,
                            Customer = row.Range[1, RoutesTable.ListColumns["Клиент"].Index].Text,
                            CustomerNumber = row.Range[1, RoutesTable.ListColumns["Номер клиента"].Index].Text,
                            Route = row.Range[1, RoutesTable.ListColumns["Маршрут"].Index].Text,
                            RouteName = row.Range[1, RoutesTable.ListColumns["Направление"].Index].Text
                        };
                        _routes.Add(route);
                    }

                }
                _routes = _routes.OrderBy(x => x.Id).ThenBy(
                                      y => y.PriorityRoute).ThenBy(y => y.PriorityPoint).ToList();
                return _routes;
            }
            set
            {
                _routes = value;
            }
        }
        static List<DeliveryPoint> _routes;


        public  List<TruckRate> RateInternationalList
        {
            get
            {
                if (_RateInternationalList == null)
                {
                    
                    _RateInternationalList = new ShefflerWB().GetTruckRateInternational();
                }
                return _RateInternationalList;
            }
        }
        private static List<TruckRate> _RateInternationalList;

        public string[] CityList
        {
            get
            {
                if (_cityList == null)
                {
                    _cityList = (from LR in RateList
                                 select LR.City
                                 ).Distinct().ToArray();
                }
                return _cityList;
            }
        }
        private string[] _cityList;
        public static string[] InternationalCityList
        {
            get
            {
                if (_internationalCityList == null)
                {
                    List<TruckRate> rates = new ShefflerWB().RateInternationalList;
                    _internationalCityList = (from LR in rates
                                              select LR.City
                                 ).Distinct().ToArray();
                }
                return _internationalCityList;
            }
        }
        private static string[] _internationalCityList;

        /// <summary>
        /// Выбрать авто 
        /// </summary>
        /// <param name="totalWeight"></param>
        /// <param name="mapDelivery"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public Truck GetTruck(double totalWeight, List<DeliveryPoint> mapDelivery, string provider = "")
        {
            if (mapDelivery.Count <= 0 || totalWeight <= 0) return null;
            if (!CheckPoints(mapDelivery)) return null;  //Нет клиента

            Truck truck = null;
            List<TruckRate> rateVariants = new List<TruckRate>();
            double tonnageNeed = totalWeight / 1000 - 0.05;  /// 50kg Допустимый перегруз
            try
            {
                if (mapDelivery.FindAll(m => m.City == "MSK" || m.City == "MO").Count > 0)
                {
                    rateVariants = GetCostMskRoutes(tonnageNeed, mapDelivery); //Для Москвы и области  (первая точка с наибольшим приоритетом по таблице)
                }
                else
                {
                    bool isInternational = false;
                   
                    foreach (string city in InternationalCityList) // Nur - Sultan //Yerevan
                    {
                        string pointCity = mapDelivery[0].City ?? "";
                        if (pointCity.Contains(city))
                        {
                            isInternational = true;
                            break;
                        }
                    }
                   rateVariants = isInternational ?
                   // Для  LTL маршрутов расчет суммы за 100 кг веса + add.point
                   rateVariants = GetTruckRateInternational(totalWeight, mapDelivery) :
                   rateVariants = GetTruckRate(tonnageNeed, mapDelivery);
                }
            }
            catch
            {
                truck = new Truck()
                {
                    Cost = 0,
                    Tonnage = 0
                };
                //  truck.ShippingCompany.Name = "";
                return truck;
            }
           //RateList Вся таблица

            if (rateVariants.Count > 0)
            {
                if (provider == "")
                {
                    truck = new Truck(rateVariants.First());
                }
                else
                {
                    TruckRate providerRate = rateVariants.Find(x => x.Company == provider);
                    truck = providerRate.Company == "" ? truck : new Truck(providerRate);
                }
            }
            return truck;
        }

        private bool CheckPoints(List<DeliveryPoint> mapDelivery)
        {
            bool chk = mapDelivery.Count >0 ;           

            foreach (DeliveryPoint point in mapDelivery)
            {                                    
                 chk =RoutesList.FindAll(x => x.IdCustomer == point.IdCustomer).Count>0;
                if (!chk) { break; }
            }
            return chk;
        }

        /// <summary>
        /// Региональные перевозки
        /// </summary>
        /// <param name="rateVariants"></param>
        /// <returns></returns>
        private List<TruckRate> GetTruckRate(double tonnageNeed,
                 List<DeliveryPoint> mapDelivery)
        {
            List<TruckRate> rateVariants = new List<TruckRate>();
            int ix = 0;
            int MaxCost = 0;
            string city = "";

            /// подходящие варианты перевозчиков

            for (int i = 0; i < mapDelivery.Count; i++)
            {      //выбор дальней точки
                DeliveryPoint point = mapDelivery[i];

                //  List < DeliveryPoint > variants  
                try
                {
                    int? MaxCostPoint = 0;
                    MaxCostPoint = (from rv in RateList
                                    where rv.City == point.City &&
                                            rv.Tonnage > tonnageNeed
                                    select rv.PriceFirstPoint
                                )?.Max();
                    if (MaxCostPoint != null)
                    {
                        if (MaxCost < MaxCostPoint)
                        {
                            MaxCost = (int)MaxCostPoint;
                            ix = i;
                            city = point.City;
                        }
                    }
                }
                catch
                {
                    Debug.WriteLine($"Не удалось найти точку. Проверьте наличие в Id клиента {mapDelivery[i].IdCustomer} на Листе \"Route\"");
                    throw new Exception("Не удалось найти точку.");
                }
            }
            rateVariants = RateList.FindAll(r =>
                                        r.City == mapDelivery[ix].City &&
                                        r.Tonnage > tonnageNeed
                                        ).ToList();

            if (rateVariants.Count > 0)
            {
                //По каждому варианту фирмы с дальним городом
                for (int rateIx = 0; rateIx < rateVariants.Count; rateIx++)
                {
                    bool hasFirstpoint = false;
                    TruckRate variantRate = rateVariants[rateIx];
                    variantRate.TotalDeliveryCost = 0;
                    // считаем общую стоимость
                    for (int pointNumber = 0; pointNumber < mapDelivery.Count; pointNumber++)
                    {
                        if (mapDelivery[pointNumber].City == city && !hasFirstpoint)
                        {
                            variantRate.TotalDeliveryCost += rateVariants[rateIx].PriceFirstPoint;
                            hasFirstpoint = true;
                        }
                        else
                        {
                            //Ищем стоимость доп точки в другом городе для той же машины 
                            TruckRate addPointRate =
                                RateList.Where(x => x.Company == variantRate.Company &&
                                                    x.Tonnage == variantRate.Tonnage &&
                                                    x.City == mapDelivery[pointNumber].City).First();
                            variantRate.TotalDeliveryCost += addPointRate.PriceAddPoint;
                        }
                    }
                    rateVariants[rateIx] = variantRate;
                }
                rateVariants = rateVariants.OrderBy(r => r.TotalDeliveryCost).ToList();
            }
            return rateVariants;
        }

        /// <summary>
        /// варианты Провайдеров для авто
        /// </summary>
        /// <param name="totalWeight"></param>
        /// <param name="mapDelivery"></param>
        /// <returns></returns>
        public List<TruckRate> GetTruckRateInternational(double totalWeight, List<DeliveryPoint> mapDelivery)
        {
            int centner = (int)Math.Ceiling(totalWeight / 100); //центнеры огругление вверх
            List<TruckRate> rateVariants = new List<TruckRate>();
            double tonnageNeed =(double) centner/10 ;   //тонн 

            for (int j=0; j< RateInternationalList.Count; j++)
            {
                TruckRate rate = RateInternationalList[j];
                if (mapDelivery[0].City.Contains(rate.City) && rate.Tonnage == tonnageNeed)
                {
                    rateVariants.Add(rate);
                }
            }
              //rateVariants = RateInternationalList.FindAll(
              //                  x => mapDelivery[0].City.Contains(x.City) &&
              //                  x.Tonnage == tonnageNeed);
           
            for (int i = 0; i < rateVariants.Count; i++)
            {
                TruckRate rate = rateVariants[i];
                int addpointCost = (mapDelivery.Count - 1) * rateVariants[i].PriceAddPoint;
                rate.TotalDeliveryCost =
                    (int)Math.Ceiling(rateVariants[i].PriceFirstPoint * totalWeight / 100 + addpointCost);
                rateVariants[i] = rate;
            }
            rateVariants = rateVariants.OrderBy(r => r.TotalDeliveryCost).ToList();
            return rateVariants;
        }

        /// <summary>
        /// По МСК и МО
        /// </summary>
        /// <param name="tonnageNeed"></param>
        /// <param name="rateVariants"></param>
        /// <param name="mapDelivery"></param>
        /// <returns></returns>
        private List<TruckRate> GetCostMskRoutes(double tonnageNeed,
                  List<DeliveryPoint> mapDelivery)
        {
            List<TruckRate> rateVariants = new List<TruckRate>();
            rateVariants = RateList.FindAll(r =>
                                        r.City == mapDelivery[0].City &&
                                      r.Tonnage > tonnageNeed
                                        ).ToList();

            if (rateVariants.Count > 0)
            {
                for (int rateIx = 0; rateIx < rateVariants.Count; rateIx++)
                {
                    TruckRate variantRate = rateVariants[rateIx];
                    variantRate.TotalDeliveryCost = rateVariants[rateIx].PriceFirstPoint;
                    for (int pointNumber = 1; pointNumber < mapDelivery.Count; pointNumber++)
                    {
                        TruckRate addPointRate =
                            RateList.Where(x => x.Company == variantRate.Company &&
                                                x.Tonnage == variantRate.Tonnage &&
                                                x.City == mapDelivery[pointNumber].City).First();
                        if (addPointRate.PriceAddPoint > 0)
                            variantRate.TotalDeliveryCost += addPointRate.PriceAddPoint;
                    }
                    rateVariants[rateIx] = variantRate;
                }
                rateVariants = rateVariants.OrderBy(r => r.TotalDeliveryCost).ToList();
            }
            return rateVariants;
        }

        /// <summary>
        /// Получить вес список цен перевозчиков в формате списка         
        /// </summary>
        /// <returns></returns>
         private static List<TruckRate> GetTruckRateList()
        {
            List<TruckRate> ListRate = new List<TruckRate>();
            foreach (ListRow row in RateTable.ListRows)
            {
                string valTonnage = row.Range[1, RateTable.ListColumns["tonnage, t"].Index].Text;
                double tonnage = double.TryParse(valTonnage, out double t) ? t : 0;

                string valCity = row.Range[1, RateTable.ListColumns["City"].Index].Text;
                valCity = valCity.Trim();

                string valCompany = row.Range[1, RateTable.ListColumns["Company"].Index].Text;
                valCompany = valCompany.Trim();


                if (tonnage > 0 && !string.IsNullOrWhiteSpace(valCity))
                {

                    string strPrice = row.Range[1, RateTable.ListColumns["vehicle"].Index].Text;
                    int priceFirst = int.TryParse(strPrice, out int pf) ? pf : 0;
                    strPrice = row.Range[1, RateTable.ListColumns["add.point"].Index].Text;
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

        /// <summary>
        /// Получить таблицу международных 
        /// </summary>
        /// <returns></returns>
        internal List<TruckRate> GetTruckRateInternational()
        {
            List<TruckRate> ListRate = new List<TruckRate>();
            Worksheet sheetRoute = Globals.ThisWorkbook.Sheets["Rate Inetrnational"];
            ListObject rateTable = sheetRoute.ListObjects["TableRateInternational"];
            foreach (ListRow row in rateTable.ListRows)
            {
                string valTonnage = row.Range[1, rateTable.ListColumns["tonnage, t"].Index].Text;
                double tonnage = double.TryParse(valTonnage, out double t) ? t : 0;

                string valCity = row.Range[1, rateTable.ListColumns["City"].Index].Text;
                valCity = valCity.Trim();

                string valCompany = row.Range[1, rateTable.ListColumns["Company"].Index].Text;
                valCompany = valCompany.Trim();

                if (tonnage > 0 && !string.IsNullOrWhiteSpace(valCity))
                {
                    string strPrice = row.Range[1, rateTable.ListColumns["vehicle"].Index].Text;
                    int price = int.TryParse(strPrice, out int pf) ? pf : 0;

                    strPrice = row.Range[1, rateTable.ListColumns["add.point"].Index].Text;
                    int priceAdd = int.TryParse(strPrice, out int pa) ? pa : 0;

                    TruckRate rate = new TruckRate()
                    {
                        City = valCity,
                        Company = valCompany,
                        PriceFirstPoint = price,
                        PriceAddPoint = priceAdd,
                        PlaceShipment = row.Range[1, 1].Text,
                        Tonnage = tonnage
                    };

                    ListRate.Add(rate);
                }
            }
            return ListRate;
        }


        internal int CreateRoute(List<Order> ordersCurrentDelivery)
        {
            List<DeliveryPoint> pointMap = RoutesList;
            DeliveryPoint LastPoint = RoutesList.Last();
            int idRoute = LastPoint.Id + 1;
            int priorityRoute = LastPoint.PriorityRoute + 1;
            //Поиск подходящего максимального приоритета
            foreach (Order ord in ordersCurrentDelivery)
            {
                string customerId = ord.Customer.Id;
                List<int> routes = (from p in pointMap
                                    where p.IdCustomer == customerId
                                    select p.PriorityRoute
                                     ).Distinct().ToList();
                int maxPriority = 0;
                if (routes.Count != 0) maxPriority = routes.Max();

                priorityRoute = maxPriority > priorityRoute ? maxPriority : priorityRoute;
            }
            int point = 0;
            foreach (Order order in ordersCurrentDelivery)
            {
                ListRow row = RoutesTable.ListRows[RoutesTable.ListRows.Count];
                RoutesTable.ListRows.Add();
                row.Range[1, RoutesTable.ListColumns["Id route"].Index].Value = idRoute;
                row.Range[1, RoutesTable.ListColumns["Priority route"].Index].Value = priorityRoute;
                row.Range[1, RoutesTable.ListColumns["Priority point"].Index].Value = ++point;
                row.Range[1, RoutesTable.ListColumns["Получатель материала"].Index].Value = order.Customer.Id;
                row.Range[1, RoutesTable.ListColumns["City"].Index].Value = order.DeliveryPoint.City;

                //поиск этого же Получателя в другой строке
                DeliveryPoint findPoint = pointMap.Find(x => x.IdCustomer == order.Customer.Id && x.CityLongName != "");
                if (!string.IsNullOrWhiteSpace(findPoint.CustomerNumber))
                {
                    row.Range[1, RoutesTable.ListColumns["Город"].Index].Value = findPoint.CityLongName;
                    row.Range[1, RoutesTable.ListColumns["Маршрут"].Index].Value = findPoint.Route;
                    row.Range[1, RoutesTable.ListColumns["Направление"].Index].Value = findPoint.RouteName;
                    row.Range[1, RoutesTable.ListColumns["Клиент"].Index].Value = findPoint.Customer;
                    row.Range[1, RoutesTable.ListColumns["Номер клиента"].Index].Value = findPoint.CustomerNumber;
                    row.Range[1, RoutesTable.ListColumns["Add"].Index].Value = "Auto";
                }
            }
            RoutesList = null;
            return idRoute;
        }


        public Range GetCurrentShippingRange()
        {
            Range currentRng = null;
            string dateDelivery = DateDelivery;
            int columnDeliveryId = TotalTable.ListColumns["Дата доставки"].Index;
            foreach (ListRow row in TotalTable.ListRows)
            {
                string dateTable = row.Range[0, columnDeliveryId].Text;
                if (dateTable == dateDelivery || string.IsNullOrEmpty(dateTable))
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
        #region Вспомогательные
        /// <summary>
        /// Ищет в диапазоне текст возвращает значение ячейки по указанному смещению
        /// </summary>
        /// <param name="header"></param>
        /// <param name="rng"></param>
        /// <param name="offsetRow"></param>
        /// <param name="offsetCol"></param>
        /// <returns></returns>
        public static string FindValue(string header, Range rng, int offsetRow = 0, int offsetCol = 0)
        {
            Range findCell = rng.Find(What: header, LookIn: XlFindLookIn.xlValues);
            if (findCell == null) return "";
            findCell = findCell.Offset[offsetRow, offsetCol];
            string valueCell = findCell.Text;
            return valueCell;
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

        public static string GetProviderId(string providerName)
        {
            int colName = ProviderTable.ListColumns["Company"].Index;
            int colId = ProviderTable.ListColumns["Id"].Index;
            int colCounter = ProviderTable.ListColumns["Счетчик"].Index;
            string id = "";
            foreach (Range row in ProviderTable.DataBodyRange.Rows)
            {
                if (row.Cells[1, colName].Text == providerName)
                {
                    string ix = row.Cells[1, colId].Text;
                    int counter = int.TryParse(row.Cells[1, colCounter].Text, out int count) ? count : 0;
                    row.Cells[1, colCounter].Value = ++counter;
                    string Counter = counter.ToString();
                    Counter = new string('0', 6 - Counter.Length) + Counter;
                    id = ix + Counter;
                    break;
                }
            }

            return id;
        }

        internal bool CheckCustomerId(string id)
        {
            DeliveryPoint dp = RoutesList.Find(x => x.IdCustomer.Contains(id));
            return string.IsNullOrWhiteSpace(dp.IdCustomer);
        }

        #endregion Вспомогательные
    }
}
