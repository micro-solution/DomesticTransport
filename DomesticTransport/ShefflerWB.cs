using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace DomesticTransport
{

    /// <summary>
    /// Действия с текущей книгой
    /// </summary>
    public class ShefflerWB
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

        public static string GetTime(string city)
        {
            ListObject TimeTable = RoutesSheet.ListObjects["Timetable"];
            string time ="";
            int     isinternational = (from r in InternationalCityList
                       where r == city
                       select r).ToArray().Length;
            if (isinternational > 0)
            {
                city = "LTL";
            }
          
            foreach (ListRow row in TimeTable.ListRows)
            {
                string direction = row.Range[1, TimeTable.ListColumns["Направление"].Index].Text;
                if ( direction.Contains(city) )
                {
                 time = row.Range[1, TimeTable.ListColumns["Время погрузки"].Index].Text;
                   break;
                }
                if (direction.Contains("Межгород"))
                {
                    time = row.Range[1, TimeTable.ListColumns["Время погрузки"].Index].Text;
                }
            }  
            return time;
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

        /// <summary>
        /// Routes лист 
        /// </summary>
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
        /// Таблица городов
        /// </summary>
        public static ListObject SityTable
        {
            get
            {
                if (_sityTable == null)
                {
                    _sityTable = RoutesSheet.ListObjects["TableCity"];
                }
                return _sityTable;
            }
        }
        private static ListObject _sityTable;

        /// <summary>
        /// Прайс
        /// </summary>
        public static List<TruckRate> RateList
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
                if ((!string.IsNullOrWhiteSpace(cell)) && idOrder.Contains(cell))
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

                    _routes = _routes.OrderBy(x => x.Id).ThenBy(
                                        y => y.PriorityRoute).ThenBy(y => y.PriorityPoint).ToList();
                }
                return _routes;
           } 
            set => _routes = value;
        }
        static List<DeliveryPoint> _routes;

        /// <summary>
        ///  Список цен международных перевозок
        /// </summary>
        public static List<TruckRate> RateInternationalList
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



        /// <summary>
        /// Города Нур-Султан, Ереван
        /// </summary>
        public static string[] InternationalCityList
        {
            get
            {
                if (_internationalCityList == null)
                {
                    List<TruckRate> rates = RateInternationalList;
                    _internationalCityList = (from LR in rates
                                              select LR.City
                                 ).Distinct().ToArray();
                }
                return _internationalCityList;
            }
        }
        private static string[] _internationalCityList;


        /// <summary>
        /// Сортировка таблицы доставки
        /// </summary>
        public static void DeliveryTableSort()
        {
            Range table = DeliveryTable.Range;
            Range col1 = table.Columns[DeliveryTable.ListColumns["№ Доставки"].Index];
            table.Sort(
                Key1: col1, 
                Order1: XlSortOrder.xlAscending, 
                Header: XlYesNoGuess.xlYes, 
                Type: Type.Missing, 
                OrderCustom: Type.Missing, MatchCase: Type.Missing,
                DataOption1: XlSortDataOption.xlSortNormal, 
                Orientation: XlSortOrientation.xlSortColumns, 
                SortMethod: XlSortMethod.xlPinYin);
        }

        /// <summary>
        /// Сортировка таблицы поставки
        /// </summary>
        public static void OrderTableSort()
        {
            Range table = OrdersTable.Range;
            Range col1 = table.Columns[OrdersTable.ListColumns["№ Доставки"].Index];
            Range col2 = table.Columns[OrdersTable.ListColumns["Порядок выгрузки"].Index];
            table.Sort(
                Key1: col1, 
                Order1: XlSortOrder.xlAscending, 
                Key2: col2, 
                Order2: XlSortOrder.xlAscending,
                OrderCustom: Type.Missing, MatchCase: Type.Missing,
                Header: XlYesNoGuess.xlYes, Orientation: XlSortOrientation.xlSortColumns);
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
                    string strPrice = row.Range[1, RateTable.ListColumns["vehicle"].Index].Value.ToString();
                    decimal priceFirst = decimal.TryParse(strPrice, out decimal pf) ? pf : 0;
                    strPrice = row.Range[1, RateTable.ListColumns["add.point"].Index].Value.ToString();
                    decimal priceAdd = decimal.TryParse(strPrice, out decimal pa) ? pa : 0;
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
        /// Получить таблицу цен международных перевозок 
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
                    string strPrice = row.Range[1, rateTable.ListColumns["vehicle"].Index].Value.ToString();
                    decimal price = decimal.TryParse(strPrice, out decimal pf) ? pf : 0;

                    strPrice = row.Range[1, rateTable.ListColumns["add.point"].Index].Value.ToString();
                    decimal priceAdd = decimal.TryParse(strPrice, out decimal pa) ? pa : 0;

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

        // Добавить маршрут в таблицу
        public int CreateRoute(List<DeliveryPoint> mapDelivery)
        {

            DeliveryPoint LastPoint = RoutesList.Last();
            int idRoute = LastPoint.Id + 1;
            int priorityNewRoute = 2;
            //Поиск максимального приоритета, из всех где встречается клиент
            foreach (DeliveryPoint point in mapDelivery)
            {
                List<int> routes = (from p in RoutesList
                                    where p.IdCustomer == point.IdCustomer
                                    select p.PriorityRoute
                                     ).Distinct().ToList();
                int maxPriority = 0;
                if (routes.Count > 0) maxPriority = routes.Max();

                priorityNewRoute = maxPriority > priorityNewRoute ? maxPriority + 1 : priorityNewRoute;
            }

            int priorityPoint = 0;

            foreach (DeliveryPoint point in mapDelivery)
            {
                RoutesTable.ListRows.Add();
                ListRow row = RoutesTable.ListRows[RoutesTable.ListRows.Count];
                row.Range[1, RoutesTable.ListColumns["Id route"].Index].Value = idRoute;
                row.Range[1, RoutesTable.ListColumns["Priority route"].Index].Value = priorityNewRoute;
                row.Range[1, RoutesTable.ListColumns["Priority point"].Index].Value = ++priorityPoint;
                row.Range[1, RoutesTable.ListColumns["Получатель материала"].Index].Value = point.IdCustomer;
                row.Range[1, RoutesTable.ListColumns["City"].Index].Value = point.City;

                //поиск этого же Получателя в другой строке
                string customerName = string.IsNullOrEmpty(point.Customer) ? "" : point.Customer;
                DeliveryPoint findPoint = RoutesList.Find(x => x.IdCustomer == point.IdCustomer && x.Customer != "");
                if (!string.IsNullOrWhiteSpace(findPoint.Customer))
                {
                    row.Range[1, RoutesTable.ListColumns["Город"].Index].Value = findPoint.CityLongName;
                    row.Range[1, RoutesTable.ListColumns["Маршрут"].Index].Value = findPoint.Route;
                    row.Range[1, RoutesTable.ListColumns["Направление"].Index].Value = findPoint.RouteName;
                    row.Range[1, RoutesTable.ListColumns["Номер клиента"].Index].Value = findPoint.CustomerNumber;
                    if (!string.IsNullOrWhiteSpace(customerName)) customerName = findPoint.Customer;
                }
                row.Range[1, RoutesTable.ListColumns["Клиент"].Index].Value = point.Customer;
                row.Range[1, RoutesTable.ListColumns["Add"].Index].Value = "Auto";
            }
            RoutesList = null;
            return idRoute;
        }

        /// <summary>
        /// Собирает диапазон Отгрузки с текущей даты
        /// </summary>
        /// <returns></returns>
        public Range GetCurrentTotalRange()
        {
            Range currentRng = null;
            string dateDelivery = DateDelivery;
            int columnDelivery = TotalTable.ListColumns["Дата доставки"].Index;
            foreach (ListRow row in TotalTable.ListRows)
            {
                string dateTable = row.Range[1, columnDelivery].Text;
                if (dateTable == dateDelivery )
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
            valueCell = valueCell.Trim();
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

        /// <summary>
        /// Возвращает id увеличивает счетчик заявок провайдеру 
        /// </summary>
        /// <param name="providerName"></param>
        /// <returns></returns>
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
                    Counter = Counter.Length < 6 ? new string('0', 6 - Counter.Length) + Counter : Counter;
                    id = ix + Counter;
                    break;
                }
            }
            return id;
        }

        #endregion Вспомогательные
    }
}
