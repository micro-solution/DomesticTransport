using Microsoft.Office.Interop.Excel;

using System.Collections.Generic;
using System.Linq;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Доставка товара
    /// </summary>
    public class Delivery
    {
        /// <summary>
        /// Номер доставки
        /// </summary>
        public int Number { get; set; } = 0;

        public string Time 
        {
            get {
                   if (string.IsNullOrWhiteSpace(_timetable))
                {
                    string city = MapDelivery.Count>0 ? MapDelivery[0].City :"";
                    _timetable = ShefflerWB.GetTime(city);
                }
                return _timetable;
            } 
             set =>  _timetable = value; 
        }
        string _timetable;

        /// <summary>
        /// Дата доставки
        /// </summary>
        public  string DateDelivery {
            get {
                if (string.IsNullOrEmpty(_dateDelivery))
                {
                    _dateDelivery = ShefflerWB.DateDelivery;
                }
                return _dateDelivery;
            }
            set => _dateDelivery = value; }
        string _dateDelivery;

        public bool HasRoute { get; set; } = true;

        ///// <summary>
        ///// Информация о водителе
        ///// </summary>
        public Driver Driver { get; set; }

        /// <summary>
        /// Стоимость доставки
        /// </summary>
        public decimal Cost
        {
            get
            {
                if (_cost == 0)
                {
                    if (Truck?.ProviderCompany?.Name == "Деловые линии" || Truck == null)
                    {
                        _cost = 0;
                    }
                    else
                    {
                        _cost = Truck?.Cost ?? 0;
                    }
                }
                return _cost;
            }
            set => _cost = value;
        }
        private decimal _cost;
        /// <summary>
        /// Общий вес
        /// </summary>
        public double TotalWeight
        {
            get
            {
                double sum = 0;
                Orders.ForEach(x => sum += x.WeightBrutto == 0 ? x.WeightNetto : x.WeightBrutto);
                return sum;
            }
        }

        ///// <summary>
        ///// Стоимость товаров
        ///// </summary>
        public double CostProducts
        {
            get
            {
                double sum = 0;
                Orders.ForEach(x => sum += x.Cost);
                return sum;
            }
        }

        public List<Order> Orders
        {
            get
            {
                if (_orders == null)
                {
                    _orders = new List<Order>();
                }
                return _orders;
            }
            set => _orders = value;
        }
        private List<Order> _orders;

        /// <summary>
        /// Точки доставки
        /// </summary>
        public List<DeliveryPoint> MapDelivery
        {
            get
            {
                List<DeliveryPoint> dp = (from r in Orders
                                          select r.DeliveryPoint
                                          ).Distinct().ToList();
                dp.OrderBy(x => x.PriorityRoute).ThenBy(y => y.PriorityPoint);
                return dp;
            }
        }

        /// <summary>
        /// Приоритет в таблице отгрузки
        /// </summary>
        public int SortPriority
        {
            get
            {
                if (MapDelivery?[0].RouteName == "Сборный груз") return 9999;
                
                int i = 0;
                foreach (ListRow row in ShefflerWB.SityTable.ListRows)
                {
                    i++;
                    if (MapDelivery != null && row.Range.Cells[1, 1].Text == MapDelivery?[0].City) break;
                }
                return i;
            }
        }

        public Truck Truck
        {
            get
            {
                if (_truck == null)
                {
                    if (!string.IsNullOrWhiteSpace(MapDelivery.Find(
                                    x => x.RouteName.Contains("Сборный груз")).IdCustomer))
                    {
                        _truck = new Truck();
                        _truck.ProviderCompany.Name = "Деловые линии";
                    }
                    else
                    {
                        _truck = Truck.GetTruck(TotalWeight, MapDelivery);
                    }

                }
                return _truck;
            }
            set => _truck = value;
        }
        private Truck _truck;

        public Delivery() { }
        public Delivery(Order order)
        {
            Orders.Add(order);
        }

        /// <summary>
        /// Проверка на превышение веса
        /// </summary>
        /// <param name="order"></param>
        /// <returns></returns>
        public bool CheckDeliveryWeight(Order order)
        {
            double sum = TotalWeight + order.WeightNetto;
            return sum <= 20100;
        }
        public bool CheckDeliveryWeightLTL(Order order)
        {
            double sum = TotalWeight + order.WeightNetto;
            return sum <= 20000;
        }

        public void SaveRoute()
        {
            if (HasFullRoute(this.MapDelivery)) { return; }
            int idRoute = new ShefflerWB().CreateRoute(MapDelivery);
            foreach (Order ord in Orders)
            {
                DeliveryPoint dp = ord.DeliveryPoint;
                dp.Id = idRoute;
                ord.DeliveryPoint = dp;
            }
        }
        /// <summary>
        ///True если все точки из маршрута есть в таблице маршрутов с общим Id  
        /// </summary>
        /// <param name="mapDelivery"></param>
        /// <returns></returns>
        public static bool HasFullRoute(List<DeliveryPoint> mapDelivery)
        {
            if (mapDelivery.Count == 0) return false;
            //все Id маршрутов             
            int[] variantsId = (from r in ShefflerWB.RoutesList
                                where r.IdCustomer == mapDelivery[0].IdCustomer
                                select r.Id).Distinct().ToArray();

            if (variantsId.Length == 0 ) return false;
            bool hasRoute = false;
            for (int i = 0; i < variantsId.Length; i++)
            {
                hasRoute = true;
                foreach (DeliveryPoint point in mapDelivery)
                {
                    if (ShefflerWB.RoutesList.FindAll(x => x.Id == variantsId[i] &&
                                            x.IdCustomer == point.IdCustomer).Count == 0)
                    {
                        hasRoute = false; break; // В группе нет точки
                    }
                }
                if (hasRoute)
                {
                    break; //есть маршрут, удовлетворяет всем точкам поездки 
                }
            }
            return hasRoute;
        }

        /// <summary>
        /// Проверить наличие маршрута
        /// </summary>
        /// <param name="id"></param>
        /// <returns> true если все точки есть в таблице</returns>
        public static bool CheckCustomerRoute(string id)
        {
            DeliveryPoint dp = ShefflerWB.RoutesList.Find(x => x.IdCustomer.Contains(id));
            return string.IsNullOrWhiteSpace(dp.IdCustomer);
        }


        /// <summary>
        ///  Проверить  все ли клиенты есть в таблице
        /// </summary>
        /// <param name="mapDelivery"></param>
        /// <returns></returns>
        public static bool CheckPoints(List<DeliveryPoint> mapDelivery)
        {
            bool chk = mapDelivery.Count > 0;
            foreach (DeliveryPoint point in mapDelivery)
            {
                chk = ShefflerWB.RoutesList.FindAll(x => x.IdCustomer == point.IdCustomer).Count > 0;
                if (!chk)
                {
                    break;
                }
            }
            return chk;
        }
    }
}
