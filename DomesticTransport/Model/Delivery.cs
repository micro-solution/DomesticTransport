using System;
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
        /// Найден маршрут доставки в таблице
        /// </summary>
        public bool HasRoute { get; set; } = true;

        /// <summary>
        /// Номер доставки
        /// </summary>
        public int Number { get; set; } = 0;


        ///// <summary>
        ///// Информация о водителе
        ///// </summary>
        public Driver Driver { get; set; }

        /// <summary>
        /// Стоимость доставки
        /// </summary>
        public double Cost
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
        private double _cost;
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


        public Truck Truck
        {
            get
            {
                if (_truck == null)
                {
                    ShefflerWB workBook = new ShefflerWB();
                    _truck = Truck.GetTruck(TotalWeight, MapDelivery);
                    if (!string.IsNullOrWhiteSpace(MapDelivery.Find(
                                    x => x.RouteName.Contains("Сборный груз")).IdCustomer))
                    {
                        Truck.ProviderCompany.Name = "Деловые линии";
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
            if (!CheckPoints(MapDelivery))
            { return; }



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
