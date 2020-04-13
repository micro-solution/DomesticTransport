using System;
using System.Collections.Generic;
using System.Linq;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Класс автомобиля перевозчика
    /// </summary>
    public class Truck
    {
        public double Tonnage { get; set; }

        /// <summary>
        /// Стоимость доставки
        /// </summary>
        public double Cost
        {
           get => _cost;
    set => _cost = Math.Ceiling(value);
        }
        double _cost;
        public Provider ProviderCompany
        {
            get
            {
                if (_shippingCompany == null)
                {
                    _shippingCompany = new Provider();
                }
                return _shippingCompany;
            }
            set => _shippingCompany = value;
        }
        private Provider _shippingCompany;


        public Truck() { }

        public Truck(TruckRate truckRate)
        {
            Tonnage = truckRate.Tonnage;
            Cost = truckRate.TotalDeliveryCost;
            string companyName = truckRate.Company;
            ProviderCompany = new Provider() { Name = companyName };
        }

        /// <summary>
        /// Выбрать авто 
        /// </summary>
        /// <param name="totalWeight"></param>
        /// <param name="mapDelivery"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static Truck GetTruck(double totalWeight, List<DeliveryPoint> mapDelivery, string provider = "")
        {
            if (mapDelivery.Count <= 0 || totalWeight <= 0) return null;
            if (!Delivery.CheckPoints(mapDelivery)) return null;  //Нет клиента

            Truck truck = null;
            List<TruckRate> rateVariants = new List<TruckRate>();
            double tonnageNeed = totalWeight / 1000 - 0.05;  /// 50kg Допустимый перегруз

            try
            {
                if (mapDelivery.FindAll(m => m.City == "MSK" || m.City == "MO").Count > 0)
                {
                    rateVariants = TruckRate.GetCostMskRoutes(tonnageNeed, mapDelivery); //Для Москвы и области  (первая точка с наибольшим приоритетом по таблице)
                }
                else
                {
                    bool isInternational = false;

                    foreach (string city in ShefflerWB.InternationalCityList) // Nur - Sultan //Yerevan
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
                    rateVariants = TruckRate.GetTruckRateInternational(totalWeight, mapDelivery) :
                    rateVariants = TruckRate.GetTruckRate(tonnageNeed, mapDelivery);
                }
            }
            catch
            {
                truck = new Truck()
                {
                    Cost = 0,
                    Tonnage = 0
                };
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
    }
}
