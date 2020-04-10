using System;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Класс автомобиля перевозчика
    /// </summary>
   public class Truck
    {
   
        public string Mark { get; set; }
        public double Tonnage { get; set; }

        /// <summary>
        /// Стоимость доставки в первую точку
        /// </summary>
        public double CostFirstPoint { get; set; }

        /// <summary>
        /// Стоимость дополнительной точки
        /// </summary>
        public double CostAddPoint { get; set; }

        /// <summary>
        /// ???
        /// </summary>
        public double Cost
        { 
           get {
                return _cost;
            }
    set { _cost = Math.Ceiling(value); }
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
            set { _shippingCompany = value; }
        }
        private Provider _shippingCompany;


        public Truck() { }

        public Truck(TruckRate truckRate)
        {
            Tonnage = truckRate.Tonnage;
            CostFirstPoint = truckRate.PriceFirstPoint;
            CostAddPoint = truckRate.PriceAddPoint;
            Cost = truckRate.TotalDeliveryCost;
            string companyName = truckRate.Company;
            ProviderCompany = new Provider() { Name = companyName };
        }
    }
}
