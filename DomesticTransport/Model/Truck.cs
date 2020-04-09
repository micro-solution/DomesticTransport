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
        public int CostFirstPoint { get; set; }

        /// <summary>
        /// Стоимость дополнительной точки
        /// </summary>
        public int CostAddPoint { get; set; }

        /// <summary>
        /// ???
        /// </summary>
        public int Cost { get; set; }
        //{ 
        //   get { 
        //        _cost = 
        //    }
        //    set { }
        //}
        //int _cost;
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
