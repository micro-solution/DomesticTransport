namespace DomesticTransport.Model
{
    /// <summary>
    /// Класс автомобиля перевозчика
    /// </summary>
    class Truck
    {
        public string Number { get; set; }
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

        public ShippingCompany ShippingCompany
        {
            get
            {
                if (_shippingCompany == null)
                {
                    _shippingCompany = new ShippingCompany();
                }
                return _shippingCompany;
            }
            set { _shippingCompany = value; }
        }
        private ShippingCompany _shippingCompany;


        public Truck() { }

        public Truck(TruckRate truckRate)
        {
            Tonnage = truckRate.Tonnage;
            CostFirstPoint = truckRate.PriceFirstPoint;
            CostAddPoint = truckRate.PriceAddPoint;
            Cost = truckRate.TotalDeliveryCost;
            string companyName = truckRate.Company;
            ShippingCompany = new ShippingCompany() { Name = companyName };
        }
    }
}
