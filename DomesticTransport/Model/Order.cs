namespace DomesticTransport.Model
{
    /// <summary>
    ///  Класс заказа (позиция перевозки)
    /// </summary>
    class Order
    {
        /// <summary>
        /// Идентификатор заказа
        /// </summary>
        public string Id
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value.Length < 10 ? new string('0', 10 - value.Length) + value : value;
            }
        }
        private string _id;


        public int DeliveryNumber { get; set; }
        public int PointNumber { get; set; }

        public Customer Customer
        {
            get
            {
                if (_customer == null) _customer = new Customer();
                return _customer;
            }
            set => _customer = value;
        }
        private Customer _customer;


        public int PalletsCount { get; set; }
        public double WeightNetto { get; set; }
        public double WeightBrutto { get; set; }

        public string TransportationUnit
        {
            get => _transportationUnit;
            set
            {
                //if (_transportationUnit.Length >18)
                //{
                    _transportationUnit = value;
                //}
                //else if (!string.IsNullOrWhiteSpace(value))
                //{
                //    _transportationUnit = new string('0', 18 - value.Length) + value;
                //}
            }
        }
        private string _transportationUnit;

        public double Cost { get; set; }
        public string Route { get; set; }
        public DeliveryPoint DeliveryPoint { get; set; }
    }
}
