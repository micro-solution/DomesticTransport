namespace DomesticTransport.Model
{
    /// <summary>
    ///  Класс заказа (позиция перевозки)
    /// </summary>
    public class Order
    {
        /// <summary>
        /// Id заказа (поле Номер доставки, Delivery)
        /// </summary>
        public string Id
        {
            get => _id;
            set => _id = value.Length < 10 ? new string('0', 10 - value.Length) + value : value;
        }
        private string _id;

        // Прорядковый номер доставки
        public int DeliveryNumber { get; set; }

        ///Порядковый номер точки выгрузки
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
            set => _transportationUnit = value;
        }
        private string _transportationUnit;

        public double Cost { get; set; }
        public string Route { get; set; }
        public DeliveryPoint DeliveryPoint { get; set; }
    }
}
