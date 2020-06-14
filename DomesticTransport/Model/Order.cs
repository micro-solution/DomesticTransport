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
        /// <summary>
        /// Дата отгрузки
        /// </summary>
        public string DateDelivery
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_dateDelivery))
                {
                    _dateDelivery = ShefflerWB.DateDelivery;
                }
                return _dateDelivery;
            }
            set => _dateDelivery = value;
        }

        private string _dateDelivery;
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
        public string RouteCity
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_route) && DeliveryPoint.RouteName != "")
                {
                    _route = DeliveryPoint.RouteName;
                }
                return _route;
            }
            set => _route = value;
        }

        private string _route;
        public DeliveryPoint DeliveryPoint
        {
            get
            {
                // if (!string.IsNullOrWhiteSpace(_deliveryPoint.RouteName) && !string.IsNullOrWhiteSpace(RouteCity))
                ///{

                // 
                ///  _deliveryPoint = new DeliveryPoint(Customer.Id, RouteCity);                   

                //}
                return _deliveryPoint;
            }
            set => _deliveryPoint = value;

        }

        private DeliveryPoint _deliveryPoint;
    }
}
