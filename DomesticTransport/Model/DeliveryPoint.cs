using System;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Строка таблицы Routes (Точка Доставки или Получатель)
    /// </summary>
    public struct DeliveryPoint
    {
        public int Id { get; set; }
        public int PriorityRoute { get; set; } 
        public int PriorityPoint { get; set; } 

        public string IdCustomer
        {
            get => _idCustomer;
            set => _idCustomer = value.Length < 10 ? new string('0', 10 - value.Length) + value : value;
        }
        private string _idCustomer;

        public string City
        {
            get => _city;
            set => _city = value.Trim();
        }
        string _city;

        public string CityLongName
        {
            get => _cityLongName;
            set => _cityLongName = value.Trim();
        }
        string _cityLongName;

        public string CustomerNumber
        {
            get => _customerNumber;
            set => _customerNumber = value.Trim();
        }
        string _customerNumber;

        public string Customer
        {
            get => _customer;
            set => _customer = value.Trim();
        }
        string _customer;

        public string Route
        {
            get => _route;
            set => _route = value.Trim();
        }
        string _route;

        public string RouteName
        {
            get => _routeName ?? GetRouteName();
            set => _routeName = value.Trim();
        }
        string _routeName;
      //  private string id;

        public DeliveryPoint(string id , string routeName) : this()
        {               
            this = ShefflerWB.RoutesList.Find(x => x.IdCustomer == id   &&
                x.RouteName ==routeName);
        }

        private string GetRouteName()
        {
            string routename ="";
            string id = IdCustomer;
            DeliveryPoint dp = ShefflerWB.RoutesList.Find(x => x.IdCustomer == id && (!string.IsNullOrWhiteSpace(x.RouteName)));
            if (dp.IdCustomer != null) routename = dp.RouteName;              
            return routename;
        }
        /// <summary>
        /// Найти у клиента название маршрута
        /// </summary>
        public void SetRouteName()
        {
            string routename =RouteName;
            string id = IdCustomer;
            DeliveryPoint dp = ShefflerWB.RoutesList.Find(x => x.IdCustomer == id && (!string.IsNullOrWhiteSpace(x.RouteName)));
            if (!string.IsNullOrWhiteSpace(dp.RouteName)) routename = dp.RouteName;
            RouteName = routename;
        }
    }
}
