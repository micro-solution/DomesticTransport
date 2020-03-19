using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Строка таблицы Routes (Точка Доставки или Получатель)
    /// </summary>
    struct DeliveryPoint
    {
        public int IdRoute { get; set; } 
        public int PriorityRoute { get; set; }
        public int PriorityPoint { get; set; }

        public string IdCustomer
        {
            get { return _idCustomer; }
            set
            {
                _idCustomer = new string('0', 10 - value.Length) + value;
            }
        }
        private string _idCustomer;

        public string City 
        {
            get {
               
                return _city;
            } 
            set
            {
                _city =  value.Trim(); 
            }
        }
        string _city;
        public string CityLongName
        {
            get { return _cityLongName; }
            set { _cityLongName = value.Trim(); }
        }
        string _cityLongName;
        public string CustomerNumber
        {
            get { return _customerNumber; }
            set { _customerNumber = value.Trim(); }
        }
        string _customerNumber;
        public string Customer
        {
            get { return _customer; }
            set { _customer = value.Trim(); }
        }
        string _customer;
        public string Route
        {
            get { return _route; }
            set { _route = value.Trim(); }
        }
        string _route;
        public string RouteName
        {
            get { return _routeName; }
            set { _routeName = value.Trim(); }
        }
        string _routeName;

    }
}
