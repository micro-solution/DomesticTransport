using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{

    class Order
    {
        public string Id;
        public int PointNumber { get; set; }
        public Customer Customer 
        {
            get 
            {
                if (_customer == null) _customer = new Customer();
                return _customer;
            }

            set { _customer = value; }
        }
        private Customer _customer;


        public int PalletsCount { get; set; }
        public double WeightNetto { get; set; }

        public double WeightBrutto { get; set; }

        public string TransportationUnit
        {
            get { return _transportationUnit; }
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    _transportationUnit = new string('0', 18 - value.Length) + value;
                }
            }

        }
        private string _transportationUnit;
        public double Cost { get; set; }
        public string Route { get; set; }

        public DeliveryPoint DeliveryPoint 
        {
            get { return _deliveryPoint; } 
            set { _deliveryPoint = value; } }
        private DeliveryPoint _deliveryPoint;
        //public int Prioriy {
        //    get{
        //    if (_prioriy == 0 && Customer !=null)//!string.IsNullOrWhiteSpace(Route))
        //        {
        //            using (ShefflerWorkBook functions = new ShefflerWorkBook())
        //            {
        //                _prioriy = functions.(Customer.Id);
        //            }                     
        //        }
        //        return _prioriy; }          
        //    }
        int _prioriy = 0;
    }
}
