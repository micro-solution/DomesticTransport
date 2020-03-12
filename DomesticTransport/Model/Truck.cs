using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    class Truck
    {
        public string Number { get; set; }
        public string Mark { get; set; }
        public double Tonnage { get; set; }

        public int CostFirstPoint { get; set; }
        public int CostAddPoint { get; set; }

        public int Cost { get; set; }

        public ShippingCompany ShippingCompany 
        { get {
            if (_shippingCompany == null)
                {
                    _shippingCompany = new ShippingCompany();
                }
                return _shippingCompany;
            }
            set { } 
        }
        private ShippingCompany _shippingCompany;
   

        public Truck( )
        {

        }

        public Truck(TruckRate truckRate)
        {
          
        }
    }
}
