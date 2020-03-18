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
            set { _shippingCompany = value; } 
        }
        private ShippingCompany _shippingCompany;
   

        public Truck( )
        {

        }

        public Truck(TruckRate truckRate)
        {
            Tonnage = truckRate.Tonnage;
            CostFirstPoint = truckRate.PriceFirstPoint;
            CostAddPoint = truckRate.PriceAddPoint;
            Cost = truckRate.TotalDeliveryCost;
            string companyName =  truckRate.Company ;
            ShippingCompany = new ShippingCompany() { Name = companyName } ;
        }
    }
}
