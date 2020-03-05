using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    class Carrier
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public Truck Truck { get; set; }

        

        ShippingCompany ShippingCompany{ get; set; }


}
}
