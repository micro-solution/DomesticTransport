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
        Truck truck { get; set; }
        ShippingCompany ShippingCompany{ get; set; }


}
}
