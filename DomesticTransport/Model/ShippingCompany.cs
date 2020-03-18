using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    class ShippingCompany
    {
        private string companyName;

        public ShippingCompany()
        {   }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string  Phone { get; set; }
    }
}
