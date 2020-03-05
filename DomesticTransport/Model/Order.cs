using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    enum Status
    {
        Awaiting,
        ReadyToShip,
        Canceled,
        Done
    }
    class Order
    {
        public string Id { get; set; }


        public Customer Customer { get;set;}
        public int PalletsCount { get; set; }
        public double Weight { get; set; }

        public string TransportationUnit { get; set; }
        public double Cost { get; set; }

        public string Route { get; set; }
        public Status Status { get; set; }

    }
}
