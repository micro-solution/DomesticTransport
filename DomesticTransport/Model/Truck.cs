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
        public int Tonnage { get; set; }


        public int CostOnePoint { get; set; }
        public int CostSecondPoint { get; set; }


    }
}
