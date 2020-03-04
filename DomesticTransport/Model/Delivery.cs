using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{

    class Delivery
    {
      public  DateTime DateCreate { get { return DateTime.Now;  } }
        public Carrier Carrier { get; set; }
        public double CostDelivery {
            get
            {
                int pointCount = Invoices.Count;
                double cost = Carrier.Truck.CostOnePoint;

                if (pointCount > 1)
                {
                    cost = Carrier.Truck.CostOnePoint * (pointCount - 1);
                }
                return cost;
            }
        }

        public List<Order> Invoices
        {
            get { return _invoices; }
            set
            {
                if (_invoices == null)
                {
                    _invoices = new List<Order>();                    
                }
                _invoices = value;
             }
        }
        List<Order> _invoices;


    }
}
