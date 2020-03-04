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

        public List<Invoice> Invoices
        {
            get { return _invoices; }
            set
            {
                if (_invoices == null)
                {
                    _invoices = new List<Invoice>();                    
                }
                _invoices = value;
             }
        }
        List<Invoice> _invoices;


    }
}
