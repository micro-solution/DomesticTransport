using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    /// <summary>
    /// Строка таблицы Routes (Точка Доставки или Получатель)
    /// </summary>
    struct DeliveryPoint
    {
        public int Id { get; set; }
        public int PriorityRoute { get; set; }
        public int PriorityPoint { get; set; }

        public string IdCustomer
        {
            get { return _idCustomer; }
            set
            {
                _idCustomer = new string('0', 10 - value.Length) + value;
            }
        }
        private string _idCustomer;

        public string City { get; set; }

    }
}
