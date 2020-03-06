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
        public int Priority { get; set; }
        public string IdClient { get; set; }

    }
}
