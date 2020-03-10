using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport.Model
{
    class Delivery
    {
    public static int Number=0;
        public DateTime DateCreate { get { return DateTime.Now; } }
        public Carrier Carrier
        {
            get { if (_carrier == null)
                {
                    _carrier = new Carrier();
                }
                    return _carrier;
            }
          private set { _carrier = value; }
        }
        Carrier _carrier;


        public double CostDelivery
        {
            get
            {
                int pointCount = Orders.Count;
                double cost = Truck.CostOnePoint;

                if (pointCount > 1)
                {
                    cost = Truck.CostOnePoint * (pointCount - 1);
                }
                return cost;
            }
        }
        public double TotalWeight
        {
            get
            {
                double sum = 0;
                Orders.ForEach(x => sum += x.WeightNetto);
                return sum;
            }
        }
        public List<Order> Orders
        {
            get
            {
                if (_orders == null)
                {
                    _orders = new List<Order>();
                }
                return _orders;
            }
         private set
            {

                _orders = value;
            }
        }
        List<Order> _orders;

        public Delivery AddOrder(Order order)
        {
            Delivery delivery;
            if (this.TotalWeight + order.WeightNetto > 20200)
            {
                delivery = new Delivery();
                delivery.AddOrder(order);
            }
            else
            {
                Orders.Add(order);
                delivery = this;
            }
            return delivery;
        }

        /// <summary>
        /// Точки доставки
        /// </summary>
        public List<DeliveryPoint> MapDelivery
        {
            get
            {
                List<DeliveryPoint> dp = (from r in Orders
                                          select r.DeliveryPoint                                         
                                          ).Distinct().ToList();
                dp.OrderBy(x => x.PriorityRoute).ThenBy(y => y.PriorityPoint);
                return dp;
            }
        }


        public Truck Truck
        {
            get
            {
                ShefflerWorkBook workBook = new ShefflerWorkBook();
                _truck = workBook.GetTruck(TotalWeight, MapDelivery);
                return _truck;
            }
            private set { _truck = value; }
        }
        Truck _truck;

        Delivery() 
        {
            Number++;        
        }
        ~Delivery()
        {
            Number--;
        }

    }
}
