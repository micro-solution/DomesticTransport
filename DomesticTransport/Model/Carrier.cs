namespace DomesticTransport.Model
{
   public class Carrier
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public string CarNumber { get; set; }
        Provider ShippingCompany { get; set; }


    }
}
