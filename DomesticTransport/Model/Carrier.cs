namespace DomesticTransport.Model
{
    class Carrier
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }


        Provider ShippingCompany { get; set; }


    }
}
