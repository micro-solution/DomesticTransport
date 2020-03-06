namespace DomesticTransport.Model
{
    public class Customer
    {
        public string Id { get; set; }
        public string AddresCity { get; set; }
        public string AddresStreet { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }



        Customer() { }
     public   Customer(string id)
        {
            Id = id;
        }
    }
}