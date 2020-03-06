namespace DomesticTransport.Model
{
    public class Customer
    {
        public string Id { get; set; }
        public string Addres { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }


        Customer() { }
     public   Customer(string id)
        {
            Id = id;
        }
    }
}