namespace DomesticTransport.Model
{
    public class Customer
    {
       public string Id
        {
            get { return _id; }

            set
            {
                _id = new string('0', 10 - value.Length) + value;
            }
        }
        private string _id;
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