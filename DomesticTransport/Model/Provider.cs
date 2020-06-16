namespace DomesticTransport.Model
{
    public class Provider
    {
        public Provider() { }
        public int Id { get; set; }
        public string Name
        {
            get => _name;
            set => _name = value;
        }

        private string _name;
        public string Email { get; set; }
        public string Phone { get; set; }
    }
}
