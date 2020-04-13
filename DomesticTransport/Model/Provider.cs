namespace DomesticTransport.Model
{
    public class Provider
    {
        public Provider() { }
        public int Id { get; set; }
        public string Name
        {
            get =>
                //if (string.IsNullOrWhiteSpace(_name))
                //{
                //    _name = "Деловые линии";
                //}                       
                _name;
            set => _name = value;
        }
        string _name;
        public string Email { get; set; }
        public string Phone { get; set; }
    }
}
