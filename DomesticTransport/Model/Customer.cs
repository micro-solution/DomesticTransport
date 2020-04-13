namespace DomesticTransport.Model
{
    /// <summary>
    /// Получатель
    /// </summary>
    public class Customer
    {
        /// <summary>
        /// Идентификатор
        /// </summary>
        public string Id
        {
            get => _id;

            set => _id = value.Length < 10 ? new string('0', 10 - value.Length) + value : value;
        }
        private string _id;

        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Город
        /// </summary>
        public string AddresCity { get; set; }

        /// <summary>
        /// Улица
        /// </summary>
        public string AddresStreet { get; set; }

        /// <summary>
        /// Почта
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Телефон
        /// </summary>
        public string Phone { get; set; }

        public Customer() { }
        public Customer(string id)
        {
            Id = id;
        }
    }
}