namespace DomesticTransport.Model
{
    public struct TruckRate
    {
        public string PlaceShipment { get; set; }
        public string PlaceDelivery { get; set; }
        public string City { get; set; }
        public string Company { get; set; }
        public int PriceFirstPoint { get; set; }
        public int PriceAddPoint { get; set; }
        public double Tonnage { get; set; }
        public int TotalDeliveryCost { get; set; }
    }
}
