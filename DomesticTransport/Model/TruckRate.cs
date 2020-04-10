namespace DomesticTransport.Model
{
    public struct TruckRate
    {
        public string PlaceShipment { get; set; }
        public string PlaceDelivery { get; set; }
        public string City { get; set; }
        public string Company { get; set; }
        public double PriceFirstPoint { get; set; }
        public double PriceAddPoint { get; set; }
        public double Tonnage { get; set; }
        public double TotalDeliveryCost { get; set; }
    }
}
