using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DomesticTransport
{
    class ShefflerWorkBook 
    {

     private   List<TruckRate> RateList
        {
            get { if (_rateList == null)
                {
                    _rateList = GetTruckRateList();
                }

                return _rateList;
            }
        }
        private List<TruckRate> _rateList;
        public List<DeliveryPoint> RoutesTable
        {
            get
            {
                if (_routes == null)
                {
                    _routes = new List<DeliveryPoint>();
                    Worksheet sheetRoute = GetSheet("Routes");
                    ListObject TableRoutes = sheetRoute?.ListObjects["TableRoutes"];
                    if (TableRoutes != null)
                    {
                        foreach (ListRow row in TableRoutes.ListRows)
                        {
                            Debug.WriteLine(row.Range.Row.ToString());
                            if (row.Range[1, 1].Value == null ||
                                row.Range[1, 2].Value == null ||
                                row.Range[1, 3].Value == null ||
                                row.Range[1, 5].Value == null ||
                                row.Range[1, 9].Value == null) continue;
                            DeliveryPoint route = new DeliveryPoint()
                            {
                                IdRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Id route"].Index].Value.ToString(), out int id) ? id : 0,
                                PriorityRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority route"].Index].Value.ToString(), out int prioritRoute) ? prioritRoute : 0,
                                PriorityPoint = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority point"].Index].Value.ToString(), out int prioritPoint) ? prioritPoint : 0,
                                IdCustomer = row.Range[1, TableRoutes.ListColumns["Получатель материала"].Index].Value.ToString(),
                                City = row.Range[1, TableRoutes.ListColumns["City"].Index].Value.ToString()
                            };
                            _routes.Add(route);
                        }
                    }
                }
                return _routes;
            }
        }
        List<DeliveryPoint> _routes;




        internal Truck GetTruck(double totalWeight, List<DeliveryPoint> mapDelivery)
        {
            List<TruckRate> rates = RateList;
            Truck truck = null;
            List<TruckRate> ratesVariant = new List<TruckRate>();
            foreach (TruckRate rateRow in rates)
            {

                DeliveryPoint findPoint = mapDelivery.Find(m=>m.City ==rateRow.City) ;

                if (rateRow.Tonnage > 0 && rateRow.Tonnage > totalWeight && string.IsNullOrWhiteSpace(findPoint.City))
                {
                    ratesVariant.Add(rateRow);
                }
            }
                        
            if (ratesVariant.Count > 0)
            {
                ratesVariant = ratesVariant.OrderBy(r => r.TotalDeliveryPrise).ToList();
                truck = new Truck(ratesVariant.First());           
            }
            return truck;
        }



        /// <summary>
        /// Вернуть лист по имени
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet GetSheet(string sheetName)
        {
            try
            {
                Worksheet sh = Globals.ThisWorkbook.Sheets[sheetName];
                return sh;
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось получить лист \"{sheetName}\"");
            }

        }

        /// <summary>
        /// Получить таблицу цен перевозчиков
        /// </summary>
        /// <returns></returns>
        internal ListObject GetRateList()
        {
            Worksheet sheetRoute = GetSheet("Rate");
            return sheetRoute?.ListObjects["PriceDelivery"];
        }


        /// <summary>
        /// Получить вес список цен перевозчиков в формате списка         
        /// </summary>
        /// <returns></returns>
        internal List<TruckRate> GetTruckRateList()
        {
            List<TruckRate> ListRate = new List<TruckRate>();
            ListObject rateTable = GetRateList();
            foreach (ListRow row in rateTable.ListRows)
            {
               double tonnage = row.Range[1, rateTable.ListColumns["tonnage, t"].Index].Value ?? 0;
               //string valTonnage = row.Range[1, rate.ListColumns["tonnage, t"].Index].Value.ToString();             
               //double tonnage = double.TryParse(valTonnage, out double t) ? t : 0;

                string valCity = row.Range[1, rateTable.ListColumns["City"].Index].Text;
                valCity = valCity.Trim();

                string valCompany = row.Range[1, rateTable.ListColumns["Company"].Index].Text;
                valCompany = valCompany.Trim();


                if (tonnage > 0 && !string.IsNullOrWhiteSpace(valCity))
                {

                    string strPrice = row.Range[1, rateTable.ListColumns["vehicle"].Index].Text;
                    int priceFirst = int.TryParse(strPrice, out int pf) ? pf : 0;
                    strPrice = row.Range[1, rateTable.ListColumns["add.point"].Index].Text;
                    int priceAdd =   int.TryParse(strPrice, out int pa) ? pa : 0;
                    TruckRate rate = new TruckRate()
                    {
                        City = valCity,
                        Company = valCompany,
                        PriceFirstPoint = priceFirst,
                        PriceAddPoint = priceAdd,
                        PlaceShipment = row.Range[1, 1].Text ,
                        PlaceDelivery = row.Range[1, 2].Text,
                        Tonnage = tonnage

                    } ;

                    ListRate.Add(rate);
                }
            }

            return ListRate;
        }
    }
}
