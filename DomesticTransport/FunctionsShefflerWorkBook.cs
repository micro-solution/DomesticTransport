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
    class ShefflerWorkBook : IDisposable
    {
        


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
                                Id = int.TryParse(row.Range[1, TableRoutes.ListColumns["Id route"].Index].Value.ToString(), out int id) ? id : 0,
                                PriorityRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority route"].Index].Value.ToString(), out int prioritRoute) ? prioritRoute : 0,
                                PriorityPoint = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority point"].Index].Value.ToString(), out int prioritPoint) ? prioritPoint : 0,
                                IdCustomer = row.Range[1, TableRoutes.ListColumns["Получатель материала"].Index].Value.ToString() ,
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
       
       

        void IDisposable.Dispose()
        {

        }

        internal Truck GetTruck(double totalWeight, List<DeliveryPoint> mapDelivery)
        {

            Truck truck = new Truck();
            return truck;
        }

    private Worksheet GetSheet(string sheetName)
    {
            try
            {
            Worksheet sh = Globals.ThisWorkbook.Sheets[sheetName];
            return sh;
            }
            catch(Exception ex)
            {                
                throw new Exception($"Не удалось получить лист \"{sheetName}\"");
            }
            
    }

    }
    
}
