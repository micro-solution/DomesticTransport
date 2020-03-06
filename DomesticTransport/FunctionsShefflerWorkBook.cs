using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
    class ShefflerWorkBook : IDisposable
    {

        List<DeliveryPoint> RoutesTable
        {
            get
            {
                if (_routes == null)
                {
                    Worksheet sheetRoute = Globals.ThisWorkbook.Sheets["Routes"];
                    ListObject TableRoutes = sheetRoute?.ListObjects["TableRoutes"];
                    if (TableRoutes != null)
                    {
                        foreach (ListRow row in TableRoutes.ListRows)
                        {
                            DeliveryPoint route = new DeliveryPoint()
                            {
                                Id = int.TryParse(row.Range[1, TableRoutes.ListColumns["Id route"].Index].Value, out int id) ? id : 0,
                                PriorityRoute = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority route"].Index].Value, out int prioritRoute) ? prioritRoute : 0,
                                Priority = int.TryParse(row.Range[1, TableRoutes.ListColumns["Priority point"].Index].Value, out int prioritPoint) ? prioritPoint : 0,
                                IdClient = row.Range[1, TableRoutes.ListColumns["Получатель материала"].Index].Value
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
    }


  
          



    
}
