using DomesticTransport.Forms;
using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
    class Archive
    {
        public Archive() { }

        public static void UnoadFromArhive()
        {
            new UnloadArchive().ShowDialog();
        }





        internal static void LoadToArhive()
        {
            //ListObject total = ShefflerWB.TotalTable;
            Functions fn = new Functions();            
            List<Delivery> deliveries = fn.GetDeliveriesFromTotal();
            
            PrintArchive(deliveries);
                          //Проверить повторение заказов
            CpopyTotalPastArchive();
            ///


        }                          

        static void CpopyTotalPastArchive()
        {
            //
             ShefflerWB.TotalTable.DataBodyRange.Copy();
            ListObject arh = ShefflerWB.ArchiveTable;
            Range rng = arh.ListRows[arh.ListRows.Count].Range[1,1] ;
            rng.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
        }
        static void ClearArchive()
        {
        }

            private static void PrintArchive(List<Delivery> deliveries)
        {
            XLTable tableArchive = new XLTable();
            tableArchive.ListTable = ShefflerWB.ArchiveTable;

        }
    }
}
