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

        public static void LoadToArhive()
        {
            Functions fn = new Functions();
            List<Delivery> deliveries = fn.GetDeliveriesFromTotal();

            if (!CheckArchive(deliveries))
            {
                //Проверить повторение заказов
                CpopyTotalPastArchive();
            }
            else
            {
                PrintArchive(deliveries);
            }
        }

        static bool CheckArchive(List<Delivery> deliveries)
        {
            bool chk = false;
            foreach (Delivery delivery in deliveries)
            {
                chk = CheckDelivery(delivery);
                if (chk) break;
            }
            return chk;
        }
        static bool CheckDelivery(Delivery delivery)
        {
            bool chk = false;
            ListObject archiveTable = ShefflerWB.ArchiveTable;
            foreach (ListRow archiveRow in archiveTable.ListRows)
            {
                string idOrder = archiveRow.Range[1, archiveTable.ListColumns["Номер поставки"].Index].Text;
                idOrder = idOrder.Length < 10 ? new string('0', 10 - idOrder.Length) + idOrder : idOrder;
                chk = delivery.Orders.Find(a => a.Id == idOrder) != null;
                if (chk) break;
            }
            return chk;
        }


        static void CpopyTotalPastArchive()
        {
            ShefflerWB.TotalTable.DataBodyRange.Copy();
            ListObject arh = ShefflerWB.ArchiveTable;
            Range rng = arh.ListRows[arh.ListRows.Count].Range[1, 1];
            rng.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
        }
        static void ClearArchive()
        {

        }

        private static void PrintArchive(List<Delivery> deliveries)
        {
            XLTable tableArchive = new XLTable();
            tableArchive.ListTable = ShefflerWB.ArchiveTable;

            bool chk = false;
            foreach (Delivery delivery in deliveries)
            {
                chk = CheckDelivery(delivery);
                if (chk) {
                    //delivery 
                  
                }
            }
        }
    }
}
