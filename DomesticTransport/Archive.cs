using DomesticTransport.Forms;
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
            ListObject total = ShefflerWB.TotalTable;
            //new Functions().




        }

    }
}
