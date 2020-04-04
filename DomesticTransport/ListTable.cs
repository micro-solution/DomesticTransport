using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
public  class ListTable
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Header { get; set; }
        public string TableName { get; set; }

        private ListObject table;

        public string GetColumnSheet()
        {
            return "";
        }
        public string GetValue()
        {

            return "";
        }
    }
}
