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

        private string GetVal(ListObject table, int row, string header)
        {
            int col = table.ListColumns[header].Index;
            string value = table.ListRows[row].Range[1, col].Text;
            return value;
        }
        private string GetVal(ListObject table, Range rng, int row, string header)
        {
            int col = table.ListColumns[header].Index;
            string value = table.Range[row, col].Text;
            return value;
        }

    }
}
