using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
    public class XlTable
    {
        public ListObject ListTable { get; set; }
        public Range TableRange
        {
            get
            {
                if (_range == null && ListTable != null)
                {
                    _range = ListTable.Range;
                }
                return _range;
            }
            set => _range = value;
        }
        Range _range;

        public Range CurrentRowRange
        {
            get
            {
                if (_currentRowRange == null && CurrentRowIndex != 0)
                {
                    if (ListTable != null)
                    {
                        _currentRowRange = TableRange.Rows[CurrentRowIndex].Range;
                    }
                }
                return _currentRowRange;
            }
            set => _currentRowRange = value;
        }
        Range _currentRowRange;

        public int CurrentRowIndex { get; set; }




        public int GetColumn(string header)
        {
            int column = 0;
            if (ListTable != null)
            {
                column = ListTable.ListColumns[header].Index;
            }
            else if (TableRange != null)
            {
                Range findCl = TableRange.Find(header);
                if (findCl != null) column = findCl.Column;
            }
            return column;
        }

        public string GetValueString( string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Text;
            return str;
        }
        public double GetValueDouble( string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Value.ToString();
            double val = double.TryParse(str, out double v) ? v : 0;
            return val;
        }
        public int GetValueInt( string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Value.ToString();
            int val = int.TryParse(str, out int v) ? v : 0;
            return val;
        }
    }
}
