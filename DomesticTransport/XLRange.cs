using Microsoft.Office.Interop.Excel;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
    /// <summary>
    ///Таблица в виде диапазона
    /// </summary>
    class XLRange : XLTable
    {
        Range HeadRow
        {
            get
            {
                return TableRange.Rows[1];
            }
        }

        public Range CurrentRowRange
        {
            get
            {
                if (_currentRowRange == null )
                {
                    if (TableRange != null)
                    {
                        _currentRowRange = TableRange.Rows[CurrentRowIndex];
                    }
                }
                return _currentRowRange;
            }
            set => _currentRowRange = value;
        }
        Range _currentRowRange;
       
        override public int GetColumn(string header)
        {
            int column = 0;
            if (HeadRow != null)
            {
                for (int i = 1; i <= HeadRow.Columns.Count; i++)
                {

                    string headCell = HeadRow.Cells[1, i].Text;
                    if (headCell == header)
                    {
                        column = i;
                        break;
                    }
                }
            }
            else { throw new Exception("Не найден столбец"); }
            return column;
        }
    }
}
