using Microsoft.Office.Interop.Excel;

using System;

namespace DomesticTransport
{
    /// <summary>
    ///Таблица в виде диапазона
    /// </summary>
    class XLRange
    {

        /// <summary>
        /// Диапазон таблицы
        /// </summary>
        public Range TableRange { get; set; }
        //    Range _range;

        /// <summary>
        /// Строка для заполнения
        /// </summary>
        public Range CurrentRowRange
        {
            get
            {
                if (_currentRowRange == null && CurrentRowIndex != 0)
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






        /// <summary>
        ///Номер последней строки
        /// </summary>
        public int CurrentRowIndex
        {
            get
            {
                if (_currentRowIndex == 0)

                    if (TableRange != null)
                    {
                        _currentRowIndex = GetLastRowIndex();
                    }
                return _currentRowIndex;
            }
            set => _currentRowIndex = value;
        }
        int _currentRowIndex;



        //Get
        public string GetValueString(string header)
        {
            int column = GetColumn(header);
            return CurrentRowRange.Cells[1, column].Text;
        }
        public double GetValueDouble(string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Value?.ToString() ?? "";
            double val = double.TryParse(str, out double v) ? v : 0;
            return val;
        }
        public decimal GetValueDecimal(string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Value?.ToString() ?? "";
            decimal val = decimal.TryParse(str, out decimal v) ? v : 0;
            return val;
        }
        public int GetValueInt(string header)
        {
            int column = GetColumn(header);
            string str = CurrentRowRange.Cells[1, column].Value?.ToString() ?? "";
            int val = int.TryParse(str, out int v) ? v : 0;
            return val;
        }


        ///  Set
        public void SetValue(string header, string Value)
        {
            int column = GetColumn(header);
            CurrentRowRange.Cells[1, column].Value = Value;

        }
        public void SetValue(string header, int Value)
        {
            int column = GetColumn(header);
            CurrentRowRange.Cells[1, column].Value = Value;

        }
        public void SetValue(string header, double Value)
        {
            int column = GetColumn(header);
            CurrentRowRange.Cells[1, column].Value = Value;

        }
        public void SetValue(string header, decimal Value)
        {
            int column = GetColumn(header);
            CurrentRowRange.Cells[1, column].Value = Value;

        }

        public int GetLastRowIndex()
        {
            int ix = TableRange.Rows.Count;
            for (int i = ix; i > 0; i--)
            {
                string str = TableRange.Cells[i, 1].Text;
                if (str == "")
                { ix = i; }
                else
                {
                    ++ix;
                    break;
                }
            }

            return ix;
        }



        /// <summary>
        /// Найти\добавить последнюю строку таблицы
        /// </summary>
        public Range GetLastRow()
        {
            int ix = GetLastRowIndex();
            if (ix == 0)
            {
                ix = 1;
            }
            Worksheet sh = (Worksheet)TableRange.Parent;
            return sh.Rows[ix].Range;
        }

        /// <summary>
        /// Установить последнюю строку таблицы
        /// </summary>
        public void SetCurrentRow()
        {
            CurrentRowRange = GetLastRow();
        }


        /// <summary>
        /// 
        /// -------------------------
        /// </summary>
        Range HeadRow
        {
            get
            {
                return TableRange.Rows[1];
            }
        }



        public int GetColumn(string header)
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
