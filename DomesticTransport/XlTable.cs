using Microsoft.Office.Interop.Excel;
using System;

namespace DomesticTransport
{
   /// <summary>
   /// Таблица 
   /// </summary>
    public class XLTable
    {
        public ListObject ListTable { get; set; }
     
        /// <summary>
        /// Диапазон таблицы
        /// </summary>
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
        
        /// <summary>
        /// Строка для заполнения
        /// </summary>
        public Range CurrentRowRange
        {
            get
            {
                if (_currentRowRange == null && CurrentRowIndex != 0)
                {
                    if (ListTable != null)
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
                    _currentRowIndex = GetLastRowIndex();
                return _currentRowIndex;
            }
            set => _currentRowIndex = value;
        }
        int _currentRowIndex;


    virtual public int GetColumn(string header)
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

        /// <summary>
        /// Последняя строка таблицы
        /// </summary>
        /// <returns></returns>
        public int GetLastRowIndex()
        {
            int ix = 0;
            ix = ListTable.ListRows.Count;

            for (int i = ix; i > 0; i--)
            {
                string str = ListTable.ListRows[i].Range[1, 1].Text;
                if (str == "")
                { ix = i; }
                else
                {
                    if (ix == i)
                    {
                        ++ix;
                        ListTable.ListRows.Add();
                    }
                    break;    
                }                 
            }

            //Добавить провекрку на пустоту в строке
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
                ListTable.ListRows.Add();
                ix = 1;
            }
            return ListTable.ListRows[ix].Range;
        }

         /// <summary>
         /// Установить последнюю строку таблицы
         /// </summary>
        public void SetCurrentRow()
        {
            CurrentRowRange = GetLastRow();
        }

    }
}
