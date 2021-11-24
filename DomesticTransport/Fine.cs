using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomesticTransport
{
    internal class Fine
    {
        private ListObject table;

        public Fine()
        {
            Worksheet fineSheet = Globals.ThisWorkbook.Sheets["Штрафы"];
            table = fineSheet.ListObjects[1];
        }

        /// <summary>
        /// Добавляет дату в лист "Штрафы"
        /// </summary>
        /// <param name="DataFromList">Информация из Листа</param>
        private void AddData(object[,] DataFromList)
        {
            ListRow row = table.ListRows.AddEx();
            int column = 1;
            for (int i = 1; i < DataFromList.GetLength(1); i++)
            {
                row.Range.Cells[1, column].Value = DataFromList[1, i];
                column++;

            }
        }

        /// <summary>
        ///  Применят штраф к выбранному экспедитору
        /// </summary>
        public void SetFine()
        {
            Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;

            if (worksheet.Name=="Отгрузка" || worksheet.Name == "Текущий архив")
            {
                object[,] Data = GetData(worksheet);
                AddData(Data);
            }
            else
            {
                return;
            }
        }
        /// <summary>
        /// Берет данные для данных из текущего листа
        /// </summary>
        /// <param name="FormWorksheet">Текущий лист</param>
        /// <returns></returns>
        private object[,] GetData(Worksheet FormWorksheet)
        {
             int activeCellRow = FormWorksheet.Application.ActiveCell.Row;
            object[,] Data = FormWorksheet.Rows.Range[FormWorksheet.Rows.Cells[activeCellRow, 4], FormWorksheet.Rows.Cells[activeCellRow, 15]].Value;
            return Data;
        }
    }
}
