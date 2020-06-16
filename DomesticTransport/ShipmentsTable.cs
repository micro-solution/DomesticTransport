using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace DomesticTransport
{
    internal class ShipmentsTable
    {
        #region Столбцы таблицы
        public const int ColumnDate = 1;
        public const int ColumnTime = 2;
        public const int ColumnId = 3;
        public const int ColumnProvider = 4;
        public const int ColumnCarType = 5;
        public const int ColumnDriver = 6;
        public const int ColumnCarNumber = 7;
        public const int ColumnDriverPhone = 8;
        public const int ColumnDeliveryNumber = 9;
        public const int ColumnSity = 10;
        public const int ColumnRoute = 11;
        public const int ColumnPoint = 12;
        public const int ColumnClientId = 13;
        public const int ColumnTTN = 14;
        public const int ColumnOrderNumber = 15;
        public const int ColumnClient = 16;
        public const int ColumnWeightBrutto = 17;
        public const int ColumnWeightNetto = 18;
        public const int ColumnPalleteCount = 19;
        public const int ColumnPriceOrder = 20;
        public const int ColumnPriceDelivery = 21;
        #endregion

        /// <summary>
        /// Путь к файлу
        /// </summary>
        public string FullName
        {
            get
            {
                string path = Properties.Settings.Default.ShipmentFileFullName;
                string defaultPath = Properties.Settings.Default.SapUnloadPath;

                if (!System.IO.File.Exists(path))
                {
                    using (OpenFileDialog fileDialog = new OpenFileDialog()
                    {
                        Title = "Выберите расположение файла Shipments",
                        DefaultExt = "*.xls*",
                        CheckFileExists = true,
                        InitialDirectory = string.IsNullOrWhiteSpace(defaultPath) ? Directory.GetCurrentDirectory() : defaultPath,
                        ValidateNames = true,
                        Multiselect = false,
                        Filter = "Excel|*.xls*"
                    })
                    {
                        if (fileDialog.ShowDialog() == DialogResult.OK)
                        {
                            path = fileDialog.FileName;
                            Properties.Settings.Default.ShipmentFileFullName = path;
                            Properties.Settings.Default.Save();
                        }
                    }
                }
                return path;
            }
        }

        public Workbook Workbook;
        private Worksheet TableSheet;

        /// <summary>
        /// Следующая (пустая) строка 
        /// </summary>
        private int NextRow
        {
            get
            {
                return TableSheet.UsedRange.Row + TableSheet.UsedRange.Rows.Count;
            }
        }

        public ShipmentsTable()
        {
            Open();
        }

        /// <summary>
        /// Открытие книги
        /// </summary>
        public void Open()
        {
            if (!File.Exists(FullName)) return;
            Workbook = Globals.ThisWorkbook.Application.Workbooks.Open(FullName);
            TableSheet = Workbook.Worksheets[1];
        }

        /// <summary>
        /// Импорт доставок в таблицу
        /// </summary>
        /// <param name="deliveries"></param>
        public void ImportDeliveryes(List<Delivery> deliveries)
        {
            int iRow = NextRow;
            DateTime dateMax = DateTime.Today;
            dateMax = dateMax.AddDays(-(double)dateMax.DayOfWeek);

            Forms.ProcessBar pb = Forms.ProcessBar.Init("Экспорт в Transport Table", deliveries.Count, 1, "Экспорт");
            if (pb == null) return;
            pb.Show();
            int i = 0;
            foreach (Delivery delivery in deliveries)
            {
                if (pb == null) return;
                i++;
                if (pb.Cancel) break;
                pb.Action($"Доставка {i} из {pb.Count}");

                if (DateTime.Parse(delivery.DateDelivery) > dateMax) continue;
                TableSheet.Cells[iRow, ColumnPriceDelivery].Value = delivery.Cost;

                foreach (Order order in delivery.Orders)
                {
                    TableSheet.Cells[iRow, ColumnId].Value = delivery.Driver.Id;
                    TableSheet.Cells[iRow, ColumnProvider].Value = delivery.Truck.ProviderCompany.Name;
                    TableSheet.Cells[iRow, ColumnCarType].Value = delivery.Truck.Tonnage;
                    TableSheet.Cells[iRow, ColumnDriver].Value = delivery.Driver.Name;
                    TableSheet.Cells[iRow, ColumnDriverPhone].Value = delivery.Driver.Phone;
                    TableSheet.Cells[iRow, ColumnCarNumber].Value = delivery.Driver.CarNumber;

                    TableSheet.Cells[iRow, ColumnDate].Value = delivery.DateDelivery;
                    TableSheet.Cells[iRow, ColumnTime].Value = delivery.Time;
                    TableSheet.Cells[iRow, ColumnDeliveryNumber].Value = order.DeliveryNumber;
                    TableSheet.Cells[iRow, ColumnSity].Value = order.DeliveryPoint.City;
                    TableSheet.Cells[iRow, ColumnRoute].Value = order.RouteCity;
                    TableSheet.Cells[iRow, ColumnPoint].Value = order.PointNumber;
                    TableSheet.Cells[iRow, ColumnClientId].Value = order.Customer.Id;
                    TableSheet.Cells[iRow, ColumnTTN].Value = order.TransportationUnit;
                    TableSheet.Cells[iRow, ColumnOrderNumber].Value = order.Id;
                    TableSheet.Cells[iRow, ColumnClient].Value = order.Customer.Name;

                    TableSheet.Cells[iRow, ColumnWeightBrutto].Value = order.WeightBrutto;
                    TableSheet.Cells[iRow, ColumnWeightNetto].Value = order.WeightNetto;
                    TableSheet.Cells[iRow, ColumnPalleteCount].Value = order.PalletsCount;
                    TableSheet.Cells[iRow, ColumnPriceOrder].Value = order.Cost;
                    iRow++;
                }
            }
            pb.Close();
        }

        /// <summary>
        /// Сохранение и выход
        /// </summary>
        public void SaveAndClose()
        {
            Workbook.Close(true);
            TableSheet = null;
            Workbook = null;
        }
    }
}
