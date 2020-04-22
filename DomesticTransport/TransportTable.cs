using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DomesticTransport
{
    class TransportTable
    {

        #region Столбцы таблицы
        public const int ColumnId = 1;
        public const int ColumnProvider = 2;
        public const int ColumnCarType = 3;
        public const int ColumnDate = 4;
        public const int ColumnCarNumber = 5;
        public const int ColumnCarDriver = 6;
        public const int ColumnDateDelivery = 7;
        public const int ColumnSity = 8;
        public const int ColumnRoute = 9;
        public const int ColumnPointCount = 10;
        public const int ColumnTTNs = 11;
        public const int ColumnClients = 12;
        public const int ColumnWeightBrutto = 13;
        public const int ColumnWeightNetto = 14;
        public const int ColumnPalleteCount = 15;
        public const int ColumnPriceOrder = 16;
        public const int ColumnPriceDelivery = 17;
        #endregion

        /// <summary>
        /// Путь к файлу
        /// </summary>
        public string FullName
        {
            get
            {
                string path = Properties.Settings.Default.TransportTableFileFullName;
                string defaultPath = Properties.Settings.Default.SapUnloadPath;

                if (!System.IO.File.Exists(path))
                {
                    using (OpenFileDialog fileDialog = new OpenFileDialog()
                    {
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
                            Properties.Settings.Default.TransportTableFileFullName = path;
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

        public TransportTable()
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

        public void ImportDeliveryes(List<Delivery> deliveries)
        {
            int iRow = NextRow;

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
                List<string> sityes = new List<string>();
                List<string> routes = new List<string>();
                List<string> ttns = new List<string>();
                List<string> clients = new List<string>();

                double weightNetto = 0;
                double weightBrutto = 0;
                double palletCount = 0;
                double priceOrder = 0;

                foreach (Order order in delivery.Orders)
                {
                    weightNetto += order.WeightNetto;
                    weightBrutto += order.WeightBrutto;
                    palletCount += order.PalletsCount;
                    priceOrder += order.Cost;


                    sityes.Add(order.DeliveryPoint.City);
                    routes.Add(order.RouteCity);
                    ttns.Add(order.TransportationUnit);

                    string client = order.Customer.Name;
                    client = client.Substring(0, client.IndexOf('/'));
                    client = client.Replace(",", "");

                    clients.Add(client);
                }

                sityes = sityes.Distinct().ToList();
                routes = routes.Distinct().ToList();
                ttns = ttns.Distinct().ToList();
                clients = clients.Distinct().ToList();

                TableSheet.Cells[iRow, ColumnId].Value = delivery.Driver.Id;
                TableSheet.Cells[iRow, ColumnProvider].Value = delivery.Truck.ProviderCompany.Name;
                TableSheet.Cells[iRow, ColumnCarType].Value = delivery.Truck.Tonnage;
                TableSheet.Cells[iRow, ColumnDate].Value = delivery.DateDelivery;
                TableSheet.Cells[iRow, ColumnCarNumber].Value = delivery.Driver.CarNumber;
                TableSheet.Cells[iRow, ColumnCarDriver].Value = delivery.Driver.Name;

                TableSheet.Cells[iRow, ColumnSity].Value = string.Join(", ", sityes.Select(x => x.ToString()));
                TableSheet.Cells[iRow, ColumnRoute].Value = string.Join(", ", routes.Select(x => x.ToString()));
                TableSheet.Cells[iRow, ColumnPointCount].Value = delivery.MapDelivery.Count;
                TableSheet.Cells[iRow, ColumnTTNs].Value = string.Join(", ", ttns.Select(x => x.ToString()));
                TableSheet.Cells[iRow, ColumnClients].Value = string.Join(", ", clients.Select(x => x.ToString()));

                TableSheet.Cells[iRow, ColumnWeightBrutto].Value = weightBrutto;
                TableSheet.Cells[iRow, ColumnWeightNetto].Value = weightNetto;
                TableSheet.Cells[iRow, ColumnPalleteCount].Value = palletCount;
                TableSheet.Cells[iRow, ColumnPriceOrder].Value = priceOrder;
                TableSheet.Cells[iRow, ColumnPriceDelivery].Value = delivery.Cost;

                iRow++;
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
