using DomesticTransport.Forms;
using DomesticTransport.Model;

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DomesticTransport
{
    class TransportTable
    {
        public DateTime FirstDate { get; set; }
        public DateTime SecondDate { get; set; }
        public string Compny { get; set; }

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
                        Title = "Выберите расположение файла Transport Table",
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
            DateTime dateMax = DateTime.Today;
            SecondDate = dateMax;
            FirstDate = dateMax.AddDays(-(double)dateMax.DayOfWeek);

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

                if (DateTime.Parse(delivery.DateDelivery) > FirstDate &&
                    DateTime.Parse(delivery.DateDelivery) < SecondDate) continue;

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
                    client = client + "-" + order.Customer.Id;
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

                TableSheet.Cells[iRow, ColumnPointCount].Value = delivery.MapDelivery.Count - 1; //Доп точек
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


        public void MessageProvider()
        {
            UnloadOnDate unloadOnDate = new UnloadOnDate();
            unloadOnDate.ShowDialog();
            if (unloadOnDate.DialogResult != System.Windows.Forms.DialogResult.OK)
            { return; }

            Compny = unloadOnDate.Compny;
            FirstDate = unloadOnDate.FirstDate;
            SecondDate = unloadOnDate.SecondDate;
           
            List<Delivery> deliveries = GetDeliveries();
            string fileName = $"TransportTable_{Compny}_{DateTime.Now.ToShortDateString()}" ;
            GenerateAttachmentFile(deliveries, fileName);
        }
        private string GenerateAttachmentFile(List<Delivery> deliveries, string name)
        {
            if (deliveries.Count == 0) return "";

            string folder = GenerateFolder();
           string filename = $"{folder}\\{name}.xlsx";

            Workbook workbook = Globals.ThisWorkbook.Application.Workbooks.Add();

            Worksheet sh = workbook.Sheets[1];
            string[] headers = {
                                "ID",
                                "Перевозчик",
                                "Тип ТС, тонн" ,
                                "Дата подачи ТС" ,
                                "Номер машины",
                                "ФИО водителя",
                                 "Дата доставки",                                
                                "Город доставки" ,
                                "Направление"   ,
                                "Кол-во точек выгрузки",
                                "Номера накладных",
                                "Наименования грузополучателей",                                                               
                                "Брутто вес",
                                "Нетто вес",
                                "Кол-во паллет" ,
                                "Стоимость груза без НДС" ,
                                "Стоимость доставки без НДС",
                                "Номер счёта перевозчика",
                                "Комментарий"
                                };               

            for (int i = 1; i <= headers.Length; i++)
            {
                sh.Cells[1, i].Value = headers[i - 1];
            }
            int row = 2;
            for (int ixDelivery = 0; ixDelivery < deliveries.Count; ixDelivery++)
            {
                Delivery delivery = deliveries[ixDelivery];
                string providerName = delivery.Truck.ProviderCompany.Name;
                if (string.IsNullOrWhiteSpace(providerName)) continue;
                sh.Cells[row, 1].Value = delivery.Driver.Id;
                sh.Cells[row, 7].Value = delivery.Time;
                sh.Cells[row, 19].Value = delivery.Cost;
                sh.Cells[row, 4].Value = delivery.Driver.Name;
                sh.Cells[row, 5].Value = delivery.Driver.CarNumber;
                sh.Cells[row, 6].Value = delivery.Driver.Phone;

                for (int i = 0; i < delivery.Orders.Count; i++)
                {
                    Range rowColor = sh.Range[sh.Cells[row, 1], sh.Cells[row, headers.Length]];
                    Order order = delivery.Orders[i];
                    if (ixDelivery % 2 == 0)
                    {
                        rowColor.Interior.Color = System.Drawing.Color.FromArgb(228, 234, 245);
                    }
                    else
                    {
                        rowColor.Interior.Color = System.Drawing.Color.FromArgb(252, 253, 255);
                    }
                    sh.Cells[row, 2].Value = providerName;
                    sh.Cells[row, 3].Value = delivery.Truck.Tonnage;



                    sh.Cells[row, 8].Value = order.DeliveryPoint.City;
                    sh.Cells[row, 9].Value = order.RouteCity;

                    sh.Cells[row, 10].Value = order.PointNumber;
                    sh.Cells[row, 11].Value = order.Customer.Id;
                    sh.Cells[row, 12].Value = order.TransportationUnit;
                    sh.Cells[row, 13].Value = order.Id;
                    sh.Cells[row, 14].Value = order.Customer.Name ?? "";
                    sh.Cells[row, 15].Value = order.WeightBrutto;
                    sh.Cells[row, 16].Value = order.WeightNetto;
                    sh.Cells[row, 17].Value = order.PalletsCount;
                    sh.Cells[row, 18].Value = order.Cost;
                    row++;
                }
            }
            Range rng = sh.Range[sh.Cells[1, 1], sh.Cells[row - 1, headers.Length]];
            ListObject list =
                sh.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rng,
                XlListObjectHasHeaders: XlYesNoGuess.xlYes);
            workbook.SaveAs(filename);
            workbook.Close();
            return filename;
        }



        private List<Delivery> GetDeliveries()
        {
            List<Delivery> deliveries = new List<Delivery>();
            Open();
            XLRange table = new XLRange();
            table.TableRange = TableSheet.UsedRange;
            int CountRows = table.TableRange.Rows.Count;
            for (int i =1; i < CountRows; i++)
            {
                table.CurrentRowRange = table.TableRange.Rows[i];

                Delivery delivery = GetDeliveryTransportTable(table);
              
                if (delivery.Truck.ProviderCompany.Name != Compny) delivery = null;
                DateTime dateDelivery = DateTime.Parse(delivery.DateDelivery);
                if (dateDelivery > FirstDate && dateDelivery < SecondDate) delivery = null;
                if (delivery != null) deliveries.Add(delivery);
            }
            return deliveries;
        }

        public Delivery GetDeliveryTransportTable(XLRange table)
        {
            Delivery delivery = new Delivery();
            
            delivery.DateDelivery = table.GetValueString("Дата подачи ТС");
            delivery.DateCompleteDelivery = table.GetValueString("Дата доставки");
            delivery.Time = table.GetValueString("Время подачи ТС");
            delivery.Cost = table.GetValueDecimal("Стоимость доставки без НДС");
            delivery.CostProducts = table.GetValueDecimal("Стоимость груза без НДС");
            delivery.TotalPalletsCount = table.GetValueInt("Кол-во паллет");
            delivery.DeliveryPointsCount = table.GetValueInt("Кол-во точек выгрузки");
            delivery.TotalWeightNetto = table.GetValueDouble("Нетто вес");
            delivery.TotalWeightBrutto = table.GetValueDouble("Брутто вес");
            delivery.OrdersInfo = table.GetValueString("Наименования грузополучателей");
            delivery.TtnInfo = table.GetValueString("Номера накладных");
            delivery.RouteName = table.GetValueString("Направление");
            delivery.City = table.GetValueString("Город доставки");
            string providerName = table.GetValueString("Перевозчик");
            if (string.IsNullOrWhiteSpace(delivery.DateDelivery) ||                                            
                                            string.IsNullOrWhiteSpace(providerName)
                                            ) return null;
            Truck truck = new Truck();
            truck.Tonnage = table.GetValueDouble("Тип ТС, тонн");
            truck.ProviderCompany.Name = providerName;
            delivery.Truck = truck;

            string id = table.GetValueString("ID");
            string curNumber = table.GetValueString("Номер машины");
            string phone = table.GetValueString("Телефон водителя");
            string fio = table.GetValueString("ФИО водителя");
            if (string.IsNullOrWhiteSpace(id))
            {
                Driver driver = new Driver()
                {
                    Id = id,
                    CarNumber = curNumber,
                    Name = fio,
                    Phone = phone
                };
                delivery.Driver = driver;
            }

            return delivery;
        }
        /// <summary>
        /// Создать папку для отправки провайдерам
        /// </summary>
        /// <returns></returns>
        private string GenerateFolder()
        {
            string folder = Globals.ThisWorkbook.Path + "\\TransportTable";

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }
    }
}
