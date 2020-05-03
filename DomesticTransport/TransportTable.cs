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
        public const int ColumnAccountNumber = 18;
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
        private Worksheet TableSheetDeline;

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
        /// <summary>
        /// Следующая (пустая) строка на листе деловые линии
        /// </summary>
        private int NextRowDeline
        {
            get
            {
                return TableSheetDeline.UsedRange.Row + TableSheetDeline.UsedRange.Rows.Count;
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
            TableSheetDeline = Workbook.Worksheets[2];
        }

        /// <summary>
        /// Импорт данных из архива
        /// </summary>
        /// <param name="deliveries"></param>
        public void ImportDeliveryes(List<Delivery> deliveries)
        {
            int iRow = NextRow;
            int iRowDeline = NextRowDeline;
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
                    if (client.IndexOf('/') != -1) client = client.Substring(0, client.IndexOf('/'));
                    client = client.Replace(",", "");
                    client = client + "-" + order.Customer.Id;
                    clients.Add(client);
                }

                sityes = sityes.Distinct().ToList();
                routes = routes.Distinct().ToList();
                ttns = ttns.Distinct().ToList();
                clients = clients.Distinct().ToList();

                Worksheet worksheet;
                int row;
                if (delivery.Truck.ProviderCompany.Name == "Деловые линии") 
                {
                    worksheet = TableSheetDeline;
                    row = iRowDeline;
                }
                else
                {
                    worksheet = TableSheet;
                    row = iRow;
                }

                worksheet.Cells[row, ColumnId].Value = delivery.Driver.Id;
                worksheet.Cells[row, ColumnProvider].Value = delivery.Truck.ProviderCompany.Name;
                worksheet.Cells[row, ColumnCarType].Value = delivery.Truck.Tonnage;
                worksheet.Cells[row, ColumnDate].Value = delivery.DateDelivery;
                worksheet.Cells[row, ColumnCarNumber].Value = delivery.Driver.CarNumber;
                worksheet.Cells[row, ColumnCarDriver].Value = delivery.Driver.Name;

                worksheet.Cells[row, ColumnSity].Value = string.Join(", ", sityes.Select(x => x.ToString()));
                worksheet.Cells[row, ColumnRoute].Value = string.Join(", ", routes.Select(x => x.ToString()));

                worksheet.Cells[row, ColumnPointCount].Value = clients.Count;
                worksheet.Cells[row, ColumnTTNs].Value = string.Join(", ", ttns.Select(x => x.ToString()));
                worksheet.Cells[row, ColumnClients].Value = string.Join(", ", clients.Select(x => x.ToString()));

                worksheet.Cells[row, ColumnWeightBrutto].Value = weightBrutto;
                worksheet.Cells[row, ColumnWeightNetto].Value = weightNetto;
                worksheet.Cells[row, ColumnPalleteCount].Value = palletCount;
                worksheet.Cells[row, ColumnPriceOrder].Value = priceOrder;
                worksheet.Cells[row, ColumnPriceDelivery].Value = delivery.Cost;

                if (worksheet == TableSheet)
                {
                    iRow++;
                }
                else
                {
                    iRowDeline++;
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

        /// <summary>
        /// Закрыть без сохранения
        /// </summary>
        public void Close()
        {
            Workbook.Close(false);
            TableSheet = null;
            Workbook = null;
        }

        /// <summary>
        /// Отправка отчета провайдеру
        /// </summary>
        public void MessageProvider(DateTime dateStart, DateTime dateEnd, string provider)
        {
            // Создаем копию листа и сохраняем в отдельную книгу
            CreateReportToProvider(dateStart, dateEnd, provider);
            string message = Properties.Settings.Default.ProviderMessageReport;
            string subject = Properties.Settings.Default.ProviderSubjectReport;

            message = message.Replace("[provider]", provider);
            message = message.Replace("[dateStart]", dateStart.ToString("d"));
            message = message.Replace("[dateEnd]", dateEnd.ToString("d"));
            subject = subject.Replace("[provider]", provider);
            subject = subject.Replace("[dateStart]", dateStart.ToString("d"));
            subject = subject.Replace("[dateEnd]", dateEnd.ToString("d"));

            string path = Globals.ThisWorkbook.Path + "\\MailToProvider\\";
            List<string> attachments = new List<string>
            {
                path + provider + "_" + dateStart.ToString("d") + "-" + dateEnd.ToString("d") + ".xlsx"
            };

            Email email = new Email();
            email.MailToProvider(provider, subject, message, attachments, Email.TypeSend.Display);
            Close();
        }

        /// <summary>
        /// Подготовка отчета провайдеру
        /// </summary>
        /// <param name="dateStart"></param>
        /// <param name="dateEnd"></param>
        /// <param name="provider"></param>
        private void CreateReportToProvider(DateTime dateStart, DateTime dateEnd, string provider)
        {
            // Создаем копию листа и сохраняем в отдельную книгу
            string path = Globals.ThisWorkbook.Path + "\\MailToProvider\\";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string attachment = path + provider + "_" + dateStart.ToString("d") + "-" + dateEnd.ToString("d") + ".xlsx";
            TableSheet.Copy();
            Workbook workbook = Globals.ThisWorkbook.Application.ActiveWorkbook;
            workbook.SaveAs(attachment, XlFileFormat.xlWorkbookDefault);
            Range rangeToDelete = workbook.Sheets[1].cells[NextRow, ColumnDate];

            // удаляем лишние строки
            for (int i = 2; i < NextRow; i++)
            {
                DateTime date;

                Range rangeDate = workbook.Sheets[1].cells[i, ColumnDate];
                if (string.IsNullOrEmpty(rangeDate.Text) && IsDate(rangeDate.Value))
                {
                    date = DateTime.Parse(rangeDate.Value.ToString());
                }
                else
                {
                    if (!DateTime.TryParse(rangeDate.Text, out date))
                    {
                        rangeToDelete = Globals.ThisWorkbook.Application.Union(rangeToDelete, rangeDate);
                    }
                }

                if (date < dateStart || date > dateEnd || workbook.Sheets[1].cells[i, ColumnProvider].Text != provider)
                {
                    rangeToDelete = Globals.ThisWorkbook.Application.Union(rangeToDelete, rangeDate);
                }
            }
            rangeToDelete.EntireRow.Delete();
            workbook.Close(true);
        }

        private bool IsDate(object attemptedDate)
        {
            bool success;
            if (attemptedDate == null) return false;
            try
            {
                DateTime dtParse = DateTime.Parse(attemptedDate.ToString());
                success = true; // это дата
            }
            catch
            {
                success = false; // это не дата
            }

            return success;
        }

        /// <summary>
        /// Получение данных из писем провайдеров
        /// </summary>
        public void GetDataFromProviderFiles()
        {
            string path = Globals.ThisWorkbook.Path + "\\MailFromProviders\\" + DateTime.Today.ToString("dd.MM.yyyy") + '\\';
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Папка " + path + " отсутствует");
                return;
            }
            string[] files = Directory.GetFiles(path);
            if (files.Length == 0) return;

            ProcessBar pb = ProcessBar.Init("Сканирование вложений", files.Length, 1, "Получение данных провайдера");
            pb.Show();

            int i = 0;
            foreach (string file in files)
            {
                i++;
                FileInfo fileInfo = new FileInfo(file);
                if (pb.Cancel) break;
                pb.Action($"Вложение {i} из {pb.Count} {fileInfo.Name} ");

                if (!file.Contains(".xls")) { continue; }
                ReadMessageFile(file);
            }
            pb.Close();
        }

        /// <summary>
        /// Импорт данных из писем провайдеров
        /// </summary>
        /// <param name="file"></param>
        public void ReadMessageFile(string file)
        {
            List<string> IdNotFound = new List<string>();
            Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: file);
            Worksheet sh = wb.Sheets[1];
            FileInfo fileInfo = new FileInfo(file);
            try
            {
                if (sh.Cells[1, 1].Text != TableSheet.Cells[1, 1].Text)
                {
                    return;
                }

                int lastRow = sh.UsedRange.Row + sh.UsedRange.Rows.Count;
                for (int i = 2; i <= lastRow; i++)
                {
                    Range dateDelivery = sh.Cells[i, ColumnDateDelivery];
                    Range accountNumber = sh.Cells[i, ColumnAccountNumber];
                    string id = sh.Cells[i, ColumnId].Text;

                    if (string.IsNullOrEmpty(id)) continue;

                    Range columnId = TableSheet.Columns[ColumnId];
                    Range findIdRow = columnId.Find(id);

                    if (findIdRow == null)
                    {
                        IdNotFound.Add(id);
                        Range rowNotFound = sh.Rows[i];
                        rowNotFound.Interior.Color = 65535;
                        continue;
                    }

                    TableSheet.Cells[findIdRow.Row, ColumnDateDelivery].Value = dateDelivery.Value;
                    TableSheet.Cells[findIdRow.Row, ColumnAccountNumber].Value = accountNumber.Value;

                    TableSheet.Cells[findIdRow.Row, ColumnDateDelivery].Interior.Color = 5296274;
                    TableSheet.Cells[findIdRow.Row, ColumnAccountNumber].Interior.Color = 5296274;               
                }

                if (IdNotFound.Count > 0)
                {
                    MessageBox.Show("В файле " + fileInfo.Name + " есть строки, которые не удалось сопоставить автоматически. Они были выделены желтой заливкой в файле", 
                                    "Обратите внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch
            {
                throw new System.Exception("Не удалось прочитать таблицу в файле " + fileInfo.Name);
            }
            finally
            {
                Globals.ThisWorkbook.Application.DisplayAlerts = false;
                wb.Close(true);
                Globals.ThisWorkbook.Application.DisplayAlerts = true;
            }
        }
    }
}
