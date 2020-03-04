using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Config = DomesticTransport.Properties.Settings;
using Excel = Microsoft.Office.Interop.Excel;

namespace DomesticTransport
{
    class Functions
    {
        // private DomesticTransport.Properties.Settings 

        /// <summary>
        /// Вывести на рабочий лист доставки 
        /// </summary>
        public void SetDelivery()
        {
            string sap = SelectFile();
            if (string.IsNullOrWhiteSpace(sap)) return;
            List<Delivery> deliveries = GetDeliveries(sap);
            if (deliveries != null)
            {

                int i = 0;
                foreach (Delivery delivery in deliveries)
                {
                    ++i;
                    PrintDelivery(delivery, i);
                }

            }
        }

        private void PrintDelivery(Delivery delivery, int number)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject CarrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject InvoiciesTable = deliverySheet.ListObjects["TableInvoicies"];
            if (CarrierTable==null || InvoiciesTable == null)
            {
                MessageBox.Show("Отсутствует таблица");
                return;
            }
            
            ListRow rowCarrier =  CarrierTable.ListRows.AddEx(CarrierTable.ListRows.Count - 1);
            rowCarrier.Range[1, 1].Value = delivery.Carrier.Id ;
            rowCarrier.Range[1, 2].Value = delivery.Carrier.Name;
            rowCarrier.Range[1, 3].Value = delivery.Carrier.Truck.Number;
            rowCarrier.Range[1, 4].Value = delivery.Carrier.Truck.Mark;
            rowCarrier.Range[1, 5].Value = delivery.Carrier.Truck.Tonnage;
            rowCarrier.Range[1, 6].Value = delivery.Carrier.Name;

            ListRow rowInvoice;
            int i=0;
            foreach (Invoice invoice in delivery.Invoices)
            {

            rowInvoice = InvoiciesTable.ListRows.AddEx(CarrierTable.ListRows.Count - 1);
               
                rowInvoice.Range[1, ++i].Value = delivery.Carrier.Id;
                rowInvoice.Range[1, ++i].Value = invoice.Id;
                rowInvoice.Range[1, ++i].Value = invoice.Customer.Id;
                rowInvoice.Range[1, ++i].Value = "" ;
                rowInvoice.Range[1, ++i].Value = invoice.Route;
                rowInvoice.Range[1, ++i].Value = invoice.ItemsCount;
                rowInvoice.Range[1, ++i].Value = invoice.Weight;
                rowInvoice.Range[1, ++i].Value = invoice.Cost;

            }


            //int offset = number > 1 ? 20 : Config.Default.HeaderRow;
            //int tableRows = 10;
            //int tableColumns = 10;

            // Range range = deliverySheet.Range[deliverySheet.Cells[number + offset, 2],
            //                                         deliverySheet.Cells[number + offset + tableRows, 2 + tableColumns]];
            //range.Cells[1, 1].Value = "s11s";
            // range.Cells[2, 1].Value = delivery.DateCreate.ToString(); 


        }

        /// <summary>
        /// Поиск 
        /// </summary>
        /// <param name="sap"></param>
        /// <returns></returns>
        public List<Delivery> GetDeliveries(string sap)
        {
            List<Delivery> deliveries = null;
            Delivery delivery = null;
            Workbook sapBook = null;
            try
            {
                sapBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: sap);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть книгу Excel");
            }
            Worksheet sheet = sapBook.Sheets[1];
            if (sheet != null)
            {


                int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;
                int lastColumn = sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1;
                Range range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[lastRow, lastColumn]];

                //List<Invoice> sapInvoicies = new List<Invoice>();
                //foreach (Range row in range.Rows)
                //{
                //    Invoice invoice = ReadSapRow(row);
                //    sapInvoicies.Add(invoice);
                //}
                //List<Invoice> uniRoute =(List<Invoice>)sapInvoicies.GroupBy(x => x.Route);//.ToList(); //Where(x=> Gro x.Route)
                //foreach(Invoice i in uniRoute)
                //{
                //    deliveries.Add(new Delivery() { Invoices = sapInvoicies.Where(x => x.Route == i.Route).ToList() });
                //}

                foreach (Range row in range.Rows)
                {
                    Invoice invoice = ReadSapRow(row);
                    if (invoice != null)
                    {
                        delivery = deliveries?.Find(d => d.Invoices.Find(i => i.Route == invoice.Route) != null);
                        if (delivery != null)
                        {
                            delivery.Invoices.Add(invoice);
                        }
                        else
                        {
                            delivery = new Delivery();
                            delivery.Invoices = new List<Invoice>();
                            delivery.Invoices.Add(invoice);
                            if (deliveries == null) deliveries = new List<Delivery>();
                            deliveries.Add(delivery);
                        }
                    }
                }

            }
            return deliveries;
        }

        /// <summary>
        /// Собираем из строки выгруза данные для формирования доставки
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Invoice ReadSapRow(Range row)
        {
            string idDocInvoice = row.Cells[1, 3].Value;
            if (string.IsNullOrWhiteSpace(idDocInvoice)) return null;
            Invoice invoice = new Invoice();
            invoice.Id = int.TryParse(idDocInvoice, out int id) ? id : 0;
            string idCusomer = row.Cells[1, 5].Value;
            invoice.Customer = string.IsNullOrWhiteSpace(idCusomer) ? null : new Customer(idCusomer);
            invoice.Route = row.Cells[1, 11].Value;
            string itemsCount = row.Cells[1, 9].Text;
            invoice.ItemsCount = int.TryParse(itemsCount, out int count) ? count : 0;
            string weight = row.Cells[1, 8].Text;
            invoice.Weight = double.TryParse(weight, out double wgt) ? wgt : 0;
            return invoice;
        }


        /// <summary>
        /// Ищет в строке или на листе ячейку с заголовком и возвращает столбец
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="header"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private int GetColumn(Worksheet sh, string header, int row = 0)
        {
            Range findRange = row == 0 ? sh.Cells : sh.Rows[row];
            Range fcell = findRange.Find(What: header, LookIn: XlFindLookIn.xlValues);
            return fcell == null ? 0 : fcell.Column;
        }

        /// <summary>
        ///  Выбрать файл выгрузки SAP
        /// </summary>
        /// <returns></returns>
        public string SelectFile()
        {
            string sapUnload = "";
            string defaultPath = Config.Default.SapUnloadPath;

            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = string.IsNullOrWhiteSpace(defaultPath) ? Directory.GetCurrentDirectory() : defaultPath,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    sapUnload = ofd.FileName;
                    Config.Default.SapUnloadPath = new FileInfo(ofd.FileName).DirectoryName;
                    Config.Default.Save();
                }
            }
            return sapUnload;
        }
    }
}
