using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
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

            SapFiles sapFiles = new SapFiles();
            sapFiles.ShowDialog();
            if (sapFiles.DialogResult == DialogResult.OK)
            {
                string sap = "";
                string orders = "";
                try
                {
                    sap = sapFiles.ExportFile;
                    orders = sapFiles.OrderFile;
                }
                catch (Exception ex)
                {
                    return;
                }
                finally
                {
                    sapFiles.Close();
                }

                List<Delivery> deliveries = GetDeliveries(sap, orders);

                if (deliveries != null)
                {
                    foreach (Delivery delivery in deliveries)
                    {
                        PrintDelivery(delivery);
                    }

                }
            }
        }

        private void PrintDelivery(Delivery delivery)
        {
            Worksheet deliverySheet = Globals.ThisWorkbook.Sheets["Delivery"];
            ListObject CarrierTable = deliverySheet.ListObjects["TableCarrier"];
            ListObject InvoiciesTable = deliverySheet.ListObjects["TableInvoicies"];
            if (CarrierTable == null || InvoiciesTable == null)
            {
                MessageBox.Show("Отсутствует таблица");
                return;
            }
            CarrierTable.ListRows.Add(); // if (CarrierTable.ListRows.Count ==0)

            ListRow rowCarrier = CarrierTable.ListRows[CarrierTable.ListRows.Count];
            deliverySheet.Rows[rowCarrier.Range.Offset[1, 0].Row].Insert();
            //rowCarrier.Range.Offset[1, 0].Row] Insert() ;
            System.Windows.Forms.Application.DoEvents();
            // ListRow rowCarrier =  CarrierTable.ListRows.AddEx(CarrierTable.ListRows.Count - 1);
            //  rowCarrier.Range[1, 1].Value = delivery.Carrier?.Id  ?? 0 ;
            // rowCarrier.Range[1, 2].Value = delivery.Carrier?.Name ?? "";
            // rowCarrier.Range[1, 3].Value = delivery.Carrier?.Truck?.Number ?? "";
            //  rowCarrier.Range[1, 4].Value = delivery.Carrier?.Truck?.Mark ?? "";
            rowCarrier.Range[1, 5].Value = delivery.Carrier?.Truck?.Tonnage ?? 0;
            rowCarrier.Range[1, 6].Value = delivery.Carrier?.Name ?? "";

            ListRow rowInvoice;

            foreach (Order invoice in delivery.Invoices)
            {
                // if (CarrierTable.ListRows.Count == 0) CarrierTable.ListRows.AddEx()
                rowInvoice = InvoiciesTable.ListRows.Count == 0 ?
                       InvoiciesTable.ListRows.Add() :
                       InvoiciesTable.ListRows[InvoiciesTable.ListRows.Count]; // InvoiciesTable.ListRows.AddEx(InvoiciesTable.ListRows.Count - 1);
                int column = 0;
                rowInvoice.Range[1, ++column].Value = delivery.Carrier?.Id ?? 0;
                rowInvoice.Range[1, ++column].Value = invoice.Id;
                rowInvoice.Range[1, ++column].Value = invoice?.Customer.Id ?? 0;
                rowInvoice.Range[1, ++column].Value = "";
                rowInvoice.Range[1, ++column].Value = invoice.Route;
                rowInvoice.Range[1, ++column].Value = invoice.PalletsCount;
                rowInvoice.Range[1, ++column].Value = invoice.Weight;
                rowInvoice.Range[1, ++column].Value = invoice.Cost;

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
        public List<Delivery> GetDeliveries(string sap, string orders)
        {
           List<Delivery> deliveries = null;
            Delivery delivery = null;
            Workbook sapBook = null;
            Workbook orderBook = null;
            try
            {
                orderBook = Globals.ThisWorkbook.Application.Workbooks.Open(Filename: orders);
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

                foreach (Range row in range.Rows)
                {
                    Order order = ReadSapRow(row);
                    if (order != null)
                    {
                        Debug.WriteLine(row.Row+ " "+ order.TransportationUnit );
                        if (!string.IsNullOrWhiteSpace(order.TransportationUnit))
                        {
                            List<string>  orderInfo = GetOrderInfo(orderBook.Sheets[1], order.TransportationUnit);
                            if (orderInfo != null)
                            {
                                string costStr = orderInfo[1];
                                costStr= costStr.Replace("Стоимость товаров без НДС:","");
                                costStr = costStr.Replace("RUB", "");
                                costStr = costStr.Replace(".", "");                                
                                costStr = costStr.Trim();
                                order.Cost = double.TryParse(costStr, out double cost) ? cost : 0;

                               // order.Customer.   //Улица , Город
                            }
                        }



                        delivery = deliveries?.Find(d => d.Invoices.Find(i => i.Route == order.Route) != null);
                        if (delivery != null)
                        {
                            delivery.Invoices.Add(order);
                        }
                        else
                        {
                            delivery = new Delivery();
                            delivery.Invoices = new List<Order>();
                            delivery.Invoices.Add(order);
                            if (deliveries == null) deliveries = new List<Delivery>();
                            deliveries.Add(delivery);
                        }
                    }

                    //Комплектовать авто


                }

            }
            return deliveries;
        }

        /// <summary>
        /// Собираем из строки выгруза данные для формирования доставки
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private Order ReadSapRow(Range row)
        {
            /// ТТН
            Order order = new Order();
            Debug.WriteLine(row.Row);
          order.TransportationUnit = row.Cells[1, 4].Text ;
          string idCusomer = row.Cells[1, 5].Text;
          order.Id = row.Cells[1, 7].Text;
            if (string.IsNullOrWhiteSpace(idCusomer) || string.IsNullOrWhiteSpace(order.Id))
            {
                return null;
            }            
                order.Customer = new Customer(idCusomer);                       

            string weight = row.Cells[1, 8].Text;
            order.Weight = double.TryParse(weight, out double wgt) ? wgt : 0;

            string palletsCount = row.Cells[1, 9].Text;
            order.PalletsCount = int.TryParse(palletsCount, out int count) ? count : 0;
           
            order.Route = row.Cells[1, 11].Text;


                //order.Customer = string.IsNullOrWhiteSpace(idCusomer) ? null : new Customer(idCusomer);
            //string idDocInvoice = row.Cells[1, 3].Value;
          //  if (string.IsNullOrWhiteSpace(idDocInvoice)) return null;
          //  order.Id = int.TryParse(idDocInvoice, out int id) ? id : 0;
            return order;
        }

        private List<string>  GetOrderInfo(Worksheet sheet, string transportationUnit)
        {
            Range findRange = sheet.Columns[1];
            //string search = "№ ТТН:" + new string('0', 18 - transportationUnit.Length) + transportationUnit;
            string search = new string('0', 18 - transportationUnit.Length) + transportationUnit;
            Range fcell = findRange.Find(What: search, LookIn: XlFindLookIn.xlValues);
          
            if (fcell == null && fcell.Value.Trim().Contains("ТТН:")) return null;

            int rowStart = fcell.Row;
            int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            int rowEnd = rowStart;
            List<string> info = new List<string>();
            do
            {
                fcell = findRange.Cells[rowEnd++, 1];
                string cellText = fcell.Value;
                cellText.Trim();
                cellText = cellText.Replace("\t","");
                cellText = cellText.Replace(";;;", "");
                if (string.IsNullOrEmpty(cellText.Replace(";", ""))) break;
                info.Add(cellText);                
            }
            while (rowEnd <= lastRow);
            return info; //findRange[findRange.Cells[rowStart, 1], findRange.Cells[rowEnd, 1]];
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
            Range findRange = row == 0 ? sh.UsedRange : sh.Rows[row];
            Range fcell = findRange.Find(What: header, LookIn: XlFindLookIn.xlValues);
            return fcell == null ? 0 : fcell.Column;
        }


    }
}
