
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DomesticTransport
{
    class Email
    {
        public Outlook.Application OutlookApp
        {
            get
            {
                if (_outlookApp == null)
                {
                    _outlookApp = new Outlook.Application();
                }
                return _outlookApp;
            }
        }
        private Outlook.Application _outlookApp;

         ///
        /// <param name="addres">Email</param>        
        /// <param name="subject">Тема</param>
        /// <param name="body">Сообщение</param>
        /// <param name="copyTo">в копию</param>
         public void CreateMessage (string сompany,
                                    string date,
                                   string attachment)
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Сообщения"];
            ListObject tableEmail = messageSheet.ListObjects["TableEmail"];
            string addres = "";
            string stroka = "";
            foreach (Range row in tableEmail.DataBodyRange.Rows)
            {
                addres = row.Text;
                if (addres == сompany)
                {
                    stroka = stroka == "" ? row.Value : $"{stroka}; {addres}";
                }
                //string addres = row == null ? "" : findCell.Offset[0,1].Value;
            }

            string signature = GetHtmlBoby();
            string textMsg = messageSheet.Cells[10, 2].Text;
            string subject = messageSheet.Cells[8, 2].Text;
            string copyTo = messageSheet.Cells[9, 2].Text;
            textMsg = textMsg.Replace("[date]", date);

            string HtmlBody = "< html >< body >< div >" +
                  textMsg +
                   "<br><br>" +
               signature +
               "</div></body></html>";

                try
            {
            OutlookApp.Session.Logon();
            Outlook.MailItem mail = (Outlook.MailItem)OutlookApp.CreateItem(0);
            mail.To = addres;
            mail.Subject ="" ;
            mail.HTMLBody = HtmlBody;
            mail.BCC = "";
            mail.CC = copyTo;        
            mail.Subject = subject;
                mail.Attachments.Add( new Attachment(attachment));
            mail.Display();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }                                        
        }
         

        private string GetHtmlBoby()
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Сообщения"];
            //ListObject tableEmail = messageSheet.ListObjects["TableEmail"];
            Range range = messageSheet.Range["A1:B7"];
            string text="";
              for(int i = 1;i<=range.Rows.Count; i++)
            {
                text = range.Cells[i,2].Value;
                text += "<br>" + text;
            }

            return text;
        }

        public static void WriteReestrSignature() 
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Сообщения"];
            Range range = messageSheet.Range["A1:B7"];

            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey SignatureKey = currentUserKey.CreateSubKey("Sheffler");          
            SignatureKey.SetValue("Ответственное лицо", range.Cells[ 1, 2 ].Text);          
            SignatureKey.SetValue("Компания", range.Cells[ 2, 2 ].Text);                  
            SignatureKey.SetValue("Адрес", range.Cells[ 3, 2 ].Text);    
            SignatureKey.SetValue("Город", range.Cells[ 4, 2 ].Text);           
            SignatureKey.SetValue("Тел", range.Cells[ 5, 2 ].Text);
            SignatureKey.SetValue("Моб", range.Cells[ 6, 2 ].Text);
            SignatureKey.SetValue("Email", range.Cells[ 7, 2 ].Text);

            SignatureKey.Close();

        }
        public static string ReadReestrSignature()
        {
            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey SignatureKey = currentUserKey.OpenSubKey("Sheffler");

            string signature =
            "<br>" + SignatureKey.GetValue("Ответственное лицо").ToString()
            + "<br>" + SignatureKey.GetValue("Компания").ToString()
            + "<br>" + SignatureKey.GetValue("Адрес").ToString()
            + "<br>" + SignatureKey.GetValue("Город").ToString()
            + "<br>" + SignatureKey.GetValue("Тел").ToString()
            + "<br>" + SignatureKey.GetValue("Тел").ToString()
            + "<br>" + SignatureKey.GetValue("Моб").ToString();


            SignatureKey.Close();
            return signature;
        }
    }
}
