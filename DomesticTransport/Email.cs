
using Microsoft.Office.Interop.Excel;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
         public void CreateMessage (string addres,                                                                     
                                   string subject,
                                   string body,
                                   string copyTo)
        {
            string signature = GetHtmlBoby();
                string HtmlBody =
                 "< html >< body >< div >" +
                  body +
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
    }
}
