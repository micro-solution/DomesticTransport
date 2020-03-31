
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
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
           ListObject tableEmail = messageSheet.ListObjects["TableEmail"];
            string addres = "";
            string stroka = "";
            foreach (Range row in tableEmail.DataBodyRange.Rows)
            {
                if (row.Cells[1, 2].Text == сompany)
                {
                    addres = row.Cells[1,2].Text;
                    stroka = stroka == "" ? addres : $"{stroka}; {addres}";
                }                            
            }


            string signature = ReadReestrSignature(); 
            string textMsg = messageSheet.Cells[10, 2].Text;
            string subject = messageSheet.Cells[8, 2].Text;
            string copyTo = messageSheet.Cells[9, 2].Text;
            textMsg = textMsg.Replace("[date]", date);
            string HtmlBody =
                  textMsg +
                   "<br><br>" +
               signature;            
            try
            {                                   
            OutlookApp.Session.Logon();
            Outlook.MailItem mail = (Outlook.MailItem)OutlookApp.CreateItem(0);
            mail.To = addres;             
            mail.HTMLBody = HtmlBody;
            mail.BCC = "";
            mail.CC = copyTo;        
            mail.Subject = subject;
            mail.Attachments.Add(attachment, Outlook.OlAttachmentType.olByValue);
            mail.Display();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }                                        
        }         

       
        public static void WriteReestrSignature() 
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            Range range = messageSheet.Range["A1:B7"];

            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey SignatureKey = currentUserKey.CreateSubKey("Sheffler");
            string name = range.Cells[1, 2].Text;
            //if (string.IsNullOrWhiteSpace(name))
            //{
            //    MessageBox.Show("Заполните информацию об отправителе.");
            //}
            SignatureKey.SetValue("Ответственное лицо", name);
            SignatureKey.SetValue("Компания", range.Cells[ 2, 2 ].Text);                  
            SignatureKey.SetValue("Адрес", range.Cells[ 3, 2 ].Text);    
            SignatureKey.SetValue("Город", range.Cells[ 4, 2 ].Text);           
            SignatureKey.SetValue("Тел", range.Cells[ 5, 2 ].Text);
            SignatureKey.SetValue("Моб", range.Cells[ 6, 2 ].Text);
            SignatureKey.SetValue("Mail", range.Cells[ 7, 2 ].Text);
            SignatureKey.Close();
        }

        public static string ReadReestrSignature()
        {
            RegistryKey currentUserKey = Registry.CurrentUser;
            
            RegistryKey SignatureKey = currentUserKey.OpenSubKey("Sheffler");
            if (SignatureKey == null) 
            {
                WriteReestrSignature();
                SignatureKey = currentUserKey.OpenSubKey("Sheffler");
            }

          string name = SignatureKey.GetValue("Ответственное лицо").ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                WriteReestrSignature();
                name = SignatureKey.GetValue("Ответственное лицо").ToString();                
            }
            string signature =
            "<br>" + name
            + "<br>" + SignatureKey.GetValue("Компания").ToString()
            + "<br>" + SignatureKey.GetValue("Адрес").ToString()
            + "<br>" + SignatureKey.GetValue("Город").ToString()
            + "<br>" + SignatureKey.GetValue("Тел").ToString()
            + "<br>" + SignatureKey.GetValue("Моб").ToString()
            + "<br>" + SignatureKey.GetValue("Mail").ToString();

            SignatureKey.Close();
            return signature;
        }
    }
}
