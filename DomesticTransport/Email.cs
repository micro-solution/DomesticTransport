
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace DomesticTransport
{
    class Email
    {

        public enum TypeSend
        {
            Save,
            Display,
            Send
        }
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
        public void CreateMessage(string сompany,
                                   string date,
                                  string attachment,
                                  string subject)
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            ListObject tableEmail = messageSheet.ListObjects["TableEmail"];
            string addres = "";
            foreach (Range row in tableEmail.DataBodyRange.Rows)
            {
                if (row.Cells[1, 1].Text == сompany)
                {
                    string stroka = row.Cells[1, 2].Text;
                    addres = stroka == "" ? addres : $"{stroka}; {addres}";
                }
            }
            string signature = ReadSignature(Properties.Settings.Default.Signature);
            string textMsg = messageSheet.Cells[10, 2].Text;
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        public string GetAdressProvider(string company)
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            ListObject tableEmail = messageSheet.ListObjects["TableEmail"];
            string addres = "";
            foreach (Range row in tableEmail.DataBodyRange.Rows)
            {
                if (row.Cells[1, 1].Text == company)
                {
                    string stroka = row.Cells[1, 2].Text;
                    addres = stroka == "" ? addres : $"{stroka}; {addres}";
                }
            }
            return addres;
        }

        public void MailToProvider(string сompany, string subject, string message, List<string> attachments, TypeSend typeSend)
        {
            string addres = GetAdressProvider(сompany);
            string copyTo = GetCopyProviderEmails();
            CreateMail(addres, copyTo, subject, message, attachments, typeSend);
        }

        private string GetCopyProviderEmails()
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            string copyTo = messageSheet.Cells[9, 2].Text;
            return copyTo;
        }
        public void CreateMessage2(string сompany,
                                   string date,
                                  string attachment,
                                  string subject,
                                  string message)
        {
            string addres = GetAdressProvider(сompany);
            string signature = ReadSignature(Properties.Settings.Default.Signature);
            string copyTo = GetCopyProviderEmails();
            string HtmlBody =
                  message +
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        /// <summary>
        /// Создание сообщения
        /// </summary>
        /// <param name="to"></param>
        /// <param name="copy"></param>
        /// <param name="subject"></param>
        /// <param name="message"></param>
        /// <param name="attachments"></param>
        /// <param name="typeSend"></param>
        public void CreateMail(string to, string copy, string subject, string message, List<string> attachments, TypeSend typeSend = TypeSend.Display)
        {
            string signature = ReadSignature(Properties.Settings.Default.Signature);
            message += "<br><br>" + signature;

            try
            {
                OutlookApp.Session.Logon();
                Outlook.MailItem mail = (Outlook.MailItem)OutlookApp.CreateItem(0);
                mail.To = to;
                mail.HTMLBody = message;
                mail.BCC = "";
                mail.CC = copy;
                mail.Subject = subject;
                foreach (string attach in attachments)
                {
                    mail.Attachments.Add(attach, Outlook.OlAttachmentType.olByValue);
                }

                switch (typeSend)
                {
                    case TypeSend.Save:
                        mail.Save();
                        break;
                    case TypeSend.Display:
                        mail.Display();
                        break;
                    case TypeSend.Send:
                        mail.Send();
                        break;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        /// <summary>
        /// Получение подписи из настроек
        /// </summary>
        /// <param name="signatureName"></param>
        /// <returns></returns>
        private string ReadSignature(string signatureName = "")
        {
            try
            {
                string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
                string signature = string.Empty;
                DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

                if (diInfo.Exists)
                {
                    FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                    foreach (FileInfo file in fiSignature)
                    {
                        string fileName = file.Name.Replace(file.Extension, string.Empty);
                        if (signatureName == "") signatureName = fileName;
                        if (signatureName == fileName)
                        {
                            StreamReader sr = new StreamReader(file.FullName, Encoding.Default);
                            signature = sr.ReadToEnd();
                            signature = signature.Replace(fileName + ".files/", appDataDir + "/" + fileName + ".files/");
                            sr.Close();
                            break;
                        }
                    }
                }
                return signature;
            }
            catch
            {
                return "";
            } 
        }

        public static void WriteReestrSignature()
        {
            Worksheet messageSheet = Globals.ThisWorkbook.Sheets["Mail"];
            Range range = messageSheet.Range["A1:B7"];

            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey SignatureKey = currentUserKey.CreateSubKey("Sheffler");
            string name = range.Cells[1, 2].Text;
            SignatureKey.SetValue("Ответственное лицо", name);
            SignatureKey.SetValue("Компания", range.Cells[2, 2].Text);
            SignatureKey.SetValue("Адрес", range.Cells[3, 2].Text);
            SignatureKey.SetValue("Город", range.Cells[4, 2].Text);
            SignatureKey.SetValue("Тел", range.Cells[5, 2].Text);
            SignatureKey.SetValue("Моб", range.Cells[6, 2].Text);
            SignatureKey.SetValue("Mail", range.Cells[7, 2].Text);
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
