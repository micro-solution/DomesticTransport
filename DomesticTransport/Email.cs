
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
        /// <summary>
        /// Тип отправки письма
        /// </summary>
        public enum TypeSend
        {
            Save,
            Display,
            Send
        }

        private Outlook.Application OutlookApp
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

        /// <summary>
        /// Получение списка получателей провайдера
        /// </summary>
        /// <param name="company"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Сообщения провайдерам
        /// </summary>
        /// <param name="сompany"></param>
        /// <param name="subject"></param>
        /// <param name="message"></param>
        /// <param name="attachments"></param>
        /// <param name="typeSend"></param>
        public void MailToProvider(string сompany, string subject, string message, List<string> attachments, TypeSend typeSend)
        {
            string addres = GetAdressProvider(сompany);
            string copyTo = Properties.Settings.Default.ProviderLettersCopy;
            CreateMail(addres, copyTo, subject, message, attachments, typeSend);
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
       
    }
}
