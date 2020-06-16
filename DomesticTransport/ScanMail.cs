using DomesticTransport.Forms;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace DomesticTransport
{
    /// <summary>
    /// Сканирование писем
    /// </summary>
    internal class ScanMail
    {
        public Outlook.Application OutlookApp
        {
            get
            {
                if (outlookApp == null)
                {
                    outlookApp = new Outlook.Application();
                }
                return outlookApp;
            }
        }
        private Outlook.Application outlookApp;

        private readonly List<string> FolderNames = new List<string>();

        public ScanMail()
        {
            FolderNames = Properties.Settings.Default.OutlookFolders.Split(';').ToList();
        }

        /// <summary>
        /// Сохранение вложений в выбранных папках, полученные текущей датой
        /// </summary>
        public int SaveAttachments()
        {
            MessageDate messageDate = new MessageDate();
            messageDate.ShowDialog();
            if (messageDate.DialogResult != DialogResult.OK) return 0;
            //messageDate.date

            int count = 0;
            OutlookApp.Session.Logon();
            foreach (string folderName in FolderNames)
            {
                Outlook.Folder folder = GetFolder(folderName);
                if (folder == null)
                {
                    continue;
                }

                ProcessBar pb = Forms.ProcessBar.Init("Сканирование папки " + folder.Name, folder.Items.Count, 1, folder.Name);
                pb.Show();
                foreach (object item in folder.Items)
                {
                    if (pb.Cancel) break;

                    if (!(item is Outlook.MailItem mail))
                    {
                        pb.Action();
                        continue;
                    }

                    pb.Action(mail.ReceivedTime.Date.ToString());
                    if (mail.Attachments.Count == 0) continue;
                    if (mail.ReceivedTime.Date < messageDate.DateStart || mail.ReceivedTime.Date > messageDate.DateEnd) continue;

                    string path = Globals.ThisWorkbook.Path + "\\MailFromProviders\\" + DateTime.Today.ToString("dd.MM.yyyy") + '\\';
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    foreach (Outlook.Attachment attach in mail.Attachments)
                    {
                        if (!attach.FileName.Contains("xls")) continue;
                        attach.SaveAsFile(path + attach.FileName);
                        count++;
                    }
                }
                pb.Close();
            }
            return count;
        }


        /// <summary>
        /// Получение папки по пути к ней
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        private Outlook.Folder GetFolder(string folderPath)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders =
                    folderPath.Split(backslash.ToCharArray());
                folder =
                    OutlookApp.Application.Session.Folders[folders[0]]
                    as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]]
                            as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch { return null; }
        }

        /// <summary>
        /// Получение данных из писем провайдеров
        /// </summary>
        public void GetDataFromProviderFiles()
        {
            string path = Globals.ThisWorkbook.Path + "\\MailFromProviders\\" + DateTime.Today.ToString("dd.MM.yyyy") + '\\';
            if (!Directory.Exists(path))
            {
                //MessageBox.Show("Папка " + path + " отсутствует");
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
                pb.Action($"Вложение {i + 1} из {pb.Count} {fileInfo.Name} ");

                if (!file.Contains(".xls")) { continue; }
                new Functions().ReadMessageFile(file);
            }
            pb.Close();
        }

    }
}
