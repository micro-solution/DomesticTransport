using System;
using System.Collections.Generic;
using System.Linq;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace DomesticTransport
{
    /// <summary>
    /// Класс сканирования писем
    /// </summary>
    class ScanMail
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
        public void SaveAttachments()
        {
            OutlookApp.Session.Logon();
            foreach (string folderName in FolderNames)
            {
                Outlook.Folder folder = GetFolder(folderName);
                if (folder == null)
                {
                    continue;
                }

                Forms.ProcessBar pb = Forms.ProcessBar.Init("Сканирование папки " + folder.Name, folder.Items.Count, 1, folder.Name);
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
                    if (mail.ReceivedTime.Date != DateTime.Today) continue;

                    string path = Globals.ThisWorkbook.Path + "\\MailAttachments\\" + DateTime.Today.ToString("dd.MM.yyyy") + '\\';
                    if (!System.IO.Directory.Exists(path))
                    {
                        System.IO.Directory.CreateDirectory(path);
                    }

                    foreach (Outlook.Attachment attach in mail.Attachments)
                    {
                        if (!attach.FileName.Contains("xls")) continue;
                        attach.SaveAsFile(path + attach.FileName);
                    }
                }
                pb.Close();
            }
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

        public void GetMessage()
        {
            
        }

    }
}
