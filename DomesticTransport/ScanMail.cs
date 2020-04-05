using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DomesticTransport
{
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

        

        public void GetFolders(Outlook.MAPIFolder folder)
        {
            if (folder.Folders.Count == 0)
            {
                MessageBox.Show(folder.Name);
            }
            else
            {
                foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                {
                    GetFolders(subFolder);
                }
            }
        }
    }
}
