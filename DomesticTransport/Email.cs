
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


        public void Init()
        {
            //Outlook.Inspectors inspectors;
            //inspectors = new Application.Inspectors ;
            //inspectors.NewInspector +=
            //new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            // Get the Application object
            //   Outlook.Application application =new Outlook.Application();

            // Get the Inspector object
            Outlook.Inspectors inspectors = OutlookApp.Inspectors;

            // Get the active Inspector object
            Outlook.Inspector activeInspector = OutlookApp.ActiveInspector();
            if (activeInspector != null)
            {
                // Get the title of the active item when the Outlook start.
                MessageBox.Show("Active inspector: " + activeInspector.Caption);
            }

            // Get the Explorer objects
            Outlook.Explorers explorers = OutlookApp.Explorers;

            // Get the active Explorer object
            Outlook.Explorer activeExplorer = OutlookApp.ActiveExplorer();
            if (activeExplorer != null)
            {
                // Get the title of the active folder when the Outlook start.
                MessageBox.Show("Active explorer: " + activeExplorer.Caption);
            }
        }

        public void Send_Email (string addres, string text, string body,
                              string copyTo  )
        {
            Outlook.MailItem mail = (Outlook.MailItem)OutlookApp.CreateItem(
                                    Outlook.OlItemType.olMailItem);
            mail.To = addres;
            mail.Subject = text;
            mail.Body = body;
            mail.BCC = "";
            mail.CC = "";


        }
        public void Receive()
        {
            try
            {
                Outlook.NameSpace _ns = OutlookApp.GetNamespace("MAPI");
                Outlook.MAPIFolder inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                _ns.SendAndReceive(true);
                foreach (Outlook.MailItem item in inbox.Items)
                {
                    Debug.WriteLine("" + item.Subject + item.SenderName +
                                item.HTMLBody + item.SentOn.ToLongDateString() +
                                item.SentOn.ToLongTimeString());
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        public void CreateMessage (string addres,
                                   string text,                                    
                                   string subject,
                                   string body,
                                   string copyTo)
        {
            OutlookApp.Session.Logon();
            Outlook.MailItem mail = (Outlook.MailItem)OutlookApp.CreateItem(0);
         
         
            mail.To = addres;
            mail.Subject = text;
            mail.HTMLBody = body;
            mail.BCC = "";
            mail.CC = copyTo;
            mail.Subject = subject;
            mail.Display();
        }
    }
}
