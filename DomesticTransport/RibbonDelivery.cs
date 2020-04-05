using DomesticTransport.Forms;

using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace DomesticTransport
{
    public partial class RibbonDelivery
    {
        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.SetDelivery();
        }

        private void btnChangePoint_Click(object sender, RibbonControlEventArgs e)
        {
           
        }

        private void btnSendShippingCompany_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.CreateMasseges();
        }

        private void btnAcept_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.AcceptDelivery();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.AddAuto();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.DeleteAuto();
        }

        private void btnChangeSet_Click(object sender, RibbonControlEventArgs e)
        {
            //ChangeDelivery chengeDelivery = new ChangeDelivery();
            // chengeDelivery.ShowDialog();
            Functions functions = new Functions();
            functions.СhangeDelivery();
        }

        private void BtnLoadAllOrders_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.LoadAllOrders();
        }

        private void btnReadForms_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.GetOrdersFromFiles();
        }

        private void btnAccept_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();
            functions.CopyDelivery();
        }

        private void btnSaveSignature_Click(object sender, RibbonControlEventArgs e)
        {
            Email.WriteReestrSignature();
        }

        private void btnAboutProgrramm_Click(object sender, RibbonControlEventArgs e)
        {
            About about = new About();
            about.ShowDialog();
        }

        /// <summary>
        /// Выбор папки для сканирования писем от провайдеров
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSelectFoldersOutlook_Click(object sender, RibbonControlEventArgs e)
        {
            OutlookFoldersSelect foldersSelect = new OutlookFoldersSelect();
            foldersSelect.ShowDialog();
        }

        /// <summary>
        /// Сканирование писем
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReadCarrierInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            if (Properties.Settings.Default.OutlookFolders == "")
            {
                MessageBox.Show("Задайте папки для сканирования почты", "Необходима настройка программы", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ScanMail scanMail = new ScanMail();
            scanMail.SaveAttachments();
        }
    }
}
