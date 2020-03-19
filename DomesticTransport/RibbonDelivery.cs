using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DomesticTransport.Forms;
using Microsoft.Office.Tools.Ribbon;

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
            ChangeRoute changeRoute = new ChangeRoute();
            changeRoute.ShowDialog();
        }

        private void btnSendShippingCompany_Click(object sender, RibbonControlEventArgs e)
        {
            MessageCarrier messageCarrier = new MessageCarrier();
            messageCarrier.ShowDialog();
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
    }
}
