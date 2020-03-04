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

        private void btnChangeSet_Click(object sender, RibbonControlEventArgs e)
        {
            ChengeDelivery chengeDelivery = new ChengeDelivery();
            chengeDelivery.Show();
        }
    }
}
