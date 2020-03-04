using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;


namespace DomesticTransport
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            Functions functions = new Functions();          
            functions.SetDelivery();   
        }
    }
}
