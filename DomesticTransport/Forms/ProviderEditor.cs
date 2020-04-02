using DomesticTransport.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class ProviderEditor : Form
    {
         public double Weight { get; set; }
        public List<DeliveryPoint> MapDpelivery { get; set; }


        public ProviderEditor()
        {
            InitializeComponent();
            tbWeight.Text = Weight.ToString();
        }

        private void Provider_Load(object sender, EventArgs e)
        {
            Worksheet sh = Globals.ThisWorkbook.Sheets["Rate"];
            ListObject providerTable = sh.ListObjects["ProviderTable"];
            int iProviler = 0;
            foreach (Range row in providerTable.DataBodyRange.Rows)
            {
                ++iProviler;
                string name = row.Cells[1,1].Text;
                lvProvider.Items.Add( name );
            }

        }
    }
}
