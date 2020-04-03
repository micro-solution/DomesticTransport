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
        public string ProviderName { get; set; }
        public double CostDelivery { get; set; }

        public Delivery DeliveryTarget{ get; set; }
        


        public ProviderEditor()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        private void Provider_Load(object sender, EventArgs e)
        {
            Worksheet sh = Globals.ThisWorkbook.Sheets["Rate"];
            ListObject providerTable = sh.ListObjects["ProviderTable"];
            ShefflerWorkBook shefflerWorkBook = new ShefflerWorkBook();
             List<DeliveryPoint> mapDpelivery = DeliveryTarget?.MapDelivery;

            int iProviler = 0;
            tbWeight.Text = Weight.ToString();
            foreach (Range row in providerTable.DataBodyRange.Rows)
            {
                string name = row.Cells[1,1].Text;
                lvProvider.Items.Add( name );
                Truck truck = shefflerWorkBook.GetTruck(Weight, mapDpelivery, name);
                lvProvider.Items[iProviler].SubItems.Add(truck.Cost.ToString());
                iProviler++;
            }
            if (mapDpelivery != null && mapDpelivery.Count>0)
            {
                for (int i = 0; i < mapDpelivery.Count; i++)
                {
                    int row = i + 1;
                    lvMap.Items.Add(row.ToString());
                    lvMap.Items[i].SubItems.Add (mapDpelivery[i].City);
                }
                    
            }
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Hide();

        }

        private void lvProvider_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvProvider.SelectedItems.Count > 0)
            {
                ProviderName = lvProvider.SelectedItems[0].Text;
                string cost = lvProvider.SelectedItems[0].SubItems[1].Text;
                CostDelivery = double.TryParse(cost, out double ct) ? ct : 0 ;
            }
        }
    }
}
