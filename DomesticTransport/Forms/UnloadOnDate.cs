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
    public partial class UnloadOnDate : Form
    {
        public DateTime FirstDate { get; set; }
        public DateTime SecondDate { get; set; }
        public string Compny { get; set; }
        public UnloadOnDate()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }
        private void UnloadOnDate_Load(object sender, EventArgs e)
        {
            DateTime dateMax = DateTime.Today;
            dateMax = dateMax.AddDays(-(double)dateMax.DayOfWeek);
            dateTimePicker1.Value = dateMax;
            dateTimePicker2.Value = DateTime.Today;
            foreach (ListRow row in ShefflerWB.ProviderTable.ListRows)
            {
                int col = ShefflerWB.ProviderTable.ListColumns["Company"].Index;
                string compny = row.Range[1, col].Text;
                if (compny !="Деловые линии" && !string.IsNullOrWhiteSpace(compny))
                {
                    cbxProvider.Items.Add(compny);
                }
            }
            cbxProvider.SelectedItem = cbxProvider.Items[0];
            Compny = cbxProvider.Items[0].ToString();
        }

        private void btnAcept_Click(object sender, EventArgs e)
        {
            FirstDate = dateTimePicker1.Value;
            SecondDate = dateTimePicker2.Value;
            Compny = cbxProvider.Text;
            DialogResult = DialogResult.OK;
            Hide();
        }

        private void btnСancel_Click(object sender, EventArgs e)
        {              
           Close();
        }

    }
}
