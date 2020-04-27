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
    public partial class MessageDate : Form
    {
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public MessageDate()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
            DateStart = DateTime.Today;
            DateEnd = DateTime.Today;
        }

        private void btnScan_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            DateStart = dt1.Value;
            DateEnd = dt2.Value;
            Hide();
        }

  

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void MessageDate_Load(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today;
            dt2.Value = DateTime.Today;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today.AddDays(-1);
            dt2.Value = DateTime.Today;
        }

        private void btnToday_Click_1(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today;
            dt2.Value = DateTime.Today;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today.AddDays(-1);
            dt2.Value = DateTime.Today;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today.AddDays(-1);
            dt2.Value = DateTime.Today.AddDays(-1);
        }
    }
}
