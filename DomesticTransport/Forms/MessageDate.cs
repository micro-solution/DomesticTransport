using System;
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

        private void BtnScan_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            DateStart = dt1.Value;
            DateEnd = dt2.Value;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void MessageDate_Load(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today;
            dt2.Value = DateTime.Today;
        }

        private void BtnToday_Click_1(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today;
            dt2.Value = DateTime.Today;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today.AddDays(-1);
            dt2.Value = DateTime.Today;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Today.AddDays(-1);
            dt2.Value = DateTime.Today.AddDays(-1);
        }
    }
}
