using System;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class Calendar : Form
    {
        public string DateDelivery { get; set; }

        public Calendar()
        {
            InitializeComponent();
        }

        private void BtnFormula_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
            Close();
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {
            DateDelivery = tbDate.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void CalendarControl_DateChanged(object sender, DateRangeEventArgs e)
        {
            tbDate.Text = e.Start.ToShortDateString();
        }

        private void TbDate_TextChanged(object sender, EventArgs e)
        {

            if (DateTime.TryParse(tbDate.Text, out DateTime dt))
            {
                if (dt >= calendarControl.MinDate && dt <= calendarControl.MaxDate)
                {
                    DateDelivery = dt.ToShortDateString();
                    calendarControl.SetSelectionRange(dt, dt);
                }
            }
        }

        private void Calendar_Load(object sender, EventArgs e)
        {
            tbDate.Text = ShefflerWB.DateDelivery;
        }
    }
}
