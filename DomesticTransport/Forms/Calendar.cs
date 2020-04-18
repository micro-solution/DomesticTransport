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
    public partial class Calendar : Form
    {
        public string DateDelivery    { get; set; }
     

        public Calendar()
        {
            InitializeComponent();
        }

        private void btnFormula_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
            Hide();
        }

        private void btnAcept_Click(object sender, EventArgs e)
        {
            DateDelivery =  tbDate.Text;
            DialogResult = DialogResult.OK;
            Hide();
        }
       
        private void calendarControl_DateChanged(object sender, DateRangeEventArgs e)
        {             
            tbDate.Text = e.Start.ToShortDateString();
        }

        private void tbDate_TextChanged(object sender, EventArgs e)
        {
            
           if ( DateTime.TryParse(tbDate.Text, out DateTime dt))
            {
             if (dt >= calendarControl.MinDate && dt <= calendarControl.MaxDate)
                {

                DateDelivery =  dt.ToShortDateString() ;                
                calendarControl.SetSelectionRange(dt, dt);
             
                }
            }
        }

        private void Calendar_Load(object sender, EventArgs e)
        {
         
            tbDate.Text = DateTime.Today.AddDays(1).ToShortDateString();
        }
    }
}
