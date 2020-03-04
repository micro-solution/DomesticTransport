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
    public partial class ChangeDelivery : Form
    {
        public ChangeDelivery()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
            
        }

        private void Accept_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.None;
            Hide();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
