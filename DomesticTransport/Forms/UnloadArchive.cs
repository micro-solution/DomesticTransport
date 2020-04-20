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
    public partial class  UnloadArchive : Form
    {
        public UnloadArchive()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnAcept_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
