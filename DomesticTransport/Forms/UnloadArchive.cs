using System;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class UnloadArchive : Form
    {
        public UnloadArchive()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
