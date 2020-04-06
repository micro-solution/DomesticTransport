using System;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class Signature : Form
    {
        public Signature()
        {
            InitializeComponent();
            TextBoxSignature.Text = Properties.Settings.Default.Signature;
        }

        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ButtonOk_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Signature = TextBoxSignature.Text;
            Properties.Settings.Default.Save();
            Close();
        }
    }
}
