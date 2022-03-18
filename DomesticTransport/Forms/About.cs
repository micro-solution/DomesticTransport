using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
            string productVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion;
            lbVer.Text = productVersion;
        }

        private void LinkMS_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://micro-solution.ru");
        }
    }
}