using System;
using System.IO;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class Settings : Form
    {

        private string _pathTransortTable = Properties.Settings.Default.TransportTableFileFullName;
        private string _pathHelper;

        public Settings()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            tbTransortTable.Text = _pathTransortTable;
            _pathHelper = Properties.Settings.Default.HelpPath;
            string defaultHalper = Globals.ThisWorkbook.Path + @"\help.docx";
            defaultHalper = File.Exists(defaultHalper) ? defaultHalper : "";
            _pathHelper = string.IsNullOrWhiteSpace(_pathHelper) ? defaultHalper : _pathHelper;
        }

        private void BtnOFD_Click(object sender, EventArgs e)
        {
            string defaultPath = Properties.Settings.Default.SapUnloadPath;
            using (OpenFileDialog fileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = string.IsNullOrWhiteSpace(defaultPath) ? Directory.GetCurrentDirectory() : defaultPath,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    _pathTransortTable = fileDialog.FileName;
                    tbTransortTable.Text = _pathTransortTable;
                }
            }
        }

        private void BtnAcept_Click(object sender, EventArgs e)
        {
            _pathTransortTable = tbTransortTable.Text;
            if (File.Exists(_pathTransortTable))
            {
                Properties.Settings.Default.TransportTableFileFullName = _pathTransortTable;
            }

            if (File.Exists(_pathHelper))
            {
                Properties.Settings.Default.HelpPath = _pathHelper;
            }

            Properties.Settings.Default.Save();
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
