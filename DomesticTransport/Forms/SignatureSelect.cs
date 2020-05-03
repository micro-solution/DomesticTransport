using System;
using System.IO;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class SignatureSelect : Form
    {
        public SignatureSelect()
        {
            InitializeComponent();
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                foreach (FileInfo file in fiSignature)
                {
                    string fileName = file.Name.Replace(file.Extension, string.Empty);
                    comboBoxSignatures.Items.Add(fileName);
                }
            }
            comboBoxSignatures.Text = Properties.Settings.Default.Signature;
        }

        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Кнопка сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSave_Click(object sender, EventArgs e)
        {
            if (comboBoxSignatures.SelectedItem == null)
            {
                MessageBox.Show("Необходимо выбрать подпись из списка", "Выбор подписи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Properties.Settings.Default.Signature = comboBoxSignatures.SelectedItem.ToString();
            Properties.Settings.Default.Save();
            Close();
        }
    }
}
