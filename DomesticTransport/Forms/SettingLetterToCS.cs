using System;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class SettingLetterToCS : Form
    {
        public SettingLetterToCS()
        {
            InitializeComponent();

            textBoxTo.Text = Properties.Settings.Default.SettingCSLetterTo;
            textBoxCopy.Text = Properties.Settings.Default.SettingCSLetterCopy;
            textBoxSubject.Text = Properties.Settings.Default.SettingCSLetterSubject;
            textBoxMessage.Text = Properties.Settings.Default.SettingCSLetterMessage;
        }

        /// <summary>
        /// Отмена сохранения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Сохранение настроек
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonOk_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.SettingCSLetterTo = textBoxTo.Text;
            Properties.Settings.Default.SettingCSLetterCopy = textBoxCopy.Text;
            Properties.Settings.Default.SettingCSLetterSubject = textBoxSubject.Text;
            Properties.Settings.Default.SettingCSLetterMessage = textBoxMessage.Text;

            Properties.Settings.Default.Save();
            Close();
        }
    }
}
