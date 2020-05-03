using System;
using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class SettingLetters : Form
    {
        public SettingLetters()
        {
            InitializeComponent();

            textBoxCSTo.Text = Properties.Settings.Default.SettingCSLetterTo;
            textBoxCSCopy.Text = Properties.Settings.Default.SettingCSLetterCopy;
            textBoxCSSubject.Text = Properties.Settings.Default.SettingCSLetterSubject;
            textBoxCSMessage.Text = Properties.Settings.Default.SettingCSLetterMessage;

            textBoxProviderCopy.Text = Properties.Settings.Default.ProviderLettersCopy;
            textBoxProviderOrderSubject.Text = Properties.Settings.Default.ProviderSubjectOrder;
            textBoxProviderOrderMessage.Text = Properties.Settings.Default.ProviderMessageOrder;

            textBoxProviderAddSubject.Text = Properties.Settings.Default.ProviderSubjectAdd;
            textBoxProviderAddMessage.Text = Properties.Settings.Default.ProviderMessageAdd;

            textBoxProviderMessageSubject.Text = Properties.Settings.Default.ProviderSubjectReport;
            textBoxProviderReportMessage.Text = Properties.Settings.Default.ProviderMessageReport;

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
            Properties.Settings.Default.SettingCSLetterTo = textBoxCSTo.Text;
            Properties.Settings.Default.SettingCSLetterCopy = textBoxCSCopy.Text;
            Properties.Settings.Default.SettingCSLetterSubject = textBoxCSSubject.Text;
            Properties.Settings.Default.SettingCSLetterMessage = textBoxCSMessage.Text;

            Properties.Settings.Default.ProviderLettersCopy = textBoxProviderCopy.Text;
            Properties.Settings.Default.ProviderSubjectOrder = textBoxProviderOrderSubject.Text;
            Properties.Settings.Default.ProviderMessageOrder = textBoxProviderOrderMessage.Text;

            Properties.Settings.Default.ProviderSubjectAdd = textBoxProviderAddSubject.Text;
            Properties.Settings.Default.ProviderMessageAdd = textBoxProviderAddMessage.Text;

            Properties.Settings.Default.ProviderSubjectReport = textBoxProviderMessageSubject.Text;
            Properties.Settings.Default.ProviderMessageReport = textBoxProviderReportMessage.Text;

            Properties.Settings.Default.Save();
            Close();
        }
    }
}
