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
            
            #region Kirirll 15.11.21
            textBoxCSLETo.Text = Properties.Settings.Default.SettingCSLELetterTo;
            textBoxCSLECopy.Text = Properties.Settings.Default.SettingCSLELetterCopy;
            textBoxCSLESubject.Text = Properties.Settings.Default.SettingCSLELetterSubject;
            textBoxCSLEMessage.Text = Properties.Settings.Default.SettingCSLELetterMessage;

            textBoxCSStorekeeperTo.Text = Properties.Settings.Default.SettingCSLetterStorekeeperTo;
            textBoxCSStorekeeperCopy.Text = Properties.Settings.Default.SettingCSLetterStorekeeperCopy;
            textBoxCSStorekeeperSubject.Text = Properties.Settings.Default.SettingCSLetterStorekeeperSubject;
            textBoxCSStorekeeperMessage.Text = Properties.Settings.Default.SettingCSLetterStorekeeperMessage;
            #endregion

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
            #region Kirill
            Properties.Settings.Default.SettingCSLELetterTo = textBoxCSLETo.Text;
            Properties.Settings.Default.SettingCSLELetterCopy = textBoxCSLECopy.Text;
            Properties.Settings.Default.SettingCSLELetterSubject = textBoxCSLESubject.Text;
            Properties.Settings.Default.SettingCSLELetterMessage = textBoxCSLEMessage.Text;

            Properties.Settings.Default.SettingCSLetterStorekeeperTo = textBoxCSStorekeeperTo .Text;
            Properties.Settings.Default.SettingCSLetterStorekeeperCopy = textBoxCSStorekeeperCopy.Text;
            Properties.Settings.Default.SettingCSLetterStorekeeperSubject = textBoxCSStorekeeperSubject.Text;
            Properties.Settings.Default.SettingCSLetterStorekeeperMessage = textBoxCSStorekeeperMessage.Text;
            #endregion

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
