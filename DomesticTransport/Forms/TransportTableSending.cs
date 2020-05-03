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
    public partial class TransportTableSending : Form
    {
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public string Provider { get; set; }

        public TransportTableSending()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        /// <summary>
        /// Загрузка формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UnloadOnDate_Load(object sender, EventArgs e)
        {
            monthCalendarEnd.SetDate(DateTime.Today);
            monthCalendarStart.SetDate(DateTime.Today);
            foreach (ListRow row in ShefflerWB.ProviderTable.ListRows)
            {
                int col = ShefflerWB.ProviderTable.ListColumns["Company"].Index;
                string compny = row.Range[1, col].Text;
                if (compny !="Деловые линии" && !string.IsNullOrWhiteSpace(compny))
                {
                    ComboboxProvider.Items.Add(compny);
                }
            }
        }

        /// <summary>
        /// Кнопка выбора
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAccept_Click(object sender, EventArgs e)
        {
            if (monthCalendarStart.SelectionStart == null)
            {
                MessageBox.Show("Выберите дату начала периода", "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (monthCalendarEnd.SelectionStart == null)
            {
                MessageBox.Show("Выберите дату окончания периода", "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (ComboboxProvider.SelectedItem == null)
            {
                MessageBox.Show("Выберите провайдера", "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (monthCalendarStart.SelectionStart > monthCalendarEnd.SelectionStart)
            {
                MessageBox.Show("Дата начала отчета не может быть позже даты завершения", "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DateStart = monthCalendarStart.SelectionStart;
            DateEnd = monthCalendarEnd.SelectionStart;
            Provider = ComboboxProvider.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// Кнопка отмены
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonСancel_Click(object sender, EventArgs e)
        {              
           Close();
        }

    }
}
