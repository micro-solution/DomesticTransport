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
    public partial class SelectProvider : Form
    {
        public string Provider { get; set; }

        public SelectProvider()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
            FillCombobox();
        }

        private void ButtonAccept_Click(object sender, EventArgs e)
        {
            if (ComboboxProvider.SelectedItem == null)
            {
                MessageBox.Show("Выберите провайдера", "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Provider = ComboboxProvider.Text;
            DialogResult = DialogResult.OK;
            Close();
        }
 

        /// <summary>
        /// Загрузка формы
        /// </summary>
        private void FillCombobox()
        {
            ComboboxProvider.Items.Add("Отправить всем");
            foreach (ListRow row in ShefflerWB.ProviderTable.ListRows)
            {
                int col = ShefflerWB.ProviderTable.ListColumns["Company"].Index;
                string compny = row.Range[1, col].Text;
                if (compny != "Деловые линии" && !string.IsNullOrWhiteSpace(compny))
                {
                    ComboboxProvider.Items.Add(compny);
                }
            }
        }

        private void ButtonСancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
