using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using Config = DomesticTransport.Properties.Settings;

using System.Windows.Forms;

namespace DomesticTransport
{
    public partial class SapFiles : Form
    {
      public string ExportFile { 
            get {
                if (!CheckPath(tbExport.Text))
                {
                throw new   FileNotFoundException("Файла не существует");
                }                
               else { return tbExport.Text; }
            } 
        }
        public string OrderFile {
            get
            {
                if (!CheckPath(tbOrders.Text))
                {
                    throw new FileNotFoundException("Файла не существует");
                }
                else   { return tbOrders.Text; }
            }
        }      





        public SapFiles()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        private void Accept_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Hide();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.None;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        ///  Выбрать файл выгрузки SAP
        /// </summary>
        /// <returns></returns>
        public string SelectFile()
        {
            string sapUnload = "";
            string defaultPath = Config.Default.SapUnloadPath;

            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls*",
                CheckFileExists = true,
                InitialDirectory = string.IsNullOrWhiteSpace(defaultPath) ? Directory.GetCurrentDirectory() : defaultPath,
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel|*.xls*"
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    sapUnload = ofd.FileName;
                    Config.Default.SapUnloadPath = new FileInfo(ofd.FileName).DirectoryName;
                    Config.Default.Save();
                }
            }
            return sapUnload;
        }

        private bool CheckPath(string path)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("Указан неверный путь к файлу!", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }
    }
}
