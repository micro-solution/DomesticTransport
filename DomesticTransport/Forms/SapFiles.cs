﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using Config = DomesticTransport.Properties.Settings;

using System.Windows.Forms;
using DomesticTransport.Properties;

namespace DomesticTransport
{
    public partial class SapFiles : Form
    {
        public string ExportFile
        {
            get
            {
                CheckPath(tbExport.Text);
                return tbExport.Text;
            }
        }
        public string OrderFile
        {
            get
            {
               // CheckPath(tbOrders.Text);
                return tbOrders.Text;
            }
        }

        public SapFiles()
        {
            InitializeComponent();
            DialogResult = DialogResult.None;
        }

        /// <summary>
        /// Кнопка выбрать папку
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            tbOrders.Text = SelectFile();
        }
        /// <summary>
        /// Кнопка выбрать папку
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            tbExport.Text = SelectFile();
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
                    FileInfo fi = new FileInfo(ofd.FileName);
                    if (fi.DirectoryName != Config.Default.SapUnloadPath)
                    {
                        Config.Default.SapUnloadPath = fi.DirectoryName;
                        Config.Default.Save();
                    }
                }
            }
            return sapUnload;
        }

        private bool CheckPath(string path)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show("Указан неверный путь к файлу!", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                throw new FileNotFoundException("Файла не существует");
            }
            return true;
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

        private void SapFiles_Load(object sender, EventArgs e)
        {
            string path = Settings.Default.SapUnloadPath;
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path);
                foreach (string file in files)
                {

                    FileInfo fi = new FileInfo(file);
                    if (!fi.Name.Contains("~$") &&
                        (fi.Extension.ToLower().Contains("xls") | 
                         fi.Extension.ToLower().Contains("csv")))
                    {

                        if (fi.Name.Contains("Export"))
                        {
                            tbExport.Text = file;
                        }
                        //if (file.Contains("orders"))
                        //{
                        //    tbOrders.Text = file;
                        //}
                        //if (tbOrders.Text != "" && tbExport.Text != "") break;
                    }
                }
            }
        }
    }

}
