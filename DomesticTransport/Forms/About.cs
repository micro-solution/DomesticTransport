﻿using System.Windows.Forms;

namespace DomesticTransport.Forms
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://micro-solution.ru");
        }
    }
}
