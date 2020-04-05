﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;


namespace DomesticTransport.Forms
{
    /// <summary>
    /// Форма выбора папок для сканирования
    /// </summary>
    public partial class OutlookFoldersSelect : Form
    {
        private Outlook.Application OutlookApp
        {
            get
            {
                if (outlookApp == null)
                {
                    outlookApp = new Outlook.Application();
                }
                return outlookApp;
            }
        }
        private Outlook.Application outlookApp;

        private readonly List<string> SelectedFolder = new List<string>();
        public OutlookFoldersSelect()
        {
            InitializeComponent();

            SelectedFolder = Properties.Settings.Default.OutlookFolders.Split(';').ToList();
            FillFolders();
        }

        /// <summary>
        /// Заполнение списка корневых папок
        /// </summary>
        private void FillFolders()
        {
            foreach (var folder in GetMainFolders())
            {
                TreeNode folderNode = new TreeNode { Text = folder.Name };
                int num = TreeViewFolders.Nodes.Add(folderNode);
                FillTreeNode(folderNode, folder);
                foreach (string path in SelectedFolder)
                {
                    if (path == TreeViewFolders.Nodes[num].FullPath) folderNode.Checked = true;
                }
            }
        }

        /// <summary>
        /// Заполнение списка подпапок
        /// </summary>
        /// <param name="folderNode"></param>
        /// <param name="folder"></param>
        private void FillTreeNode(TreeNode folderNode, Outlook.MAPIFolder folder)
        {
            foreach (Outlook.MAPIFolder subfolder in folder.Folders)
            {
                TreeNode subfolderNode = new TreeNode { Text = subfolder.Name };
                int num = folderNode.Nodes.Add(subfolderNode);
                foreach (string path in SelectedFolder)
                {
                    if (path == folderNode.Nodes[num].FullPath) folderNode.Nodes[num].Checked = true;
                }
                FillTreeNode(subfolderNode, subfolder);
            }
        }

        /// <summary>
        /// Получает список корневых папок
        /// </summary>
        /// <returns></returns>
        private List<Outlook.MAPIFolder> GetMainFolders()
        {
            OutlookApp.Session.Logon();
            List<Outlook.MAPIFolder> folders = new List<Outlook.MAPIFolder>();
            foreach (Outlook.Folder folder in OutlookApp.GetNamespace("MAPI").Folders)
            {
                folders.Add(folder);
            }
            return folders;
        }

        /// <summary>
        /// Сохранение настроек
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonOk_Click(object sender, System.EventArgs e)
        {
            SelectedFolder.Clear();
            foreach (TreeNode node in TreeViewFolders.Nodes)
            {
                CheckNodesRecursive(node);
            }

            if (SelectedFolder.Count > 0)
            {
                Properties.Settings.Default.OutlookFolders = String.Join(";", SelectedFolder.ToArray());
            }
            else
            {
                Properties.Settings.Default.OutlookFolders = "";
            }

            Properties.Settings.Default.Save();
            Close();
        }

        /// <summary>
        /// Рекурсивная проверка выбранных элементов (папок)
        /// </summary>
        /// <param name="parentNode"></param>
        private void CheckNodesRecursive(TreeNode parentNode)
        {
            foreach (TreeNode subNode in parentNode.Nodes)
            {
                if (subNode.Checked)
                {
                    SelectedFolder.Add(subNode.FullPath);
                }
                CheckNodesRecursive(subNode);
            }
        }

        /// <summary>
        /// Кнопка отмены
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
