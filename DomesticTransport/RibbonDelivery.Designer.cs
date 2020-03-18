﻿namespace DomesticTransport
{
    partial class RibbonDelivery : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonDelivery()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonDelivery));
            this.ShefflerRibbon = this.Factory.CreateRibbonTab();
            this.groupGeneral = this.Factory.CreateRibbonGroup();
            this.groupEdit = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.settings = this.Factory.CreateRibbonGroup();
            this.about = this.Factory.CreateRibbonGroup();
            this.btnStart = this.Factory.CreateRibbonButton();
            this.btnAcept = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.btnChangeSet = this.Factory.CreateRibbonButton();
            this.btnChangePoint = this.Factory.CreateRibbonButton();
            this.btnSendShippingCompany = this.Factory.CreateRibbonButton();
            this.btnReadCarrierInvoice = this.Factory.CreateRibbonButton();
            this.btnSetts = this.Factory.CreateRibbonButton();
            this.btnAboutProgrramm = this.Factory.CreateRibbonButton();
            this.BtnLoadAllOrders = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.ShefflerRibbon.SuspendLayout();
            this.groupGeneral.SuspendLayout();
            this.groupEdit.SuspendLayout();
            this.group1.SuspendLayout();
            this.settings.SuspendLayout();
            this.about.SuspendLayout();
            this.SuspendLayout();
            // 
            // ShefflerRibbon
            // 
            this.ShefflerRibbon.Groups.Add(this.groupGeneral);
            this.ShefflerRibbon.Groups.Add(this.groupEdit);
            this.ShefflerRibbon.Groups.Add(this.group1);
            this.ShefflerRibbon.Groups.Add(this.settings);
            this.ShefflerRibbon.Groups.Add(this.about);
            this.ShefflerRibbon.Label = "Шеффлер";
            this.ShefflerRibbon.Name = "ShefflerRibbon";
            this.ShefflerRibbon.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // groupGeneral
            // 
            this.groupGeneral.Items.Add(this.btnStart);
            this.groupGeneral.Items.Add(this.BtnLoadAllOrders);
            this.groupGeneral.Items.Add(this.btnAcept);
            this.groupGeneral.Label = "Список";
            this.groupGeneral.Name = "groupGeneral";
            // 
            // groupEdit
            // 
            this.groupEdit.Items.Add(this.button1);
            this.groupEdit.Items.Add(this.button2);
            this.groupEdit.Items.Add(this.button3);
            this.groupEdit.Items.Add(this.btnChangeSet);
            this.groupEdit.Label = "Редактирование";
            this.groupEdit.Name = "groupEdit";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnChangePoint);
            this.group1.Items.Add(this.btnSendShippingCompany);
            this.group1.Items.Add(this.btnReadCarrierInvoice);
            this.group1.Label = "Сообщения";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // settings
            // 
            this.settings.Items.Add(this.btnSetts);
            this.settings.Label = "Настройки";
            this.settings.Name = "settings";
            this.settings.Visible = false;
            // 
            // about
            // 
            this.about.Items.Add(this.btnAboutProgrramm);
            this.about.Label = "Справка";
            this.about.Name = "about";
            this.about.Visible = false;
            // 
            // btnStart
            // 
            this.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStart.Image = ((System.Drawing.Image)(resources.GetObject("btnStart.Image")));
            this.btnStart.Label = "Формировать список доставок";
            this.btnStart.Name = "btnStart";
            this.btnStart.ShowImage = true;
            this.btnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStart_Click);
            // 
            // btnAcept
            // 
            this.btnAcept.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAcept.Image = ((System.Drawing.Image)(resources.GetObject("btnAcept.Image")));
            this.btnAcept.Label = "Принять ";
            this.btnAcept.Name = "btnAcept";
            this.btnAcept.ShowImage = true;
            this.btnAcept.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAcept_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Добавить авто";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "Удалить авто";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // btnChangeSet
            // 
            this.btnChangeSet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChangeSet.Image = ((System.Drawing.Image)(resources.GetObject("btnChangeSet.Image")));
            this.btnChangeSet.Label = "Пересчитать маршруты";
            this.btnChangeSet.Name = "btnChangeSet";
            this.btnChangeSet.ShowImage = true;
            this.btnChangeSet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeSet_Click);
            // 
            // btnChangePoint
            // 
            this.btnChangePoint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChangePoint.Image = ((System.Drawing.Image)(resources.GetObject("btnChangePoint.Image")));
            this.btnChangePoint.Label = "Изменить маршрут";
            this.btnChangePoint.Name = "btnChangePoint";
            this.btnChangePoint.ShowImage = true;
            this.btnChangePoint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangePoint_Click);
            // 
            // btnSendShippingCompany
            // 
            this.btnSendShippingCompany.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSendShippingCompany.Image = ((System.Drawing.Image)(resources.GetObject("btnSendShippingCompany.Image")));
            this.btnSendShippingCompany.Label = "Заказать перевозку";
            this.btnSendShippingCompany.Name = "btnSendShippingCompany";
            this.btnSendShippingCompany.ShowImage = true;
            this.btnSendShippingCompany.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendShippingCompany_Click);
            // 
            // btnReadCarrierInvoice
            // 
            this.btnReadCarrierInvoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadCarrierInvoice.Image = ((System.Drawing.Image)(resources.GetObject("btnReadCarrierInvoice.Image")));
            this.btnReadCarrierInvoice.Label = "Сканировать ответ";
            this.btnReadCarrierInvoice.Name = "btnReadCarrierInvoice";
            this.btnReadCarrierInvoice.ShowImage = true;
            // 
            // btnSetts
            // 
            this.btnSetts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetts.Image = ((System.Drawing.Image)(resources.GetObject("btnSetts.Image")));
            this.btnSetts.Label = "Настройки";
            this.btnSetts.Name = "btnSetts";
            this.btnSetts.ShowImage = true;
            // 
            // btnAboutProgrramm
            // 
            this.btnAboutProgrramm.Image = ((System.Drawing.Image)(resources.GetObject("btnAboutProgrramm.Image")));
            this.btnAboutProgrramm.Label = "О программе";
            this.btnAboutProgrramm.Name = "btnAboutProgrramm";
            this.btnAboutProgrramm.ShowImage = true;
            // 
            // BtnLoadAllOrders
            // 
            this.BtnLoadAllOrders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadAllOrders.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadAllOrders.Image")));
            this.BtnLoadAllOrders.Label = "Принять ";
            this.BtnLoadAllOrders.Name = "BtnLoadAllOrders";
            this.BtnLoadAllOrders.ShowImage = true;
            this.BtnLoadAllOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAcept_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Изменить доставки";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChangeSet_Click);
            // 
            // RibbonDelivery
            // 
            this.Name = "RibbonDelivery";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ShefflerRibbon);
            this.ShefflerRibbon.ResumeLayout(false);
            this.ShefflerRibbon.PerformLayout();
            this.groupGeneral.ResumeLayout(false);
            this.groupGeneral.PerformLayout();
            this.groupEdit.ResumeLayout(false);
            this.groupEdit.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.settings.ResumeLayout(false);
            this.settings.PerformLayout();
            this.about.ResumeLayout(false);
            this.about.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGeneral;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangePoint;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAcept;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChangeSet;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab ShefflerRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendShippingCompany;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadCarrierInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup settings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetts;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup about;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAboutProgrramm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadAllOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDelivery Ribbon
        {
            get { return this.GetRibbon<RibbonDelivery>(); }
        }
    }
}
