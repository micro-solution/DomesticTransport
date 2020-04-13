namespace DomesticTransport
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
            this.btnStart = this.Factory.CreateRibbonButton();
            this.btnReadForms = this.Factory.CreateRibbonButton();
            this.BtnLoadAllOrders = this.Factory.CreateRibbonButton();
            this.groupEdit = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.btnChangeSet = this.Factory.CreateRibbonButton();
            this.btnAccept = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSendShippingCompany = this.Factory.CreateRibbonButton();
            this.btnReadCarrierInvoice = this.Factory.CreateRibbonButton();
            this.settings = this.Factory.CreateRibbonGroup();
            this.btnSaveSignature = this.Factory.CreateRibbonButton();
            this.ButtonSelectFoldersOutlook = this.Factory.CreateRibbonButton();
            this.about = this.Factory.CreateRibbonGroup();
            this.btnAboutProgrramm = this.Factory.CreateRibbonButton();
            this.btnSetts = this.Factory.CreateRibbonButton();
            this.btnSaveRoute = this.Factory.CreateRibbonButton();
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
            this.ShefflerRibbon.Label = "Schaeffler";
            this.ShefflerRibbon.Name = "ShefflerRibbon";
            this.ShefflerRibbon.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // groupGeneral
            // 
            this.groupGeneral.Items.Add(this.btnStart);
            this.groupGeneral.Items.Add(this.btnReadForms);
            this.groupGeneral.Items.Add(this.BtnLoadAllOrders);
            this.groupGeneral.Label = "Загрузка заказов";
            this.groupGeneral.Name = "groupGeneral";
            // 
            // btnStart
            // 
            this.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStart.Image = ((System.Drawing.Image)(resources.GetObject("btnStart.Image")));
            this.btnStart.Label = "Export from SAP";
            this.btnStart.Name = "btnStart";
            this.btnStart.ShowImage = true;
            this.btnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStart_Click);
            // 
            // btnReadForms
            // 
            this.btnReadForms.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadForms.Image = ((System.Drawing.Image)(resources.GetObject("btnReadForms.Image")));
            this.btnReadForms.Label = "Order from CS";
            this.btnReadForms.Name = "btnReadForms";
            this.btnReadForms.ShowImage = true;
            this.btnReadForms.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadForms_Click);
            // 
            // BtnLoadAllOrders
            // 
            this.BtnLoadAllOrders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadAllOrders.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadAllOrders.Image")));
            this.BtnLoadAllOrders.Label = "Загрузить All Orders";
            this.BtnLoadAllOrders.Name = "BtnLoadAllOrders";
            this.BtnLoadAllOrders.ShowImage = true;
            this.BtnLoadAllOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadAllOrders_Click);
            // 
            // groupEdit
            // 
            this.groupEdit.Items.Add(this.button1);
            this.groupEdit.Items.Add(this.button2);
            this.groupEdit.Items.Add(this.btnSaveRoute);
            this.groupEdit.Items.Add(this.btnChangeSet);
            this.groupEdit.Items.Add(this.btnAccept);
            this.groupEdit.Label = "Редактирование";
            this.groupEdit.Name = "groupEdit";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Добавить авто";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
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
            // btnAccept
            // 
            this.btnAccept.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAccept.Image = ((System.Drawing.Image)(resources.GetObject("btnAccept.Image")));
            this.btnAccept.Label = "Перенести в отгрузки";
            this.btnAccept.Name = "btnAccept";
            this.btnAccept.ShowImage = true;
            this.btnAccept.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAccept_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSendShippingCompany);
            this.group1.Items.Add(this.btnReadCarrierInvoice);
            this.group1.Label = "Сообщения";
            this.group1.Name = "group1";
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
            this.btnReadCarrierInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadCarrierInvoice_Click_1);
            // 
            // settings
            // 
            this.settings.Items.Add(this.btnSaveSignature);
            this.settings.Items.Add(this.ButtonSelectFoldersOutlook);
            this.settings.Label = "Настройки";
            this.settings.Name = "settings";
            // 
            // btnSaveSignature
            // 
            this.btnSaveSignature.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveSignature.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveSignature.Image")));
            this.btnSaveSignature.Label = "Сохранить подпись";
            this.btnSaveSignature.Name = "btnSaveSignature";
            this.btnSaveSignature.ShowImage = true;
            this.btnSaveSignature.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSignature_Click);
            // 
            // ButtonSelectFoldersOutlook
            // 
            this.ButtonSelectFoldersOutlook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSelectFoldersOutlook.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSelectFoldersOutlook.Image")));
            this.ButtonSelectFoldersOutlook.Label = "Папки с письмами";
            this.ButtonSelectFoldersOutlook.Name = "ButtonSelectFoldersOutlook";
            this.ButtonSelectFoldersOutlook.ShowImage = true;
            this.ButtonSelectFoldersOutlook.SuperTip = "Выбор папок outlook, в которые сохраняются письма от провайдеров с информацией о " +
    "водителях";
            this.ButtonSelectFoldersOutlook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSelectFoldersOutlook_Click);
            // 
            // about
            // 
            this.about.Items.Add(this.btnAboutProgrramm);
            this.about.Items.Add(this.btnSetts);
            this.about.Label = "Справка";
            this.about.Name = "about";
            // 
            // btnAboutProgrramm
            // 
            this.btnAboutProgrramm.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAboutProgrramm.Image = ((System.Drawing.Image)(resources.GetObject("btnAboutProgrramm.Image")));
            this.btnAboutProgrramm.Label = "О программе";
            this.btnAboutProgrramm.Name = "btnAboutProgrramm";
            this.btnAboutProgrramm.ShowImage = true;
            this.btnAboutProgrramm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAboutProgrramm_Click);
            // 
            // btnSetts
            // 
            this.btnSetts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetts.Image = ((System.Drawing.Image)(resources.GetObject("btnSetts.Image")));
            this.btnSetts.Label = "Настройки";
            this.btnSetts.Name = "btnSetts";
            this.btnSetts.ShowImage = true;
            this.btnSetts.Visible = false;
            // 
            // btnSaveRoute
            // 
            this.btnSaveRoute.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveRoute.Label = "Сохранить маршрут";
            this.btnSaveRoute.Name = "btnSaveRoute";
            this.btnSaveRoute.ShowImage = true;
            this.btnSaveRoute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveRoute_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadForms;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAccept;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveSignature;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSelectFoldersOutlook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveRoute;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDelivery Ribbon
        {
            get { return this.GetRibbon<RibbonDelivery>(); }
        }
    }
}
