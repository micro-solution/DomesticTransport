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
            this.BtnExportFromSap = this.Factory.CreateRibbonButton();
            this.ButtonOrderFromCS = this.Factory.CreateRibbonButton();
            this.BtnLoadAllOrders = this.Factory.CreateRibbonButton();
            this.groupEdit = this.Factory.CreateRibbonGroup();
            this.ButtonAddAuto = this.Factory.CreateRibbonButton();
            this.ButtonDeleteAuto = this.Factory.CreateRibbonButton();
            this.BtnRecalcilate = this.Factory.CreateRibbonButton();
            this.BtnFillTable = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BtnSendShippingCompany = this.Factory.CreateRibbonButton();
            this.BtnReadCarrierInvoice = this.Factory.CreateRibbonButton();
            this.ButtonSendToCS = this.Factory.CreateRibbonButton();
            this.settings = this.Factory.CreateRibbonGroup();
            this.BtnSaveSignature = this.Factory.CreateRibbonButton();
            this.ButtonSelectFoldersOutlook = this.Factory.CreateRibbonButton();
            this.ButtonSettingLetterCS = this.Factory.CreateRibbonButton();
            this.about = this.Factory.CreateRibbonGroup();
            this.BtnAboutProgrramm = this.Factory.CreateRibbonButton();
            this.btnSetts = this.Factory.CreateRibbonButton();
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
            this.groupGeneral.Items.Add(this.BtnExportFromSap);
            this.groupGeneral.Items.Add(this.ButtonOrderFromCS);
            this.groupGeneral.Items.Add(this.BtnLoadAllOrders);
            this.groupGeneral.Label = "Загрузка заказов";
            this.groupGeneral.Name = "groupGeneral";
            // 
            // BtnExportFromSap
            // 
            this.BtnExportFromSap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnExportFromSap.Image = ((System.Drawing.Image)(resources.GetObject("BtnExportFromSap.Image")));
            this.BtnExportFromSap.Label = "Export from SAP";
            this.BtnExportFromSap.Name = "BtnExportFromSap";
            this.BtnExportFromSap.ScreenTip = "Загрузка файла из SAP";
            this.BtnExportFromSap.ShowImage = true;
            this.BtnExportFromSap.SuperTip = "Выберите файл Excel из SAP и нажмите кнопку ОК";
            this.BtnExportFromSap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExportFromSap_Click);
            // 
            // ButtonOrderFromCS
            // 
            this.ButtonOrderFromCS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonOrderFromCS.Image = ((System.Drawing.Image)(resources.GetObject("ButtonOrderFromCS.Image")));
            this.ButtonOrderFromCS.Label = "Order from CS";
            this.ButtonOrderFromCS.Name = "ButtonOrderFromCS";
            this.ButtonOrderFromCS.ScreenTip = "Загрузка заявки от customer servises";
            this.ButtonOrderFromCS.ShowImage = true;
            this.ButtonOrderFromCS.SuperTip = "Нажмите на кнопку и выберите файл Excel от CS ";
            this.ButtonOrderFromCS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonOrderFromCS_Click);
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
            this.groupEdit.Items.Add(this.ButtonAddAuto);
            this.groupEdit.Items.Add(this.ButtonDeleteAuto);
            this.groupEdit.Items.Add(this.BtnRecalcilate);
            this.groupEdit.Items.Add(this.BtnFillTable);
            this.groupEdit.Label = "Редактирование";
            this.groupEdit.Name = "groupEdit";
            // 
            // ButtonAddAuto
            // 
            this.ButtonAddAuto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonAddAuto.Image = ((System.Drawing.Image)(resources.GetObject("ButtonAddAuto.Image")));
            this.ButtonAddAuto.Label = "Добавить авто";
            this.ButtonAddAuto.Name = "ButtonAddAuto";
            this.ButtonAddAuto.ScreenTip = "Добавление новой машины без товаров";
            this.ButtonAddAuto.ShowImage = true;
            this.ButtonAddAuto.SuperTip = "При необходимости разделить груз на несколько машин";
            this.ButtonAddAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAddAuto_Click);
            // 
            // ButtonDeleteAuto
            // 
            this.ButtonDeleteAuto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonDeleteAuto.Image = ((System.Drawing.Image)(resources.GetObject("ButtonDeleteAuto.Image")));
            this.ButtonDeleteAuto.Label = "Удалить авто";
            this.ButtonDeleteAuto.Name = "ButtonDeleteAuto";
            this.ButtonDeleteAuto.ScreenTip = "Удаляет выбранное авто";
            this.ButtonDeleteAuto.ShowImage = true;
            this.ButtonDeleteAuto.SuperTip = "Выделите 1 авто, которое необходимо удалить. Можно выбрать одну ячейку или строку" +
    " целиком";
            this.ButtonDeleteAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonDeleteAuto_Click);
            // 
            // BtnRecalcilate
            // 
            this.BtnRecalcilate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnRecalcilate.Image = ((System.Drawing.Image)(resources.GetObject("BtnRecalcilate.Image")));
            this.BtnRecalcilate.Label = "Пересчитать маршруты";
            this.BtnRecalcilate.Name = "BtnRecalcilate";
            this.BtnRecalcilate.ScreenTip = "Пересчет транспорта";
            this.BtnRecalcilate.ShowImage = true;
            this.BtnRecalcilate.SuperTip = "Пересчитывает стоимость транспорта, а также выбирает оптимального провайдера";
            this.BtnRecalcilate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRecalcilate_Click);
            // 
            // BtnFillTable
            // 
            this.BtnFillTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnFillTable.Image = ((System.Drawing.Image)(resources.GetObject("BtnFillTable.Image")));
            this.BtnFillTable.Label = "Перенести в отгрузки";
            this.BtnFillTable.Name = "BtnFillTable";
            this.BtnFillTable.ScreenTip = "Перенос данных в таблицу открузки";
            this.BtnFillTable.ShowImage = true;
            this.BtnFillTable.SuperTip = "Переносит данные из таблиц товары и доставки в таблицу отгрузки";
            this.BtnFillTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFillTable_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.BtnSendShippingCompany);
            this.group1.Items.Add(this.BtnReadCarrierInvoice);
            this.group1.Items.Add(this.ButtonSendToCS);
            this.group1.Label = "Сообщения";
            this.group1.Name = "group1";
            // 
            // BtnSendShippingCompany
            // 
            this.BtnSendShippingCompany.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSendShippingCompany.Image = ((System.Drawing.Image)(resources.GetObject("BtnSendShippingCompany.Image")));
            this.BtnSendShippingCompany.Label = "Заказать перевозку";
            this.BtnSendShippingCompany.Name = "BtnSendShippingCompany";
            this.BtnSendShippingCompany.ScreenTip = "Подготовка писем провайдерам";
            this.BtnSendShippingCompany.ShowImage = true;
            this.BtnSendShippingCompany.SuperTip = "Создает письма провайдером, со списком отгрузки для дальнейшего заполнения данных" +
    " о перевозчике";
            this.BtnSendShippingCompany.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSendShippingCompany_Click);
            // 
            // BtnReadCarrierInvoice
            // 
            this.BtnReadCarrierInvoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnReadCarrierInvoice.Image = ((System.Drawing.Image)(resources.GetObject("BtnReadCarrierInvoice.Image")));
            this.BtnReadCarrierInvoice.Label = "Сканировать ответ";
            this.BtnReadCarrierInvoice.Name = "BtnReadCarrierInvoice";
            this.BtnReadCarrierInvoice.ScreenTip = "Сканирование писем от провайдеров";
            this.BtnReadCarrierInvoice.ShowImage = true;
            this.BtnReadCarrierInvoice.SuperTip = "Скарирует письма от провайдеров и переносит в таблицу отгрузки данные о перевозчи" +
    "ках";
            this.BtnReadCarrierInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnReadCarrierInvoice_Click);
            // 
            // ButtonSendToCS
            // 
            this.ButtonSendToCS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSendToCS.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSendToCS.Image")));
            this.ButtonSendToCS.Label = "Отправить в CS";
            this.ButtonSendToCS.Name = "ButtonSendToCS";
            this.ButtonSendToCS.ScreenTip = "Отправить в CS";
            this.ButtonSendToCS.ShowImage = true;
            this.ButtonSendToCS.SuperTip = "Подготовка письма с данными об отгрузке для Custom Servises";
            this.ButtonSendToCS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSendToCS_Click);
            // 
            // settings
            // 
            this.settings.Items.Add(this.BtnSaveSignature);
            this.settings.Items.Add(this.ButtonSelectFoldersOutlook);
            this.settings.Items.Add(this.ButtonSettingLetterCS);
            this.settings.Label = "Настройки";
            this.settings.Name = "settings";
            // 
            // BtnSaveSignature
            // 
            this.BtnSaveSignature.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSaveSignature.Image = ((System.Drawing.Image)(resources.GetObject("BtnSaveSignature.Image")));
            this.BtnSaveSignature.Label = "Сохранить подпись";
            this.BtnSaveSignature.Name = "BtnSaveSignature";
            this.BtnSaveSignature.ScreenTip = "Сохранить подпись";
            this.BtnSaveSignature.ShowImage = true;
            this.BtnSaveSignature.SuperTip = "Заполните данные подписи на листе Mail и нажмите сохранить";
            this.BtnSaveSignature.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSaveSignature_Click);
            // 
            // ButtonSelectFoldersOutlook
            // 
            this.ButtonSelectFoldersOutlook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSelectFoldersOutlook.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSelectFoldersOutlook.Image")));
            this.ButtonSelectFoldersOutlook.Label = "Папки с письмами";
            this.ButtonSelectFoldersOutlook.Name = "ButtonSelectFoldersOutlook";
            this.ButtonSelectFoldersOutlook.ScreenTip = "Выбор папок сканирования";
            this.ButtonSelectFoldersOutlook.ShowImage = true;
            this.ButtonSelectFoldersOutlook.SuperTip = "Выбор папок outlook, в которые сохраняются письма от провайдеров с информацией о " +
    "водителях";
            this.ButtonSelectFoldersOutlook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSelectFoldersOutlook_Click);
            // 
            // ButtonSettingLetterCS
            // 
            this.ButtonSettingLetterCS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSettingLetterCS.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSettingLetterCS.Image")));
            this.ButtonSettingLetterCS.Label = "Письмо CS";
            this.ButtonSettingLetterCS.Name = "ButtonSettingLetterCS";
            this.ButtonSettingLetterCS.ScreenTip = "Настройки письма для CS";
            this.ButtonSettingLetterCS.ShowImage = true;
            this.ButtonSettingLetterCS.SuperTip = "Настройка шаблона письма, которое отправляется Customer Servises с файлом отгрузк" +
    "и";
            this.ButtonSettingLetterCS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSettingLetterCS_Click);
            // 
            // about
            // 
            this.about.Items.Add(this.BtnAboutProgrramm);
            this.about.Items.Add(this.btnSetts);
            this.about.Label = "Справка";
            this.about.Name = "about";
            // 
            // BtnAboutProgrramm
            // 
            this.BtnAboutProgrramm.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnAboutProgrramm.Image = ((System.Drawing.Image)(resources.GetObject("BtnAboutProgrramm.Image")));
            this.BtnAboutProgrramm.Label = "О программе";
            this.BtnAboutProgrramm.Name = "BtnAboutProgrramm";
            this.BtnAboutProgrramm.ScreenTip = "Некотороые сведения о программе";
            this.BtnAboutProgrramm.ShowImage = true;
            this.BtnAboutProgrramm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAboutProgrramm_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnExportFromSap;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEdit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRecalcilate;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab ShefflerRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSendShippingCompany;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnReadCarrierInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup settings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetts;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup about;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAboutProgrramm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonAddAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonDeleteAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadAllOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonOrderFromCS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFillTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSaveSignature;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSelectFoldersOutlook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSendToCS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSettingLetterCS;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDelivery Ribbon
        {
            get { return this.GetRibbon<RibbonDelivery>(); }
        }
    }
}
