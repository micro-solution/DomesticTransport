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
            this.splitButtonUpdateAuto = this.Factory.CreateRibbonSplitButton();
            this.ButtonUpdateAutoMain = this.Factory.CreateRibbonButton();
            this.ButtonUpdateAutoSecond = this.Factory.CreateRibbonButton();
            this.ButtonAddAuto = this.Factory.CreateRibbonButton();
            this.ButtonDeleteAuto = this.Factory.CreateRibbonButton();
            this.btnSaveRoute = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.BtnRecalcilate = this.Factory.CreateRibbonButton();
            this.btnNunerateDeliveries = this.Factory.CreateRibbonButton();
            this.BtnFillTable = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BtnSendShippingCompany = this.Factory.CreateRibbonButton();
            this.BtnReadCarrierInvoice = this.Factory.CreateRibbonButton();
            this.MenuSendTotal = this.Factory.CreateRibbonMenu();
            this.ButtonTotalToProviders = this.Factory.CreateRibbonButton();
            this.ButtonSendToCS = this.Factory.CreateRibbonButton();
            this.ButtonToArchive = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.ToTransportTable = this.Factory.CreateRibbonButton();
            this.ButtonSendTransportTable = this.Factory.CreateRibbonButton();
            this.ButtonScanTransportTable = this.Factory.CreateRibbonButton();
            this.settings = this.Factory.CreateRibbonGroup();
            this.ButtonSettingLetter = this.Factory.CreateRibbonButton();
            this.BtnSaveSignature = this.Factory.CreateRibbonButton();
            this.btnDate = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.ButtonSelectFoldersOutlook = this.Factory.CreateRibbonButton();
            this.BtnAboutProgrramm = this.Factory.CreateRibbonButton();
            this.btnChangeRoute = this.Factory.CreateRibbonSplitButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.ShefflerRibbon.SuspendLayout();
            this.groupGeneral.SuspendLayout();
            this.groupEdit.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.settings.SuspendLayout();
            this.SuspendLayout();
            // 
            // ShefflerRibbon
            // 
            this.ShefflerRibbon.Groups.Add(this.groupGeneral);
            this.ShefflerRibbon.Groups.Add(this.groupEdit);
            this.ShefflerRibbon.Groups.Add(this.group1);
            this.ShefflerRibbon.Groups.Add(this.group3);
            this.ShefflerRibbon.Groups.Add(this.settings);
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
            this.BtnExportFromSap.ScreenTip = "Загрузка поставок из файла SAP";
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
            this.ButtonOrderFromCS.ScreenTip = "Загрузка заявки от CS IND";
            this.ButtonOrderFromCS.ShowImage = true;
            this.ButtonOrderFromCS.SuperTip = "Нажмите на кнопку и выберите файл Excel с заявкой от CS IND";
            this.ButtonOrderFromCS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonOrderFromCS_Click);
            // 
            // BtnLoadAllOrders
            // 
            this.BtnLoadAllOrders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadAllOrders.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadAllOrders.Image")));
            this.BtnLoadAllOrders.Label = "Загрузить All Orders";
            this.BtnLoadAllOrders.Name = "BtnLoadAllOrders";
            this.BtnLoadAllOrders.ScreenTip = "Загрузка данных по поставкам";
            this.BtnLoadAllOrders.ShowImage = true;
            this.BtnLoadAllOrders.SuperTip = "Выберите файл с выгрузкой из SAP с информацией о собранных поставках для загрузки" +
    " брутто веса, количества паллет, стоимости поставки";
            this.BtnLoadAllOrders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadAllOrders_Click);
            // 
            // groupEdit
            // 
            this.groupEdit.Items.Add(this.splitButtonUpdateAuto);
            this.groupEdit.Items.Add(this.ButtonAddAuto);
            this.groupEdit.Items.Add(this.ButtonDeleteAuto);
            this.groupEdit.Items.Add(this.btnSaveRoute);
            this.groupEdit.Items.Add(this.separator1);
            this.groupEdit.Items.Add(this.BtnRecalcilate);
            this.groupEdit.Items.Add(this.btnNunerateDeliveries);
            this.groupEdit.Items.Add(this.BtnFillTable);
            this.groupEdit.Label = "Формирование транспорта";
            this.groupEdit.Name = "groupEdit";
            // 
            // splitButtonUpdateAuto
            // 
            this.splitButtonUpdateAuto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonUpdateAuto.Image = ((System.Drawing.Image)(resources.GetObject("splitButtonUpdateAuto.Image")));
            this.splitButtonUpdateAuto.Items.Add(this.ButtonUpdateAutoMain);
            this.splitButtonUpdateAuto.Items.Add(this.ButtonUpdateAutoSecond);
            this.splitButtonUpdateAuto.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonUpdateAuto.Label = "Обновить авто";
            this.splitButtonUpdateAuto.Name = "splitButtonUpdateAuto";
            this.splitButtonUpdateAuto.ScreenTip = "Обновление поставок";
            this.splitButtonUpdateAuto.SuperTip = "Программа пересчитывает все поставки. По умолчанию используются только основные м" +
    "аршруты";
            this.splitButtonUpdateAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SplitButtonUpdateAuto_Click);
            // 
            // ButtonUpdateAutoMain
            // 
            this.ButtonUpdateAutoMain.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonUpdateAutoMain.Image = ((System.Drawing.Image)(resources.GetObject("ButtonUpdateAutoMain.Image")));
            this.ButtonUpdateAutoMain.Label = "Используя основные маршруты";
            this.ButtonUpdateAutoMain.Name = "ButtonUpdateAutoMain";
            this.ButtonUpdateAutoMain.ScreenTip = "По основным маршрутам";
            this.ButtonUpdateAutoMain.ShowImage = true;
            this.ButtonUpdateAutoMain.SuperTip = "Пересчет всех данных и формирование нового списка доставок. Используются только о" +
    "сновные маршруты";
            this.ButtonUpdateAutoMain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonUpdateAutoMain_Click);
            // 
            // ButtonUpdateAutoSecond
            // 
            this.ButtonUpdateAutoSecond.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonUpdateAutoSecond.Image = ((System.Drawing.Image)(resources.GetObject("ButtonUpdateAutoSecond.Image")));
            this.ButtonUpdateAutoSecond.Label = "Объединить неукомплектованные авто";
            this.ButtonUpdateAutoSecond.Name = "ButtonUpdateAutoSecond";
            this.ButtonUpdateAutoSecond.ShowImage = true;
            this.ButtonUpdateAutoSecond.SuperTip = "Пересчет всех данных и формирование нового списка доставок. Программа пытается до" +
    "укомплектовать транспорт с учетом второстепенных маршрутов";
            this.ButtonUpdateAutoSecond.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonUpdateAutoSecond_Click);
            // 
            // ButtonAddAuto
            // 
            this.ButtonAddAuto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonAddAuto.Image = ((System.Drawing.Image)(resources.GetObject("ButtonAddAuto.Image")));
            this.ButtonAddAuto.Label = "Добавить авто";
            this.ButtonAddAuto.Name = "ButtonAddAuto";
            this.ButtonAddAuto.ScreenTip = "Добавление новой машины";
            this.ButtonAddAuto.ShowImage = true;
            this.ButtonAddAuto.SuperTip = "Используется при необходимости разделить доставку на несколько машин. Выделите яч" +
    "ейки с нужными поставками и нажмите кнопку Добавить авто";
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
            // btnSaveRoute
            // 
            this.btnSaveRoute.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveRoute.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveRoute.Image")));
            this.btnSaveRoute.Label = "Сохранить маршрут";
            this.btnSaveRoute.Name = "btnSaveRoute";
            this.btnSaveRoute.ScreenTip = "Сохраняет отредактированные маршруты, если их еще нет в таблице";
            this.btnSaveRoute.ShowImage = true;
            this.btnSaveRoute.SuperTip = "Измените маршруты на листе Delivery и нажмите эту кнопку, чтобы сохранить новые м" +
    "аршруты в таблицу Routes";
            this.btnSaveRoute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSaveRoute_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // BtnRecalcilate
            // 
            this.BtnRecalcilate.Image = ((System.Drawing.Image)(resources.GetObject("BtnRecalcilate.Image")));
            this.BtnRecalcilate.Label = "Пересчитать стоимость";
            this.BtnRecalcilate.Name = "BtnRecalcilate";
            this.BtnRecalcilate.ScreenTip = "Пересчет стоимости доставок";
            this.BtnRecalcilate.ShowImage = true;
            this.BtnRecalcilate.SuperTip = "По сформированным маршрутам определятся оптимальный провайдер и рассчитывается ст" +
    "оимось перевозки";
            this.BtnRecalcilate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnChangeRoute_Click);
            // 
            // btnNunerateDeliveries
            // 
            this.btnNunerateDeliveries.Image = ((System.Drawing.Image)(resources.GetObject("btnNunerateDeliveries.Image")));
            this.btnNunerateDeliveries.Label = "Сортировка отгрузок";
            this.btnNunerateDeliveries.Name = "btnNunerateDeliveries";
            this.btnNunerateDeliveries.ScreenTip = "Сортировка отгрузок";
            this.btnNunerateDeliveries.ShowImage = true;
            this.btnNunerateDeliveries.SuperTip = "Сортирует отгрузки и восстанавливает нумерацию (МСК -> Регионы -> СНГ -> LTL -> С" +
    "борный груз)";
            this.btnNunerateDeliveries.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NunerateDeliveries);
            // 
            // BtnFillTable
            // 
            this.BtnFillTable.Image = ((System.Drawing.Image)(resources.GetObject("BtnFillTable.Image")));
            this.BtnFillTable.Label = "Перенести в Отгрузки";
            this.BtnFillTable.Name = "BtnFillTable";
            this.BtnFillTable.ScreenTip = "Перенос данных в таблицу Открузки";
            this.BtnFillTable.ShowImage = true;
            this.BtnFillTable.SuperTip = "Переносит данные из таблиц Транспорт и Поставки в таблицу Отгрузки";
            this.BtnFillTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFillTable_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.BtnSendShippingCompany);
            this.group1.Items.Add(this.BtnReadCarrierInvoice);
            this.group1.Items.Add(this.MenuSendTotal);
            this.group1.Items.Add(this.ButtonToArchive);
            this.group1.Label = "Отгрузки";
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
            this.BtnSendShippingCompany.SuperTip = "Создает письма провайдерам со списком отгрузок для дальнейшего заполнения данных " +
    "о перевозке";
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
            this.BtnReadCarrierInvoice.SuperTip = "Скарирует письма от провайдеров и переносит заполненные ими данные в таблицу Отгр" +
    "узки";
            this.BtnReadCarrierInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnReadCarrierInvoice_Click);
            // 
            // MenuSendTotal
            // 
            this.MenuSendTotal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MenuSendTotal.Image = ((System.Drawing.Image)(resources.GetObject("MenuSendTotal.Image")));
            this.MenuSendTotal.Items.Add(this.ButtonTotalToProviders);
            this.MenuSendTotal.Items.Add(this.ButtonSendToCS);
            this.MenuSendTotal.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MenuSendTotal.Label = "Отправка Отгрузок";
            this.MenuSendTotal.Name = "MenuSendTotal";
            this.MenuSendTotal.ShowImage = true;
            // 
            // ButtonTotalToProviders
            // 
            this.ButtonTotalToProviders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonTotalToProviders.Image = ((System.Drawing.Image)(resources.GetObject("ButtonTotalToProviders.Image")));
            this.ButtonTotalToProviders.Label = "Провайдерам";
            this.ButtonTotalToProviders.Name = "ButtonTotalToProviders";
            this.ButtonTotalToProviders.ShowImage = true;
            this.ButtonTotalToProviders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTotalToProviders_Click);
            // 
            // ButtonSendToCS
            // 
            this.ButtonSendToCS.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSendToCS.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSendToCS.Image")));
            this.ButtonSendToCS.Label = "В CS и WH";
            this.ButtonSendToCS.Name = "ButtonSendToCS";
            this.ButtonSendToCS.ScreenTip = "Отправить отгрузки в CS и WH";
            this.ButtonSendToCS.ShowImage = true;
            this.ButtonSendToCS.SuperTip = "Подготовка письма с данными об отгрузках для CS и WH";
            this.ButtonSendToCS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSendToCS_Click);
            // 
            // ButtonToArchive
            // 
            this.ButtonToArchive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonToArchive.Image = ((System.Drawing.Image)(resources.GetObject("ButtonToArchive.Image")));
            this.ButtonToArchive.Label = "Сохранить отгрузки";
            this.ButtonToArchive.Name = "ButtonToArchive";
            this.ButtonToArchive.ShowImage = true;
            this.ButtonToArchive.SuperTip = "Сохранение листа Отгрузки во временный архив";
            this.ButtonToArchive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveToArchive_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.ToTransportTable);
            this.group3.Items.Add(this.ButtonSendTransportTable);
            this.group3.Items.Add(this.ButtonScanTransportTable);
            this.group3.Label = "Отчеты";
            this.group3.Name = "group3";
            // 
            // ToTransportTable
            // 
            this.ToTransportTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ToTransportTable.Image = ((System.Drawing.Image)(resources.GetObject("ToTransportTable.Image")));
            this.ToTransportTable.Label = "Перенести в отчет";
            this.ToTransportTable.Name = "ToTransportTable";
            this.ToTransportTable.ScreenTip = "Перенести в отчет";
            this.ToTransportTable.ShowImage = true;
            this.ToTransportTable.SuperTip = "Перенос данных в отчетные таблицы Shipments и Transport Table. Переносятся данные" +
    " за предыдущую неделю и старше";
            this.ToTransportTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveToTransportTable_Click);
            // 
            // ButtonSendTransportTable
            // 
            this.ButtonSendTransportTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonSendTransportTable.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSendTransportTable.Image")));
            this.ButtonSendTransportTable.Label = "Отчет провайдеру";
            this.ButtonSendTransportTable.Name = "ButtonSendTransportTable";
            this.ButtonSendTransportTable.ShowImage = true;
            this.ButtonSendTransportTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSendTransportTable_Click);
            // 
            // ButtonScanTransportTable
            // 
            this.ButtonScanTransportTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonScanTransportTable.Image = ((System.Drawing.Image)(resources.GetObject("ButtonScanTransportTable.Image")));
            this.ButtonScanTransportTable.Label = "Сканировать ответ";
            this.ButtonScanTransportTable.Name = "ButtonScanTransportTable";
            this.ButtonScanTransportTable.ShowImage = true;
            this.ButtonScanTransportTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonScanTransportTable_Click);
            // 
            // settings
            // 
            this.settings.Items.Add(this.ButtonSettingLetter);
            this.settings.Items.Add(this.BtnSaveSignature);
            this.settings.Items.Add(this.btnDate);
            this.settings.Items.Add(this.btnSettings);
            this.settings.Items.Add(this.ButtonSelectFoldersOutlook);
            this.settings.Items.Add(this.BtnAboutProgrramm);
            this.settings.Label = "Настройки";
            this.settings.Name = "settings";
            // 
            // ButtonSettingLetter
            // 
            this.ButtonSettingLetter.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSettingLetter.Image")));
            this.ButtonSettingLetter.Label = "Настройка писем";
            this.ButtonSettingLetter.Name = "ButtonSettingLetter";
            this.ButtonSettingLetter.ShowImage = true;
            this.ButtonSettingLetter.SuperTip = "Настройка шаблонов писем";
            this.ButtonSettingLetter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSettingLetterCS_Click);
            // 
            // BtnSaveSignature
            // 
            this.BtnSaveSignature.Image = ((System.Drawing.Image)(resources.GetObject("BtnSaveSignature.Image")));
            this.BtnSaveSignature.Label = "Подпись";
            this.BtnSaveSignature.Name = "BtnSaveSignature";
            this.BtnSaveSignature.ScreenTip = "Сохранить подпись";
            this.BtnSaveSignature.ShowImage = true;
            this.BtnSaveSignature.SuperTip = "Заполните данные подписи на листе Mail и нажмите сохранить";
            this.BtnSaveSignature.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSaveSignature_Click);
            // 
            // btnDate
            // 
            this.btnDate.Image = ((System.Drawing.Image)(resources.GetObject("btnDate.Image")));
            this.btnDate.Label = "Выбрать дату";
            this.btnDate.Name = "btnDate";
            this.btnDate.ShowImage = true;
            this.btnDate.SuperTip = "Установить дату отгрузки";
            this.btnDate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDate_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSettings.Image")));
            this.btnSettings.Label = "Настройки";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSettings_Click);
            // 
            // ButtonSelectFoldersOutlook
            // 
            this.ButtonSelectFoldersOutlook.Image = ((System.Drawing.Image)(resources.GetObject("ButtonSelectFoldersOutlook.Image")));
            this.ButtonSelectFoldersOutlook.Label = "Папки с письмами";
            this.ButtonSelectFoldersOutlook.Name = "ButtonSelectFoldersOutlook";
            this.ButtonSelectFoldersOutlook.ScreenTip = "Выбор папок сканирования";
            this.ButtonSelectFoldersOutlook.ShowImage = true;
            this.ButtonSelectFoldersOutlook.SuperTip = "Выбор папок outlook, в которые сохраняются письма с заявками от провайдеров";
            this.ButtonSelectFoldersOutlook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSelectFoldersOutlook_Click);
            // 
            // BtnAboutProgrramm
            // 
            this.BtnAboutProgrramm.Image = ((System.Drawing.Image)(resources.GetObject("BtnAboutProgrramm.Image")));
            this.BtnAboutProgrramm.Label = "О программе";
            this.BtnAboutProgrramm.Name = "BtnAboutProgrramm";
            this.BtnAboutProgrramm.ShowImage = true;
            this.BtnAboutProgrramm.SuperTip = "Некотороые сведения о программе";
            this.BtnAboutProgrramm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAboutProgrramm_Click);
            // 
            // btnChangeRoute
            // 
            this.btnChangeRoute.Label = "";
            this.btnChangeRoute.Name = "btnChangeRoute";
            // 
            // button3
            // 
            this.button3.Label = "";
            this.button3.Name = "button3";
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.settings.ResumeLayout(false);
            this.settings.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAboutProgrramm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonAddAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonDeleteAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadAllOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonOrderFromCS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSaveSignature;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSelectFoldersOutlook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSendToCS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSettingLetter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveRoute;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNunerateDeliveries;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btnChangeRoute;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonUpdateAutoMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonUpdateAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonUpdateAutoSecond;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MenuSendTotal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonTotalToProviders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonToArchive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ToTransportTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonSendTransportTable;
        protected internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFillTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonScanTransportTable;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDelivery Ribbon
        {
            get { return this.GetRibbon<RibbonDelivery>(); }
        }
    }
}
