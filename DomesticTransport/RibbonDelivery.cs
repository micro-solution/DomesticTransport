using DomesticTransport.Forms;

using Microsoft.Office.Tools.Ribbon;

using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace DomesticTransport
{
    public partial class RibbonDelivery
    {
        /// <summary>
        /// Кнопка Export from SAP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportFromSap_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().ExportFromSAP();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка отправки писем провайдерам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSendShippingCompany_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().SendEmailToProviderAdoutOrders();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка создания нового авто
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAddAuto_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().AddAuto();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка удаления авто
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonDeleteAuto_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().DeleteAuto();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка пересчитать стоимость
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnChangeRoute_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().СhangeDelivery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка загрузки файла All Orders
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadAllOrders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().LoadAllOrders();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка загрузки файла от CS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonOrderFromCS_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().ExportFromCS();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка перенести в отгрузки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFillTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().UpdateTotal();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка сохранения подписи
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSaveSignature_Click(object sender, RibbonControlEventArgs e)
        {
            SignatureSelect signatureSelect = new SignatureSelect();
            signatureSelect.ShowDialog();
        }

        /// <summary>
        /// Кнопка о программе
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAboutProgrramm_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new About().ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Выбор папки для сканирования писем от провайдеров
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSelectFoldersOutlook_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new OutlookFoldersSelect().ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка сканирования писем от провайдеров
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>   
        private void BtnReadCarrierInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                if (Properties.Settings.Default.OutlookFolders == "")
                {
                    MessageBox.Show("Задайте папки для сканирования почты", "Необходима настройка программы", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                ScanMail scanMail = new ScanMail();
                if (scanMail.SaveAttachments() == 0)
                {
                    MessageBox.Show("Письма не обнаружены", "Сканирование почты", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    scanMail.GetDataFromProviderFiles();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Отправка файла отгрузки в CS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSendToCS_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().CreateLetterToCS();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Настройки письма для CS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSettingLetterCS_Click(object sender, RibbonControlEventArgs e)
        {
            SettingLetters setting = new SettingLetters();
            setting.ShowDialog();
        }

        /// <summary>
        /// Кнопка сохранить маршрут
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSaveRoute_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().SaveRoute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Сквозное нумерование доставок 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NunerateDeliveries(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().RenumberDeliveries();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Обновление всех маршрутов по основным маршрутам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonUpdateAutoMain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().UpdateAutoMain();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Сплитбатон обновления авто
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SplitButtonUpdateAuto_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().UpdateAutoMain();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка обновления авто по второстепенным маршрутам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonUpdateAutoSecond_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().UpdateAutoSecond();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Отправка таблицы Отгрузка провайдерам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonTotalToProviders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Functions().SendEmailToProviderAdoutAdding();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        private void btnDate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                new ShefflerWB().SetDateCell();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Сохранение листа отгрузки во временный архив
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveToArchive_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                Archive.LoadToArhive();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Сохранение временного архива в TransportTable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveToTransportTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                Archive.ToTransportTableAndShepments();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }

        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new Settings().ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        /// <summary>
        /// Кнопка отправки очета провайдеру
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSendTransportTable_Click(object sender, RibbonControlEventArgs e)
        {
            TransportTableSending tableSending = new TransportTableSending();
            tableSending.ShowDialog();
            if (tableSending.DialogResult != DialogResult.OK) return;

            try
            {
                ShefflerWB.ExcelOptimizateOn();
                new TransportTable().MessageProvider(tableSending.DateStart, tableSending.DateEnd, tableSending.Provider);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }

        private void helper_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (File.Exists(Properties.Settings.Default.HelpPath))
                {
                    Process.Start(Properties.Settings.Default.HelpPath);
                }
                else
                {
                    new Settings().ShowDialog();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// Кнопка сканирования почты для импорта отчетов провайдера
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonScanTransportTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                if (Properties.Settings.Default.OutlookFolders == "")
                {
                    MessageBox.Show("Задайте папки для сканирования почты", "Необходима настройка программы", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                ScanMail scanMail = new ScanMail();
                
                if (scanMail.SaveAttachments() == 0)
                {
                    MessageBox.Show("Письма не обнаружены", "Сканирование почты", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    TransportTable transportTable = new TransportTable();
                    transportTable.GetDataFromProviderFiles();
                    transportTable.SaveAndClose();
                    MessageBox.Show("Данные импортированы. Изменения выделены цветом.", "Импорт данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ShefflerWB.ExcelOptimizateOff();
            }
        }
    }
}
