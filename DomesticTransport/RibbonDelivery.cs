using DomesticTransport.Forms;

using Microsoft.Office.Tools.Ribbon;

using System;
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
                new Functions().CreateMasseges();
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
        /// Пересчет маршрутов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRecalcilate_Click(object sender, RibbonControlEventArgs e)
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
            try
            {
                ShefflerWB.ExcelOptimizateOn();
                Email.WriteReestrSignature();
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
                    MessageBox.Show("Сегодня письма не обнаружены", "Сканирование почты", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            SettingLetterToCS setting = new SettingLetterToCS();
            setting.ShowDialog();
        }

        private void btnSaveRoute_Click(object sender, RibbonControlEventArgs e)
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
    }
}
