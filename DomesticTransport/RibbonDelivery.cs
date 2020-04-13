using DomesticTransport.Forms;

using Microsoft.Office.Tools.Ribbon;

using System;
using System.Windows.Forms;

namespace DomesticTransport
{
    public partial class RibbonDelivery
    {
        private void BtnStart_Click(object sender, RibbonControlEventArgs e)
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

        private void btnSendShippingCompany_Click(object sender, RibbonControlEventArgs e)
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



        private void button1_Click(object sender, RibbonControlEventArgs e)
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

        private void button2_Click(object sender, RibbonControlEventArgs e)
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

        private void btnChangeSet_Click(object sender, RibbonControlEventArgs e)
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

        private void btnReadForms_Click(object sender, RibbonControlEventArgs e)
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

        private void btnAccept_Click(object sender, RibbonControlEventArgs e)
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

        private void btnSaveSignature_Click(object sender, RibbonControlEventArgs e)
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

        private void btnAboutProgrramm_Click(object sender, RibbonControlEventArgs e)
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
        /// Сканирование писем
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>   
        private void btnReadCarrierInvoice_Click_1(object sender, RibbonControlEventArgs e)
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
                    scanMail.GetMessage();
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
    }
}
