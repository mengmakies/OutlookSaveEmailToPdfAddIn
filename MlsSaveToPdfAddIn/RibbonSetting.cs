using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using NReco.PdfGenerator;
using Exception = System.Exception;
namespace MlsSaveToPdfAddIn
{

    public partial class RibbonSetting
    {
        // PDF保存目录
        private const string DIR_PDF = @"c:\SaveToPdf\";

        BackgroundWorker mBckWorker = null;
        private ProgressBar mProgressBar = null;
        private FrmLoading mForm = null;
        private int mMailCount = 0;

        private void RibbonSetting_Load(object sender, RibbonUIEventArgs e)
        {
            mBckWorker = new BackgroundWorker();
            mBckWorker.DoWork += MBckWorker_DoWork;
            mBckWorker.RunWorkerCompleted += MBckWorker_RunWorkerCompleted;
            mBckWorker.ProgressChanged += MBckWorker_ProgressChanged;
            mBckWorker.WorkerReportsProgress = true;

            mProgressBar = new ProgressBar();
            mProgressBar.Visible = true;
            Screen scr = Screen.PrimaryScreen;
            mProgressBar.Left = (scr.WorkingArea.Width - mProgressBar.Width) / 2;
            mProgressBar.Top = (scr.WorkingArea.Height - mProgressBar.Height) / 2;
            mProgressBar.Text = "邮件转换中，请稍后...";

            mForm = new FrmLoading()
            {
                Msg = "邮件转换中，请稍后...",
                StartPosition = FormStartPosition.CenterScreen
            };
        }

        private void MBckWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int index = e.ProgressPercentage;
            if (index > 0)
            {
                mForm.Msg = string.Format(@"邮件转换中（{0}/{1}），请稍后...", index, mMailCount);
            }
        }

        private void MBckWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var explorer = base.Context as Explorer;
                int i = 1;
                mMailCount = explorer.CurrentFolder.Items.Count;
                foreach (var item in explorer.CurrentFolder.Items)
                {
                    if (!(item is MailItem)) continue;
                    var mail = item as MailItem;

                    if (string.IsNullOrEmpty(mail.HTMLBody)) continue;

                    var fileName = Path.Combine(DIR_PDF, string.Format(@"{0}_{1}.pdf", mail.ReceivedTime.ToString("yyyyMMddHHmmss"), mail.Subject));

                    var html = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                 <!DOCTYPE html 
                     PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""
                    ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
                 <html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">
                    <head>
                        <title>Minimal XHTML 1.0 Document with W3C DTD</title>
                    </head>
                  <body>
                    " + mail.HTMLBody + "</body></html>";

                    var pdfDoc = new HtmlToPdfConverter
                    {
                        Margins = new PageMargins { Bottom = 50, Left = 50, Right = 50, Top = 50 },
                        Size = PageSize.A4,
                        PageHeaderHtml = string.Format(@"<h1>{0}</h1><h4>{1}({2})</h4><h5>收件时间：{3}</h5>", mail.Subject, mail.SenderName, mail.SenderEmailAddress, mail.ReceivedTime),
                        PageFooterHtml = string.Format(@"<h3><a href='http://www.mln.com'>{0}</a></h3>", "@Power by Mln.COM")
                    };

                    var pdfBytes = pdfDoc.GeneratePdf(html);

                    using (var file = new FileStream(fileName, FileMode.Create))
                    {
                        file.Write(pdfBytes, 0, (int)pdfBytes.Length);
                    }

                    mBckWorker.ReportProgress(i++);
                    

                }
                e.Result = "";
            }
            catch (Exception ex)
            {
                e.Result = ex.Message;
            }
        }
        
        private void MBckWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            mForm.Hide();
            string result = e.Result.ToString();
            if (result == "")
            {
                MessageBox.Show("转换PDF成功！PDF所在目录：" + DIR_PDF);
            }
            else
            {
                MessageBox.Show("转换出错！" + result);
            }
        }

        private void btnSaveToPdf_Click(object sender, RibbonControlEventArgs e)
        {
            var explorer = base.Context as Explorer;
            if (explorer == null) return;

            int count = explorer.CurrentFolder.Items.Count;
            if (count < 1)
            {
                MessageBox.Show("请选择要转换pdf的文件夹");
                return;
            }

            if (!Directory.Exists(DIR_PDF))
            {
                Directory.CreateDirectory(DIR_PDF);
            }

            mBckWorker.RunWorkerAsync();
            mForm.ShowDialog();
        }

    }
}
