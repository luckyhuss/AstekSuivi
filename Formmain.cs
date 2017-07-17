using AstekSuivi.Model;
using AstekSuivi.Service;
using OfficeOpenXml;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AstekSuivi
{
    public partial class FormMain : Form
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _application = new Microsoft.Office.Interop.Outlook.Application();
        private readonly string pathLot2 = ConfigurationManager.AppSettings["Path.Lot2"];
        private Microsoft.Office.Interop.Outlook.MailItem __mail = null;
        private string mailNameForExcel = string.Empty;
        private string mailBodyDelimiter = "◆";
        private int mailBodyLength = 250;

        private enum ExcelColumns
        {
            Semestre = 1,
            Mois =2,
            Date_Demande = 3,
            Demandeur_Destinataires = 4,
            Sujet = 5,
            Demande = 6,
            Mail = 7,
            Nom_KPI = 8,
            Etat = 9,
            Conso = 10, // Lot 2.3
            Vendue = 11 // Lot 2.3
        }

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void FormMain_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                StringBuilder sbRecipients = null;
                mailNameForExcel = string.Empty;
                // supports a drop of a Outlook message
                foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in _application.ActiveExplorer().Selection)
                {
                    __mail = mailItem;
                    DateTime dateToken = mailItem.SentOn;
                    textBoxMailDate.Text = dateToken.ToString();
                    textBoxSender.Text = mailItem.SenderEmailAddress;
                    
                    sbRecipients = new StringBuilder();
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in mailItem.Recipients)
                    {
                        sbRecipients.Append(recipient.Address).Append("; ");
                    }
                    textBoxRecipients.Text = sbRecipients.ToString().Trim();

                    mailNameForExcel = Path.Combine(dateToken.Year.ToString(), dateToken.ToString("yyyyMMdd-HHmmss") + ".msg");
                    textBoxFilenameMail.Tag = Path.Combine(pathLot2, mailNameForExcel);
                    textBoxFilenameExcel.Tag = String.Format(pathLot2, "{0}", ConfigurationManager.AppSettings["File.Suivi"]);

                    // set remaining text area
                    string fwdEmailAddress = SetProject(mailItem.Subject, mailItem.Body);

                    /// 17/07/2017 - Forward email to collegues                    
                    Microsoft.Office.Interop.Outlook.MailItem fwdMailItem = mailItem.Forward();
                    
                    fwdMailItem.Recipients.Add(fwdEmailAddress);
                    fwdMailItem.Display(true);
                    
                    // first mail item only is considered -> so break from loop
                    break;
                }
            }
            else
            {
                MessageBox.Show(this, "This is not an Outlook item.\nPlease drag and drop an outlook item", "File error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string SetProject(string subject, string body)
        {
            comboBoxProject.SelectedIndex = -1;
            string fwdEmailAddress = string.Empty;

            // check project in email

            // check this first because of my signature in mails
            if (subject.ToLower().Contains("siclop") || body.ToLower().Contains("siclop\r\n"))
            {
                comboBoxProject.Text = "SICLOP";
                fwdEmailAddress = "cfooahpiang@astek.mu";
            }

            if (subject.ToLower().Contains("aspin") || body.ToLower().Contains("aspin\r\n"))
            {
                comboBoxProject.Text = "ASPIN";
                fwdEmailAddress = "smaregadee@astek.mu";
            }

            if (subject.ToLower().Contains("spid") || body.ToLower().Contains("spid\r\n"))
            {
                comboBoxProject.Text = "SPID";
                fwdEmailAddress = "smaregadee@astek.mu";
            }           

            if (subject.ToLower().Contains("scoop") || body.ToLower().Contains("scoop\r\n"))
            {
                comboBoxProject.Text = "SCOOP";
                fwdEmailAddress = "cfooahpiang@astek.mu";
            }            

            textBoxMailSubject.Text = subject;
            // remove white spaces / empty lines
            textBoxMailBody.Text = Regex.Replace(body, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);

            /// 17/07/2017 - OCEANE mail
            // serach for "Commentaire de nos services :"
            int index = textBoxMailBody.Text.IndexOf("Commentaire de nos services :\r\n");

            if (index != -1)
            {
                // text found => there mail from OCEANE
                textBoxMailBody.Text = textBoxMailBody.Text.Substring(index, textBoxMailBody.Text.Length - index - 1);

                // replace "Commentaire de nos services :"
                textBoxMailBody.Text = textBoxMailBody.Text.Replace("Commentaire de nos services :\r\n", string.Empty);

                // search "Nous continuons à traiter ce ticket"
                index = textBoxMailBody.Text.IndexOf("Nous continuons à traiter ce ticket");

                if (index != -1)
                {
                    // text found
                    textBoxMailBody.Text = textBoxMailBody.Text.Substring(0, index - 1);
                }
            }

            if (textBoxMailBody.Text.Length < mailBodyLength)
            {
                textBoxMailBody.Text = textBoxMailBody.Text.Insert(textBoxMailBody.Text.Length, mailBodyDelimiter);
            }
            else
            {
                textBoxMailBody.Text = textBoxMailBody.Text.Insert(mailBodyLength, mailBodyDelimiter);
            }

            return fwdEmailAddress;
        }

        private void LoadSettings()
        {
            var settingsData = File.ReadAllLines(ConfigurationManager.AppSettings["File.Settings"], Encoding.Default);

            foreach (var entry in settingsData)
            {
                // menu_type|menu_title|menu_link

                // ignore commmented lines in settings.ini file
                if (entry[0].Equals('#'))
                {
                    continue;
                }

                var namevalue = entry.Split('|');

                CustomContextMenu ccm = new CustomContextMenu(namevalue[0], namevalue[2]);
                var tsi = contextMenuStripMain.Items.Add(namevalue[1]);
                tsi.Tag = ccm;
                contextMenuStripMain.Items.Insert(0, tsi);
            }
        }


        private void LoadControls()
        {
            comboBoxProject.Items.AddRange(ConfigurationManager.AppSettings["Project.Name"].Split(';'));

            // rhombus
            labelChar.Text = mailBodyDelimiter;

            // reset
            textBoxFilenameMail.Tag = textBoxFilenameExcel.Tag = string.Empty;

            buttonAdd.Enabled = buttonOpenExcel.Enabled = false;
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            // load settings
            LoadSettings();

            // load all UI controls
            LoadControls();

            WindowState = FormWindowState.Minimized;

            // display balloon
            notifyIconMain.BalloonTipText = "Application running ..";
            notifyIconMain.ShowBalloonTip(500);
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (!radioButtonLot21.Checked && !radioButtonLot23.Checked)
            {
                MessageBox.Show(this, "Indicate the Lot", "Details missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (__mail != null)
            {
                // alert truncate body
                int index = textBoxMailBody.Text.IndexOf(mailBodyDelimiter);
                string subBody = textBoxMailBody.Text.Substring(0, index);

                DialogResult result = MessageBox.Show(this, "Body will be truncated : \n\n" + subBody, "Truncate body", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Cancel) return;

                // truncate mail body
                textBoxMailBody.Text = subBody;

                // check if folder exists                
                if (!Directory.Exists(Path.GetDirectoryName(textBoxFilenameMail.Text)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(textBoxFilenameMail.Text));
                }

                // save mail on server
                __mail.SaveAs(textBoxFilenameMail.Text);

                // touch file with sentOn date
                FileInfo mail = new FileInfo(textBoxFilenameMail.Text);
                mail.CreationTime = mail.LastWriteTime = __mail.SentOn;
                
                // update excel suivi
                WriteToExcel();

                MessageBox.Show(this, "Suivi has been updated", "Entry saved", MessageBoxButtons.OK, MessageBoxIcon.Information);

                ResetControls();
            }
            else
            {
                MessageBox.Show(this, "Mail is not valid", "Details missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetControls()
        {            
            textBoxFilenameExcel.Text = textBoxFilenameMail.Text = textBoxMailBody.Text = textBoxMailDate.Text
                = textBoxMailSubject.Text = textBoxRecipients.Text = textBoxSender.Text = string.Empty;

            comboBoxProject.SelectedIndex = -1;
            buttonAdd.Enabled = false;
        }

        private void WriteToExcel()
        {
            if (String.IsNullOrEmpty(textBoxFilenameExcel.Text)) return;

            FileInfo file = new FileInfo(textBoxFilenameExcel.Text);
            ExcelPackage package = new ExcelPackage(file);

            string worksheetName = string.Empty;
            string excelParamName = string.Empty;
            if (radioButtonLot21.Checked)
            {
                // Lot 2.1
                worksheetName = radioButtonLot21.Text;
                excelParamName = "URL_LOT21";
            }
            else if (radioButtonLot23.Checked)
            {
                // Lot 2.3
                worksheetName = radioButtonLot23.Text;
                excelParamName = "URL_LOT23";
            } else
                return;

            // get sheet from Excel file
            ExcelWorksheet sheet = package.Workbook.Worksheets[worksheetName];

            if (sheet == null)
            {
                MessageBox.Show(this, String.Format("Sheet {0} does not exist", worksheetName), "Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int nextRow = sheet.Dimension.End.Row + 1;

            // check semester >= july
            sheet.Cells[nextRow, (int)ExcelColumns.Semestre].Value =
                (__mail.SentOn.Month >= 7) ? String.Format("S2 {0}", __mail.SentOn.Year) : String.Format("S1 {0}", __mail.SentOn.Year);
            sheet.Cells[nextRow, (int)ExcelColumns.Mois].Value = __mail.SentOn.ToString("MMM-yyyy");
            sheet.Cells[nextRow, (int)ExcelColumns.Date_Demande].Value = textBoxMailDate.Text;
            sheet.Cells[nextRow, (int)ExcelColumns.Demandeur_Destinataires].Value = textBoxSender.Text;
            sheet.Cells[nextRow, (int)ExcelColumns.Demandeur_Destinataires].AddComment(String.Format("To : {0}", textBoxRecipients.Text), "abuchoo");
            sheet.Cells[nextRow, (int)ExcelColumns.Sujet].Value = textBoxMailSubject.Text;
            sheet.Cells[nextRow, (int)ExcelColumns.Demande].Value = textBoxMailBody.Text;

            // contruct dynamic hyperlink to mail object on server
            string mailHyperlink = String.Format("HYPERLINK(CONCATENATE({0},\"{1}\"),\"{2}\")", excelParamName, mailNameForExcel, Path.GetFileName(mailNameForExcel));
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Formula = mailHyperlink;
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Style.Font.UnderLine = true;
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

            sheet.Cells[nextRow, (int)ExcelColumns.Nom_KPI].Value = String.Format("[{0}] {1}", __mail.SentOn.ToString("dd/MM"), textBoxMailSubject.Text);
            sheet.Cells[nextRow, (int)ExcelColumns.Etat].Value = "En cours";
            
            package.Save();
        }

        private void comboBoxProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxProject.SelectedIndex == -1)
            {
                buttonAdd.Enabled = buttonOpenExcel.Enabled= false;
                return;
            }
            else if (!string.IsNullOrEmpty(textBoxFilenameMail.Text))
            {
                buttonAdd.Enabled = buttonOpenExcel.Enabled = true;
            } else
            {
                buttonOpenExcel.Enabled = true;
            }

            textBoxFilenameExcel.Tag = String.Format(pathLot2, "{0}", ConfigurationManager.AppSettings["File.Suivi"]);
            textBoxFilenameExcel.Text = String.Format(textBoxFilenameExcel.Tag.ToString(), comboBoxProject.Text);
            
            // mail msg
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), 
                comboBoxProject.Text, radioButtonLot21.Checked ? radioButtonLot21.Text : radioButtonLot23.Text);
            
            // lot 2.3 by default
            radioButtonLot23.Checked = true;
        }

        private void radioButtonLot21_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), comboBoxProject.Text, radioButtonLot21.Text);
        }

        private void radioButtonLot23_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), comboBoxProject.Text, radioButtonLot23.Text);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FormMain_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {                
                this.ShowInTaskbar = false;
            }
            else
            {
                this.ShowInTaskbar = true;
            }
        }

        private void notifyIconMain_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) return;

            if (FormWindowState.Minimized == this.WindowState)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void contextMenuStripMain_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            CustomContextMenu ccm = e.ClickedItem.Tag as CustomContextMenu;
            Launcher.LaunchControl(ccm);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.WindowState = FormWindowState.Minimized;
        }

        private void labelChar_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(mailBodyDelimiter);
            // display balloon
            notifyIconMain.BalloonTipText = "Copied to clipboard ..";
            notifyIconMain.ShowBalloonTip(500);
        }

        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxFilenameExcel.Text))
            {
                Process.Start(textBoxFilenameExcel.Text);
            }
        }

        //private void GetAttachmentsInfo(Microsoft.Office.Interop.Outlook.MailItem pMailItem)
        //{
        //    if (pMailItem.Attachments != null)
        //    {
        //        for (int i = 0; i < pMailItem.Attachments.Count; i++)
        //        {
        //            Microsoft.Office.Interop.Outlook.Attachment currentAttachment = pMailItem.Attachments[i + 1];
        //            if (currentAttachment != null)
        //            {
        //                string strFile = Path.Combine(@"c:\temp", FixFileName(currentAttachment.FileName));
        //                currentAttachment.SaveAsFile(strFile);
        //                Marshal.ReleaseComObject(currentAttachment);
        //            }
        //        }
        //    }
        //}    
    }
}
