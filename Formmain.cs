﻿using OfficeOpenXml;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        private string delimiter = "◆";

        private enum ExcelColumns
        {
            Semestre = 1,
            Date_Demande = 2,
            Demandeur = 3,
            Destinataires = 3,
            Sujet = 4,
            Demande = 5,
            Mail = 6,
            Nom_KPI = 7,
            Etat = 8,
            Conso = 9, // Lot 2.3
            Vendue = 10 // Lot 2.3
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
            try
            {
                if (e.Data.GetDataPresent("FileGroupDescriptor"))
                {
                    StringBuilder sbRecipients = null;
                    // supports a drop of a Outlook message
                    foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in _application.ActiveExplorer().Selection)
                    {
                        __mail = mailItem;
                        DateTime dateToken = mailItem.SentOn;
                        textBoxMailDate.Text = dateToken.ToString();
                        textBoxSender.Text = mailItem.SenderEmailAddress;

                        SetProject(mailItem.Subject, mailItem.Body);

                        sbRecipients = new StringBuilder();
                        foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in mailItem.Recipients)
                            sbRecipients.Append(recipient.Address).Append("; ");
                        textBoxRecipients.Text = sbRecipients.ToString();

                        textBoxFilenameMail.Tag = Path.Combine(pathLot2, dateToken.Year.ToString(), dateToken.ToString("yyyyMMdd-HHmmss") + ".msg");
                        textBoxFilenameExcel.Tag = String.Format(pathLot2, "{0}", ConfigurationManager.AppSettings["File.Suivi"]);

                        // first item only is considered
                        break;
                    }
                } else
                {
                    MessageBox.Show(this, "This is not an Outlook item.\nPlease drag and drop an outlook item", "File error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }

        private void SetProject(string subject, string body)
        {
            comboBoxProject.SelectedIndex = -1;

            // check if ASPIN or SPID
            if (subject.ToLower().Contains("aspin") || body.ToLower().Contains("aspin"))
            {
                comboBoxProject.Text = "ASPIN";
            }

            if (subject.ToLower().Contains("spid") || body.ToLower().Contains("spid"))
            {
                comboBoxProject.Text = "SPID";
            }

            textBoxMailSubject.Text = subject;
            // remove white spaces / empty lines
            textBoxMailBody.Text = Regex.Replace(body, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
            textBoxMailBody.Text = textBoxMailBody.Text.Insert(250, delimiter);
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

        private void FormMain_Load(object sender, EventArgs e)
        {
            buttonAdd.Enabled = false;
            textBoxFilenameMail.Tag = textBoxFilenameExcel.Tag = string.Empty;
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
                int index = textBoxMailBody.Text.IndexOf(delimiter);
                string subBody = textBoxMailBody.Text.Substring(0, index);

                DialogResult result = MessageBox.Show(this, "Body will be truncated : \n\n" + subBody, "Truncate body", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

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

                // update excel suivi
                WriteToExcel();
            }
            else
            {
                MessageBox.Show(this, "Mail is not valid", "Details missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WriteToExcel()
        {
            if (String.IsNullOrEmpty(textBoxFilenameExcel.Text)) return;

            FileInfo file = new FileInfo(textBoxFilenameExcel.Text);
            ExcelPackage package = new ExcelPackage(file);
            ExcelWorksheet sheet = package.Workbook.Worksheets["Lot2.1"];
            
            int nextRow = sheet.Dimension.End.Row + 1;

            // check semester
            sheet.Cells[nextRow, (int)ExcelColumns.Semestre].Value = (__mail.SentOn.Month >= 7) ? String.Format("S2 {0}", __mail.SentOn.Year) : String.Format("S1 {0}", __mail.SentOn.Year);
            sheet.Cells[nextRow, (int)ExcelColumns.Date_Demande].Value = textBoxMailDate.Text;
            sheet.Cells[nextRow, (int)ExcelColumns.Demandeur].Value = textBoxSender.Text;
            sheet.Cells[nextRow, (int)ExcelColumns.Destinataires].AddComment(String.Format("To : {0}", textBoxRecipients.Text), "abuchoo");
            sheet.Cells[nextRow, (int)ExcelColumns.Sujet].Value = textBoxMailSubject.Text;            
            sheet.Cells[nextRow, (int)ExcelColumns.Demande].Value = textBoxMailBody.Text;

            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Hyperlink = new Uri(textBoxFilenameMail.Text);
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Value = Path.GetFileName(textBoxFilenameMail.Text);
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Style.Font.UnderLine = true;
            sheet.Cells[nextRow, (int)ExcelColumns.Mail].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

            //sheet.Cells[nextRow, (int)ExcelColumns.Mail].Formula = "HYPERLINK(\"" + textBoxFilenameMail.Text + "\",\"" + Path.GetFileName(textBoxFilenameMail.Text) + "\")";

            sheet.Cells[nextRow, (int)ExcelColumns.Nom_KPI].Value = String.Format("{0} : {1}", __mail.SentOn.ToString("dd/MM"), textBoxMailSubject.Text);
            sheet.Cells[nextRow, (int)ExcelColumns.Etat].Value = "En cours";

            package.Save();            
        }

        private void comboBoxProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), 
                comboBoxProject.Text, radioButtonLot21.Checked ? radioButtonLot21.Tag : radioButtonLot23.Tag);
            textBoxFilenameExcel.Text = String.Format(textBoxFilenameExcel.Tag.ToString(), comboBoxProject.Text);

            buttonAdd.Enabled = true;
            radioButtonLot21.Checked = true;
        }

        private void radioButtonLot21_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), comboBoxProject.Text, radioButtonLot21.Tag);
        }

        private void radioButtonLot23_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilenameMail.Text = String.Format(textBoxFilenameMail.Tag.ToString(), comboBoxProject.Text, radioButtonLot23.Tag);
        }        
    }
}
