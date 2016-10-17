using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AstekSuivi
{
    public partial class FormMain : Form
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _application = new Microsoft.Office.Interop.Outlook.Application();
        private readonly string pathLot2 = ConfigurationManager.AppSettings["Path.Lot2"];
        private Microsoft.Office.Interop.Outlook.MailItem __mail = null;

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
                        textBoxMailDate.Text = mailItem.SentOn.ToString();
                        textBoxSender.Text = mailItem.SenderEmailAddress;

                        SetProject(mailItem.Subject, mailItem.Body);

                        sbRecipients = new StringBuilder();
                        foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in mailItem.Recipients)
                            sbRecipients.Append(recipient.Address).Append("; ");
                        textBoxRecipients.Text = sbRecipients.ToString();

                        textBoxFilename.Tag = Path.Combine(pathLot2, DateTime.Today.Year.ToString(), DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".msg");

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
            textBoxMailBody.Text = body.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine);
        }

        private void GetAttachmentsInfo(Microsoft.Office.Interop.Outlook.MailItem pMailItem)
        {
            if (pMailItem.Attachments != null)
            {
                for (int i = 0; i < pMailItem.Attachments.Count; i++)
                {
                    Microsoft.Office.Interop.Outlook.Attachment currentAttachment = pMailItem.Attachments[i + 1];
                    if (currentAttachment != null)
                    {
                        string strFile = Path.Combine(@"c:\temp", FixFileName(currentAttachment.FileName));
                        currentAttachment.SaveAsFile(strFile);
                        Marshal.ReleaseComObject(currentAttachment);
                    }
                }
            }
        }

        private string FixFileName(string pFileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            if (pFileName.IndexOfAny(invalidChars) >= 0)
            {
                pFileName = invalidChars.Aggregate(pFileName, (current, invalidChar) => current.Replace(invalidChar, Convert.ToChar("_")));
            }
            return pFileName;
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            buttonAdd.Enabled = false;
            textBoxFilename.Tag = string.Empty;
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
                // check if folder exists                
                if (!Directory.Exists(Path.GetDirectoryName(textBoxFilename.Text)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(textBoxFilename.Text));
                }
                __mail.SaveAs(textBoxFilename.Text);
            }
            else
            {
                MessageBox.Show(this, "Mail is not valid", "Details missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBoxProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxFilename.Text = String.Format(
                textBoxFilename.Tag.ToString(), 
                comboBoxProject.Text, radioButtonLot21.Checked ? radioButtonLot21.Tag : radioButtonLot23.Tag);
            buttonAdd.Enabled = true;
            radioButtonLot21.Checked = true;
        }

        private void radioButtonLot21_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilename.Text = String.Format(textBoxFilename.Tag.ToString(), comboBoxProject.Text, radioButtonLot21.Tag);            
        }

        private void radioButtonLot23_CheckedChanged(object sender, EventArgs e)
        {
            textBoxFilename.Text = String.Format(textBoxFilename.Tag.ToString(), comboBoxProject.Text, radioButtonLot23.Tag);
        }
    }
}
