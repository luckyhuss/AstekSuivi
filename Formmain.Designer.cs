namespace AstekSuivi
{
    partial class FormMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonAdd = new System.Windows.Forms.Button();
            this.textBoxMailBody = new System.Windows.Forms.TextBox();
            this.textBoxMailSubject = new System.Windows.Forms.TextBox();
            this.textBoxMailDate = new System.Windows.Forms.TextBox();
            this.textBoxSender = new System.Windows.Forms.TextBox();
            this.textBoxRecipients = new System.Windows.Forms.TextBox();
            this.textBoxFilenameMail = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxProject = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.radioButtonLot21 = new System.Windows.Forms.RadioButton();
            this.radioButtonLot23 = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.textBoxFilenameExcel = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonAdd
            // 
            this.buttonAdd.Location = new System.Drawing.Point(331, 425);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(140, 23);
            this.buttonAdd.TabIndex = 0;
            this.buttonAdd.Text = "Add";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // textBoxMailBody
            // 
            this.textBoxMailBody.Location = new System.Drawing.Point(12, 90);
            this.textBoxMailBody.Multiline = true;
            this.textBoxMailBody.Name = "textBoxMailBody";
            this.textBoxMailBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxMailBody.Size = new System.Drawing.Size(459, 198);
            this.textBoxMailBody.TabIndex = 1;
            // 
            // textBoxMailSubject
            // 
            this.textBoxMailSubject.Location = new System.Drawing.Point(54, 64);
            this.textBoxMailSubject.Multiline = true;
            this.textBoxMailSubject.Name = "textBoxMailSubject";
            this.textBoxMailSubject.ReadOnly = true;
            this.textBoxMailSubject.Size = new System.Drawing.Size(417, 20);
            this.textBoxMailSubject.TabIndex = 2;
            // 
            // textBoxMailDate
            // 
            this.textBoxMailDate.Location = new System.Drawing.Point(54, 12);
            this.textBoxMailDate.Name = "textBoxMailDate";
            this.textBoxMailDate.ReadOnly = true;
            this.textBoxMailDate.Size = new System.Drawing.Size(120, 20);
            this.textBoxMailDate.TabIndex = 3;
            // 
            // textBoxSender
            // 
            this.textBoxSender.Location = new System.Drawing.Point(222, 12);
            this.textBoxSender.Name = "textBoxSender";
            this.textBoxSender.ReadOnly = true;
            this.textBoxSender.Size = new System.Drawing.Size(249, 20);
            this.textBoxSender.TabIndex = 4;
            // 
            // textBoxRecipients
            // 
            this.textBoxRecipients.Location = new System.Drawing.Point(54, 38);
            this.textBoxRecipients.Multiline = true;
            this.textBoxRecipients.Name = "textBoxRecipients";
            this.textBoxRecipients.ReadOnly = true;
            this.textBoxRecipients.Size = new System.Drawing.Size(417, 20);
            this.textBoxRecipients.TabIndex = 5;
            // 
            // textBoxFilenameMail
            // 
            this.textBoxFilenameMail.Location = new System.Drawing.Point(61, 321);
            this.textBoxFilenameMail.Name = "textBoxFilenameMail";
            this.textBoxFilenameMail.ReadOnly = true;
            this.textBoxFilenameMail.Size = new System.Drawing.Size(401, 20);
            this.textBoxFilenameMail.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Date :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(180, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "From :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "To :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Subject";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 324);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Mail : ";
            // 
            // comboBoxProject
            // 
            this.comboBoxProject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxProject.FormattingEnabled = true;
            this.comboBoxProject.Items.AddRange(new object[] {
            "ASPIN",
            "SPID"});
            this.comboBoxProject.Location = new System.Drawing.Point(61, 294);
            this.comboBoxProject.Name = "comboBoxProject";
            this.comboBoxProject.Size = new System.Drawing.Size(121, 21);
            this.comboBoxProject.TabIndex = 12;
            this.comboBoxProject.SelectedIndexChanged += new System.EventHandler(this.comboBoxProject_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 297);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "Project :";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(188, 297);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(31, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Lot : ";
            // 
            // radioButtonLot21
            // 
            this.radioButtonLot21.AutoSize = true;
            this.radioButtonLot21.Location = new System.Drawing.Point(225, 295);
            this.radioButtonLot21.Name = "radioButtonLot21";
            this.radioButtonLot21.Size = new System.Drawing.Size(40, 17);
            this.radioButtonLot21.TabIndex = 15;
            this.radioButtonLot21.Tag = "OCEANE";
            this.radioButtonLot21.Text = "2.1";
            this.radioButtonLot21.UseVisualStyleBackColor = true;
            this.radioButtonLot21.CheckedChanged += new System.EventHandler(this.radioButtonLot21_CheckedChanged);
            // 
            // radioButtonLot23
            // 
            this.radioButtonLot23.AutoSize = true;
            this.radioButtonLot23.Location = new System.Drawing.Point(271, 295);
            this.radioButtonLot23.Name = "radioButtonLot23";
            this.radioButtonLot23.Size = new System.Drawing.Size(40, 17);
            this.radioButtonLot23.TabIndex = 16;
            this.radioButtonLot23.Tag = "Lot 2.3";
            this.radioButtonLot23.Text = "2.3";
            this.radioButtonLot23.UseVisualStyleBackColor = true;
            this.radioButtonLot23.CheckedChanged += new System.EventHandler(this.radioButtonLot23_CheckedChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(9, 350);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(39, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Excel :";
            // 
            // textBoxFilenameExcel
            // 
            this.textBoxFilenameExcel.Location = new System.Drawing.Point(61, 347);
            this.textBoxFilenameExcel.Name = "textBoxFilenameExcel";
            this.textBoxFilenameExcel.ReadOnly = true;
            this.textBoxFilenameExcel.Size = new System.Drawing.Size(401, 20);
            this.textBoxFilenameExcel.TabIndex = 17;
            // 
            // FormMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 493);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBoxFilenameExcel);
            this.Controls.Add(this.radioButtonLot23);
            this.Controls.Add(this.radioButtonLot21);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBoxProject);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxFilenameMail);
            this.Controls.Add(this.textBoxRecipients);
            this.Controls.Add(this.textBoxSender);
            this.Controls.Add(this.textBoxMailDate);
            this.Controls.Add(this.textBoxMailSubject);
            this.Controls.Add(this.textBoxMailBody);
            this.Controls.Add(this.buttonAdd);
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Astek Suivi Mensuel";
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.FormMain_DragDrop);
            this.DragOver += new System.Windows.Forms.DragEventHandler(this.FormMain_DragOver);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.TextBox textBoxMailBody;
        private System.Windows.Forms.TextBox textBoxMailSubject;
        private System.Windows.Forms.TextBox textBoxMailDate;
        private System.Windows.Forms.TextBox textBoxSender;
        private System.Windows.Forms.TextBox textBoxRecipients;
        private System.Windows.Forms.TextBox textBoxFilenameMail;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxProject;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.RadioButton radioButtonLot21;
        private System.Windows.Forms.RadioButton radioButtonLot23;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBoxFilenameExcel;
    }
}

