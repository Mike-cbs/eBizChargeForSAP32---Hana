namespace Setup
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.textPin = new System.Windows.Forms.TextBox();
            this.textSourceKey = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textLicense = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textDatabase = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textDBUser = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textDBPassword = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cbDBType = new System.Windows.Forms.ComboBox();
            this.cbURL = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtSiteUserPassword = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtSiteUser = new System.Windows.Forms.TextBox();
            this.txtCompanyDBName = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textInstalledFolder = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(109, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pin:";
            this.label1.Visible = false;
            // 
            // textPin
            // 
            this.textPin.Location = new System.Drawing.Point(154, 25);
            this.textPin.Name = "textPin";
            this.textPin.Size = new System.Drawing.Size(332, 20);
            this.textPin.TabIndex = 1;
            this.textPin.Visible = false;
            // 
            // textSourceKey
            // 
            this.textSourceKey.Location = new System.Drawing.Point(154, 55);
            this.textSourceKey.Name = "textSourceKey";
            this.textSourceKey.Size = new System.Drawing.Size(332, 20);
            this.textSourceKey.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(69, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Security Key:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(57, 91);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Gateway URL:";
            this.label3.Visible = false;
            // 
            // textLicense
            // 
            this.textLicense.Location = new System.Drawing.Point(154, 146);
            this.textLicense.Name = "textLicense";
            this.textLicense.Size = new System.Drawing.Size(332, 20);
            this.textLicense.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(41, 123);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Database Server:";
            // 
            // textDatabase
            // 
            this.textDatabase.Location = new System.Drawing.Point(154, 116);
            this.textDatabase.Name = "textDatabase";
            this.textDatabase.Size = new System.Drawing.Size(332, 20);
            this.textDatabase.TabIndex = 4;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(50, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(81, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "License Server:";
            // 
            // textDBUser
            // 
            this.textDBUser.Location = new System.Drawing.Point(154, 176);
            this.textDBUser.Name = "textDBUser";
            this.textDBUser.Size = new System.Drawing.Size(332, 20);
            this.textDBUser.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 187);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 10;
            // 
            // textDBPassword
            // 
            this.textDBPassword.Location = new System.Drawing.Point(154, 206);
            this.textDBPassword.Name = "textDBPassword";
            this.textDBPassword.PasswordChar = '*';
            this.textDBPassword.Size = new System.Drawing.Size(332, 20);
            this.textDBPassword.TabIndex = 7;
            this.textDBPassword.UseSystemPasswordChar = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(60, 212);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(74, 13);
            this.label7.TabIndex = 12;
            this.label7.Text = "DB Password:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(458, 410);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Setup";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(84, 183);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(47, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "DB User";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(48, 238);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(83, 13);
            this.label11.TabIndex = 20;
            this.label11.Text = "Database Type:";
            this.label11.Visible = false;
            // 
            // cbDBType
            // 
            this.cbDBType.FormattingEnabled = true;
            this.cbDBType.Items.AddRange(new object[] {
            "Hana"});
            this.cbDBType.Location = new System.Drawing.Point(154, 236);
            this.cbDBType.Name = "cbDBType";
            this.cbDBType.Size = new System.Drawing.Size(121, 21);
            this.cbDBType.TabIndex = 8;
            this.cbDBType.Visible = false;
            // 
            // cbURL
            // 
            this.cbURL.AutoCompleteCustomSource.AddRange(new string[] {
            " SandBox: https://sandbox.ebizcharge.com/soap/gate/CCEBDC0A ",
            " Production:https://secure.ebizcharge.com/soap/gate/CCEBDC0A"});
            this.cbURL.FormattingEnabled = true;
            this.cbURL.Items.AddRange(new object[] {
            "SandBox: https://sandbox.ebizcharge.com/soap/gate/CCEBDC0A ",
            "Production:https://secure.ebizcharge.com/soap/gate/CCEBDC0A"});
            this.cbURL.Location = new System.Drawing.Point(154, 85);
            this.cbURL.Name = "cbURL";
            this.cbURL.Size = new System.Drawing.Size(332, 21);
            this.cbURL.TabIndex = 3;
            this.cbURL.Visible = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(60, 276);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(48, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "B1 User:";
            // 
            // txtSiteUserPassword
            // 
            this.txtSiteUserPassword.Location = new System.Drawing.Point(156, 299);
            this.txtSiteUserPassword.Name = "txtSiteUserPassword";
            this.txtSiteUserPassword.PasswordChar = '*';
            this.txtSiteUserPassword.Size = new System.Drawing.Size(332, 20);
            this.txtSiteUserPassword.TabIndex = 22;
            this.txtSiteUserPassword.UseSystemPasswordChar = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(20, 306);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(93, 13);
            this.label9.TabIndex = 23;
            this.label9.Text = "Bi User Password:";
            // 
            // txtSiteUser
            // 
            this.txtSiteUser.Location = new System.Drawing.Point(156, 269);
            this.txtSiteUser.Name = "txtSiteUser";
            this.txtSiteUser.Size = new System.Drawing.Size(332, 20);
            this.txtSiteUser.TabIndex = 21;
            // 
            // txtCompanyDBName
            // 
            this.txtCompanyDBName.Location = new System.Drawing.Point(156, 327);
            this.txtCompanyDBName.Name = "txtCompanyDBName";
            this.txtCompanyDBName.Size = new System.Drawing.Size(332, 20);
            this.txtCompanyDBName.TabIndex = 25;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(2, 334);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(134, 13);
            this.label12.TabIndex = 26;
            this.label12.Text = "Company Database Name:";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Setup.Resource1.Image1;
            this.pictureBox1.Location = new System.Drawing.Point(154, 81);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(243, 230);
            this.pictureBox1.TabIndex = 27;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // textInstalledFolder
            // 
            this.textInstalledFolder.Location = new System.Drawing.Point(156, 358);
            this.textInstalledFolder.Name = "textInstalledFolder";
            this.textInstalledFolder.Size = new System.Drawing.Size(332, 20);
            this.textInstalledFolder.TabIndex = 34;
            this.textInstalledFolder.Visible = false;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(54, 365);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(76, 13);
            this.label13.TabIndex = 35;
            this.label13.Text = "Add-on Folder:";
            this.label13.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(571, 454);
            this.Controls.Add(this.textInstalledFolder);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.txtCompanyDBName);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtSiteUserPassword);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtSiteUser);
            this.Controls.Add(this.cbURL);
            this.Controls.Add(this.cbDBType);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textDBPassword);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textDBUser);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textDatabase);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textLicense);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textSourceKey);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textPin);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "eBizCharge For SAP Setup";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textPin;
        private System.Windows.Forms.TextBox textSourceKey;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textLicense;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textDatabase;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textDBUser;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textDBPassword;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbDBType;
        private System.Windows.Forms.ComboBox cbURL;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtSiteUserPassword;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtSiteUser;
        private System.Windows.Forms.TextBox txtCompanyDBName;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textInstalledFolder;
        private System.Windows.Forms.Label label13;
    }
}