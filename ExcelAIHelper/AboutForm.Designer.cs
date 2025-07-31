namespace ExcelAIHelper
{
    partial class AboutForm
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
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lblWebsite = new System.Windows.Forms.Label();
            this.linkWebsite = new System.Windows.Forms.LinkLabel();
            this.lblContact = new System.Windows.Forms.Label();
            this.lblContactInfo = new System.Windows.Forms.Label();
            this.picQRCode = new System.Windows.Forms.PictureBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblDescription = new System.Windows.Forms.Label();
            this.panelMain = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.picQRCode)).BeginInit();
            this.panelMain.SuspendLayout();
            this.SuspendLayout();
            
            // 
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.Controls.Add(this.lblTitle);
            this.panelMain.Controls.Add(this.lblVersion);
            this.panelMain.Controls.Add(this.lblDescription);
            this.panelMain.Controls.Add(this.lblWebsite);
            this.panelMain.Controls.Add(this.linkWebsite);
            this.panelMain.Controls.Add(this.lblContact);
            this.panelMain.Controls.Add(this.lblContactInfo);
            this.panelMain.Controls.Add(this.picQRCode);
            this.panelMain.Location = new System.Drawing.Point(12, 12);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(460, 320);
            this.panelMain.TabIndex = 0;
            
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("微软雅黑", 16F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.lblTitle.Location = new System.Drawing.Point(20, 20);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(180, 30);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Excel AI 助手";
            
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.lblVersion.ForeColor = System.Drawing.Color.Gray;
            this.lblVersion.Location = new System.Drawing.Point(25, 60);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(79, 20);
            this.lblVersion.TabIndex = 1;
            this.lblVersion.Text = "版本: v1.0";
            
            // 
            // lblDescription
            // 
            this.lblDescription.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDescription.ForeColor = System.Drawing.Color.DimGray;
            this.lblDescription.Location = new System.Drawing.Point(25, 90);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new System.Drawing.Size(280, 40);
            this.lblDescription.TabIndex = 2;
            this.lblDescription.Text = "一款智能的Excel插件，集成AI助手功能，\r\n帮助您更高效地处理数据和公式。";
            
            // 
            // lblWebsite
            // 
            this.lblWebsite.AutoSize = true;
            this.lblWebsite.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblWebsite.Location = new System.Drawing.Point(25, 150);
            this.lblWebsite.Name = "lblWebsite";
            this.lblWebsite.Size = new System.Drawing.Size(56, 17);
            this.lblWebsite.TabIndex = 3;
            this.lblWebsite.Text = "官网:";
            
            // 
            // linkWebsite
            // 
            this.linkWebsite.AutoSize = true;
            this.linkWebsite.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.linkWebsite.Location = new System.Drawing.Point(85, 150);
            this.linkWebsite.Name = "linkWebsite";
            this.linkWebsite.Size = new System.Drawing.Size(180, 17);
            this.linkWebsite.TabIndex = 4;
            this.linkWebsite.TabStop = true;
            this.linkWebsite.Text = "https://www.aiexcel.cn";
            this.linkWebsite.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkWebsite_LinkClicked);
            
            // 
            // lblContact
            // 
            this.lblContact.AutoSize = true;
            this.lblContact.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblContact.Location = new System.Drawing.Point(25, 180);
            this.lblContact.Name = "lblContact";
            this.lblContact.Size = new System.Drawing.Size(68, 17);
            this.lblContact.TabIndex = 5;
            this.lblContact.Text = "联系方式:";
            
            // 
            // lblContactInfo
            // 
            this.lblContactInfo.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblContactInfo.Location = new System.Drawing.Point(25, 205);
            this.lblContactInfo.Name = "lblContactInfo";
            this.lblContactInfo.Size = new System.Drawing.Size(280, 60);
            this.lblContactInfo.TabIndex = 6;
            this.lblContactInfo.Text = "邮箱: support@aiexcel.cn\r\n电话: 400-123-4567\r\n技术支持: tech@aiexcel.cn";
            
            // 
            // picQRCode
            // 
            this.picQRCode.BackColor = System.Drawing.Color.White;
            this.picQRCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picQRCode.Location = new System.Drawing.Point(330, 150);
            this.picQRCode.Name = "picQRCode";
            this.picQRCode.Size = new System.Drawing.Size(100, 100);
            this.picQRCode.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picQRCode.TabIndex = 7;
            this.picQRCode.TabStop = false;
            
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(397, 350);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            
            // 
            // AboutForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(484, 392);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.panelMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "关于 Excel AI 助手";
            this.Load += new System.EventHandler(this.AboutForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picQRCode)).EndInit();
            this.panelMain.ResumeLayout(false);
            this.panelMain.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.Label lblDescription;
        private System.Windows.Forms.Label lblWebsite;
        private System.Windows.Forms.LinkLabel linkWebsite;
        private System.Windows.Forms.Label lblContact;
        private System.Windows.Forms.Label lblContactInfo;
        private System.Windows.Forms.PictureBox picQRCode;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Panel panelMain;
    }
}