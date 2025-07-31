namespace ExcelAIHelper
{
    partial class ChatPaneControl
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.rtbChatHistory = new System.Windows.Forms.RichTextBox();
            this.txtUserInput = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.pnlInput = new System.Windows.Forms.Panel();
            this.pnlPreview = new System.Windows.Forms.Panel();
            this.rtbPreview = new System.Windows.Forms.RichTextBox();
            this.lblPreview = new System.Windows.Forms.Label();
            this.btnConfirm = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.pnlInput.SuspendLayout();
            this.pnlPreview.SuspendLayout();
            this.SuspendLayout();
            // 
            // rtbChatHistory
            // 
            this.rtbChatHistory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbChatHistory.BackColor = System.Drawing.SystemColors.Window;
            this.rtbChatHistory.Location = new System.Drawing.Point(0, 0);
            this.rtbChatHistory.Name = "rtbChatHistory";
            this.rtbChatHistory.ReadOnly = true;
            this.rtbChatHistory.Size = new System.Drawing.Size(350, 300);
            this.rtbChatHistory.TabIndex = 0;
            this.rtbChatHistory.Text = "欢迎使用 Excel AI 助手！请输入您的问题或指令。";
            // 
            // txtUserInput
            // 
            this.txtUserInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtUserInput.Location = new System.Drawing.Point(3, 3);
            this.txtUserInput.Multiline = true;
            this.txtUserInput.Name = "txtUserInput";
            this.txtUserInput.Size = new System.Drawing.Size(263, 44);
            this.txtUserInput.TabIndex = 1;
            this.txtUserInput.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUserInput_KeyDown);
            // 
            // btnSend
            // 
            this.btnSend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSend.Location = new System.Drawing.Point(272, 3);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 44);
            this.btnSend.TabIndex = 2;
            this.btnSend.Text = "发送";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // pnlInput
            // 
            this.pnlInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlInput.Controls.Add(this.txtUserInput);
            this.pnlInput.Controls.Add(this.btnSend);
            this.pnlInput.Location = new System.Drawing.Point(0, 300);
            this.pnlInput.Name = "pnlInput";
            this.pnlInput.Size = new System.Drawing.Size(350, 50);
            this.pnlInput.TabIndex = 3;
            // 
            // pnlPreview
            // 
            this.pnlPreview.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlPreview.Controls.Add(this.btnCancel);
            this.pnlPreview.Controls.Add(this.btnConfirm);
            this.pnlPreview.Controls.Add(this.lblPreview);
            this.pnlPreview.Controls.Add(this.rtbPreview);
            this.pnlPreview.Location = new System.Drawing.Point(0, 180);
            this.pnlPreview.Name = "pnlPreview";
            this.pnlPreview.Size = new System.Drawing.Size(350, 120);
            this.pnlPreview.TabIndex = 4;
            this.pnlPreview.Visible = false;
            // 
            // rtbPreview
            // 
            this.rtbPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbPreview.BackColor = System.Drawing.Color.LightYellow;
            this.rtbPreview.Location = new System.Drawing.Point(3, 23);
            this.rtbPreview.Name = "rtbPreview";
            this.rtbPreview.ReadOnly = true;
            this.rtbPreview.Size = new System.Drawing.Size(344, 65);
            this.rtbPreview.TabIndex = 0;
            this.rtbPreview.Text = "";
            // 
            // lblPreview
            // 
            this.lblPreview.AutoSize = true;
            this.lblPreview.Location = new System.Drawing.Point(3, 7);
            this.lblPreview.Name = "lblPreview";
            this.lblPreview.Size = new System.Drawing.Size(94, 13);
            this.lblPreview.TabIndex = 1;
            this.lblPreview.Text = "操作预览（确认执行）";
            // 
            // btnConfirm
            // 
            this.btnConfirm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnConfirm.Location = new System.Drawing.Point(272, 94);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(75, 23);
            this.btnConfirm.TabIndex = 2;
            this.btnConfirm.Text = "确认";
            this.btnConfirm.UseVisualStyleBackColor = true;
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(191, 94);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ChatPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlPreview);
            this.Controls.Add(this.pnlInput);
            this.Controls.Add(this.rtbChatHistory);
            this.Name = "ChatPaneControl";
            this.Size = new System.Drawing.Size(350, 350);
            this.pnlInput.ResumeLayout(false);
            this.pnlInput.PerformLayout();
            this.pnlPreview.ResumeLayout(false);
            this.pnlPreview.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbChatHistory;
        private System.Windows.Forms.TextBox txtUserInput;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Panel pnlInput;
        private System.Windows.Forms.Panel pnlPreview;
        private System.Windows.Forms.RichTextBox rtbPreview;
        private System.Windows.Forms.Label lblPreview;
        private System.Windows.Forms.Button btnConfirm;
        private System.Windows.Forms.Button btnCancel;
    }
}