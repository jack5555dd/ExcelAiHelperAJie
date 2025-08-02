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
            this.pnlModeSelector = new System.Windows.Forms.Panel();
            this.rbVstoMode = new System.Windows.Forms.RadioButton();
            this.rbVbaMode = new System.Windows.Forms.RadioButton();
            this.rbChatOnlyMode = new System.Windows.Forms.RadioButton();
            this.lblMode = new System.Windows.Forms.Label();
            this.pnlVbaCode = new System.Windows.Forms.Panel();
            this.rtbVbaCode = new System.Windows.Forms.RichTextBox();
            this.lblVbaCode = new System.Windows.Forms.Label();
            this.btnExecuteVba = new System.Windows.Forms.Button();
            this.btnCopyVbaCode = new System.Windows.Forms.Button();
            this.btnToggleVbaCode = new System.Windows.Forms.Button();
            this.pnlInput.SuspendLayout();
            this.pnlPreview.SuspendLayout();
            this.pnlModeSelector.SuspendLayout();
            this.pnlVbaCode.SuspendLayout();
            this.SuspendLayout();
            // 
            // rtbChatHistory
            // 
            this.rtbChatHistory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbChatHistory.BackColor = System.Drawing.SystemColors.Window;
            this.rtbChatHistory.Location = new System.Drawing.Point(0, 35);
            this.rtbChatHistory.Name = "rtbChatHistory";
            this.rtbChatHistory.ReadOnly = true;
            this.rtbChatHistory.Size = new System.Drawing.Size(350, 200);
            this.rtbChatHistory.TabIndex = 0;
            this.rtbChatHistory.Text = "欢迎使用 Excel AI 助手！请输入您的问题或指令。";
            // 
            // pnlModeSelector
            // 
            this.pnlModeSelector.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlModeSelector.Controls.Add(this.rbChatOnlyMode);
            this.pnlModeSelector.Controls.Add(this.rbVbaMode);
            this.pnlModeSelector.Controls.Add(this.rbVstoMode);
            this.pnlModeSelector.Controls.Add(this.lblMode);
            this.pnlModeSelector.Location = new System.Drawing.Point(0, 0);
            this.pnlModeSelector.Name = "pnlModeSelector";
            this.pnlModeSelector.Size = new System.Drawing.Size(350, 35);
            this.pnlModeSelector.TabIndex = 5;
            // 
            // lblMode
            // 
            this.lblMode.AutoSize = true;
            this.lblMode.Location = new System.Drawing.Point(3, 10);
            this.lblMode.Name = "lblMode";
            this.lblMode.Size = new System.Drawing.Size(43, 13);
            this.lblMode.TabIndex = 0;
            this.lblMode.Text = "模式：";
            // 
            // rbVstoMode
            // 
            this.rbVstoMode.AutoSize = true;
            this.rbVstoMode.Checked = true;
            this.rbVstoMode.Location = new System.Drawing.Point(52, 8);
            this.rbVstoMode.Name = "rbVstoMode";
            this.rbVstoMode.Size = new System.Drawing.Size(53, 17);
            this.rbVstoMode.TabIndex = 1;
            this.rbVstoMode.TabStop = true;
            this.rbVstoMode.Text = "VSTO";
            this.rbVstoMode.UseVisualStyleBackColor = true;
            this.rbVstoMode.CheckedChanged += new System.EventHandler(this.rbVstoMode_CheckedChanged);
            // 
            // rbVbaMode
            // 
            this.rbVbaMode.AutoSize = true;
            this.rbVbaMode.Location = new System.Drawing.Point(111, 8);
            this.rbVbaMode.Name = "rbVbaMode";
            this.rbVbaMode.Size = new System.Drawing.Size(45, 17);
            this.rbVbaMode.TabIndex = 2;
            this.rbVbaMode.Text = "VBA";
            this.rbVbaMode.UseVisualStyleBackColor = true;
            this.rbVbaMode.CheckedChanged += new System.EventHandler(this.rbVbaMode_CheckedChanged);
            // 
            // rbChatOnlyMode
            // 
            this.rbChatOnlyMode.AutoSize = true;
            this.rbChatOnlyMode.Location = new System.Drawing.Point(162, 8);
            this.rbChatOnlyMode.Name = "rbChatOnlyMode";
            this.rbChatOnlyMode.Size = new System.Drawing.Size(61, 17);
            this.rbChatOnlyMode.TabIndex = 3;
            this.rbChatOnlyMode.Text = "仅聊天";
            this.rbChatOnlyMode.UseVisualStyleBackColor = true;
            this.rbChatOnlyMode.CheckedChanged += new System.EventHandler(this.rbChatOnlyMode_CheckedChanged);

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
            // pnlVbaCode
            // 
            this.pnlVbaCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlVbaCode.Controls.Add(this.btnToggleVbaCode);
            this.pnlVbaCode.Controls.Add(this.btnCopyVbaCode);
            this.pnlVbaCode.Controls.Add(this.btnExecuteVba);
            this.pnlVbaCode.Controls.Add(this.lblVbaCode);
            this.pnlVbaCode.Controls.Add(this.rtbVbaCode);
            this.pnlVbaCode.Location = new System.Drawing.Point(0, 235);
            this.pnlVbaCode.Name = "pnlVbaCode";
            this.pnlVbaCode.Size = new System.Drawing.Size(350, 65);
            this.pnlVbaCode.TabIndex = 6;
            this.pnlVbaCode.Visible = false;
            // 
            // lblVbaCode
            // 
            this.lblVbaCode.AutoSize = true;
            this.lblVbaCode.Location = new System.Drawing.Point(3, 3);
            this.lblVbaCode.Name = "lblVbaCode";
            this.lblVbaCode.Size = new System.Drawing.Size(58, 13);
            this.lblVbaCode.TabIndex = 0;
            this.lblVbaCode.Text = "VBA代码:";
            // 
            // rtbVbaCode
            // 
            this.rtbVbaCode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbVbaCode.BackColor = System.Drawing.Color.LightCyan;
            this.rtbVbaCode.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtbVbaCode.Location = new System.Drawing.Point(3, 19);
            this.rtbVbaCode.Name = "rtbVbaCode";
            this.rtbVbaCode.ReadOnly = true;
            this.rtbVbaCode.Size = new System.Drawing.Size(344, 20);
            this.rtbVbaCode.TabIndex = 1;
            this.rtbVbaCode.Text = "";
            // 
            // btnExecuteVba
            // 
            this.btnExecuteVba.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExecuteVba.Location = new System.Drawing.Point(272, 42);
            this.btnExecuteVba.Name = "btnExecuteVba";
            this.btnExecuteVba.Size = new System.Drawing.Size(75, 20);
            this.btnExecuteVba.TabIndex = 2;
            this.btnExecuteVba.Text = "执行VBA";
            this.btnExecuteVba.UseVisualStyleBackColor = true;
            this.btnExecuteVba.Click += new System.EventHandler(this.btnExecuteVba_Click);
            // 
            // btnCopyVbaCode
            // 
            this.btnCopyVbaCode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyVbaCode.Location = new System.Drawing.Point(191, 42);
            this.btnCopyVbaCode.Name = "btnCopyVbaCode";
            this.btnCopyVbaCode.Size = new System.Drawing.Size(75, 20);
            this.btnCopyVbaCode.TabIndex = 3;
            this.btnCopyVbaCode.Text = "复制代码";
            this.btnCopyVbaCode.UseVisualStyleBackColor = true;
            this.btnCopyVbaCode.Click += new System.EventHandler(this.btnCopyVbaCode_Click);
            // 
            // btnToggleVbaCode
            // 
            this.btnToggleVbaCode.Location = new System.Drawing.Point(67, 1);
            this.btnToggleVbaCode.Name = "btnToggleVbaCode";
            this.btnToggleVbaCode.Size = new System.Drawing.Size(50, 17);
            this.btnToggleVbaCode.TabIndex = 4;
            this.btnToggleVbaCode.Text = "展开";
            this.btnToggleVbaCode.UseVisualStyleBackColor = true;
            this.btnToggleVbaCode.Click += new System.EventHandler(this.btnToggleVbaCode_Click);
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
            this.pnlPreview.Location = new System.Drawing.Point(0, 150);
            this.pnlPreview.Name = "pnlPreview";
            this.pnlPreview.Size = new System.Drawing.Size(350, 85);
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
            this.Controls.Add(this.pnlVbaCode);
            this.Controls.Add(this.pnlInput);
            this.Controls.Add(this.rtbChatHistory);
            this.Controls.Add(this.pnlModeSelector);
            this.Name = "ChatPaneControl";
            this.Size = new System.Drawing.Size(350, 350);
            this.pnlInput.ResumeLayout(false);
            this.pnlInput.PerformLayout();
            this.pnlPreview.ResumeLayout(false);
            this.pnlPreview.PerformLayout();
            this.pnlModeSelector.ResumeLayout(false);
            this.pnlModeSelector.PerformLayout();
            this.pnlVbaCode.ResumeLayout(false);
            this.pnlVbaCode.PerformLayout();
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
        private System.Windows.Forms.Panel pnlModeSelector;
        private System.Windows.Forms.RadioButton rbVstoMode;
        private System.Windows.Forms.RadioButton rbVbaMode;
        private System.Windows.Forms.RadioButton rbChatOnlyMode;
        private System.Windows.Forms.Label lblMode;
        private System.Windows.Forms.Panel pnlVbaCode;
        private System.Windows.Forms.RichTextBox rtbVbaCode;
        private System.Windows.Forms.Label lblVbaCode;
        private System.Windows.Forms.Button btnExecuteVba;
        private System.Windows.Forms.Button btnCopyVbaCode;
        private System.Windows.Forms.Button btnToggleVbaCode;
    }
}