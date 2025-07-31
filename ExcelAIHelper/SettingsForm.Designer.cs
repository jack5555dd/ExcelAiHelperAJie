namespace ExcelAIHelper
{
    partial class SettingsForm
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageGeneral = new System.Windows.Forms.TabPage();
            this.tabPageApi = new System.Windows.Forms.TabPage();
            this.tabPageAdvanced = new System.Windows.Forms.TabPage();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            
            // API设置控件
            this.lblProvider = new System.Windows.Forms.Label();
            this.cmbProvider = new System.Windows.Forms.ComboBox();
            this.lblModel = new System.Windows.Forms.Label();
            this.txtModel = new System.Windows.Forms.TextBox();
            this.lblApiKey = new System.Windows.Forms.Label();
            this.txtApiKey = new System.Windows.Forms.TextBox();
            this.lblBaseUrl = new System.Windows.Forms.Label();
            this.txtBaseUrl = new System.Windows.Forms.TextBox();
            this.lblTimeout = new System.Windows.Forms.Label();
            this.numTimeout = new System.Windows.Forms.NumericUpDown();
            this.btnTestConnection = new System.Windows.Forms.Button();
            
            // 通用设置控件
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cmbLanguage = new System.Windows.Forms.ComboBox();
            this.chkAutoSave = new System.Windows.Forms.CheckBox();
            this.chkShowTips = new System.Windows.Forms.CheckBox();
            
            this.tabControl1.SuspendLayout();
            this.tabPageApi.SuspendLayout();
            this.tabPageGeneral.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeout)).BeginInit();
            this.SuspendLayout();
            
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageGeneral);
            this.tabControl1.Controls.Add(this.tabPageApi);
            this.tabControl1.Controls.Add(this.tabPageAdvanced);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(560, 380);
            this.tabControl1.TabIndex = 0;
            
            // 
            // tabPageGeneral
            // 
            this.tabPageGeneral.Controls.Add(this.lblLanguage);
            this.tabPageGeneral.Controls.Add(this.cmbLanguage);
            this.tabPageGeneral.Controls.Add(this.chkAutoSave);
            this.tabPageGeneral.Controls.Add(this.chkShowTips);
            this.tabPageGeneral.Location = new System.Drawing.Point(4, 25);
            this.tabPageGeneral.Name = "tabPageGeneral";
            this.tabPageGeneral.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGeneral.Size = new System.Drawing.Size(552, 351);
            this.tabPageGeneral.TabIndex = 0;
            this.tabPageGeneral.Text = "通用";
            this.tabPageGeneral.UseVisualStyleBackColor = true;
            
            // 
            // tabPageApi
            // 
            this.tabPageApi.Controls.Add(this.lblProvider);
            this.tabPageApi.Controls.Add(this.cmbProvider);
            this.tabPageApi.Controls.Add(this.lblModel);
            this.tabPageApi.Controls.Add(this.txtModel);
            this.tabPageApi.Controls.Add(this.lblApiKey);
            this.tabPageApi.Controls.Add(this.txtApiKey);
            this.tabPageApi.Controls.Add(this.lblBaseUrl);
            this.tabPageApi.Controls.Add(this.txtBaseUrl);
            this.tabPageApi.Controls.Add(this.lblTimeout);
            this.tabPageApi.Controls.Add(this.numTimeout);
            this.tabPageApi.Controls.Add(this.btnTestConnection);
            this.tabPageApi.Location = new System.Drawing.Point(4, 25);
            this.tabPageApi.Name = "tabPageApi";
            this.tabPageApi.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageApi.Size = new System.Drawing.Size(552, 351);
            this.tabPageApi.TabIndex = 1;
            this.tabPageApi.Text = "API设置";
            this.tabPageApi.UseVisualStyleBackColor = true;
            
            // 
            // tabPageAdvanced
            // 
            this.tabPageAdvanced.Location = new System.Drawing.Point(4, 25);
            this.tabPageAdvanced.Name = "tabPageAdvanced";
            this.tabPageAdvanced.Size = new System.Drawing.Size(552, 351);
            this.tabPageAdvanced.TabIndex = 2;
            this.tabPageAdvanced.Text = "高级";
            this.tabPageAdvanced.UseVisualStyleBackColor = true;
            
            // API设置控件布局
            // 
            // lblProvider
            // 
            this.lblProvider.AutoSize = true;
            this.lblProvider.Location = new System.Drawing.Point(20, 25);
            this.lblProvider.Name = "lblProvider";
            this.lblProvider.Size = new System.Drawing.Size(67, 15);
            this.lblProvider.TabIndex = 0;
            this.lblProvider.Text = "服务商:";
            
            // 
            // cmbProvider
            // 
            this.cmbProvider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbProvider.FormattingEnabled = true;
            this.cmbProvider.Items.AddRange(new object[] { "DeepSeek", "OpenAI", "Azure OpenAI", "其他" });
            this.cmbProvider.Location = new System.Drawing.Point(120, 22);
            this.cmbProvider.Name = "cmbProvider";
            this.cmbProvider.Size = new System.Drawing.Size(200, 23);
            this.cmbProvider.TabIndex = 1;
            
            // 
            // lblModel
            // 
            this.lblModel.AutoSize = true;
            this.lblModel.Location = new System.Drawing.Point(20, 65);
            this.lblModel.Name = "lblModel";
            this.lblModel.Size = new System.Drawing.Size(52, 15);
            this.lblModel.TabIndex = 2;
            this.lblModel.Text = "模型:";
            
            // 
            // txtModel
            // 
            this.txtModel.Location = new System.Drawing.Point(120, 62);
            this.txtModel.Name = "txtModel";
            this.txtModel.Size = new System.Drawing.Size(200, 25);
            this.txtModel.TabIndex = 3;
            
            // 
            // lblApiKey
            // 
            this.lblApiKey.AutoSize = true;
            this.lblApiKey.Location = new System.Drawing.Point(20, 105);
            this.lblApiKey.Name = "lblApiKey";
            this.lblApiKey.Size = new System.Drawing.Size(67, 15);
            this.lblApiKey.TabIndex = 4;
            this.lblApiKey.Text = "API密钥:";
            
            // 
            // txtApiKey
            // 
            this.txtApiKey.Location = new System.Drawing.Point(120, 102);
            this.txtApiKey.Name = "txtApiKey";
            this.txtApiKey.PasswordChar = '*';
            this.txtApiKey.Size = new System.Drawing.Size(400, 25);
            this.txtApiKey.TabIndex = 5;
            
            // 
            // lblBaseUrl
            // 
            this.lblBaseUrl.AutoSize = true;
            this.lblBaseUrl.Location = new System.Drawing.Point(20, 145);
            this.lblBaseUrl.Name = "lblBaseUrl";
            this.lblBaseUrl.Size = new System.Drawing.Size(82, 15);
            this.lblBaseUrl.TabIndex = 6;
            this.lblBaseUrl.Text = "API地址:";
            
            // 
            // txtBaseUrl
            // 
            this.txtBaseUrl.Location = new System.Drawing.Point(120, 142);
            this.txtBaseUrl.Name = "txtBaseUrl";
            this.txtBaseUrl.Size = new System.Drawing.Size(400, 25);
            this.txtBaseUrl.TabIndex = 7;
            
            // 
            // lblTimeout
            // 
            this.lblTimeout.AutoSize = true;
            this.lblTimeout.Location = new System.Drawing.Point(20, 185);
            this.lblTimeout.Name = "lblTimeout";
            this.lblTimeout.Size = new System.Drawing.Size(97, 15);
            this.lblTimeout.TabIndex = 8;
            this.lblTimeout.Text = "超时时间(秒):";
            
            // 
            // numTimeout
            // 
            this.numTimeout.Location = new System.Drawing.Point(120, 182);
            this.numTimeout.Maximum = new decimal(new int[] { 300, 0, 0, 0 });
            this.numTimeout.Minimum = new decimal(new int[] { 5, 0, 0, 0 });
            this.numTimeout.Name = "numTimeout";
            this.numTimeout.Size = new System.Drawing.Size(100, 25);
            this.numTimeout.TabIndex = 9;
            this.numTimeout.Value = new decimal(new int[] { 15, 0, 0, 0 });
            
            // 
            // btnTestConnection
            // 
            this.btnTestConnection.Location = new System.Drawing.Point(120, 220);
            this.btnTestConnection.Name = "btnTestConnection";
            this.btnTestConnection.Size = new System.Drawing.Size(120, 30);
            this.btnTestConnection.TabIndex = 10;
            this.btnTestConnection.Text = "测试连接";
            this.btnTestConnection.UseVisualStyleBackColor = true;
            this.btnTestConnection.Click += new System.EventHandler(this.btnTestConnection_Click);
            
            // 通用设置控件布局
            // 
            // lblLanguage
            // 
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(20, 25);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(52, 15);
            this.lblLanguage.TabIndex = 0;
            this.lblLanguage.Text = "语言:";
            
            // 
            // cmbLanguage
            // 
            this.cmbLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLanguage.FormattingEnabled = true;
            this.cmbLanguage.Items.AddRange(new object[] { "简体中文", "English" });
            this.cmbLanguage.Location = new System.Drawing.Point(120, 22);
            this.cmbLanguage.Name = "cmbLanguage";
            this.cmbLanguage.Size = new System.Drawing.Size(150, 23);
            this.cmbLanguage.TabIndex = 1;
            
            // 
            // chkAutoSave
            // 
            this.chkAutoSave.AutoSize = true;
            this.chkAutoSave.Location = new System.Drawing.Point(20, 65);
            this.chkAutoSave.Name = "chkAutoSave";
            this.chkAutoSave.Size = new System.Drawing.Size(104, 19);
            this.chkAutoSave.TabIndex = 2;
            this.chkAutoSave.Text = "自动保存设置";
            this.chkAutoSave.UseVisualStyleBackColor = true;
            
            // 
            // chkShowTips
            // 
            this.chkShowTips.AutoSize = true;
            this.chkShowTips.Location = new System.Drawing.Point(20, 95);
            this.chkShowTips.Name = "chkShowTips";
            this.chkShowTips.Size = new System.Drawing.Size(89, 19);
            this.chkShowTips.TabIndex = 3;
            this.chkShowTips.Text = "显示提示";
            this.chkShowTips.UseVisualStyleBackColor = true;
            
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(335, 410);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(420, 410);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(505, 410);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 30);
            this.btnApply.TabIndex = 3;
            this.btnApply.Text = "应用";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(584, 461);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "设置";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPageApi.ResumeLayout(false);
            this.tabPageApi.PerformLayout();
            this.tabPageGeneral.ResumeLayout(false);
            this.tabPageGeneral.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeout)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageGeneral;
        private System.Windows.Forms.TabPage tabPageApi;
        private System.Windows.Forms.TabPage tabPageAdvanced;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnApply;
        
        // API设置控件
        private System.Windows.Forms.Label lblProvider;
        private System.Windows.Forms.ComboBox cmbProvider;
        private System.Windows.Forms.Label lblModel;
        private System.Windows.Forms.TextBox txtModel;
        private System.Windows.Forms.Label lblApiKey;
        private System.Windows.Forms.TextBox txtApiKey;
        private System.Windows.Forms.Label lblBaseUrl;
        private System.Windows.Forms.TextBox txtBaseUrl;
        private System.Windows.Forms.Label lblTimeout;
        private System.Windows.Forms.NumericUpDown numTimeout;
        private System.Windows.Forms.Button btnTestConnection;
        
        // 通用设置控件
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cmbLanguage;
        private System.Windows.Forms.CheckBox chkAutoSave;
        private System.Windows.Forms.CheckBox chkShowTips;
    }
}