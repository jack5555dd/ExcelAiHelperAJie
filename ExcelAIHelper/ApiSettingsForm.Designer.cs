namespace ExcelAIHelper
{
    partial class ApiSettingsForm
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

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.txtApiKey = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblApiKey = new System.Windows.Forms.Label();
            this.lblApiEndpoint = new System.Windows.Forms.Label();
            this.txtApiEndpoint = new System.Windows.Forms.TextBox();
            this.cboModel = new System.Windows.Forms.ComboBox();
            this.lblModel = new System.Windows.Forms.Label();
            this.lblTemperature = new System.Windows.Forms.Label();
            this.nudTemperature = new System.Windows.Forms.NumericUpDown();
            this.lblMaxTokens = new System.Windows.Forms.Label();
            this.nudMaxTokens = new System.Windows.Forms.NumericUpDown();
            this.btnTestConnection = new System.Windows.Forms.Button();
            this.btnNetworkDiag = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nudTemperature)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaxTokens)).BeginInit();
            this.SuspendLayout();
            // 
            // txtApiKey
            // 
            this.txtApiKey.Location = new System.Drawing.Point(120, 15);
            this.txtApiKey.Name = "txtApiKey";
            this.txtApiKey.Size = new System.Drawing.Size(360, 20);
            this.txtApiKey.TabIndex = 0;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(405, 195);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnTestConnection
            // 
            this.btnTestConnection.Location = new System.Drawing.Point(300, 195);
            this.btnTestConnection.Name = "btnTestConnection";
            this.btnTestConnection.Size = new System.Drawing.Size(90, 23);
            this.btnTestConnection.TabIndex = 6;
            this.btnTestConnection.Text = "测试连接";
            this.btnTestConnection.UseVisualStyleBackColor = true;
            this.btnTestConnection.Click += new System.EventHandler(this.btnTestConnection_Click);
            // 
            // btnNetworkDiag
            // 
            this.btnNetworkDiag.Location = new System.Drawing.Point(190, 195);
            this.btnNetworkDiag.Name = "btnNetworkDiag";
            this.btnNetworkDiag.Size = new System.Drawing.Size(90, 23);
            this.btnNetworkDiag.TabIndex = 7;
            this.btnNetworkDiag.Text = "网络诊断";
            this.btnNetworkDiag.UseVisualStyleBackColor = true;
            this.btnNetworkDiag.Click += new System.EventHandler(this.btnNetworkDiag_Click);
            // 
            // lblApiKey
            // 
            this.lblApiKey.AutoSize = true;
            this.lblApiKey.Location = new System.Drawing.Point(15, 18);
            this.lblApiKey.Name = "lblApiKey";
            this.lblApiKey.Size = new System.Drawing.Size(50, 13);
            this.lblApiKey.TabIndex = 6;
            this.lblApiKey.Text = "API Key:";
            // 
            // lblApiEndpoint
            // 
            this.lblApiEndpoint.AutoSize = true;
            this.lblApiEndpoint.Location = new System.Drawing.Point(15, 55);
            this.lblApiEndpoint.Name = "lblApiEndpoint";
            this.lblApiEndpoint.Size = new System.Drawing.Size(72, 13);
            this.lblApiEndpoint.TabIndex = 7;
            this.lblApiEndpoint.Text = "API Endpoint:";
            // 
            // txtApiEndpoint
            // 
            this.txtApiEndpoint.Location = new System.Drawing.Point(120, 52);
            this.txtApiEndpoint.Name = "txtApiEndpoint";
            this.txtApiEndpoint.Size = new System.Drawing.Size(360, 20);
            this.txtApiEndpoint.TabIndex = 1;
            // 
            // cboModel
            // 
            this.cboModel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModel.FormattingEnabled = true;
            this.cboModel.Items.AddRange(new object[] {
            "deepseek-chat",
            "deepseek-coder",
            "deepseek-lite",
            "other"});
            this.cboModel.Location = new System.Drawing.Point(120, 89);
            this.cboModel.Name = "cboModel";
            this.cboModel.Size = new System.Drawing.Size(360, 21);
            this.cboModel.TabIndex = 2;
            // 
            // lblModel
            // 
            this.lblModel.AutoSize = true;
            this.lblModel.Location = new System.Drawing.Point(15, 92);
            this.lblModel.Name = "lblModel";
            this.lblModel.Size = new System.Drawing.Size(39, 13);
            this.lblModel.TabIndex = 9;
            this.lblModel.Text = "Model:";
            // 
            // lblTemperature
            // 
            this.lblTemperature.AutoSize = true;
            this.lblTemperature.Location = new System.Drawing.Point(15, 129);
            this.lblTemperature.Name = "lblTemperature";
            this.lblTemperature.Size = new System.Drawing.Size(70, 13);
            this.lblTemperature.TabIndex = 10;
            this.lblTemperature.Text = "Temperature:";
            // 
            // nudTemperature
            // 
            this.nudTemperature.DecimalPlaces = 1;
            this.nudTemperature.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudTemperature.Location = new System.Drawing.Point(120, 127);
            this.nudTemperature.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nudTemperature.Name = "nudTemperature";
            this.nudTemperature.Size = new System.Drawing.Size(120, 20);
            this.nudTemperature.TabIndex = 3;
            this.nudTemperature.Value = new decimal(new int[] {
            7,
            0,
            0,
            65536});
            // 
            // lblMaxTokens
            // 
            this.lblMaxTokens.AutoSize = true;
            this.lblMaxTokens.Location = new System.Drawing.Point(15, 166);
            this.lblMaxTokens.Name = "lblMaxTokens";
            this.lblMaxTokens.Size = new System.Drawing.Size(68, 13);
            this.lblMaxTokens.TabIndex = 12;
            this.lblMaxTokens.Text = "Max Tokens:";
            // 
            // nudMaxTokens
            // 
            this.nudMaxTokens.Increment = new decimal(new int[] {
            256,
            0,
            0,
            0});
            this.nudMaxTokens.Location = new System.Drawing.Point(120, 164);
            this.nudMaxTokens.Maximum = new decimal(new int[] {
            8192,
            0,
            0,
            0});
            this.nudMaxTokens.Minimum = new decimal(new int[] {
            256,
            0,
            0,
            0});
            this.nudMaxTokens.Name = "nudMaxTokens";
            this.nudMaxTokens.Size = new System.Drawing.Size(120, 20);
            this.nudMaxTokens.TabIndex = 4;
            this.nudMaxTokens.Value = new decimal(new int[] {
            2048,
            0,
            0,
            0});
            // 
            // ApiSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 231);
            this.Controls.Add(this.nudMaxTokens);
            this.Controls.Add(this.lblMaxTokens);
            this.Controls.Add(this.nudTemperature);
            this.Controls.Add(this.lblTemperature);
            this.Controls.Add(this.lblModel);
            this.Controls.Add(this.cboModel);
            this.Controls.Add(this.txtApiEndpoint);
            this.Controls.Add(this.lblApiEndpoint);
            this.Controls.Add(this.lblApiKey);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnTestConnection);
            this.Controls.Add(this.btnNetworkDiag);
            this.Controls.Add(this.txtApiKey);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ApiSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "API 设置";
            this.Load += new System.EventHandler(this.ApiSettingsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.nudTemperature)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaxTokens)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox txtApiKey;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblApiKey;
        private System.Windows.Forms.Label lblApiEndpoint;
        private System.Windows.Forms.TextBox txtApiEndpoint;
        private System.Windows.Forms.ComboBox cboModel;
        private System.Windows.Forms.Label lblModel;
        private System.Windows.Forms.Label lblTemperature;
        private System.Windows.Forms.NumericUpDown nudTemperature;
        private System.Windows.Forms.Label lblMaxTokens;
        private System.Windows.Forms.NumericUpDown nudMaxTokens;
        private System.Windows.Forms.Button btnTestConnection;
        private System.Windows.Forms.Button btnNetworkDiag;
    }
}