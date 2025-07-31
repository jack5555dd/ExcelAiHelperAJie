namespace ExcelAIHelper
{
    partial class ExampleForm
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
            this.lblDescription = new System.Windows.Forms.Label();
            this.listExamples = new System.Windows.Forms.ListBox();
            this.lblExampleTitle = new System.Windows.Forms.Label();
            this.txtExampleDescription = new System.Windows.Forms.TextBox();
            this.linkExample = new System.Windows.Forms.LinkLabel();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnVisitExample = new System.Windows.Forms.Button();
            this.panelMain = new System.Windows.Forms.Panel();
            this.panelMain.SuspendLayout();
            this.SuspendLayout();
            
            // 
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.Controls.Add(this.lblTitle);
            this.panelMain.Controls.Add(this.lblDescription);
            this.panelMain.Controls.Add(this.listExamples);
            this.panelMain.Controls.Add(this.lblExampleTitle);
            this.panelMain.Controls.Add(this.txtExampleDescription);
            this.panelMain.Controls.Add(this.linkExample);
            this.panelMain.Controls.Add(this.btnVisitExample);
            this.panelMain.Location = new System.Drawing.Point(12, 12);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(560, 400);
            this.panelMain.TabIndex = 0;
            
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("微软雅黑", 14F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.lblTitle.Location = new System.Drawing.Point(20, 20);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(126, 25);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "使用示例";
            
            // 
            // lblDescription
            // 
            this.lblDescription.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDescription.ForeColor = System.Drawing.Color.DimGray;
            this.lblDescription.Location = new System.Drawing.Point(25, 55);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new System.Drawing.Size(500, 20);
            this.lblDescription.TabIndex = 1;
            this.lblDescription.Text = "以下是Excel AI助手的一些使用示例，点击查看详细说明和在线演示。";
            
            // 
            // listExamples
            // 
            this.listExamples.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.listExamples.FormattingEnabled = true;
            this.listExamples.ItemHeight = 17;
            this.listExamples.Location = new System.Drawing.Point(25, 85);
            this.listExamples.Name = "listExamples";
            this.listExamples.Size = new System.Drawing.Size(200, 276);
            this.listExamples.TabIndex = 2;
            this.listExamples.SelectedIndexChanged += new System.EventHandler(this.listExamples_SelectedIndexChanged);
            
            // 
            // lblExampleTitle
            // 
            this.lblExampleTitle.AutoSize = true;
            this.lblExampleTitle.Font = new System.Drawing.Font("微软雅黑", 11F, System.Drawing.FontStyle.Bold);
            this.lblExampleTitle.Location = new System.Drawing.Point(250, 85);
            this.lblExampleTitle.Name = "lblExampleTitle";
            this.lblExampleTitle.Size = new System.Drawing.Size(103, 19);
            this.lblExampleTitle.TabIndex = 3;
            this.lblExampleTitle.Text = "选择示例查看";
            
            // 
            // txtExampleDescription
            // 
            this.txtExampleDescription.BackColor = System.Drawing.Color.WhiteSmoke;
            this.txtExampleDescription.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtExampleDescription.Location = new System.Drawing.Point(250, 115);
            this.txtExampleDescription.Multiline = true;
            this.txtExampleDescription.Name = "txtExampleDescription";
            this.txtExampleDescription.ReadOnly = true;
            this.txtExampleDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtExampleDescription.Size = new System.Drawing.Size(280, 180);
            this.txtExampleDescription.TabIndex = 4;
            this.txtExampleDescription.Text = "请从左侧列表中选择一个示例来查看详细说明。";
            
            // 
            // linkExample
            // 
            this.linkExample.AutoSize = true;
            this.linkExample.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.linkExample.Location = new System.Drawing.Point(250, 310);
            this.linkExample.Name = "linkExample";
            this.linkExample.Size = new System.Drawing.Size(200, 17);
            this.linkExample.TabIndex = 5;
            this.linkExample.TabStop = true;
            this.linkExample.Text = "https://examples.aiexcel.cn";
            this.linkExample.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkExample_LinkClicked);
            
            // 
            // btnVisitExample
            // 
            this.btnVisitExample.Location = new System.Drawing.Point(250, 340);
            this.btnVisitExample.Name = "btnVisitExample";
            this.btnVisitExample.Size = new System.Drawing.Size(100, 30);
            this.btnVisitExample.TabIndex = 6;
            this.btnVisitExample.Text = "在线演示";
            this.btnVisitExample.UseVisualStyleBackColor = true;
            this.btnVisitExample.Click += new System.EventHandler(this.btnVisitExample_Click);
            
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(497, 430);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "关闭";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            
            // 
            // ExampleForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(584, 472);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.panelMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExampleForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "使用示例";
            this.Load += new System.EventHandler(this.ExampleForm_Load);
            this.panelMain.ResumeLayout(false);
            this.panelMain.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDescription;
        private System.Windows.Forms.ListBox listExamples;
        private System.Windows.Forms.Label lblExampleTitle;
        private System.Windows.Forms.TextBox txtExampleDescription;
        private System.Windows.Forms.LinkLabel linkExample;
        private System.Windows.Forms.Button btnVisitExample;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Panel panelMain;
    }
}