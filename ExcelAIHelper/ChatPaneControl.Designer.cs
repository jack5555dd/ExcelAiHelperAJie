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
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // richTextBox
            // 
            this.richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox.Location = new System.Drawing.Point(0, 0);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(300, 400);
            this.richTextBox.TabIndex = 0;
            this.richTextBox.Text = "AI聊天助手面板（占位）";
            // 
            // ChatPaneControl
            // 
            this.Controls.Add(this.richTextBox);
            this.Name = "ChatPaneControl";
            this.Size = new System.Drawing.Size(300, 400);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox;
    }
}