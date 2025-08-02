using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAIHelper.Services;

namespace ExcelAIHelper
{
    /// <summary>
    /// VBE权限设置指导对话框
    /// </summary>
    public partial class VbePermissionDialog : Form
    {
        public VbePermissionDialog()
        {
            InitializeComponent();
            LoadInstructions();
        }

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.rtbInstructions = new System.Windows.Forms.RichTextBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.picIcon = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picIcon)).BeginInit();
            this.SuspendLayout();
            
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(60, 20);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(200, 20);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "VBA功能需要额外权限";
            
            // 
            // picIcon
            // 
            this.picIcon.Location = new System.Drawing.Point(20, 15);
            this.picIcon.Name = "picIcon";
            this.picIcon.Size = new System.Drawing.Size(32, 32);
            this.picIcon.TabIndex = 1;
            this.picIcon.TabStop = false;
            
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(20, 60);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(300, 13);
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "当前状态：VBE访问权限未启用";
            
            // 
            // rtbInstructions
            // 
            this.rtbInstructions.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbInstructions.BackColor = System.Drawing.SystemColors.Window;
            this.rtbInstructions.Location = new System.Drawing.Point(20, 85);
            this.rtbInstructions.Name = "rtbInstructions";
            this.rtbInstructions.ReadOnly = true;
            this.rtbInstructions.Size = new System.Drawing.Size(540, 280);
            this.rtbInstructions.TabIndex = 3;
            this.rtbInstructions.Text = "";
            
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefresh.Location = new System.Drawing.Point(400, 380);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 23);
            this.btnRefresh.TabIndex = 4;
            this.btnRefresh.Text = "重新检查";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(485, 380);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            
            // 
            // VbePermissionDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(580, 420);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.rtbInstructions);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.picIcon);
            this.Controls.Add(this.lblTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "VbePermissionDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "VBA权限设置指导";
            ((System.ComponentModel.ISupportInitialize)(this.picIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.RichTextBox rtbInstructions;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.PictureBox picIcon;

        /// <summary>
        /// 加载设置指导内容
        /// </summary>
        private void LoadInstructions()
        {
            try
            {
                // 设置图标
                picIcon.Image = SystemIcons.Information.ToBitmap();
                
                // 检查当前状态
                bool vbeAccessEnabled = ExecutionModeManager.CheckVbeAccess();
                
                if (vbeAccessEnabled)
                {
                    lblStatus.Text = "当前状态：✅ VBE访问权限已启用";
                    lblStatus.ForeColor = Color.Green;
                    lblTitle.Text = "VBA功能已可用";
                    btnRefresh.Text = "确定";
                }
                else
                {
                    lblStatus.Text = "当前状态：❌ VBE访问权限未启用";
                    lblStatus.ForeColor = Color.Red;
                }
                
                // 加载详细说明
                LoadDetailedInstructions(vbeAccessEnabled);
            }
            catch (Exception ex)
            {
                rtbInstructions.Text = $"加载指导内容时出错：{ex.Message}";
            }
        }

        /// <summary>
        /// 加载详细的设置指导
        /// </summary>
        private void LoadDetailedInstructions(bool vbeAccessEnabled)
        {
            rtbInstructions.Clear();
            
            if (vbeAccessEnabled)
            {
                // VBE访问已启用
                AppendFormattedText("🎉 恭喜！VBA功能已可用", Color.Green, true);
                AppendFormattedText("\n\n", Color.Black, false);
                AppendFormattedText("您的Excel已正确配置VBE访问权限，现在可以使用AI-VBA功能了。", Color.Black, false);
                AppendFormattedText("\n\n", Color.Black, false);
                AppendFormattedText("VBA模式功能：", Color.Blue, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("• AI将根据您的需求生成VBA代码\n", Color.Black, false);
                AppendFormattedText("• 代码会经过安全扫描确保安全性\n", Color.Black, false);
                AppendFormattedText("• 您可以预览代码后选择执行\n", Color.Black, false);
                AppendFormattedText("• 支持复制代码到剪贴板\n", Color.Black, false);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("现在您可以关闭此对话框，在聊天界面中选择VBA模式开始使用。", Color.Green, false);
            }
            else
            {
                // VBE访问未启用，显示设置指导
                AppendFormattedText("要启用VBA功能，请按以下步骤设置：", Color.Blue, true);
                AppendFormattedText("\n\n", Color.Black, false);
                
                AppendFormattedText("📋 详细设置步骤：", Color.DarkBlue, true);
                AppendFormattedText("\n\n", Color.Black, false);
                
                AppendFormattedText("1. 打开Excel选项", Color.Black, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("   • 点击Excel菜单栏的\"文件\"选项\n", Color.Black, false);
                AppendFormattedText("   • 在左侧菜单中选择\"选项\"\n", Color.Black, false);
                AppendFormattedText("\n", Color.Black, false);
                
                AppendFormattedText("2. 进入信任中心", Color.Black, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("   • 在Excel选项对话框左侧选择\"信任中心\"\n", Color.Black, false);
                AppendFormattedText("   • 点击右侧的\"信任中心设置\"按钮\n", Color.Black, false);
                AppendFormattedText("\n", Color.Black, false);
                
                AppendFormattedText("3. 配置宏设置", Color.Black, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("   • 在信任中心对话框左侧选择\"宏设置\"\n", Color.Black, false);
                AppendFormattedText("   • 勾选\"信任对VBA项目对象模型的访问\"", Color.Red, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("   • 点击\"确定\"保存设置\n", Color.Black, false);
                AppendFormattedText("\n", Color.Black, false);
                
                AppendFormattedText("4. 重启Excel", Color.Black, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("   • 关闭Excel应用程序\n", Color.Black, false);
                AppendFormattedText("   • 重新打开Excel使设置生效\n", Color.Black, false);
                AppendFormattedText("\n", Color.Black, false);
                
                AppendFormattedText("⚠️ 重要提示：", Color.Orange, true);
                AppendFormattedText("\n", Color.Black, false);
                AppendFormattedText("• 此设置需要管理员权限，如果无法修改请联系系统管理员\n", Color.Orange, false);
                AppendFormattedText("• 启用此选项是安全的，它只允许受信任的加载项访问VBA环境\n", Color.Orange, false);
                AppendFormattedText("• 我们的AI-VBA功能包含完整的安全扫描机制\n", Color.Orange, false);
                AppendFormattedText("\n", Color.Black, false);
                
                AppendFormattedText("完成设置后，点击\"重新检查\"按钮验证配置。", Color.Green, false);
            }
        }

        /// <summary>
        /// 添加格式化文本到RichTextBox
        /// </summary>
        private void AppendFormattedText(string text, Color color, bool bold)
        {
            rtbInstructions.SelectionStart = rtbInstructions.TextLength;
            rtbInstructions.SelectionLength = 0;
            rtbInstructions.SelectionColor = color;
            rtbInstructions.SelectionFont = new Font(rtbInstructions.Font, bold ? FontStyle.Bold : FontStyle.Regular);
            rtbInstructions.AppendText(text);
        }

        /// <summary>
        /// 重新检查按钮点击事件
        /// </summary>
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                // 刷新VBA状态
                ExecutionModeManager.RefreshVbaStatus();
                
                // 重新加载指导内容
                LoadInstructions();
                
                // 如果VBA已可用，可以关闭对话框
                if (ExecutionModeManager.IsVbaEnabled)
                {
                    MessageBox.Show("✅ VBA功能已启用！现在可以使用AI-VBA模式了。", 
                                  "设置成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("VBE访问权限仍未启用，请按照指导完成设置后重试。", 
                                  "权限检查", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"检查权限时出错：{ex.Message}", 
                              "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关闭按钮点击事件
        /// </summary>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// 显示VBE权限设置对话框
        /// </summary>
        /// <param name="parent">父窗口</param>
        /// <returns>对话框结果</returns>
        public static DialogResult ShowVbePermissionDialog(IWin32Window parent = null)
        {
            using (var dialog = new VbePermissionDialog())
            {
                return dialog.ShowDialog(parent);
            }
        }
    }
}