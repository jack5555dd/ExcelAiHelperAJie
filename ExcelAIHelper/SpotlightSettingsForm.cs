using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAIHelper
{
    public partial class SpotlightSettingsForm : Form
    {
        private Panel pnlColor;
        private ComboBox cboMode;
        
        public SpotlightSettingsForm()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }
        
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "聚光灯设置";
            this.Size = new Size(400, 250);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            
            // Color label
            var lblColor = new Label
            {
                Text = "颜色:",
                Location = new Point(20, 30),
                Size = new Size(60, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblColor);
            
            // Color panel (shows current color)
            pnlColor = new Panel
            {
                Location = new Point(90, 30),
                Size = new Size(200, 23),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = SpotlightManager.CurrentHighlightColor,
                Cursor = Cursors.Hand
            };
            pnlColor.Click += PnlColor_Click;
            this.Controls.Add(pnlColor);
            
            // Mode label
            var lblMode = new Label
            {
                Text = "显示项:",
                Location = new Point(20, 70),
                Size = new Size(60, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblMode);
            
            // Mode combo box
            cboMode = new ComboBox
            {
                Location = new Point(90, 70),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboMode.Items.Add("全部(行+列)");
            cboMode.Items.Add("行");
            cboMode.Items.Add("列");
            cboMode.SelectedIndex = (int)SpotlightManager.CurrentMode;
            this.Controls.Add(cboMode);
            
            // Status label
            var lblStatus = new Label
            {
                Text = $"当前状态: {(SpotlightManager.IsActive ? "开启" : "关闭")}",
                Location = new Point(20, 110),
                Size = new Size(200, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SpotlightManager.IsActive ? Color.Green : Color.Gray
            };
            this.Controls.Add(lblStatus);
            
            // OK button
            var btnOK = new Button
            {
                Text = "确定",
                Location = new Point(135, 160),
                Size = new Size(75, 23),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);
            
            // Cancel button
            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(220, 160),
                Size = new Size(75, 23),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);
            
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
            
            this.ResumeLayout(false);
        }
        
        private void LoadCurrentSettings()
        {
            // 设置已在InitializeComponent中完成
        }
        
        private void PnlColor_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = pnlColor.BackColor;
                colorDialog.FullOpen = true;
                
                // 添加一些常用颜色
                colorDialog.CustomColors = new int[] {
                    ColorTranslator.ToOle(Color.Yellow),
                    ColorTranslator.ToOle(Color.LightBlue),
                    ColorTranslator.ToOle(Color.LightGreen),
                    ColorTranslator.ToOle(Color.LightPink),
                    ColorTranslator.ToOle(Color.LightGray),
                    ColorTranslator.ToOle(Color.Orange)
                };
                
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    pnlColor.BackColor = colorDialog.Color;
                }
            }
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // 应用设置到SpotlightManager
                SpotlightManager.SetHighlightColor(pnlColor.BackColor);
                SpotlightManager.SetMode((SpotlightMode)cboMode.SelectedIndex);
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}