using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAIHelper
{
    public partial class QuickInputSettingsForm : Form
    {
        private TextBox txtRange;
        private ComboBox cboInputType;
        private DataGridView dgvCustomMappings;
        private Timer selectionTimer;
        private Excel.AppEvents_SheetSelectionChangeEventHandler _selectionChangeHandler;
        private Dictionary<string, string> _customMappings = new Dictionary<string, string>();
        private Label lblStatus;
        private Label descLabel;
        
        public QuickInputSettingsForm()
        {
            InitializeComponent();
            LoadDefaultSettings();
        }
        
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "快速录入设置";
            this.Size = new Size(500, 450);
            this.StartPosition = FormStartPosition.Manual;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.TopMost = true; // 始终显示在Excel上方
            
            // 设置窗口位置到屏幕右侧
            this.Load += (s, e) => {
                var screen = Screen.PrimaryScreen.WorkingArea;
                this.Location = new Point(screen.Width - this.Width - 20, 100);
            };
            
            // 创建主面板
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15)
            };
            this.Controls.Add(mainPanel);
            
            // 标题区域
            var titlePanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(470, 50),
                BackColor = Color.FromArgb(0, 120, 215)
            };
            
            var titleIcon = new Label
            {
                Text = "⚡",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 12),
                Size = new Size(30, 25),
                TextAlign = ContentAlignment.MiddleCenter
            };
            titlePanel.Controls.Add(titleIcon);
            
            var titleLabel = new Label
            {
                Text = "快速录入设置",
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(50, 15),
                Size = new Size(150, 20)
            };
            titlePanel.Controls.Add(titleLabel);
            
            mainPanel.Controls.Add(titlePanel);
            
            // 说明文字
            descLabel = new Label
            {
                Text = "💡 数据说明：快速录入1，显示'男'；快速录入2，显示'女'",
                Location = new Point(15, 65),
                Size = new Size(450, 25),
                Font = new Font("Microsoft YaHei", 9),
                ForeColor = Color.Gray,
                BackColor = Color.FromArgb(248, 249, 250),
                TextAlign = ContentAlignment.MiddleLeft
            };
            mainPanel.Controls.Add(descLabel);
            
            // 应用区域
            var lblRange = new Label
            {
                Text = "应用区域:",
                Location = new Point(15, 105),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei", 9)
            };
            mainPanel.Controls.Add(lblRange);
            
            txtRange = new TextBox
            {
                Location = new Point(100, 105),
                Size = new Size(200, 23),
                Text = "C8:G8",
                Font = new Font("Microsoft YaHei", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240)
            };
            mainPanel.Controls.Add(txtRange);
            
            var lblRangeNote = new Label
            {
                Text = "📍 实时跟踪选区变化",
                Location = new Point(310, 105),
                Size = new Size(150, 23),
                Font = new Font("Microsoft YaHei", 8),
                ForeColor = Color.Green,
                TextAlign = ContentAlignment.MiddleLeft
            };
            mainPanel.Controls.Add(lblRangeNote);
            
            // 录入类型
            var lblType = new Label
            {
                Text = "录入类型:",
                Location = new Point(15, 140),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei", 9)
            };
            mainPanel.Controls.Add(lblType);
            
            cboInputType = new ComboBox
            {
                Location = new Point(100, 140),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9)
            };
            cboInputType.Items.AddRange(new string[] {
                "序号 (1, 2, 3...)",
                "字母 (A, B, C...)",
                "日期序列",
                "自定义快捷输入",
                "数值序列"
            });
            cboInputType.SelectedIndex = 3; // 默认选择自定义快捷输入
            cboInputType.SelectedIndexChanged += CboInputType_SelectedIndexChanged;
            mainPanel.Controls.Add(cboInputType);
            
            // 自定义映射表格
            var lblMappings = new Label
            {
                Text = "快捷输入映射:",
                Location = new Point(15, 175),
                Size = new Size(100, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei", 9)
            };
            mainPanel.Controls.Add(lblMappings);
            
            dgvCustomMappings = new DataGridView
            {
                Location = new Point(15, 200),
                Size = new Size(450, 150),
                Font = new Font("Microsoft YaHei", 9),
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            
            // 添加列
            dgvCustomMappings.Columns.Add("Input", "快捷输入");
            dgvCustomMappings.Columns.Add("Output", "显示内容");
            dgvCustomMappings.Columns["Input"].Width = 100;
            dgvCustomMappings.Columns["Output"].Width = 200;
            
            // 添加默认数据
            dgvCustomMappings.Rows.Add("1", "男");
            dgvCustomMappings.Rows.Add("2", "女");
            dgvCustomMappings.Rows.Add("3", "未知");
            
            dgvCustomMappings.CellValueChanged += DgvCustomMappings_CellValueChanged;
            mainPanel.Controls.Add(dgvCustomMappings);
            
            // 状态显示区域
            var statusPanel = new Panel
            {
                Location = new Point(15, 315),
                Size = new Size(450, 30),
                BackColor = Color.FromArgb(248, 249, 250)
            };
            mainPanel.Controls.Add(statusPanel);
            
            lblStatus = new Label
            {
                Text = "📋 状态: 未设置快速录入",
                Location = new Point(15, 5),
                Size = new Size(420, 20),
                Font = new Font("Microsoft YaHei", 9),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusPanel.Controls.Add(lblStatus);
            
            // 按钮区域
            var buttonPanel = new Panel
            {
                Location = new Point(15, 355),
                Size = new Size(450, 40),
                BackColor = Color.FromArgb(248, 249, 250)
            };
            mainPanel.Controls.Add(buttonPanel);
            
            var btnConfirm = new Button
            {
                Text = "✅ 确认设置",
                Location = new Point(15, 8),
                Size = new Size(100, 25),
                Font = new Font("Microsoft YaHei", 9),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                UseVisualStyleBackColor = false
            };
            btnConfirm.Click += BtnConfirm_Click;
            buttonPanel.Controls.Add(btnConfirm);
            
            var btnCancel = new Button
            {
                Text = "🚫 撤销快速录入",
                Location = new Point(125, 8),
                Size = new Size(120, 25),
                Font = new Font("Microsoft YaHei", 9),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                UseVisualStyleBackColor = false
            };
            btnCancel.Click += BtnCancel_Click;
            buttonPanel.Controls.Add(btnCancel);
            
            var btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(360, 8),
                Size = new Size(75, 25),
                Font = new Font("Microsoft YaHei", 9),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                UseVisualStyleBackColor = false,
                DialogResult = DialogResult.Cancel
            };
            btnClose.Click += (s, e) => this.Close();
            buttonPanel.Controls.Add(btnClose);
            
            this.CancelButton = btnClose;
            
            // 初始化定时器
            selectionTimer = new Timer
            {
                Interval = 500, // 500ms检查一次
                Enabled = false
            };
            selectionTimer.Tick += SelectionTimer_Tick;
            
            // 窗口激活事件处理
            this.Activated += (s, e) => {
                // 窗口激活时更新状态
                UpdateStatusDisplay();
            };
            
            // 防止窗口失去焦点时隐藏
            this.Deactivate += (s, e) => {
                // 延迟一点时间再检查，避免快速切换时的问题
                var timer = new Timer { Interval = 100 };
                timer.Tick += (sender, args) => {
                    timer.Stop();
                    timer.Dispose();
                    // 确保窗口仍然可见
                    if (!this.IsDisposed && this.Visible)
                    {
                        this.TopMost = true;
                    }
                };
                timer.Start();
            };
            
            this.ResumeLayout(false);
        }
        
        private void LoadDefaultSettings()
        {
            try
            {
                // 获取当前选中区域
                UpdateSelectedRange();
                
                // 启动选区跟踪
                StartSelectionTracking();
                
                // 更新状态显示
                UpdateStatusDisplay();
                
                // 更新说明文字
                UpdateDescriptionText();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadDefaultSettings error: {ex.Message}");
                txtRange.Text = "A1:A10";
            }
        }
        
        private void UpdateStatusDisplay()
        {
            try
            {
                if (QuickInputManager.IsActive)
                {
                    lblStatus.Text = "✅ 状态: 快速录入已激活";
                    lblStatus.ForeColor = Color.Green;
                }
                else
                {
                    lblStatus.Text = "📋 状态: 未设置快速录入";
                    lblStatus.ForeColor = Color.Gray;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateStatusDisplay error: {ex.Message}");
            }
        }
        
        private void UpdateSelectedRange()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.Selection as Excel.Range;
                if (selection != null)
                {
                    txtRange.Text = selection.Address[false, false];
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateSelectedRange error: {ex.Message}");
            }
        }
        
        private void StartSelectionTracking()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                _selectionChangeHandler = new Excel.AppEvents_SheetSelectionChangeEventHandler(OnSelectionChange);
                app.SheetSelectionChange += _selectionChangeHandler;
                selectionTimer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"StartSelectionTracking error: {ex.Message}");
            }
        }
        
        private void StopSelectionTracking()
        {
            try
            {
                selectionTimer.Stop();
                if (_selectionChangeHandler != null)
                {
                    var app = Globals.ThisAddIn.Application;
                    app.SheetSelectionChange -= _selectionChangeHandler;
                    _selectionChangeHandler = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"StopSelectionTracking error: {ex.Message}");
            }
        }
        
        private void OnSelectionChange(object sh, Excel.Range target)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => OnSelectionChange(sh, target)));
                    return;
                }
                
                if (target != null)
                {
                    txtRange.Text = target.Address[false, false];
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnSelectionChange error: {ex.Message}");
            }
        }
        
        private void SelectionTimer_Tick(object sender, EventArgs e)
        {
            UpdateSelectedRange();
        }
        
        private void CboInputType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 根据选择的类型显示/隐藏自定义映射表格
            bool showMappings = cboInputType.SelectedIndex == 3; // 自定义快捷输入
            dgvCustomMappings.Visible = showMappings;
            
            // 更新说明文字
            UpdateDescriptionText();
            
            if (showMappings)
            {
                UpdateCustomMappings();
            }
        }
        
        /// <summary>
        /// 根据选择的录入类型更新说明文字
        /// </summary>
        private void UpdateDescriptionText()
        {
            try
            {
                switch (cboInputType.SelectedIndex)
                {
                    case 0: // 序号
                        descLabel.Text = "💡 数据说明：在选定区域填充数字序列 1, 2, 3, 4...";
                        break;
                    case 1: // 字母
                        descLabel.Text = "💡 数据说明：在选定区域填充字母序列 A, B, C, D...";
                        break;
                    case 2: // 日期序列
                        descLabel.Text = "💡 数据说明：在选定区域填充日期序列，从今天开始递增";
                        break;
                    case 3: // 自定义快捷输入
                        descLabel.Text = "💡 数据说明：快速录入1，显示'男'；快速录入2，显示'女'";
                        break;
                    case 4: // 数值序列
                        descLabel.Text = "💡 数据说明：在选定区域填充数值序列 1.0, 2.0, 3.0...";
                        break;
                    default:
                        descLabel.Text = "💡 数据说明：请选择录入类型";
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateDescriptionText error: {ex.Message}");
            }
        }
        
        private void DgvCustomMappings_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            UpdateCustomMappings();
        }
        
        private void UpdateCustomMappings()
        {
            try
            {
                _customMappings.Clear();
                
                foreach (DataGridViewRow row in dgvCustomMappings.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                    {
                        string input = row.Cells[0].Value.ToString().Trim();
                        string output = row.Cells[1].Value.ToString().Trim();
                        
                        if (!string.IsNullOrEmpty(input) && !string.IsNullOrEmpty(output))
                        {
                            _customMappings[input] = output;
                        }
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"Updated mappings: {_customMappings.Count} items");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateCustomMappings error: {ex.Message}");
            }
        }
        
        private void BtnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取当前选中的区域
                string targetRange = txtRange.Text;
                if (string.IsNullOrEmpty(targetRange))
                {
                    MessageBox.Show("请先在Excel中选择要应用快速录入的区域", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 根据选择的录入类型执行不同的操作
                int selectedType = cboInputType.SelectedIndex;
                string typeName = cboInputType.SelectedItem.ToString();
                
                // 确认对话框
                DialogResult result;
                
                if (selectedType == 3) // 自定义快捷输入
                {
                    // 更新自定义映射
                    UpdateCustomMappings();
                    
                    result = MessageBox.Show(
                        $"确认要为区域 {targetRange} 设置快速录入吗？\n\n" +
                        $"录入类型: {typeName}\n" +
                        $"设置的映射关系:\n" +
                        string.Join("\n", _customMappings.Select(kv => $"  {kv.Key} → {kv.Value}")),
                        "确认设置", 
                        MessageBoxButtons.YesNo, 
                        MessageBoxIcon.Question);
                }
                else
                {
                    result = MessageBox.Show(
                        $"确认要为区域 {targetRange} 填充数据吗？\n\n" +
                        $"录入类型: {typeName}\n" +
                        $"将会在选定区域填充相应的序列数据。",
                        "确认设置", 
                        MessageBoxButtons.YesNo, 
                        MessageBoxIcon.Question);
                }
                
                if (result != DialogResult.Yes)
                    return;
                
                // 根据类型执行相应操作
                switch (selectedType)
                {
                    case 0: // 序号 (1, 2, 3...)
                        FillNumberSequence(targetRange);
                        break;
                    case 1: // 字母 (A, B, C...)
                        FillLetterSequence(targetRange);
                        break;
                    case 2: // 日期序列
                        FillDateSequence(targetRange);
                        break;
                    case 3: // 自定义快捷输入
                        QuickInputManager.Start(_customMappings, targetRange);
                        break;
                    case 4: // 数值序列
                        FillNumericSequence(targetRange);
                        break;
                }
                
                // 更新状态显示
                if (selectedType == 3)
                {
                    lblStatus.Text = $"✅ 状态: 已设置快速录入 (区域: {targetRange})";
                    lblStatus.ForeColor = Color.Green;
                    MessageBox.Show($"快速录入已设置完成！\n应用区域: {targetRange}\n现在您可以在该区域输入数字，系统会自动转换为对应的文本。", 
                        "设置完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    lblStatus.Text = $"✅ 状态: 已填充 {typeName} (区域: {targetRange})";
                    lblStatus.ForeColor = Color.Green;
                    MessageBox.Show($"数据填充完成！\n应用区域: {targetRange}\n类型: {typeName}", 
                        "填充完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                // 关闭窗口
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                if (!QuickInputManager.IsActive)
                {
                    MessageBox.Show("当前没有活动的快速录入设置", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // 确认撤销
                var result = MessageBox.Show(
                    "确认要撤销当前的快速录入设置吗？\n撤销后，数字输入将不再自动转换。",
                    "确认撤销", 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // 停止快速录入功能
                    QuickInputManager.Stop();
                    
                    // 更新状态显示
                    lblStatus.Text = "📋 状态: 已撤销快速录入";
                    lblStatus.ForeColor = Color.Orange;
                    
                    MessageBox.Show("快速录入已撤销", "撤销完成", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"撤销失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 填充数字序列 (1, 2, 3...)
        /// </summary>
        private void FillNumberSequence(string targetRange)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var worksheet = app.ActiveSheet;
                var range = worksheet.Range[targetRange];
                
                int counter = 1;
                foreach (Excel.Range cell in range.Cells)
                {
                    cell.Value = counter;
                    counter++;
                }
                
                System.Diagnostics.Debug.WriteLine($"Filled number sequence in range {targetRange}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FillNumberSequence error: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 填充字母序列 (A, B, C...)
        /// </summary>
        private void FillLetterSequence(string targetRange)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var worksheet = app.ActiveSheet;
                var range = worksheet.Range[targetRange];
                
                int counter = 0;
                foreach (Excel.Range cell in range.Cells)
                {
                    // 生成字母序列：A, B, C, ..., Z, AA, AB, AC...
                    string letter = GetColumnName(counter);
                    cell.Value = letter;
                    counter++;
                }
                
                System.Diagnostics.Debug.WriteLine($"Filled letter sequence in range {targetRange}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FillLetterSequence error: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 填充日期序列
        /// </summary>
        private void FillDateSequence(string targetRange)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var worksheet = app.ActiveSheet;
                var range = worksheet.Range[targetRange];
                
                DateTime startDate = DateTime.Today;
                int counter = 0;
                
                foreach (Excel.Range cell in range.Cells)
                {
                    DateTime currentDate = startDate.AddDays(counter);
                    cell.Value = currentDate.ToString("yyyy-MM-dd");
                    counter++;
                }
                
                System.Diagnostics.Debug.WriteLine($"Filled date sequence in range {targetRange}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FillDateSequence error: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 填充数值序列 (1.0, 2.0, 3.0...)
        /// </summary>
        private void FillNumericSequence(string targetRange)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var worksheet = app.ActiveSheet;
                var range = worksheet.Range[targetRange];
                
                double counter = 1.0;
                foreach (Excel.Range cell in range.Cells)
                {
                    cell.Value = counter;
                    counter += 1.0;
                }
                
                System.Diagnostics.Debug.WriteLine($"Filled numeric sequence in range {targetRange}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FillNumericSequence error: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 获取Excel列名 (A, B, C, ..., Z, AA, AB...)
        /// </summary>
        private string GetColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex >= 0)
            {
                columnName = (char)('A' + (columnIndex % 26)) + columnName;
                columnIndex = columnIndex / 26 - 1;
            }
            return columnName;
        }
        

        
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                // 停止选区跟踪
                StopSelectionTracking();
                
                // 释放定时器
                if (selectionTimer != null)
                {
                    selectionTimer.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnFormClosing error: {ex.Message}");
            }
            
            base.OnFormClosing(e);
        }
    }
    
    /// <summary>
    /// 快速录入管理器
    /// </summary>
    public static class QuickInputManager
    {
        private static bool _isActive = false;
        private static Dictionary<string, string> _mappings = new Dictionary<string, string>();
        private static Excel.AppEvents_SheetChangeEventHandler _changeHandler;
        private static string _targetRange = "";
        private static Excel.Range _targetRangeObject = null;
        
        public static bool IsActive => _isActive;
        
        public static void Start(Dictionary<string, string> mappings, string targetRange = "")
        {
            if (_isActive) Stop(); // 如果已经激活，先停止
            
            try
            {
                _mappings = new Dictionary<string, string>(mappings);
                _targetRange = targetRange;
                
                // 如果指定了目标区域，获取区域对象
                if (!string.IsNullOrEmpty(_targetRange))
                {
                    var app = Globals.ThisAddIn.Application;
                    var worksheet = app.ActiveSheet;
                    _targetRangeObject = worksheet.Range[_targetRange];
                }
                
                var app2 = Globals.ThisAddIn.Application;
                _changeHandler = new Excel.AppEvents_SheetChangeEventHandler(OnSheetChange);
                app2.SheetChange += _changeHandler;
                
                _isActive = true;
                
                System.Diagnostics.Debug.WriteLine($"QuickInputManager started with {_mappings.Count} mappings, target range: {_targetRange}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"QuickInputManager.Start error: {ex.Message}");
                throw;
            }
        }
        
        public static void Stop()
        {
            if (!_isActive) return;
            
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (_changeHandler != null)
                {
                    app.SheetChange -= _changeHandler;
                    _changeHandler = null;
                }
                
                _isActive = false;
                _mappings.Clear();
                _targetRange = "";
                _targetRangeObject = null;
                
                System.Diagnostics.Debug.WriteLine("QuickInputManager stopped");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"QuickInputManager.Stop error: {ex.Message}");
            }
        }
        
        private static void OnSheetChange(object sh, Excel.Range target)
        {
            try
            {
                if (!_isActive || target == null) return;
                
                // 如果指定了目标区域，检查变化的单元格是否在目标区域内
                if (_targetRangeObject != null)
                {
                    var intersection = Globals.ThisAddIn.Application.Intersect(target, _targetRangeObject);
                    if (intersection == null) return; // 不在目标区域内，忽略
                }
                
                // 获取单元格的值
                var cellValue = target.Value?.ToString();
                if (string.IsNullOrEmpty(cellValue)) return;
                
                // 检查是否有对应的映射
                if (_mappings.ContainsKey(cellValue))
                {
                    // 暂时禁用事件处理，避免递归
                    var app = Globals.ThisAddIn.Application;
                    app.EnableEvents = false;
                    
                    try
                    {
                        // 替换为映射的值
                        target.Value = _mappings[cellValue];
                        System.Diagnostics.Debug.WriteLine($"Quick input: {cellValue} -> {_mappings[cellValue]} in range {_targetRange}");
                    }
                    finally
                    {
                        // 重新启用事件处理
                        app.EnableEvents = true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnSheetChange error: {ex.Message}");
                
                // 确保事件处理被重新启用
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    app.EnableEvents = true;
                }
                catch { }
            }
        }
    }
}