using System;
using System.Windows.Forms;

namespace ExcelAIHelper
{
    public partial class ApiSettingsForm : Form
    {
        public ApiSettingsForm()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Save API key (the only setting available in the generated Settings class)
            Properties.Settings.Default.ApiKey = txtApiKey.Text;
            
            // For other settings, we'll need to implement custom storage later
            // For now, just store the API key
            
            Properties.Settings.Default.Save();
            MessageBox.Show("API Key 已保存");
            Close();
        }

        private void ApiSettingsForm_Load(object sender, EventArgs e)
        {
            // Load API key
            txtApiKey.Text = Properties.Settings.Default.ApiKey;
            
            // Set defaults for other fields
            txtApiEndpoint.Text = "https://api.deepseek.com";
            
            // Set model dropdown
            cboModel.SelectedIndex = 0; // Default to first item
            
            // Set numeric values
            nudTemperature.Value = 0.7M;
            nudMaxTokens.Value = 2048;
        }

        private async void btnTestConnection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtApiKey.Text.Trim()))
            {
                MessageBox.Show("请先输入API密钥", "测试连接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Disable the button during test
            btnTestConnection.Enabled = false;
            btnTestConnection.Text = "测试中...";

            try
            {
                // Temporarily save current settings to test
                string originalApiKey = Properties.Settings.Default.ApiKey;
                Properties.Settings.Default.ApiKey = txtApiKey.Text.Trim();

                // Create a test client
                using (var testClient = new ExcelAIHelper.Services.DeepSeekClient())
                {
                    // First test basic connectivity
                    var (basicSuccess, basicMessage) = await testClient.TestBasicConnectivityAsync();
                    
                    if (!basicSuccess)
                    {
                        MessageBox.Show($"基础网络连接失败:\n{basicMessage}\n\n请检查:\n1. 网络连接是否正常\n2. 防火墙设置\n3. 代理配置", 
                                      "网络连接失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    
                    // Then test API authentication
                    var (success, message) = await testClient.TestConnectionAsync();
                    
                    if (success)
                    {
                        MessageBox.Show($"✓ {message}\n✓ {basicMessage}", "连接测试成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"API认证测试失败:\n{message}", "连接测试失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Restore original API key
                Properties.Settings.Default.ApiKey = originalApiKey;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"测试过程中发生错误:\n{ex.Message}", "测试错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Re-enable the button
                btnTestConnection.Enabled = true;
                btnTestConnection.Text = "测试连接";
            }
        }

        private async void btnNetworkDiag_Click(object sender, EventArgs e)
        {
            btnNetworkDiag.Enabled = false;
            btnNetworkDiag.Text = "诊断中...";

            try
            {
                string diagnostics = await ExcelAIHelper.Services.NetworkDiagnostics.GetNetworkDiagnosticsAsync();
                
                // Show diagnostics in a scrollable message box
                var diagForm = new Form
                {
                    Text = "网络诊断结果",
                    Size = new System.Drawing.Size(500, 400),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.Sizable
                };

                var textBox = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    ReadOnly = true,
                    Dock = DockStyle.Fill,
                    Text = diagnostics,
                    Font = new System.Drawing.Font("Consolas", 9)
                };

                var okButton = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    Size = new System.Drawing.Size(75, 23)
                };
                okButton.Location = new System.Drawing.Point(diagForm.Width - 95, diagForm.Height - 55);

                diagForm.Controls.Add(textBox);
                diagForm.Controls.Add(okButton);
                diagForm.AcceptButton = okButton;

                diagForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"网络诊断失败:\n{ex.Message}", "诊断错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnNetworkDiag.Enabled = true;
                btnNetworkDiag.Text = "网络诊断";
            }
        }
    }
}