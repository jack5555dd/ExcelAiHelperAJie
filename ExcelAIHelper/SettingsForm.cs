using System;
using System.Configuration;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace ExcelAIHelper
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            try
            {
                // 加载API设置
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                
                // 服务商
                string provider = GetConfigValue("ApiProvider", "DeepSeek");
                if (cmbProvider.Items.Contains(provider))
                {
                    cmbProvider.SelectedItem = provider;
                }
                else
                {
                    cmbProvider.SelectedIndex = 0;
                }
                
                // 模型
                txtModel.Text = GetConfigValue("ApiModel", "deepseek-chat");
                
                // API密钥
                txtApiKey.Text = GetConfigValue("ApiKey", "");
                
                // API地址
                txtBaseUrl.Text = GetConfigValue("ApiBaseUrl", "https://api.deepseek.com");
                
                // 超时时间
                if (int.TryParse(GetConfigValue("ApiTimeout", "15"), out int timeout))
                {
                    numTimeout.Value = Math.Max(5, Math.Min(300, timeout));
                }
                
                // 加载通用设置
                string language = GetConfigValue("Language", "简体中文");
                if (cmbLanguage.Items.Contains(language))
                {
                    cmbLanguage.SelectedItem = language;
                }
                else
                {
                    cmbLanguage.SelectedIndex = 0;
                }
                
                chkAutoSave.Checked = bool.Parse(GetConfigValue("AutoSave", "true"));
                chkShowTips.Checked = bool.Parse(GetConfigValue("ShowTips", "true"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载设置时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetConfigValue(string key, string defaultValue)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var setting = config.AppSettings.Settings[key];
                return setting?.Value ?? defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        private void SaveSettings()
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                
                // 保存API设置
                SetConfigValue(config, "ApiProvider", cmbProvider.SelectedItem?.ToString() ?? "DeepSeek");
                SetConfigValue(config, "ApiModel", txtModel.Text.Trim());
                SetConfigValue(config, "ApiKey", txtApiKey.Text.Trim());
                SetConfigValue(config, "ApiBaseUrl", txtBaseUrl.Text.Trim());
                SetConfigValue(config, "ApiTimeout", numTimeout.Value.ToString());
                
                // 保存通用设置
                SetConfigValue(config, "Language", cmbLanguage.SelectedItem?.ToString() ?? "简体中文");
                SetConfigValue(config, "AutoSave", chkAutoSave.Checked.ToString());
                SetConfigValue(config, "ShowTips", chkShowTips.Checked.ToString());
                
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
                
                MessageBox.Show("设置已保存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetConfigValue(Configuration config, string key, string value)
        {
            if (config.AppSettings.Settings[key] == null)
            {
                config.AppSettings.Settings.Add(key, value);
            }
            else
            {
                config.AppSettings.Settings[key].Value = value;
            }
        }

        private async void btnTestConnection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtApiKey.Text))
            {
                MessageBox.Show("请先输入API密钥", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtBaseUrl.Text))
            {
                MessageBox.Show("请先输入API地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnTestConnection.Enabled = false;
            btnTestConnection.Text = "测试中...";

            try
            {
                await TestApiConnection();
                MessageBox.Show("API连接测试成功！", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"API连接测试失败: {ex.Message}", "测试结果", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnTestConnection.Enabled = true;
                btnTestConnection.Text = "测试连接";
            }
        }

        private async Task TestApiConnection()
        {
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds((double)numTimeout.Value);
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {txtApiKey.Text.Trim()}");
                client.DefaultRequestHeaders.Add("User-Agent", "ExcelAIHelper/1.0");

                var requestData = new
                {
                    model = txtModel.Text.Trim(),
                    messages = new[]
                    {
                        new { role = "user", content = "Hello" }
                    },
                    max_tokens = 10,
                    temperature = 0.1
                };

                var json = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string apiUrl = txtBaseUrl.Text.Trim().TrimEnd('/') + "/v1/chat/completions";
                var response = await client.PostAsync(apiUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"HTTP {(int)response.StatusCode}: {errorContent}");
                }

                var responseContent = await response.Content.ReadAsStringAsync();
                var responseObj = JsonConvert.DeserializeObject<dynamic>(responseContent);
                
                if (responseObj?.choices == null || responseObj.choices.Count == 0)
                {
                    throw new Exception("API响应格式不正确");
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (ValidateSettings())
            {
                SaveSettings();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            if (ValidateSettings())
            {
                SaveSettings();
            }
        }

        private bool ValidateSettings()
        {
            // 验证API设置
            if (string.IsNullOrWhiteSpace(txtModel.Text))
            {
                MessageBox.Show("请输入模型名称", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabPageApi;
                txtModel.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtApiKey.Text))
            {
                MessageBox.Show("请输入API密钥", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabPageApi;
                txtApiKey.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtBaseUrl.Text))
            {
                MessageBox.Show("请输入API地址", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabPageApi;
                txtBaseUrl.Focus();
                return false;
            }

            // 验证URL格式
            if (!Uri.TryCreate(txtBaseUrl.Text.Trim(), UriKind.Absolute, out Uri uri) || 
                (uri.Scheme != "http" && uri.Scheme != "https"))
            {
                MessageBox.Show("API地址格式不正确，请输入有效的HTTP或HTTPS地址", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabPageApi;
                txtBaseUrl.Focus();
                return false;
            }

            return true;
        }
    }
}