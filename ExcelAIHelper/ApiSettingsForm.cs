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
            Properties.Settings.Default.ApiKey = txtApiKey.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("API Key 已保存");
            Close();
        }

        private void ApiSettingsForm_Load(object sender, EventArgs e)
        {
            txtApiKey.Text = Properties.Settings.Default.ApiKey;
        }
    }
}