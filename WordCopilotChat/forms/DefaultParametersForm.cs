using System;
using System.Drawing;
using System.Windows.Forms;
using WordCopilotChat.services;

namespace WordCopilotChat
{
    public partial class DefaultParametersForm : Form
    {
        private AppSettingsService _appSettingsService;

        public DefaultParametersForm(AppSettingsService appSettingsService)
        {
            _appSettingsService = appSettingsService;
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            try
            {
                // 加载通用默认参数
                nudDefaultTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("default_temperature", 0.7);
                var defaultMaxTokens = _appSettingsService.GetSetting("default_max_tokens");
                if (!string.IsNullOrEmpty(defaultMaxTokens) && int.TryParse(defaultMaxTokens, out int defaultMaxTokensValue))
                {
                    nudDefaultMaxTokens.Value = defaultMaxTokensValue;
                }
                else
                {
                    nudDefaultMaxTokens.Value = nudDefaultMaxTokens.Minimum; // 留空显示为最小值
                }
                nudDefaultTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("default_top_p", 0.9);
                
                // 加载Chat模式参数（创意写作）
                nudChatTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("chat_temperature", 0.7);
                var chatMaxTokens = _appSettingsService.GetSetting("chat_max_tokens");
                if (!string.IsNullOrEmpty(chatMaxTokens) && int.TryParse(chatMaxTokens, out int chatMaxTokensValue))
                {
                    nudChatMaxTokens.Value = chatMaxTokensValue;
                }
                else
                {
                    nudChatMaxTokens.Value = nudChatMaxTokens.Minimum;
                }
                nudChatTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("chat_top_p", 0.9);
                
                // 加载Agent模式参数（严格工具调用）
                nudAgentTemp.Value = (decimal)_appSettingsService.GetDoubleSetting("agent_temperature", 0.3);
                var agentMaxTokens = _appSettingsService.GetSetting("agent_max_tokens");
                if (!string.IsNullOrEmpty(agentMaxTokens) && int.TryParse(agentMaxTokens, out int agentMaxTokensValue))
                {
                    nudAgentMaxTokens.Value = agentMaxTokensValue;
                }
                else
                {
                    nudAgentMaxTokens.Value = nudAgentMaxTokens.Minimum;
                }
                nudAgentTopP.Value = (decimal)_appSettingsService.GetDoubleSetting("agent_top_p", 0.85);

                // 加载自动压缩阈值（百分比）
                var compressPct = _appSettingsService.GetIntSetting("context_compress_threshold_pct", 90);
                if (compressPct < 50) compressPct = 50;
                if (compressPct > 95) compressPct = 95;
                nudContextCompressThreshold.Value = compressPct;
                
                System.Diagnostics.Debug.WriteLine("默认参数设置已加载");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载当前设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 保存通用默认参数
                _appSettingsService.UpdateSetting("default_temperature", nudDefaultTemp.Value.ToString());
                // max_tokens为null时保存空字符串，由服务商自动处理
                _appSettingsService.UpdateSetting("default_max_tokens", 
                    nudDefaultMaxTokens.Value == nudDefaultMaxTokens.Minimum ? "" : nudDefaultMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("default_top_p", nudDefaultTopP.Value.ToString());
                
                // 保存Chat模式参数
                _appSettingsService.UpdateSetting("chat_temperature", nudChatTemp.Value.ToString());
                _appSettingsService.UpdateSetting("chat_max_tokens", 
                    nudChatMaxTokens.Value == nudChatMaxTokens.Minimum ? "" : nudChatMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("chat_top_p", nudChatTopP.Value.ToString());
                
                // 保存Agent模式参数
                _appSettingsService.UpdateSetting("agent_temperature", nudAgentTemp.Value.ToString());
                _appSettingsService.UpdateSetting("agent_max_tokens", 
                    nudAgentMaxTokens.Value == nudAgentMaxTokens.Minimum ? "" : nudAgentMaxTokens.Value.ToString());
                _appSettingsService.UpdateSetting("agent_top_p", nudAgentTopP.Value.ToString());

                // 保存自动压缩阈值（百分比）
                _appSettingsService.UpdateSetting("context_compress_threshold_pct", nudContextCompressThreshold.Value.ToString());
                
                System.Diagnostics.Debug.WriteLine("默认参数设置已保存到数据库");
                MessageBox.Show("参数设置已保存成功！\n\n提示：max_tokens留空时将由AI服务商自动处理。", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "确定要重置为系统推荐值吗？\n\n" +
                "Chat模式（创意写作）：温度0.7，适合内容生成\n" +
                "Agent模式（工具调用）：温度0.3，确保准确执行\n" +
                "max_tokens：留空，由AI服务商自动处理", 
                "确认重置", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                // 重置为系统推荐值（优化后的参数）
                nudDefaultTemp.Value = 0.7m;
                nudDefaultMaxTokens.Value = nudDefaultMaxTokens.Minimum; // 留空
                nudDefaultTopP.Value = 0.9m;
                
                // Chat模式：适合创意写作
                nudChatTemp.Value = 0.7m;
                nudChatMaxTokens.Value = nudChatMaxTokens.Minimum; // 留空
                nudChatTopP.Value = 0.9m;
                
                // Agent模式：严格遵守提示词，准确调用工具
                nudAgentTemp.Value = 0.3m;
                nudAgentMaxTokens.Value = nudAgentMaxTokens.Minimum; // 留空
                nudAgentTopP.Value = 0.85m;

                // 自动压缩阈值（默认90%）
                nudContextCompressThreshold.Value = 90;
                
                MessageBox.Show("已重置为系统推荐参数！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
} 