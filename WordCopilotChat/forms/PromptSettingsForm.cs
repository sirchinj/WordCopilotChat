using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordCopilotChat.services;
using WordCopilotChat.utils;
using System.IO;
using Newtonsoft.Json;

namespace WordCopilotChat
{
    public partial class PromptSettingsForm : Form
    {
        private bool _hasChanges = false;

        public PromptSettingsForm()
        {
            InitializeComponent();
            
            // 设置TextBox属性以优化文本显示
            ConfigureTextBoxes();
            
            // 订阅Tab切换事件
            tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged;
            
            LoadPrompts();
            
            // 初始化Token估算显示
            UpdateTokenEstimate();
            
            // 调试信息：检查文本内容
            System.Diagnostics.Debug.WriteLine("=== TextBox配置完成 ===");
            System.Diagnostics.Debug.WriteLine($"txtChatPrompt.WordWrap: {txtChatPrompt.WordWrap}");
            System.Diagnostics.Debug.WriteLine($"txtChatPrompt.Multiline: {txtChatPrompt.Multiline}");
            System.Diagnostics.Debug.WriteLine($"txtChatPrompt.AcceptsReturn: {txtChatPrompt.AcceptsReturn}");
            System.Diagnostics.Debug.WriteLine($"文本长度: {txtChatPrompt.Text.Length}");
            System.Diagnostics.Debug.WriteLine($"包含\\r\\n: {txtChatPrompt.Text.Contains("\r\n")}");
            System.Diagnostics.Debug.WriteLine($"包含\\n: {txtChatPrompt.Text.Contains("\n")}");
            System.Diagnostics.Debug.WriteLine("=== 数据库状态检查 ===");
            System.Diagnostics.Debug.WriteLine($"PromptService可用模式: {string.Join(", ", PromptService.GetAvailableModes())}");
            System.Diagnostics.Debug.WriteLine("=== 初始化完成 ===");
        }
        
        private void ConfigureTextBoxes()
        {
            // 配置所有TextBox的显示属性
            ConfigureTextBox(txtChatPrompt);
            ConfigureTextBox(txtAgentPrompt);
            ConfigureTextBox(txtWelcomePrompt);
            ConfigureTextBox(txtCompressPrompt);
        }
        
        private void ConfigureTextBox(TextBox textBox)
        {
            // 确保多行模式
            textBox.Multiline = true;
            
            // 启用自动换行
            textBox.WordWrap = true;
            
            // 设置滚动条
            textBox.ScrollBars = ScrollBars.Vertical;
            
            // 设置字体以提高可读性
            textBox.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            
            // 设置选择时的行为
            textBox.HideSelection = false;
            
            // 设置文本对齐方式
            textBox.TextAlign = HorizontalAlignment.Left;
            
            // 确保接受Tab键和换行符
            textBox.AcceptsTab = true;
            textBox.AcceptsReturn = true;
            
            // 设置右键菜单
            textBox.ShortcutsEnabled = true;
        }
        
        /// <summary>
        /// 估算文本的Token数（简单估算：中文约1.5字符/token，英文约4字符/token）
        /// </summary>
        private int EstimateTokenCount(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;
                
            int chineseCount = 0;
            int otherCount = 0;
            
            foreach (char c in text)
            {
                if (c >= 0x4E00 && c <= 0x9FFF)
                    chineseCount++;
                else
                    otherCount++;
            }
            
            return Math.Max(1, (int)(chineseCount / 1.5) + (int)(otherCount / 4.0));
        }
        
        /// <summary>
        /// 更新当前Tab的Token估算显示
        /// </summary>
        private void UpdateTokenEstimate()
        {
            string content = GetCurrentContent();
            int tokenCount = EstimateTokenCount(content);
            
            // 更新对应Tab的描述标签
            if (tabControl1.SelectedTab == tabPageChat)
            {
                lblChatDesc.Text = $"智能问答模式用于普通对话。该模式下AI不会主动调用Word工具。\n【Token估算: ~{tokenCount} tokens】";
            }
            else if (tabControl1.SelectedTab == tabPageAgent)
            {
                lblAgentDesc.Text = $"智能体模式具备智能代理能力，可以主动调用Word工具完成复杂文档操作。\n【Token估算: ~{tokenCount} tokens】";
            }
            else if (tabControl1.SelectedTab == tabPageWelcome)
            {
                // 欢迎页内容不参与对话请求，不需要 Token 估算提示
                lblWelcomeDesc.Text = "欢迎页在用户首次打开聊天界面时显示，支持Markdown格式。";
            }
            else if (tabControl1.SelectedTab == tabPageCompress)
            {
                lblCompressDesc.Text = $"当对话历史Token占用过高时，系统会自动调用此提示词进行压缩总结。\n【Token估算: ~{tokenCount} tokens】";
            }
        }

        private void LoadPrompts()
        {
            try
            {
                // 加载智能问答模式提示词并转换换行符
                txtChatPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("chat"));
                
                // 加载智能体模式提示词并转换换行符
                txtAgentPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("chat-agent"));
                
                // 加载欢迎页提示词并转换换行符
                txtWelcomePrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("welcome"));
                
                // 加载上下文压缩提示词并转换换行符
                txtCompressPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("compress-context"));
                
                // 强制刷新显示
                RefreshTextBoxes();
                
                _hasChanges = false;
                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载提示词失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 标准化换行符为Windows格式
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>标准化后的文本</returns>
        private string NormalizeLineEndings(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
                
            // 先统一为\n，然后转换为\r\n
            return text.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        /// <summary>
        /// 保存前预处理：
        /// - 统一换行符（入库使用 \n，避免跨平台差异）；
        /// - 移除不可见/控制字符；
        /// - 将智能引号“”统一为标准双引号"；
        /// - 保护占位符 ${{docs}} 不被破坏；
        /// - 不对普通双引号做重复转义（数据库存原始语义，运行期安全）。
        /// </summary>
        private string PrepareForSave(string text)
        {
            if (text == null)
                return string.Empty;

            string result = text;
            
            // 1) 统一换行：入库使用 \n，读取时再标准化为 Windows 的 \r\n
            result = result.Replace("\r\n", "\n").Replace("\r", "\n");

            // 2) 移除常见不可见字符（零宽空格、BOM、NULL）
            result = result
                .Replace("\uFEFF", string.Empty)   // BOM
                .Replace("\u200B", string.Empty)   // ZERO WIDTH SPACE
                .Replace("\0", string.Empty);      // NULL

            // 3) 统一智能引号为标准引号
            result = result
                .Replace('\u201C'.ToString(), "\"") // “
                .Replace('\u201D'.ToString(), "\""); // ”

            // 4) 保护占位符：保持原样（这里不做额外处理，仅作为显式注释）
            // ${{docs}} 将保持不变

            // 5) 不对双引号做 "" 形式的重复转义：数据库与运行期都按原义存储/使用

            return result;
        }

        /// <summary>
        /// 强制刷新所有TextBox的显示
        /// </summary>
        private void RefreshTextBoxes()
        {
            txtChatPrompt.Invalidate();
            txtChatPrompt.Refresh();
            
            txtAgentPrompt.Invalidate();
            txtAgentPrompt.Refresh();
            
            txtWelcomePrompt.Invalidate();
            txtWelcomePrompt.Refresh();
            
            txtCompressPrompt.Invalidate();
            txtCompressPrompt.Refresh();
        }

        private void SavePrompts()
        {
            try
            {
                // 仅保存当前Tab
                var current = tabControl1.SelectedTab;
                if (current == tabPageChat)
                {
                    var chatToSave = PrepareForSave(txtChatPrompt.Text);
                    PromptService.SetPrompt("chat", chatToSave);
                }
                else if (current == tabPageAgent)
                {
                    var agentToSave = PrepareForSave(txtAgentPrompt.Text);
                    PromptService.SetPrompt("chat-agent", agentToSave);
                }
                else if (current == tabPageWelcome)
                {
                    var welcomeToSave = PrepareForSave(txtWelcomePrompt.Text);
                    PromptService.SetPrompt("welcome", welcomeToSave);
                }
                else if (current == tabPageCompress)
                {
                    var compressToSave = PrepareForSave(txtCompressPrompt.Text);
                    PromptService.SetPrompt("compress-context", compressToSave);
                }
                
                // 刷新PromptService缓存
                PromptService.RefreshCache();
                
                _hasChanges = false;
                UpdateButtonStates();
                
                MessageBox.Show("当前页提示词保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存提示词失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetToDefaults()
        {
            var current = tabControl1.SelectedTab;
            string targetType = current == tabPageChat ? "chat" 
                : current == tabPageAgent ? "chat-agent" 
                : current == tabPageCompress ? "compress-context"
                : "welcome";
            var result = MessageBox.Show($"确定要重置当前页（{targetType}）提示词为默认值吗？此操作不可撤销。",
                "确认重置", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                try
                {
                    PromptService.ResetToDefault(targetType);
                    // 仅刷新当前页内容
                    if (targetType == "chat")
                        txtChatPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("chat"));
                    else if (targetType == "chat-agent")
                        txtAgentPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("chat-agent"));
                    else if (targetType == "compress-context")
                        txtCompressPrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("compress-context"));
                    else
                        txtWelcomePrompt.Text = NormalizeLineEndings(PromptService.GetPrompt("welcome"));
                    _hasChanges = false;
                    UpdateButtonStates();
                    MessageBox.Show("当前页提示词已重置为默认值", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"重置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpdateButtonStates()
        {
            btnSave.Enabled = _hasChanges;
            btnReset.Enabled = true;
        }

        private void OnTextChanged(object sender, EventArgs e)
        {
            _hasChanges = true;
            UpdateButtonStates();
            UpdateTokenEstimate();
        }
        
        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateTokenEstimate();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SavePrompts();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            ResetToDefaults();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_hasChanges)
            {
                var result = MessageBox.Show("存在未保存的更改，确定要关闭吗？", 
                    "确认关闭", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                if (result == DialogResult.No)
                    return;
            }
            
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (_hasChanges)
            {
                SavePrompts();
            }
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // ============ 导出/导入（当前Tab） ============
        private string GetCurrentType()
        {
            if (tabControl1.SelectedTab == tabPageChat) return "chat";
            if (tabControl1.SelectedTab == tabPageAgent) return "chat-agent";
            if (tabControl1.SelectedTab == tabPageCompress) return "compress-context";
            return "welcome";
        }

        private string GetCurrentContent()
        {
            if (tabControl1.SelectedTab == tabPageChat) return txtChatPrompt.Text;
            if (tabControl1.SelectedTab == tabPageAgent) return txtAgentPrompt.Text;
            if (tabControl1.SelectedTab == tabPageCompress) return txtCompressPrompt.Text;
            return txtWelcomePrompt.Text;
        }

        private void SetCurrentContent(string text)
        {
            if (tabControl1.SelectedTab == tabPageChat) txtChatPrompt.Text = text;
            else if (tabControl1.SelectedTab == tabPageAgent) txtAgentPrompt.Text = text;
            else if (tabControl1.SelectedTab == tabPageCompress) txtCompressPrompt.Text = text;
            else txtWelcomePrompt.Text = text;
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            var type = GetCurrentType();
            // 获取导出密码
            string password = PasswordDialog.ShowExportDialog(this);
            if (string.IsNullOrEmpty(password)) return;

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "WordCopilot提示词|*.wcp";
                saveFileDialog.Title = "导出提示词";
                saveFileDialog.FileName = $"WordCopilot_Prompt_{type}_{DateTime.Now:yyyyMMdd_HHmmss}.wcp";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var exportData = PromptService.BuildExportData(type);
                        var json = JsonConvert.SerializeObject(exportData, Formatting.Indented);
                        var encrypted = EncryptionUtils.Encrypt(json, password);
                        File.WriteAllText(saveFileDialog.FileName, encrypted);
                        MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            var type = GetCurrentType();
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "WordCopilot提示词|*.wcp|所有文件|*.*";
                openFileDialog.Title = "导入提示词";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string password = PasswordDialog.ShowImportDialog(this);
                        if (string.IsNullOrEmpty(password)) return;

                        string encrypted = File.ReadAllText(openFileDialog.FileName);
                        string json;
                        try
                        {
                            json = EncryptionUtils.Decrypt(encrypted, password);
                        }
                        catch (System.Security.Cryptography.CryptographicException)
                        {
                            MessageBox.Show("密码错误，无法解密文件", "解密失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        var data = JsonConvert.DeserializeObject<PromptExportData>(json);
                        if (data?.Items == null || data.Items.Count == 0)
                        {
                            MessageBox.Show("文件中没有有效提示词", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // 找到与当前类型匹配的条目
                        var item = data.Items.Find(i => string.Equals(i.PromptType, type, StringComparison.OrdinalIgnoreCase));
                        if (item == null)
                        {
                            MessageBox.Show("文件中不包含当前页的提示词类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // 应用并刷新
                        if (PromptService.ApplyImportItem(item))
                        {
                            SetCurrentContent(NormalizeLineEndings(item.PromptContent));
                            _hasChanges = false;
                            UpdateButtonStates();
                            MessageBox.Show("导入并保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("导入失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_hasChanges && e.CloseReason == CloseReason.UserClosing)
            {
                var result = MessageBox.Show("存在未保存的更改，确定要关闭吗？", 
                    "确认关闭", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }
            
            base.OnFormClosing(e);
        }

        private void BtnPreviewChat_Click(object sender, EventArgs e)
        {
            ShowPromptPreview("智能问答模式提示词", txtChatPrompt.Text);
        }

        private void BtnPreviewAgent_Click(object sender, EventArgs e)
        {
            ShowPromptPreview("智能体模式提示词", txtAgentPrompt.Text);
        }

        private void BtnPreviewWelcome_Click(object sender, EventArgs e)
        {
            ShowPromptPreview("欢迎页内容", txtWelcomePrompt.Text);
        }

        private void BtnPreviewCompress_Click(object sender, EventArgs e)
        {
            ShowPromptPreview("上下文压缩提示词", txtCompressPrompt.Text);
        }

        private void ShowPromptPreview(string title, string content)
        {
            var previewForm = new Form
            {
                Text = title,
                Size = new Size(600, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow
            };

            var textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                WordWrap = true,  // 启用自动换行
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = NormalizeLineEndings(content),  // 标准化换行符
                Font = new Font("Microsoft YaHei", 9F),
                HideSelection = false  // 保持选择高亮显示
            };

            previewForm.Controls.Add(textBox);
            previewForm.ShowDialog(this);
        }


    }
} 