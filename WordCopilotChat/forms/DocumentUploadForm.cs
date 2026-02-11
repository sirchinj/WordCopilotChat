using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordCopilotChat.models;
using WordCopilotChat.services;
using WordCopilotChat.utils;

namespace WordCopilotChat
{
    public partial class DocumentUploadForm : Form
    {
        private DocumentService _documentService;
        private string _selectedFilePath;
        private List<DocumentHeading> _parsedHeadings;

        public DocumentUploadForm()
        {
            InitializeComponent();
            _documentService = new DocumentService();
        }

        private void DocumentUploadForm_Load(object sender, EventArgs e)
        {
            UpdateUI();
        }

        /// <summary>
        /// 更新界面状态
        /// </summary>
        private void UpdateUI()
        {
            bool hasFile = !string.IsNullOrEmpty(_selectedFilePath);
            buttonParse.Enabled = hasFile;
            buttonUpload.Enabled = hasFile && _parsedHeadings != null && _parsedHeadings.Count > 0;

            if (hasFile)
            {
                labelSelectedFile.Text = Path.GetFileName(_selectedFilePath);
                labelFileSize.Text = $"大小: {GetFileSizeText(new FileInfo(_selectedFilePath).Length)}";
            }
            else
            {
                labelSelectedFile.Text = "未选择文件";
                labelFileSize.Text = "";
            }

            // 更新标题数量显示
            if (_parsedHeadings != null)
            {
                this.Text = $"上传文档 - 已解析 {_parsedHeadings.Count} 个标题";
            }
            else
            {
                this.Text = "上传文档";
            }
        }

        /// <summary>
        /// 获取文件大小文本
        /// </summary>
        private string GetFileSizeText(long bytes)
        {
            if (bytes < 1024)
                return $"{bytes} B";
            else if (bytes < 1024 * 1024)
                return $"{bytes / 1024:F1} KB";
            else
                return $"{bytes / (1024 * 1024):F1} MB";
        }

        #region 事件处理

        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "选择要上传的文档";
                openFileDialog.Filter = "支持的文档格式|*.docx;*.doc;*.md|Word文档 (*.docx;*.doc)|*.docx;*.doc|Markdown文档 (*.md)|*.md|所有文件|*.*";
                openFileDialog.FilterIndex = 1;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedFilePath = openFileDialog.FileName;
                    
                    // 验证文件格式
                    if (!DocumentParser.IsSupportedFileType(_selectedFilePath))
                    {
                        MessageBox.Show("不支持的文件格式！请选择 .docx、.doc 或 .md 文件。",
                            "格式错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _selectedFilePath = null;
                        return;
                    }

                    // 检查文件是否已存在
                    string fileName = Path.GetFileName(_selectedFilePath);
                    if (_documentService.IsDocumentExists(fileName))
                    {
                        var result = MessageBox.Show($"文档 \"{fileName}\" 已存在，是否要替换？",
                            "文档已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        
                        if (result == DialogResult.No)
                        {
                            _selectedFilePath = null;
                            UpdateUI();
                            return;
                        }
                    }

                    // 重置解析结果
                    _parsedHeadings = null;
                    listViewHeadings.Items.Clear();
                    labelParseResult.Text = "";

                    UpdateUI();
                }
            }
        }

        private async void buttonParse_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedFilePath)) return;

            buttonParse.Enabled = false;
            progressBar.Visible = true;
            textBoxParseLog.Visible = true;
            textBoxParseLog.Clear();
            progressBar.Style = ProgressBarStyle.Marquee;
            labelParseResult.Text = "正在解析文档...";

            try
            {
                // 在UI线程上显示开始信息
                labelParseResult.Text = "开始解析文档，请稍候...";
                LogParseProgress("开始解析文档...");
                Application.DoEvents(); // 强制更新UI

                // 异步解析文档
                await Task.Run(() =>
                {
                    string fileType = DocumentParser.GetFileType(_selectedFilePath);
                    DocumentParseResult result;

                    // 在UI线程上更新状态
                    this.Invoke((MethodInvoker)delegate {
                        labelParseResult.Text = $"正在解析 {fileType.ToUpper()} 文档...";
                        LogParseProgress($"文档类型: {fileType.ToUpper()}");
                        LogParseProgress($"文件路径: {Path.GetFileName(_selectedFilePath)}");
                    });

                    if (fileType == "md")
                    {
                        this.Invoke((MethodInvoker)delegate {
                            LogParseProgress("开始解析Markdown文档...");
                        });
                        result = DocumentParser.ParseMarkdownDocument(_selectedFilePath);
                    }
                    else
                    {
                        this.Invoke((MethodInvoker)delegate {
                            LogParseProgress("开始解析Word文档...");
                            LogParseProgress("正在打开Word应用程序...");
                        });
                        
                        // 使用自定义的解析方法，带进度回调
                        result = ParseWordDocumentWithProgress(_selectedFilePath);
                    }

                    if (result.Success)
                    {
                        _parsedHeadings = result.Headings;
                        
                        // 在UI线程上更新进度
                        this.Invoke((MethodInvoker)delegate {
                            labelParseResult.Text = $"解析完成，找到 {_parsedHeadings.Count} 个标题，正在构建预览...";
                            LogParseProgress($"解析完成！共找到 {_parsedHeadings.Count} 个标题");
                            LogParseProgress("正在构建预览界面...");
                        });
                    }
                    else
                    {
                        throw new Exception(result.Message);
                    }
                });

                // 显示解析结果
                DisplayParseResult();
                labelParseResult.Text = $"解析完成！找到 {_parsedHeadings.Count} 个标题。";
                labelParseResult.ForeColor = Color.Green;
                LogParseProgress("预览界面构建完成！");

                // 如果没有找到标题，给出提示
                if (_parsedHeadings.Count == 0)
                {
                    labelParseResult.Text = "未找到标题！请确保文档有明显的标题格式。";
                    labelParseResult.ForeColor = Color.Orange;
                    LogParseProgress("警告：未找到任何标题！");
                    
                    MessageBox.Show("未在文档中找到标题！\n\n系统会尝试以下方式识别标题：\n" +
                                  "1. Word标题样式（标题1、标题2等）\n" +
                                  "2. 字体较大的文本（比正文大1.5磅以上）\n" +
                                  "3. Markdown标题标记（#、##等）\n\n" +
                                  "请确保文档中的标题有明显的格式特征。", 
                                  "解析结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                labelParseResult.Text = $"解析失败: {ex.Message}";
                labelParseResult.ForeColor = Color.Red;
                _parsedHeadings = null;
                LogParseProgress($"错误: {ex.Message}");
                
                // 显示详细错误信息
                MessageBox.Show($"文档解析失败：\n\n{ex.Message}\n\n请检查：\n" +
                              "• 文件是否损坏\n" +
                              "• 是否安装了Microsoft Office\n" +
                              "• 文件是否被其他程序占用", 
                              "解析错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                buttonParse.Enabled = true;
                UpdateUI();
            }
        }

        private async void buttonUpload_Click(object sender, EventArgs e)
        {
            if (_parsedHeadings == null || _parsedHeadings.Count == 0)
            {
                MessageBox.Show("请先解析文档！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 检查文档数量限制
            if (_documentService.IsDocumentLimitReached())
            {
                MessageBox.Show("已达到文档数量限制，请先删除一些文档或增加限制。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            buttonUpload.Enabled = false;
            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Marquee;
            labelParseResult.Text = "正在上传文档...";

            try
            {
                await Task.Run(() =>
                {
                    // 创建文档记录
                    var fileInfo = new FileInfo(_selectedFilePath);
                    var document = new Document
                    {
                        FileName = fileInfo.Name,
                        FilePath = _selectedFilePath,
                        FileType = DocumentParser.GetFileType(_selectedFilePath),
                        FileSize = fileInfo.Length,
                        TotalHeadings = _parsedHeadings.Count
                    };

                    // 保存文档
                    if (!_documentService.AddDocument(document))
                    {
                        throw new Exception("保存文档信息失败");
                    }

                    // 获取刚保存的文档ID
                    var savedDoc = _documentService.GetAllDocuments()
                        .FirstOrDefault(d => d.FileName == document.FileName);
                    
                    if (savedDoc == null)
                    {
                        throw new Exception("无法获取已保存的文档信息");
                    }

                    // 设置文档ID并保存标题
                    foreach (var heading in _parsedHeadings)
                    {
                        heading.DocumentId = savedDoc.Id;
                    }

                    if (!_documentService.AddDocumentHeadings(_parsedHeadings))
                    {
                        throw new Exception("保存文档标题失败");
                    }
                });

                labelParseResult.Text = "文档上传成功！";
                labelParseResult.ForeColor = Color.Green;
                
                MessageBox.Show("文档上传成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                labelParseResult.Text = $"上传失败: {ex.Message}";
                labelParseResult.ForeColor = Color.Red;
                MessageBox.Show($"上传文档失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                buttonUpload.Enabled = true;
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // 右键菜单事件处理
        private void toolStripMenuItemViewContent_Click(object sender, EventArgs e)
        {
            if (listViewHeadings.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先选择一个标题！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            var selectedItem = listViewHeadings.SelectedItems[0];
            if (selectedItem.Tag is DocumentHeading heading)
            {
                ShowHeadingContent(heading);
            }
        }
        
        private void toolStripMenuItemRemove_Click(object sender, EventArgs e)
        {
            RemoveSelectedHeadings();
        }

        private void toolStripMenuItemSelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewHeadings.Items)
            {
                item.Selected = true;
            }
        }

        private void toolStripMenuItemUnselectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewHeadings.Items)
            {
                item.Selected = false;
            }
        }
        
        // 双击标题查看内容
        private void listViewHeadings_DoubleClick(object sender, EventArgs e)
        {
            if (listViewHeadings.SelectedItems.Count == 0)
                return;
            
            var selectedItem = listViewHeadings.SelectedItems[0];
            if (selectedItem.Tag is DocumentHeading heading)
            {
                ShowHeadingContent(heading);
            }
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 记录解析进度日志
        /// </summary>
        private void LogParseProgress(string message)
        {
            if (InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate { LogParseProgress(message); });
                return;
            }

            string timeStamp = DateTime.Now.ToString("HH:mm:ss");
            textBoxParseLog.AppendText($"[{timeStamp}] {message}\r\n");
            textBoxParseLog.ScrollToCaret();
            Application.DoEvents();
        }

                 /// <summary>
         /// 带进度回调的Word文档解析
         /// </summary>
         private DocumentParseResult ParseWordDocumentWithProgress(string filePath)
         {
             var result = new DocumentParseResult();
             
             try
             {
                 // 调用带进度回调的解析方法
                 result = DocumentParser.ParseWordDocument(filePath, quickMode: false, progressCallback: (message) =>
                 {
                     this.Invoke((MethodInvoker)delegate {
                         LogParseProgress(message);
                     });
                 });

                 this.Invoke((MethodInvoker)delegate {
                     LogParseProgress("Word文档解析完成");
                 });
             }
             catch (Exception ex)
             {
                 this.Invoke((MethodInvoker)delegate {
                     LogParseProgress($"解析失败: {ex.Message}");
                 });
                 result.Success = false;
                 result.Message = ex.Message;
             }

             return result;
         }

        /// <summary>
        /// 移除选中的标题
        /// </summary>
        private void RemoveSelectedHeadings()
        {
            if (listViewHeadings.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先选择要剔除的标题！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 收集要删除的标题（包括子标题）
            var selectedHeadings = new List<DocumentHeading>();
            foreach (ListViewItem selectedItem in listViewHeadings.SelectedItems)
            {
                if (selectedItem.Tag is DocumentHeading heading)
                {
                    selectedHeadings.Add(heading);
                }
            }

            // 递归收集所有子标题
            var allHeadingsToRemove = new HashSet<DocumentHeading>();
            foreach (var heading in selectedHeadings)
            {
                CollectHeadingAndChildren(heading, allHeadingsToRemove);
            }

            var totalCount = allHeadingsToRemove.Count;
            var childCount = totalCount - selectedHeadings.Count;

            string message = $"确定要剔除选中的 {selectedHeadings.Count} 个标题吗？";
            if (childCount > 0)
            {
                message += $"\n\n注意：这将同时剔除 {childCount} 个子标题，总共 {totalCount} 个标题。";
            }
            message += "\n\n被剔除的标题将不会被上传到系统中。";

            var result = MessageBox.Show(message, "确认剔除", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // 收集要删除的界面项
                var itemsToRemove = new List<ListViewItem>();
                foreach (ListViewItem item in listViewHeadings.Items)
                {
                    if (item.Tag is DocumentHeading heading && allHeadingsToRemove.Contains(heading))
                    {
                        itemsToRemove.Add(item);
                    }
                }

                // 从界面中移除
                foreach (var item in itemsToRemove)
                {
                    listViewHeadings.Items.Remove(item);
                }

                // 从数据中移除
                foreach (var heading in allHeadingsToRemove)
                {
                    _parsedHeadings.Remove(heading);
                }

                // 重新排序
                for (int i = 0; i < _parsedHeadings.Count; i++)
                {
                    _parsedHeadings[i].OrderIndex = i + 1;
                }

                // 更新界面状态
                UpdateUI();
                MessageBox.Show($"已成功剔除 {totalCount} 个标题（包括 {childCount} 个子标题）！", "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                LogParseProgress($"用户手动剔除了 {totalCount} 个标题（包括 {childCount} 个子标题）");
            }
        }

        /// <summary>
        /// 递归收集标题及其所有子标题
        /// </summary>
        private void CollectHeadingAndChildren(DocumentHeading heading, HashSet<DocumentHeading> result)
        {
            result.Add(heading);
            
            // 查找所有子标题（级别更高且在当前标题之后的标题）
            int currentIndex = _parsedHeadings.IndexOf(heading);
            if (currentIndex >= 0)
            {
                for (int i = currentIndex + 1; i < _parsedHeadings.Count; i++)
                {
                    var nextHeading = _parsedHeadings[i];
                    
                    // 如果遇到同级或更高级的标题，停止搜索
                    if (nextHeading.HeadingLevel <= heading.HeadingLevel)
                    {
                        break;
                    }
                    
                    // 这是一个子标题，递归收集它的子标题
                    CollectHeadingAndChildren(nextHeading, result);
                    
                    // 跳过已经处理的标题
                    while (i + 1 < _parsedHeadings.Count && result.Contains(_parsedHeadings[i + 1]))
                    {
                        i++;
                    }
                }
            }
        }

        /// <summary>
        /// 显示解析结果
        /// </summary>
        private void DisplayParseResult()
        {
            listViewHeadings.Items.Clear();

            if (_parsedHeadings == null) return;

            foreach (var heading in _parsedHeadings)
            {
                var item = new ListViewItem(GetLevelIndent(heading.HeadingLevel) + heading.HeadingText);
                item.SubItems.Add($"H{heading.HeadingLevel}");
                item.SubItems.Add(heading.Content?.Length.ToString() ?? "0");
                item.Tag = heading;

                // 根据级别设置不同颜色
                switch (heading.HeadingLevel)
                {
                    case 1:
                        item.ForeColor = Color.DarkBlue;
                        item.Font = new Font(listViewHeadings.Font, FontStyle.Bold);
                        break;
                    case 2:
                        item.ForeColor = Color.DarkGreen;
                        break;
                    case 3:
                        item.ForeColor = Color.DarkOrange;
                        break;
                    default:
                        item.ForeColor = Color.Black;
                        break;
                }

                listViewHeadings.Items.Add(item);
            }
        }

        /// <summary>
        /// 获取层级缩进
        /// </summary>
        private string GetLevelIndent(int level)
        {
            return new string(' ', (level - 1) * 4);
        }
        
        /// <summary>
        /// 显示标题内容
        /// </summary>
        private void ShowHeadingContent(DocumentHeading heading)
        {
            if (heading == null) return;
            
            string content = heading.Content ?? "(无内容)";
            
            // 创建一个新窗体显示内容
            var contentForm = new Form
            {
                Text = $"标题内容 - {heading.HeadingText}",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable
            };
            
            var textBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Text = content,
                Font = new Font("Consolas", 9F)
            };
            
            var panel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };
            
            var btnClose = new Button
            {
                Text = "关闭",
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(700, 8),
                Size = new Size(80, 25)
            };
            btnClose.Click += (s, e) => contentForm.Close();
            
            var lblInfo = new Label
            {
                Text = $"级别: H{heading.HeadingLevel}  |  内容长度: {content.Length} 字符  |  (双击标题可查看内容)",
                AutoSize = false,
                Size = new Size(680, 25),
                Location = new Point(10, 10),
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            panel.Controls.Add(btnClose);
            panel.Controls.Add(lblInfo);
            contentForm.Controls.Add(textBox);
            contentForm.Controls.Add(panel);
            
            contentForm.ShowDialog(this);
        }

        #endregion
    }
} 