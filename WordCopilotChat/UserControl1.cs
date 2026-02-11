using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using WordCopilotChat.utils;
using WordCopilotChat.models;
using WordCopilotChat.services;

// 使用别名解决命名冲突
using TaskAsync = System.Threading.Tasks.Task;
using Model = WordCopilotChat.models.Model;
using WordCopilot.utils;



namespace WordCopilotChat
{
    public partial class UserControl1 : UserControl
    {
        // 标记WebView2是否已经初始化
        private bool _webViewInitialized = false;

        // MarkdownToWord工具实例
        private MarkdownToWord _markdownToWord;

        // 对话历史记录 - 用于维护上下文
        private List<JObject> _conversationHistory = new List<JObject>();

        // 最大历史记录数量（避免内存占用过大）
        private const int MAX_HISTORY_COUNT = 20;

        // 加载的文档内容（用于Chat和Agent模式）
        private string _loadedDocumentContent = null;
        private DateTime _documentLoadedTime = DateTime.MinValue;

        // 用于取消API请求的CancellationTokenSource
        private CancellationTokenSource _cancellationTokenSource;

        // 用于取消获取文档标题操作的CancellationTokenSource
        private CancellationTokenSource _headingsCancellationTokenSource;

        // 用于取消“加载当前Word文档”操作的CancellationTokenSource
        private CancellationTokenSource _loadDocumentCancellationTokenSource;
        // 加载文档操作序号（用于丢弃过期结果）
        private int _loadDocumentOpId = 0;
        // 是否正在加载文档
        private volatile bool _isLoadingDocument = false;


        // 模型服务
        private ModelService _modelService;

        // 文档服务
        private DocumentService _documentService;

        // 内部请求（如上下文压缩）期间，避免向前端推送 usage，防止 Token 百分比跳变
        private int _suppressUsageToFrontendCount = 0;
        // 最近一次 usage（用于调试输出）
        private JObject _lastUsagePayload = null;

        // 应用设置服务
        private AppSettingsService _appSettingsService;

        // 当前使用的模型配置（用于用户操作反馈对话）
        private Model _currentModelConfig;
        // 最近一次对话所使用的模型请求信息（用于“强制压缩”等非发送路径触发的内部请求）
        private string _currentApiUrl;
        private string _currentApiKey;
        private string _currentModelName;
        // 最近一次对话的系统提示词与 max_tokens（用于压缩后估算并刷新 Token 使用率）
        private string _lastChatSystemPrompt;
        private int? _lastChatMaxTokens;
        // 防止并发压缩（0=未压缩中，1=压缩中）
        private int _compressingFlag = 0;

        public UserControl1()
        {
            InitializeComponent();
            _markdownToWord = new MarkdownToWord();
            _modelService = new ModelService();
            _documentService = new DocumentService();
            _appSettingsService = new AppSettingsService();

            // 订阅工具预览事件
            OpenAIUtils.OnToolPreviewReady += HandleToolPreviewFromOpenAI;

            // 订阅工具进度事件
            OpenAIUtils.OnToolProgress += HandleToolProgressFromOpenAI;

            // 订阅 usage 事件（token 用量）
            OpenAIUtils.OnUsageReady += HandleUsageFromOpenAI;
        }

        private async void UserControl1_Load(object sender, EventArgs e)
        {
            // 检测是否安装了 WebView2 运行时
            try
            {
                String browserVersion = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
                if (string.IsNullOrEmpty(browserVersion))
                {
                    MessageBox.Show("未安装 WebView2 运行时，请先安装");
                    return;
                }
                else
                {
                    try
                    {
                        // 加载web资源1
                        //// 初始化WebView2环境
                        //var environment = await CoreWebView2Environment.CreateAsync(null, null, null);
                        //await webView21.EnsureCoreWebView2Async(environment);

                        //// vsto插件的路径
                        //string pluginDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        //string localHtmlPath = Path.Combine(pluginDirectory, "wwwroot", "index.html");

                        //// 设置允许开发者工具
                        //webView21.CoreWebView2.Settings.AreDevToolsEnabled = true;

                        //// 设置WebMessage接收处理程序
                        //webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                        //// 加载HTML页面
                        //this.webView21.Source = new Uri($"file:///{localHtmlPath}");

                        //加载web资源方案2
                        // 为本插件显式指定独立的 WebView2 用户数据目录，避免 0x800700AA
                        string userDataFolder = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "WordCopilotChat",
                            "WebView2",
                            "UserData");
                        Directory.CreateDirectory(userDataFolder);
                        // 初始化WebView2环境
                        var environment = await CoreWebView2Environment.CreateAsync(
                            browserExecutableFolder: null,
                            userDataFolder: userDataFolder,
                            options: null
                        );
                        await webView21.EnsureCoreWebView2Async(environment);
                        // vsto插件的路径
                        string pluginDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        string localHtmlPath = Path.Combine(pluginDirectory, "wwwroot", "index.html");

                        // 设置允许开发者工具
                        webView21.CoreWebView2.Settings.AreDevToolsEnabled = false;

                        // 设置WebMessage接收处理程序
                        webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;


                        // 优先使用虚拟主机映射以保证静态资源的可靠加载
                        string wwwRoot = Path.Combine(pluginDirectory, "wwwroot");
                        if (Directory.Exists(wwwRoot))
                        {
                            webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                                hostName: "app.local",
                                folderPath: wwwRoot,
                                accessKind: CoreWebView2HostResourceAccessKind.Allow);
                            webView21.CoreWebView2.Navigate("https://app.local/index.html");
                        }
                        else
                        {
                            // 回退到本地文件（从路径构造 Uri，可自动处理如 # -> %23 等特殊字符）
                            if (!File.Exists(localHtmlPath))
                            {
                                MessageBox.Show($"未找到前端文件: {localHtmlPath}");
                                return;
                            }
                            var localUri = new Uri(localHtmlPath);
                            this.webView21.Source = localUri;
                        }





                        Debug.WriteLine("WebView2 运行时版本: " + browserVersion);

                        // 标记WebView2已初始化
                        _webViewInitialized = true;

                        // 添加导航完成事件
                        webView21.NavigationCompleted += WebView21_NavigationCompleted;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"初始化WebView2时出错: {ex.Message}");
                        Debug.WriteLine($"初始化WebView2详细错误: {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"检测WebView2时出错: {ex.Message}");
                Debug.WriteLine($"检测WebView2详细错误: {ex}");
            }
        }

        private void WebView21_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                Debug.WriteLine("导航完成，页面加载成功");

                // 页面加载成功后，等待页面ready消息，然后发送欢迎消息内容和模型列表
                // 在CoreWebView2_WebMessageReceived的ready消息处理中发送欢迎消息和模型列表
            }
            else
            {
                Debug.WriteLine($"导航完成，但页面加载失败: {e.WebErrorStatus}");
            }
        }

        // 处理从JavaScript接收的WebMessage
        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取JSON字符串并解析
                string jsonString = e.WebMessageAsJson;
                Debug.WriteLine($"收到WebMessage完整内容: {jsonString}");

                // 使用Newtonsoft.Json解析
                JObject jsonObject = JObject.Parse(jsonString);
                string messageType = jsonObject["type"]?.ToString();
                string messageContent = jsonObject["message"]?.ToString();
                string chatMode = jsonObject["mode"]?.ToString(); // 新增：获取聊天模式
                JArray enabledToolsArray = jsonObject["enabledTools"] as JArray; // 新增：获取启用的工具列表
                int selectedModelId = jsonObject["selectedModelId"]?.ToObject<int>() ?? 0; // 新增：获取选中的模型ID
                JArray contextsArray = jsonObject["contexts"] as JArray; // 新增：获取上下文列表
                JArray imagesArray = jsonObject["images"] as JArray; // 新增：获取图片列表

                // 根据消息类型处理
                if (!string.IsNullOrEmpty(messageType))
                {
                    switch (messageType)
                    {
                        case "ready":
                            // JavaScript通知已准备好
                            Debug.WriteLine("WebView2页面已准备就绪");

                            // 发送欢迎消息的HTML内容到前端
                            SendWelcomeMessageToWebView();

                            // 发送模型列表到前端
                            SendModelListToWebView();
                            break;

                        case "userMessage":
                            // 处理用户消息，传递聊天模式和启用的工具列表
                            string[] enabledTools = null;
                            if (enabledToolsArray != null)
                            {
                                enabledTools = enabledToolsArray.ToObject<string[]>();
                            }

                            // 处理上下文内容
                            List<JObject> contexts = null;
                            if (contextsArray != null)
                            {
                                contexts = contextsArray.ToObject<List<JObject>>();
                            }

                            // 处理图片数据
                            List<JObject> images = null;
                            if (imagesArray != null)
                            {
                                images = imagesArray.ToObject<List<JObject>>();
                            }

                            HandleUserMessage(messageContent, chatMode, enabledTools, selectedModelId, contexts, images);
                            break;

                        case "loadDocument":
                            // 处理加载文档内容的请求
                            Debug.WriteLine("收到加载文档内容请求");
                            HandleLoadDocument();
                            break;

                        case "cancelLoadDocument":
                            Debug.WriteLine("收到取消加载文档请求");
                            CancelLoadDocument();
                            break;

                        case "unloadDocument":
                            Debug.WriteLine("收到卸载已加载文档请求");
                            UnloadDocument();
                            break;

                        case "clearHistory":
                            // 处理清空对话历史的请求
                            Debug.WriteLine("收到清空对话历史请求");
                            ClearConversationHistory();
                            break;

                        case "forceCompress":
                            // 强制压缩对话历史（由前端按钮触发）
                            Debug.WriteLine("收到强制压缩请求");
                            // 必须在 UI 线程执行 WebView2 相关操作
                            if (InvokeRequired)
                            {
                                BeginInvoke(new Action(async () =>
                                {
                                    await HandleForceCompressAsync();
                                }));
                            }
                            else
                            {
                                _ = HandleForceCompressAsync();
                            }
                            break;

                        case "copyToWord":
                            // 处理复制到Word的请求
                            Debug.WriteLine("收到复制到Word请求");

                            // 获取格式类型
                            string format = jsonObject["format"]?.ToString() ?? "html";

                            // 检查是否是选中内容
                            bool isSelection = jsonObject["isSelection"] != null && jsonObject["isSelection"].ToObject<bool>();
                            if (isSelection)
                            {
                                string content = jsonObject["content"]?.ToString();
                                bool hasFormula = jsonObject["hasFormula"] != null && jsonObject["hasFormula"].ToObject<bool>();

                                // 处理选中内容插入
                                InsertSelectionToWord(content, hasFormula);
                            }
                            else if (format == "formula")
                            {
                                // 处理公式插入
                                string formulaContent = jsonObject["content"]?.ToString();
                                InsertFormulaToWord(formulaContent);
                            }
                            else if (format == "text")
                            {
                                // 处理纯文本插入
                                string textContent = jsonObject["content"]?.ToString();
                                InsertTextToWord(textContent);
                            }
                            else if (jsonObject["formulas"] != null)
                            {
                                // 有公式数据，需要特殊处理
                                Debug.WriteLine($"包含公式数据: {jsonObject["formulas"].Count()} 个公式");
                                CopyContentToWord(jsonString); // 传递整个JSON字符串
                            }
                            else
                            {
                                // 普通HTML内容
                                string htmlContent = jsonObject["content"]?.ToString();
                                CopyContentToWord(htmlContent);
                            }
                            break;

                        case "insertSequence":
                            // 处理顺序插入请求
                            Debug.WriteLine("收到顺序插入请求");
                            HandleSequenceInsertion(jsonObject);
                            break;

                        case "insertMermaidImage":
                            // 处理Mermaid图片插入请求
                            Debug.WriteLine("收到Mermaid图片插入请求");
                            HandleMermaidImageInsertion(jsonObject);
                            break;

                        case "stopGeneration":
                            // 处理停止生成请求
                            Debug.WriteLine("收到停止生成请求");
                            HandleStopGeneration();
                            break;

                        case "getDocumentHeadings":
                            // 处理获取文档标题请求
                            Debug.WriteLine("收到获取文档标题请求");
                            HandleGetDocumentHeadings(jsonObject);
                            break;

                        case "applyPreviewedAction":
                            // 处理应用预览操作请求
                            Debug.WriteLine("收到应用预览操作请求");
                            HandleApplyPreviewedAction(jsonObject);
                            break;

                        case "getDocuments":
                            // 处理获取文档列表请求
                            Debug.WriteLine("收到获取文档列表请求");
                            HandleGetDocuments();
                            break;

                        case "getDocumentContent":
                            // 处理获取文档内容请求
                            Debug.WriteLine("收到获取文档内容请求");
                            HandleGetDocumentContent(jsonObject);
                            break;

                        case "updateAgentConfig":
                            // 处理Agent配置更新请求
                            Debug.WriteLine("收到Agent配置更新请求");
                            HandleUpdateAgentConfig(jsonObject);
                            break;

                        case "rejectPreviewedAction":
                            // 处理拒绝预览操作请求
                            Debug.WriteLine("收到拒绝预览操作请求");
                            HandleRejectPreviewedAction(jsonObject);
                            break;

                        default:
                            Debug.WriteLine($"收到未知类型的消息: {messageType}");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理WebMessage时出错: {ex.Message}");
            }
        }

        // 处理强制压缩请求
        private async System.Threading.Tasks.Task HandleForceCompressAsync()
        {
            try
            {
                // 避免并发压缩
                if (System.Threading.Interlocked.CompareExchange(ref _compressingFlag, 1, 0) != 0)
                {
                    Debug.WriteLine("强制压缩：已有压缩任务在进行中，忽略本次请求");
                    return;
                }

                if (_currentModelConfig?.ContextLength.HasValue != true || _currentModelConfig.ContextLength.Value <= 0)
                {
                    Debug.WriteLine("强制压缩：当前模型未配置 context_length，无法执行");
                    PostWebMessageSafe(JsonConvert.SerializeObject(new { type = "showError", message = "当前模型未配置上下文长度，无法执行强制压缩。" }));
                    return;
                }

                if (string.IsNullOrWhiteSpace(_currentApiUrl) || string.IsNullOrWhiteSpace(_currentApiKey) || string.IsNullOrWhiteSpace(_currentModelName))
                {
                    Debug.WriteLine("强制压缩：缺少最近一次模型请求信息（apiUrl/apiKey/modelName）");
                    PostWebMessageSafe(JsonConvert.SerializeObject(new { type = "showError", message = "尚未进行过有效对话，无法执行强制压缩。请先发送一条消息。" }));
                    return;
                }

                int contextLengthTokens = _currentModelConfig.ContextLength.Value * 1024;

                // 通知前端开始压缩
                var forceMsg = "正在强制压缩对话历史以节省Token...";
                // 兼容性：部分情况下 WebMessage 会延迟/被吞，这里额外用 ExecuteScript 直接触发前端展示
                try
                {
                    var forceMsgJs = JsonConvert.SerializeObject(forceMsg);
                    await ExecuteScriptOnUIThreadAsync($"try{{showCompressionBanner({forceMsgJs}, 'info');setCompressionLock(true);}}catch(e){{}}");
                }
                catch { }
                PostWebMessageSafe(JsonConvert.SerializeObject(new { type = "contextCompressing", message = forceMsg }));

                System.Threading.Interlocked.Increment(ref _suppressUsageToFrontendCount);
                try
                {
                    await CompressConversationHistory(_currentApiUrl, _currentApiKey, _currentModelName, contextLengthTokens);
                }
                finally
                {
                    System.Threading.Interlocked.Decrement(ref _suppressUsageToFrontendCount);
                }

                // 压缩完成后：主动推送一次“估算 usage”，让前端 Token 使用率立刻刷新（不依赖下一次模型回复）
                SendEstimatedUsageToFrontend(_lastChatSystemPrompt, _currentModelConfig, _lastChatMaxTokens);

                // 通知前端压缩完成（await 后可能在非 UI 线程）
                try
                {
                    var doneMsg = "对话历史已压缩";
                    var doneMsgJs = JsonConvert.SerializeObject(doneMsg);
                    await ExecuteScriptOnUIThreadAsync($"try{{showCompressionBanner({doneMsgJs}, 'success', 1200);setCompressionLock(false);}}catch(e){{}}");
                }
                catch { }
                PostWebMessageSafe(JsonConvert.SerializeObject(new { type = "contextCompressed", message = "对话历史已压缩" }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"强制压缩失败: {ex.Message}");
                PostWebMessageSafe(JsonConvert.SerializeObject(new { type = "showError", message = $"强制压缩失败: {ex.Message}" }));
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _compressingFlag, 0);
            }
        }

        // 线程安全的 WebMessage 发送（自动切换到 UI 线程）
        // 注意：CoreWebView2 的 getter 也要求在 UI 线程访问，因此不要在 UI 线程切换前读取 webView21.CoreWebView2
        private void PostWebMessageSafe(string json)
        {
            if (!_webViewInitialized || webView21 == null)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (_webViewInitialized && webView21.CoreWebView2 != null)
                        {
                            webView21.CoreWebView2.PostWebMessageAsJson(json);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"PostWebMessageSafe 出错: {ex.Message}");
                    }
                }));
                return;
            }

            try
            {
                if (webView21.CoreWebView2 != null)
                {
                    webView21.CoreWebView2.PostWebMessageAsJson(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PostWebMessageSafe 出错: {ex.Message}");
            }
        }

        // 处理选中内容插入到Word
        private void InsertSelectionToWord(string content, bool hasFormula)
        {
            try
            {
                Debug.WriteLine($"开始处理选中内容插入: {content.Substring(0, Math.Min(50, content.Length))}...");

                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => InsertSelectionToWord(content, hasFormula)));
                    return;
                }

                // 使用MarkdownToWord工具类插入选中内容
                _markdownToWord.InsertSelectedText(content, hasFormula);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入选中内容到Word时出错: {ex.Message}");
                MessageBox.Show($"插入到Word失败: {ex.Message}");
            }
        }

        // 处理纯文本插入到Word
        private void InsertTextToWord(string textContent)
        {
            try
            {
                Debug.WriteLine($"=== 开始处理文本插入 ===");
                Debug.WriteLine($"文本内容: {textContent?.Substring(0, Math.Min(100, textContent?.Length ?? 0))}...");

                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => InsertTextToWord(textContent)));
                    return;
                }

                if (string.IsNullOrEmpty(textContent))
                {
                    Debug.WriteLine("文本内容为空，跳过插入");
                    return;
                }

                // 获取Word应用程序和选择区域
                dynamic wordApp = _markdownToWord.GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法获取Word应用程序");
                    MessageBox.Show("无法连接到Word应用程序，请确保Word已打开");
                    return;
                }

                dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    MessageBox.Show("无法获取活动的Word文档");
                    return;
                }

                dynamic selection = wordApp.Selection;

                // 移动到文档末尾
                selection.EndKey(6); // 6 = wdStory，移动到文档末尾
                Debug.WriteLine($"移动到文档末尾，位置: {selection.Start}");

                // 如果文档不为空，添加换行分隔
                if (selection.Start > 0)
                {
                    selection.TypeText("\r\n\r\n");
                }

                // 插入文本内容
                selection.TypeText(textContent);
                Debug.WriteLine("=== 文本插入完成 ===");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"=== 插入文本到Word时出错 ===");
                Debug.WriteLine($"错误信息: {ex.Message}");
                MessageBox.Show($"插入文本到Word失败: {ex.Message}");
            }
        }

        // 处理公式插入到Word
        private void InsertFormulaToWord(string formulaContent)
        {
            try
            {
                Debug.WriteLine($"=== 开始处理公式插入 ===");
                Debug.WriteLine($"公式内容: {formulaContent}");
                Debug.WriteLine($"当前线程: {System.Threading.Thread.CurrentThread.ManagedThreadId}");

                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    Debug.WriteLine("需要调用到UI线程");
                    this.Invoke(new Action(() => InsertFormulaToWord(formulaContent)));
                    return;
                }

                Debug.WriteLine("开始获取Word应用程序...");

                // 使用MarkdownToWord工具类插入公式
                // 获取Word应用程序和选择区域
                dynamic wordApp = _markdownToWord.GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法获取Word应用程序");
                    MessageBox.Show("无法连接到Word应用程序，请确保Word已打开");
                    return;
                }

                Debug.WriteLine("成功获取Word应用程序，获取活动文档...");

                dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    MessageBox.Show("无法获取活动的Word文档");
                    return;
                }

                Debug.WriteLine("成功获取活动文档，获取选择区域...");

                dynamic selection = wordApp.Selection;
                Debug.WriteLine($"当前光标位置: {selection.Start}");

                // 保持用户当前选中的位置，不强制移动到文档末尾
                // 检查当前位置是否适合插入公式
                try
                {
                    // 如果当前在公式编辑状态，先退出
                    if (selection.OMaths.Count > 0)
                    {
                        Debug.WriteLine("当前在公式编辑状态，移动到合适位置");
                        selection.MoveRight(1, 1);
                        selection.Collapse(0);
                    }

                    // 检查是否在行中间，如果是则添加换行
                    var currentPara = selection.Paragraphs[1];
                    if (currentPara.Range.Text.Trim().Length > 0 && selection.Start > currentPara.Range.Start)
                    {
                        Debug.WriteLine("在行中间，添加换行开始新行");
                        selection.TypeText("\r\n");
                    }
                }
                catch (Exception posEx)
                {
                    Debug.WriteLine($"检查插入位置时出错: {posEx.Message}");
                    // 如果检查失败，确保至少有一个换行
                    try
                    {
                        selection.TypeText("\r\n");
                    }
                    catch { }
                }

                Debug.WriteLine($"准备在位置 {selection.Start} 插入公式");

                // 使用MarkdownToWord工具类插入公式，设置为显示模式
                Debug.WriteLine($"开始调用InsertFormula方法，公式: {formulaContent}, 显示模式: true");
                bool success = _markdownToWord.InsertFormula(selection, formulaContent, true);

                if (success)
                {
                    Debug.WriteLine("=== 公式插入成功 ===");
                }
                else
                {
                    Debug.WriteLine("=== 公式插入失败，已回退为文本 ===");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"=== 插入公式到Word时出错 ===");
                Debug.WriteLine($"错误信息: {ex.Message}");
                Debug.WriteLine($"错误堆栈: {ex.StackTrace}");
                MessageBox.Show($"插入公式到Word失败: {ex.Message}");
            }
        }

        // 处理来自用户的消息
        private async void HandleUserMessage(string message, string chatMode, string[] enabledTools = null, int selectedModelId = 0, List<JObject> contexts = null, List<JObject> images = null)
        {
            try
            {
                Debug.WriteLine($"处理用户消息: {message}, 模式: {chatMode}");

                // 如果有图片，输出图片信息
                if (images != null && images.Count > 0)
                {
                    Debug.WriteLine($"收到 {images.Count} 张图片");
                    foreach (var img in images)
                    {
                        string imgName = img["name"]?.ToString() ?? "未命名";
                        int imgSize = img["size"]?.ToObject<int>() ?? 0;
                        Debug.WriteLine($"  - {imgName}: {imgSize / 1024.0:F2} KB");
                    }
                }

                // 重置WordTools的取消标志（开始新的对话）
                WordTools.ResetCancelFlag();

                // 如果是智能体模式，注册Word工具
                if (chatMode == "chat-agent")
                {
                    try
                    {
                        if (enabledTools != null)
                        {
                            if (enabledTools.Length > 0)
                            {
                                // 根据用户选择的工具注册
                                WordTools.RegisterSelectedTools(enabledTools);
                                Debug.WriteLine($"智能体模式：根据用户选择注册了 {OpenAIUtils.GetToolCount()} 个Word工具");
                                Debug.WriteLine($"启用的工具: {string.Join(", ", enabledTools)}");
                            }
                            else
                            {
                                // 用户明确选择了0个工具，清空所有工具
                                OpenAIUtils.ClearTools();
                                Debug.WriteLine("智能体模式：用户选择了0个工具，已清空所有工具");
                            }
                        }
                        else
                        {
                            // 如果没有提供工具列表，注册所有工具（向后兼容，只在老版本前端时发生）
                            WordTools.RegisterAllTools();
                            Debug.WriteLine($"智能体模式：未收到工具选择信息，注册了所有 {OpenAIUtils.GetToolCount()} 个Word工具（向后兼容模式）");
                        }
                    }
                    catch (Exception toolEx)
                    {
                        Debug.WriteLine($"注册Word工具时出错: {toolEx.Message}");
                    }
                }
                else
                {
                    // 如果不是智能体模式，清空工具
                    OpenAIUtils.ClearTools();
                    Debug.WriteLine("智能问答模式：已清空所有工具");
                }

                // 注意：已加载的Word文档内容现在会添加到系统提示词中，不再重复添加到每条用户消息
                // 这样可以大幅节省Token（文档内容只出现在system里一次，而不是每条user消息都带上）
                if (!string.IsNullOrEmpty(_loadedDocumentContent) && !_loadedDocumentContent.StartsWith("错误"))
                {
                    Debug.WriteLine($"检测到已加载的文档内容，字符数: {_loadedDocumentContent.Length}，将添加到系统提示词");
                }

                // 添加用户消息到历史记录（支持多模态）
                AddUserMessageToHistory(message, images);

                // 从数据库获取选中的模型配置
                var modelConfig = GetSelectedModel(selectedModelId);
                string modelName, apiUrl, apiKey;
                bool supportsMultimodal = false;

                if (modelConfig == null)
                {
                    Debug.WriteLine("未找到选中的模型，请用户配置模型");

                    // 通知前端显示错误并提示用户配置模型
                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(async () => {
                            try
                            {
                                await webView21.ExecuteScriptAsync("showError('请先在设置中配置AI模型后再使用对话功能')");
                                await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"显示错误消息时出错: {ex.Message}");
                            }
                        }));
                    }
                    else
                    {
                        await webView21.ExecuteScriptAsync("showError('请先在设置中配置AI模型后再使用对话功能')");
                        await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                    }
                    return;
                }
                else
                {
                    modelName = ExtractModelNameFromParameters(modelConfig.Parameters);
                    apiUrl = modelConfig.BaseUrl;
                    apiKey = modelConfig.ApiKey;
                    supportsMultimodal = modelConfig.EnableMulti == 1; // 检查是否启用多模态

                    // 保存当前模型配置，用于后续的用户操作反馈对话
                    _currentModelConfig = modelConfig;
                    // 保存最近一次对话请求信息，供“强制压缩”等操作复用
                    _currentApiUrl = apiUrl;
                    _currentApiKey = apiKey;
                    _currentModelName = modelName;

                    Debug.WriteLine($"模型多模态支持: {(supportsMultimodal ? "是" : "否")}");

                    // 如果有图片但模型不支持多模态，给出警告
                    if (images != null && images.Count > 0 && !supportsMultimodal)
                    {
                        Debug.WriteLine("警告: 当前模型不支持多模态，图片将被忽略");
                        if (InvokeRequired)
                        {
                            BeginInvoke(new Action(async () => {
                                try
                                {
                                    await webView21.ExecuteScriptAsync("showError('当前模型不支持多模态，无法处理图片。请在模型列表中启用多模态支持。')");
                                    await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"显示错误消息时出错: {ex.Message}");
                                }
                            }));
                        }
                        else
                        {
                            await webView21.ExecuteScriptAsync("showError('当前模型不支持多模态，无法处理图片。请在模型列表中启用多模态支持。')");
                            await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                        }
                        return;
                    }
                }

                Debug.WriteLine($"使用模型: {modelName}, API地址: {apiUrl}");

                // 根据聊天模式调整系统提示词
                string systemPrompt = GetSystemPromptByMode(chatMode, enabledTools, contexts);

                // 获取AI参数：优先从模型配置中解析，其次从应用设置中获取
                var (temperature, maxTokens, topP) = GetAIParameters(modelConfig, chatMode);

                // 记录最近一次对话的系统提示词与 max_tokens，供压缩后刷新 Token 使用率
                _lastChatSystemPrompt = systemPrompt;
                _lastChatMaxTokens = maxTokens;

                // 检查是否需要压缩对话历史（Token管理）
                await CheckAndCompressHistoryIfNeeded(systemPrompt, modelConfig, maxTokens, apiUrl, apiKey, modelName);

                // 构建消息数组，包含系统提示词和对话历史
                JArray messages = new JArray();

                // 添加系统提示词
                messages.Add(new JObject
                {
                    ["role"] = "system",
                    ["content"] = systemPrompt
                });

                // 添加对话历史
                foreach (var historyMessage in _conversationHistory)
                {
                    messages.Add(historyMessage);
                }

                // 构建请求的基础JSON对象
                JObject baseObj = new JObject
                {
                    ["model"] = modelName,
                    ["messages"] = messages,
                    ["stream"] = true,
                    ["temperature"] = temperature,
                    ["top_p"] = topP
                };

                // max_tokens为可选参数，仅在有值时添加
                if (maxTokens.HasValue)
                {
                    baseObj["max_tokens"] = maxTokens.Value;
                }

                // 在发送请求前输出本次对话的完整 messages 内容和请求参数
                try
                {
                    var prettyMessages = JsonConvert.SerializeObject(messages, Formatting.Indented);
                    Debug.WriteLine("=== 本次请求 messages 开始 ===");
                    Debug.WriteLine(prettyMessages);
                    Debug.WriteLine("=== 本次请求 messages 结束 ===");

                    // 输出请求参数
                    Debug.WriteLine("=== 本次请求参数 ===");
                    Debug.WriteLine($"模型: {modelName}");
                    Debug.WriteLine($"温度: {temperature}");
                    Debug.WriteLine($"最大令牌: {(maxTokens.HasValue ? maxTokens.Value.ToString() : "未设置(由服务商自动处理)")}");
                    Debug.WriteLine($"Top-P: {topP}");
                    Debug.WriteLine($"聊天模式: {chatMode}");
                    Debug.WriteLine("=== 参数输出结束 ===");
                }
                catch { }

                // 序列化为最终JSON
                string json = baseObj.ToString();

                Debug.WriteLine($"发送的消息历史数量: {_conversationHistory.Count}");

                // 创建新的CancellationTokenSource
                _cancellationTokenSource?.Cancel(); // 取消之前的请求
                _cancellationTokenSource?.Dispose(); // 释放资源
                _cancellationTokenSource = new CancellationTokenSource();

                // 通知前端开始生成（必须在UI线程访问WebView2）
                await ExecuteScriptOnUIThreadAsync("startGeneratingOutline()");

                // 收集生成的内容
                StringBuilder responseContent = new StringBuilder();

                try
                {
                    // 调用API生成回复，传递CancellationToken和对话历史（用于工具链支持）
                    List<JObject> conversationForTools = null;
                    if (chatMode == "chat-agent" && OpenAIUtils.GetToolCount() > 0)
                    {
                        // 为工具链调用准备对话历史
                        conversationForTools = new List<JObject>();
                        foreach (var msg in messages)
                        {
                            conversationForTools.Add(msg as JObject);
                        }
                        Debug.WriteLine($"准备工具链调用，对话历史消息数: {conversationForTools.Count}");
                    }

                    await OpenAIUtils.OpenAIApiClientAsync(apiUrl, apiKey, json, _cancellationTokenSource.Token, content =>
                    {
                        // 将内容添加到StringBuilder
                        responseContent.Append(content);

                        // 确保在UI线程上执行WebView操作
                        if (InvokeRequired)
                        {
                            BeginInvoke(new Action(async () => {
                                try
                                {
                                    await webView21.ExecuteScriptAsync($"appendOutlineContent(`{content.Replace("`", "\\`")}`)");
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"向前端发送内容时出错: {ex.Message}");
                                }
                            }));
                        }
                        else
                        {
                            // 已在UI线程，直接执行
                            webView21.ExecuteScriptAsync($"appendOutlineContent(`{content.Replace("`", "\\`")}`)").ConfigureAwait(false);
                        }
                    }, conversationForTools);

                    // 将助手回复添加到历史记录
                    string finalResponse = responseContent.ToString();
                    if (!string.IsNullOrEmpty(finalResponse))
                    {
                        AddAssistantMessageToHistory(finalResponse);
                    }

                    // 通知前端生成完成
                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(async () => {
                            try
                            {
                                await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                                Debug.WriteLine($"对话回复生成完成 - 模式: {chatMode}, 历史记录数量: {_conversationHistory.Count}");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"通知前端生成完成时出错: {ex.Message}");
                            }
                        }));
                    }
                    else
                    {
                        // 已在UI线程，直接执行
                        await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                        Debug.WriteLine($"对话回复生成完成 - 模式: {chatMode}, 历史记录数量: {_conversationHistory.Count}");
                    }
                }
                catch (OperationCanceledException)
                {
                    // 操作被取消（用户点击停止）
                    Debug.WriteLine("API请求被用户取消");

                    // 通知前端停止生成
                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(async () => {
                            try
                            {
                                await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"通知前端停止生成时出错: {ex.Message}");
                            }
                        }));
                    }
                    else
                    {
                        await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"生成回复时出错: {ex.Message}");
                Debug.WriteLine($"异常堆栈: {ex.StackTrace}");

                // 确保在UI线程上执行WebView操作
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(async () => {
                        try
                        {
                            await webView21.ExecuteScriptAsync($"showError('生成回复时出错: {ex.Message.Replace("'", "\\'")}')");
                            // 确保通知前端生成已结束
                            await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                            Debug.WriteLine("异常处理：已通知前端生成结束");
                        }
                        catch (Exception innerEx)
                        {
                            Debug.WriteLine($"显示错误消息时出错: {innerEx.Message}");
                            MessageBox.Show($"生成回复时出错: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    // 已在UI线程，直接执行
                    await webView21.ExecuteScriptAsync($"showError('生成回复时出错: {ex.Message.Replace("'", "\\'")}')");
                    // 确保通知前端生成已结束
                    await webView21.ExecuteScriptAsync("finishGeneratingOutline()");
                    Debug.WriteLine("异常处理：已通知前端生成结束");
                }
            }
        }

        // 处理停止生成请求
        private void HandleStopGeneration()
        {
            try
            {
                Debug.WriteLine("正在停止API请求和Word工具操作...");

                // 取消当前的API请求
                _cancellationTokenSource?.Cancel();

                // 取消获取文档标题操作
                _headingsCancellationTokenSource?.Cancel();

                // 取消WordTools中的长时间运行操作
                WordTools.CancelCurrentOperation();

                Debug.WriteLine("API请求和Word工具操作已取消");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"停止生成时出错: {ex.Message}");
            }
        }

        // 处理获取文档标题请求（支持分页和取消）
        private async void HandleGetDocumentHeadings(JObject data = null)
        {
            try
            {
                // 获取分页参数
                int page = 0;
                int pageSize = 10; // 每页显示10个标题
                bool append = false;

                if (data != null)
                {
                    page = data["page"]?.ToObject<int>() ?? 0;
                    pageSize = data["pageSize"]?.ToObject<int>() ?? 10;
                    append = data["append"]?.ToObject<bool>() ?? false;
                }

                Debug.WriteLine($"开始获取文档标题... 页码: {page}, 每页: {pageSize}, 追加: {append}");

                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleGetDocumentHeadings(data)));
                    return;
                }

                // 在后台线程获取标题数据，使用专门的取消令牌
                string headingsJson = "";

                // 取消之前的标题获取操作（如果有的话）
                _headingsCancellationTokenSource?.Cancel();
                _headingsCancellationTokenSource?.Dispose();
                _headingsCancellationTokenSource = new CancellationTokenSource();

                var cancellationToken = _headingsCancellationTokenSource.Token;

                await System.Threading.Tasks.Task.Run(() =>
                {
                    headingsJson = WordTools.GetDocumentHeadingsForQuickSelector(page, pageSize, cancellationToken);
                }, cancellationToken);

                Debug.WriteLine($"获取到标题数据: {headingsJson.Substring(0, Math.Min(100, headingsJson.Length))}...");

                // 在UI线程中处理WebView2通信
                SendHeadingsToWebView(headingsJson, page, append);
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("获取文档标题被取消");
                // 发送取消信息到前端
                SendHeadingsCancelledToWebView();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档标题时出错: {ex.Message}");

                // 在UI线程中发送错误信息
                SendHeadingsErrorToWebView(ex.Message);
            }
        }

        // 发送标题数据到WebView（确保在UI线程执行）
        private void SendHeadingsToWebView(string headingsJson, int page = 0, bool append = false)
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => SendHeadingsToWebView(headingsJson, page, append)));
                    return;
                }

                // 解析JSON并发送到前端
                var headingsData = JsonConvert.DeserializeObject<dynamic>(headingsJson);

                // 发送到WebView
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var messageData = new
                    {
                        type = "documentHeadings",
                        headings = headingsData.headings,
                        page = page,
                        append = append,
                        hasMore = headingsData.hasMore ?? false,
                        total = headingsData.total ?? 0
                    };

                    string json = JsonConvert.SerializeObject(messageData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine($"文档标题数据已发送到前端 - 页码: {page}, 追加: {append}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送标题数据时出错: {ex.Message}");
            }
        }

        // 发送错误信息到WebView（确保在UI线程执行）
        private void SendHeadingsErrorToWebView(string errorMessage)
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => SendHeadingsErrorToWebView(errorMessage)));
                    return;
                }

                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var errorData = new
                    {
                        type = "documentHeadings",
                        headings = new object[0], // 空数组
                        error = errorMessage
                    };

                    string json = JsonConvert.SerializeObject(errorData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine("错误信息已发送到前端");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送错误信息失败: {ex.Message}");
            }
        }

        // 发送取消信息到WebView（确保在UI线程执行）
        private void SendHeadingsCancelledToWebView()
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => SendHeadingsCancelledToWebView()));
                    return;
                }

                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var cancelData = new
                    {
                        type = "documentHeadings",
                        headings = new object[0], // 空数组
                        cancelled = true,
                        message = "获取标题已取消"
                    };

                    string json = JsonConvert.SerializeObject(cancelData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine("取消信息已发送到前端");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送取消信息失败: {ex.Message}");
            }
        }

        // 添加用户消息到历史记录（支持多模态）
        private void AddUserMessageToHistory(string message, List<JObject> images = null)
        {
            JObject userMessage;

            // 如果有图片，构建多模态消息格式
            if (images != null && images.Count > 0)
            {
                var contentArray = new JArray();

                // 添加文本内容
                contentArray.Add(new JObject
                {
                    ["type"] = "text",
                    ["text"] = message
                });

                // 添加图片内容
                foreach (var img in images)
                {
                    string base64Data = img["base64"]?.ToString() ?? "";
                    string imageType = img["type"]?.ToString() ?? "image/png";

                    // 确保base64数据格式正确
                    if (!base64Data.StartsWith("data:"))
                    {
                        base64Data = $"data:{imageType};base64,{base64Data}";
                    }

                    contentArray.Add(new JObject
                    {
                        ["type"] = "image_url",
                        ["image_url"] = new JObject
                        {
                            ["url"] = base64Data
                        }
                    });
                }

                userMessage = new JObject
                {
                    ["role"] = "user",
                    ["content"] = contentArray
                };

                Debug.WriteLine($"添加多模态用户消息到历史，包含 {images.Count} 张图片");
            }
            else
            {
                // 纯文本消息
                userMessage = new JObject
                {
                    ["role"] = "user",
                    ["content"] = message
                };
            }

            _conversationHistory.Add(userMessage);

            // 限制历史记录数量，保留最近的对话
            TrimConversationHistory();

            Debug.WriteLine($"添加用户消息到历史，当前历史数量: {_conversationHistory.Count}");
        }

        // 添加助手消息到历史记录
        private void AddAssistantMessageToHistory(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            // 思维链不应进入后续对话上下文（节省Token、避免服务商不兼容）
            var (cleanContent, think) = ExtractThinkAndCleanContent(message);

            var assistantMessage = new JObject
            {
                ["role"] = "assistant",
                ["content"] = cleanContent
            };

            // 仅用于本地存储/日志，不发送给模型
            if (!string.IsNullOrWhiteSpace(think))
            {
                assistantMessage["think"] = think;
            }

            _conversationHistory.Add(assistantMessage);

            // 限制历史记录数量
            TrimConversationHistory();

            Debug.WriteLine($"添加助手消息到历史，当前历史数量: {_conversationHistory.Count}");
        }

        /// <summary>
        /// 从模型输出中提取思维链（&lt;think&gt;...&lt;/think&gt;），并返回去除思维链后的正文内容
        /// </summary>
        private static (string content, string think) ExtractThinkAndCleanContent(string raw)
        {
            try
            {
                if (string.IsNullOrEmpty(raw)) return ("", "");

                var thinkParts = new List<string>();

                // 1) 提取已闭合的 <think>...</think>
                try
                {
                    foreach (Match m in Regex.Matches(raw, @"<think>([\s\S]*?)</think>", RegexOptions.IgnoreCase))
                    {
                        if (m.Success && m.Groups.Count > 1)
                        {
                            var t = m.Groups[1]?.Value;
                            if (!string.IsNullOrWhiteSpace(t)) thinkParts.Add(t.Trim());
                        }
                    }
                }
                catch { }

                // 2) 移除闭合段
                string cleaned = Regex.Replace(raw, @"<think>[\s\S]*?</think>", "", RegexOptions.IgnoreCase);

                // 3) 处理未闭合的 <think>（兜底：把其后的内容作为 think）
                int openIdx = cleaned.IndexOf("<think>", StringComparison.OrdinalIgnoreCase);
                if (openIdx >= 0)
                {
                    string after = cleaned.Substring(openIdx + "<think>".Length);
                    if (!string.IsNullOrWhiteSpace(after)) thinkParts.Add(after.Trim());
                    cleaned = cleaned.Substring(0, openIdx);
                }

                // 4) 清理残留闭合标签
                cleaned = Regex.Replace(cleaned, @"</think>", "", RegexOptions.IgnoreCase);

                // 5) 工具卡片占位符属于UI层，不应进入后续对话上下文（节省Token/避免干扰模型）
                cleaned = Regex.Replace(cleaned ?? "", @"<!--\s*TOOL_CARD_[^>]*-->", "", RegexOptions.IgnoreCase);
                cleaned = Regex.Replace(cleaned ?? "", @"__TOOL_CARD_[^_]+__", "", RegexOptions.IgnoreCase);
                cleaned = Regex.Replace(cleaned ?? "", @"\*\*TOOL_CARD_[^*]+\*\*", "", RegexOptions.IgnoreCase);

                string think = string.Join("\n\n", thinkParts.Where(x => !string.IsNullOrWhiteSpace(x)));
                // think 仅用于本地日志/调试，也顺手去掉占位符
                think = Regex.Replace(think ?? "", @"<!--\s*TOOL_CARD_[^>]*-->", "", RegexOptions.IgnoreCase);
                think = Regex.Replace(think ?? "", @"__TOOL_CARD_[^_]+__", "", RegexOptions.IgnoreCase);
                think = Regex.Replace(think ?? "", @"\*\*TOOL_CARD_[^*]+\*\*", "", RegexOptions.IgnoreCase);

                return ((cleaned ?? "").Trim(), (think ?? "").Trim());
            }
            catch
            {
                // 出错时不影响主流程：宁可不提取也不要丢正文
                return ((raw ?? "").Trim(), "");
            }
        }

        // 修剪对话历史，保持在最大数量限制内
        private void TrimConversationHistory()
        {
            while (_conversationHistory.Count > MAX_HISTORY_COUNT)
            {
                // 移除最早的消息，但要保持成对的用户-助手对话
                if (_conversationHistory.Count >= 2)
                {
                    // 找到第一对完整的对话（用户+助手）并移除
                    for (int i = 0; i < _conversationHistory.Count - 1; i++)
                    {
                        if (_conversationHistory[i]["role"].ToString() == "user" &&
                            _conversationHistory[i + 1]["role"].ToString() == "assistant")
                        {
                            _conversationHistory.RemoveRange(i, 2);
                            break;
                        }
                    }

                    // 如果没找到完整对话，就移除第一条
                    if (_conversationHistory.Count > MAX_HISTORY_COUNT)
                    {
                        _conversationHistory.RemoveAt(0);
                    }
                }
                else
                {
                    _conversationHistory.RemoveAt(0);
                }
            }
        }

        // 清空对话历史（可用于重置对话）
        public void ClearConversationHistory()
        {
            _conversationHistory.Clear();
            _lastUsagePayload = null;
            System.Threading.Interlocked.Exchange(ref _suppressUsageToFrontendCount, 0);
            Debug.WriteLine("对话历史已清空");

            // 通知前端清空 token 显示
            try
            {
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var msg = new
                    {
                        type = "usageClear",
                        timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
                    };
                    string json = JsonConvert.SerializeObject(msg);
                    PostWebMessageAsJsonOnUIThread(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清空token显示通知失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理加载文档内容的请求
        /// </summary>
        private async void HandleLoadDocument()
        {
            try
            {
                // 若已加载，则视为“再次点击=卸载”
                if (!string.IsNullOrEmpty(_loadedDocumentContent) && !_loadedDocumentContent.StartsWith("错误"))
                {
                    UnloadDocument();
                    return;
                }

                // 若正在加载，则视为忽略（前端会走 cancelLoadDocument）
                if (_isLoadingDocument)
                {
                    Debug.WriteLine("文档加载已在进行中，忽略重复加载请求");
                    return;
                }

                Debug.WriteLine("开始加载文档内容...");
                _isLoadingDocument = true;

                // 发送“loading”状态给前端（对齐按钮提示）
                SendDocumentLoadStateMessage("loading", message: "正在加载文档...");

                // 在后台线程中读取文档内容
                string documentContent = null;
                DateTime loadTime = DateTime.MinValue;

                // 准备取消令牌与操作序号
                int opId = System.Threading.Interlocked.Increment(ref _loadDocumentOpId);
                try { _loadDocumentCancellationTokenSource?.Cancel(); } catch { }
                try { _loadDocumentCancellationTokenSource?.Dispose(); } catch { }
                _loadDocumentCancellationTokenSource = new CancellationTokenSource();
                var token = _loadDocumentCancellationTokenSource.Token;

                await TaskAsync.Run(() =>
                {
                    // 若已取消则尽早退出（注意：WordTools 内部不可中断，这里只能“尽早不开始”）
                    token.ThrowIfCancellationRequested();
                    documentContent = WordTools.GetFullDocumentContent(50000); // 限制最大50000字符
                    loadTime = DateTime.Now;
                }, token);

                // 若用户取消或已有更新的加载操作，则丢弃本次结果
                if (token.IsCancellationRequested || opId != _loadDocumentOpId)
                {
                    Debug.WriteLine("文档加载结果已过期或已取消，丢弃本次结果");
                    _isLoadingDocument = false;
                    return;
                }

                // 保存到实例变量（需要在UI线程）
                _loadedDocumentContent = documentContent;
                _documentLoadedTime = loadTime;
                _isLoadingDocument = false;

                // 统计文档信息
                int charCount = _loadedDocumentContent?.Length ?? 0;
                bool success = !string.IsNullOrEmpty(_loadedDocumentContent) && !_loadedDocumentContent.StartsWith("错误");

                Debug.WriteLine($"文档内容加载完成，字符数: {charCount}");

                // 切换回UI线程发送消息
                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        SendDocumentLoadedMessage(success, charCount, _documentLoadedTime, _loadedDocumentContent);
                    }));
                }
                else
                {
                    SendDocumentLoadedMessage(success, charCount, _documentLoadedTime, _loadedDocumentContent);
                }
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("加载文档内容已取消");
                _isLoadingDocument = false;
                // 取消时由 CancelLoadDocument 主动通知前端，这里不重复提示
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载文档内容失败: {ex.Message}");
                _isLoadingDocument = false;

                // 切换回UI线程发送错误消息
                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        SendDocumentLoadedErrorMessage(ex.Message);
                    }));
                }
                else
                {
                    SendDocumentLoadedErrorMessage(ex.Message);
                }
            }
        }

        /// <summary>
        /// 取消加载文档（并清空已加载内容）
        /// </summary>
        private void CancelLoadDocument()
        {
            try
            {
                // 递增操作序号，确保旧任务完成后结果会被丢弃
                System.Threading.Interlocked.Increment(ref _loadDocumentOpId);
                _isLoadingDocument = false;
                try { _loadDocumentCancellationTokenSource?.Cancel(); } catch { }

                _loadedDocumentContent = null;
                _documentLoadedTime = DateTime.MinValue;

                SendDocumentLoadStateMessage("cancelled", message: "已取消加载文档");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CancelLoadDocument 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 卸载已加载的文档内容（清空上下文）
        /// </summary>
        private void UnloadDocument()
        {
            try
            {
                // 同时取消可能进行中的加载
                System.Threading.Interlocked.Increment(ref _loadDocumentOpId);
                _isLoadingDocument = false;
                try { _loadDocumentCancellationTokenSource?.Cancel(); } catch { }

                _loadedDocumentContent = null;
                _documentLoadedTime = DateTime.MinValue;

                SendDocumentLoadStateMessage("unloaded", message: "已取消加载文档（已清空已加载内容）");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"UnloadDocument 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 发送文档加载成功消息（必须在UI线程调用）
        /// </summary>
        private void SendDocumentLoadedMessage(bool success, int charCount, DateTime loadTime, string documentContent)
        {
            try
            {
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var msg = new
                    {
                        type = "documentLoaded",
                        state = success ? "loaded" : "error",
                        success = success,
                        charCount = charCount,
                        loadTime = loadTime.ToString("yyyy-MM-dd HH:mm:ss"),
                        message = success ? $"文档已加载（{charCount} 字符）" : documentContent
                    };
                    string json = JsonConvert.SerializeObject(msg);
                    PostWebMessageAsJsonOnUIThread(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送文档加载消息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 发送文档加载错误消息（必须在UI线程调用）
        /// </summary>
        private void SendDocumentLoadedErrorMessage(string errorMessage)
        {
            try
            {
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var msg = new
                    {
                        type = "documentLoaded",
                        state = "error",
                        success = false,
                        message = $"加载失败: {errorMessage}"
                    };
                    string json = JsonConvert.SerializeObject(msg);
                    PostWebMessageAsJsonOnUIThread(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送文档加载错误消息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 向前端发送文档加载状态（loading/cancelled/unloaded）
        /// </summary>
        private void SendDocumentLoadStateMessage(string state, string message = null)
        {
            try
            {
                if (!_webViewInitialized || webView21?.CoreWebView2 == null) return;
                var msg = new
                {
                    type = "documentLoaded",
                    state = state,
                    success = state == "loaded",
                    message = message ?? ""
                };
                string json = JsonConvert.SerializeObject(msg);
                PostWebMessageAsJsonOnUIThread(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送文档加载状态失败: {ex.Message}");
            }
        }

        // Token估算（简单粗略估算：中文约1.5字符/token，英文约4字符/token）
        private int EstimateTokenCount(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;

            // 统计中文字符和非中文字符
            int chineseCount = 0;
            int otherCount = 0;

            foreach (char c in text)
            {
                if (c >= 0x4E00 && c <= 0x9FFF) // 中文字符范围
                {
                    chineseCount++;
                }
                else
                {
                    otherCount++;
                }
            }

            // 中文约1.5字符/token，英文约4字符/token
            int estimatedTokens = (int)(chineseCount / 1.5) + (int)(otherCount / 4.0);
            return Math.Max(1, estimatedTokens); // 至少返回1
        }

        // 估算整个对话历史的Token数
        private int EstimateHistoryTokens()
        {
            int total = 0;
            foreach (var msg in _conversationHistory)
            {
                string content = msg["content"]?.ToString() ?? "";
                total += EstimateTokenCount(content);

                // 如果有tool_calls，也要计算其Token
                var toolCalls = msg["tool_calls"] as JArray;
                if (toolCalls != null)
                {
                    foreach (var tc in toolCalls)
                    {
                        string args = tc["function"]?["arguments"]?.ToString() ?? "";
                        total += EstimateTokenCount(args);
                    }
                }
            }
            return total;
        }

        /// <summary>
        /// 压缩后向前端推送“估算 usage”，用于刷新 Token 使用率显示。
        /// 说明：压缩属于内部请求，压缩请求本身的 usage 会被 suppress；压缩完成后需要补一次“对当前会话上下文的估算”。
        /// </summary>
        private void SendEstimatedUsageToFrontend(string systemPrompt, Model modelConfig, int? maxTokens)
        {
            try
            {
                if (modelConfig?.ContextLength.HasValue != true || modelConfig.ContextLength.Value <= 0)
                {
                    return;
                }

                int contextLengthTokens = modelConfig.ContextLength.Value * 1024;
                int systemTokens = EstimateTokenCount(systemPrompt ?? "");
                int historyTokens = EstimateHistoryTokens();
                int promptTokens = systemTokens + historyTokens;

                var messageData = new
                {
                    type = "usage",
                    prompt_tokens = promptTokens,
                    completion_tokens = 0,
                    total_tokens = promptTokens,
                    max_tokens = maxTokens,
                    context_length = contextLengthTokens,
                    used_for_max_tokens = 0,
                    remaining_tokens = (int?)null,
                    estimated = true,
                    timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
                };

                PostWebMessageSafe(JsonConvert.SerializeObject(messageData));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SendEstimatedUsageToFrontend 出错: {ex.Message}");
            }
        }

        // 检查并在需要时压缩对话历史
        private async System.Threading.Tasks.Task CheckAndCompressHistoryIfNeeded(string systemPrompt, Model modelConfig, int? maxTokens, string apiUrl, string apiKey, string modelName)
        {
            try
            {
                // 1. 检查模型是否配置了上下文长度
                if (!modelConfig.ContextLength.HasValue || modelConfig.ContextLength.Value <= 0)
                {
                    Debug.WriteLine("模型未配置上下文长度（ContextLength），跳过自动压缩");
                    return;
                }

                // 2. 历史记录太少也跳过压缩（至少要有2条消息：1用户+1助手）
                if (_conversationHistory.Count < 2)
                {
                    return;
                }

                // 3. 计算上下文长度（单位转换：k -> tokens）
                int contextLengthTokens = modelConfig.ContextLength.Value * 1024;

                // 3.1 优先参考上一轮真实 usage（尤其是 Agent 模式包含 tools schema 时，本地估算会明显偏小）
                // 注意：usage 来自上一轮模型回复，total_tokens 通常包含 prompt+completion，作为保守判断可直接使用
                int? lastTotalTokens = null;
                try
                {
                    lastTotalTokens = _lastUsagePayload?["total_tokens"]?.ToObject<int?>();
                }
                catch { }

                // 4. 估算当前Token使用量
                int systemTokens = EstimateTokenCount(systemPrompt);
                int historyTokens = EstimateHistoryTokens();
                int reservedOutputTokens = maxTokens ?? 2000; // 为输出预留空间（默认2000）
                int totalEstimated = systemTokens + historyTokens + reservedOutputTokens;

                // 留出安全边界：当估算的输入Token超过阈值比例时触发压缩（阈值可在“默认AI参数设置”里配置）
                var thresholdPct = _appSettingsService?.GetIntSetting("context_compress_threshold_pct", 70) ?? 70;
                // 容错：限制区间[50%, 90%]
                if (thresholdPct < 50) thresholdPct = 50;
                if (thresholdPct > 90) thresholdPct = 90;
                double threshold = contextLengthTokens * (thresholdPct / 100.0);

                Debug.WriteLine($"🔍 Token估算 - 系统提示词: {systemTokens}, 历史对话: {historyTokens}, 预留输出: {reservedOutputTokens}, 总计: {totalEstimated}");
                Debug.WriteLine($"📊 上下文限制: {contextLengthTokens} tokens ({modelConfig.ContextLength}k), 阈值: {threshold:F0}（{thresholdPct}%）");
                if (lastTotalTokens.HasValue)
                {
                    Debug.WriteLine($"📌 上一轮真实 usage.total_tokens: {lastTotalTokens.Value}（用于辅助判断自动压缩）");
                }

                // 同步写入 openai_usage 日志（便于复盘“为什么触发压缩/阈值是多少”）
                try
                {
                    var debugLine1 = $"🔍 Token估算 - 系统提示词: {systemTokens}, 历史对话: {historyTokens}, 预留输出: {reservedOutputTokens}, 总计: {totalEstimated}";
                    var debugLine2 = $"📊 上下文限制: {contextLengthTokens} tokens ({modelConfig.ContextLength}k), 阈值: {threshold:F0}（{thresholdPct}%）";

                    OpenAIUtils.LogUsageDebug(
                        "context_compress_check",
                        new JObject
                        {
                            ["model"] = modelName,
                            ["context_length_k"] = modelConfig.ContextLength,
                            ["context_length_tokens"] = contextLengthTokens,
                            ["threshold_pct"] = thresholdPct,
                            ["threshold_tokens"] = (int)Math.Round(threshold),
                            ["system_tokens_est"] = systemTokens,
                            ["history_tokens_est"] = historyTokens,
                            ["reserved_output_tokens"] = reservedOutputTokens,
                            ["total_estimated_tokens"] = totalEstimated,
                            ["would_trigger"] = totalEstimated > threshold,
                            ["line1"] = debugLine1,
                            ["line2"] = debugLine2
                        });
                }
                catch { }

                // 触发条件：
                // - 若上一轮真实 total_tokens 已超过阈值，则优先触发（避免 Agent/tools 导致的估算偏差漏触发）
                // - 否则使用本地估算进行判断
                bool wouldTriggerByLastUsage = lastTotalTokens.HasValue && lastTotalTokens.Value > threshold;
                if (wouldTriggerByLastUsage || totalEstimated > threshold)
                {
                    Debug.WriteLine($"⚠️ Token即将超限，触发自动压缩");
                    Debug.WriteLine($"   当前: {totalEstimated} tokens, 阈值: {threshold:F0} tokens, 上下文窗口: {contextLengthTokens} tokens");

                    // 通知前端开始压缩
                    var bannerMsg = "对话历史较长，正在自动压缩以节省Token...";
                    var notifyData = new
                    {
                        type = "contextCompressing",
                        message = bannerMsg
                    };
                    // 兼容性：部分情况下 WebMessage 会延迟/被吞，这里额外用 ExecuteScript 直接触发前端展示
                    try
                    {
                        var bannerMsgJs = JsonConvert.SerializeObject(bannerMsg);
                        await ExecuteScriptOnUIThreadAsync($"try{{showCompressionBanner({bannerMsgJs}, 'info');setCompressionLock(true);}}catch(e){{}}");
                    }
                    catch { }
                    PostWebMessageSafe(JsonConvert.SerializeObject(notifyData));
                    // 压缩属于内部请求：为了避免 Token 显示“跳变”，压缩期间不向前端推送 usage（但仍会在VS输出中记录）
                    System.Threading.Interlocked.Increment(ref _suppressUsageToFrontendCount);
                    try
                    {
                        await CompressConversationHistory(apiUrl, apiKey, modelName, contextLengthTokens);
                    }
                    finally
                    {
                        System.Threading.Interlocked.Decrement(ref _suppressUsageToFrontendCount);
                    }

                    // 压缩完成后：推送一次“估算 usage”，让前端 Token 使用率立刻刷新
                    SendEstimatedUsageToFrontend(systemPrompt, modelConfig, maxTokens);

                    // 通知前端压缩完成（await 后可能在非 UI 线程）
                    try
                    {
                        var doneMsg = "对话历史已压缩";
                        var doneMsgJs = JsonConvert.SerializeObject(doneMsg);
                        await ExecuteScriptOnUIThreadAsync($"try{{showCompressionBanner({doneMsgJs}, 'success', 1200);setCompressionLock(false);}}catch(e){{}}");
                    }
                    catch { }
                    var completeData = new
                    {
                        type = "contextCompressed",
                        message = "对话历史已压缩",
                        originalCount = _conversationHistory.Count
                    };
                    PostWebMessageSafe(JsonConvert.SerializeObject(completeData));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"检查和压缩对话历史时出错: {ex.Message}");
                // 压缩失败不应阻止正常对话，仅记录日志
            }
        }

        /// <summary>
        /// 在 UI 线程执行 WebView2 ExecuteScriptAsync
        /// </summary>
        private System.Threading.Tasks.Task<string> ExecuteScriptOnUIThreadAsync(string script)
        {
            try
            {
                if (!_webViewInitialized || webView21 == null)
                {
                    return System.Threading.Tasks.Task.FromResult<string>(null);
                }

                if (InvokeRequired)
                {
                    var tcs = new System.Threading.Tasks.TaskCompletionSource<string>();
                    BeginInvoke(new Action(async () =>
                    {
                        try
                        {
                            var result = await webView21.ExecuteScriptAsync(script);
                            tcs.TrySetResult(result);
                        }
                        catch (Exception ex)
                        {
                            tcs.TrySetException(ex);
                        }
                    }));
                    return tcs.Task;
                }

                return webView21.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                return System.Threading.Tasks.Task.FromException<string>(ex);
            }
        }

        /// <summary>
        /// 在 UI 线程发送 PostWebMessageAsJson（CoreWebView2 只能在 UI 线程访问）
        /// </summary>
        private void PostWebMessageAsJsonOnUIThread(string json)
        {
            // 统一走线程安全实现（避免在非 UI 线程提前读取 CoreWebView2）
            PostWebMessageSafe(json);
        }

        // 压缩对话历史
        private async System.Threading.Tasks.Task CompressConversationHistory(string apiUrl, string apiKey, string modelName, int contextLengthTokens = 0)
        {
            try
            {
                if (_conversationHistory.Count < 2)
                {
                    Debug.WriteLine("对话历史太短，无需压缩");
                    return;
                }

                // 获取压缩提示词
                var compressPrompt = PromptService.GetPrompt("compress-context");
                if (string.IsNullOrEmpty(compressPrompt))
                {
                    Debug.WriteLine("未找到压缩提示词，跳过压缩");
                    return;
                }

                // 保留最近的1条用户消息（当前提问），其余的全部进行压缩
                // 原因：如果保留"用户+助手"对，上一轮的助手回复会不断累积而不被压缩
                var recentMessages = new List<JObject>();
                var oldMessages = new List<JObject>();

                // 从后往前找，保留最后1条用户消息
                for (int i = _conversationHistory.Count - 1; i >= 0; i--)
                {
                    var msg = _conversationHistory[i];
                    string role = msg["role"]?.ToString() ?? "";

                    if (role == "user" && recentMessages.Count == 0)
                    {
                        // 找到最后一条用户消息，保留它
                        recentMessages.Insert(0, msg);
                        // 之前的所有消息都需要压缩
                        for (int j = 0; j < i; j++)
                        {
                            oldMessages.Add(_conversationHistory[j]);
                        }
                        break;
                    }
                }

                if (oldMessages.Count == 0)
                {
                    Debug.WriteLine("没有需要压缩的历史记录");
                    return;
                }

                // 构建压缩请求的对话历史文本
                StringBuilder historyText = new StringBuilder();
                foreach (var msg in oldMessages)
                {
                    string role = msg["role"]?.ToString() ?? "unknown";
                    string content = msg["content"]?.ToString() ?? "";

                    if (role == "user")
                    {
                        historyText.AppendLine($"[用户]: {content}");
                    }
                    else if (role == "assistant")
                    {
                        // 保留助手的文本回复
                        if (!string.IsNullOrEmpty(content))
                        {
                            historyText.AppendLine($"[助手]: {content}");
                        }

                        // 处理工具调用请求（Agent模式的核心）
                        var toolCalls = msg["tool_calls"] as JArray;
                        if (toolCalls != null && toolCalls.Count > 0)
                        {
                            historyText.AppendLine("[工具调用请求]:");
                            foreach (var tc in toolCalls)
                            {
                                string toolName = tc["function"]?["name"]?.ToString() ?? "未知工具";
                                string toolArgs = tc["function"]?["arguments"]?.ToString() ?? "{}";

                                // 根据工具类型决定保留多少参数信息
                                if (IsImportantToolForHistory(toolName))
                                {
                                    // 重要工具（如编辑类工具）保留完整参数摘要
                                    historyText.AppendLine($"  - 工具: {toolName}");
                                    historyText.AppendLine($"    参数: {FormatToolArgumentsForHistory(toolName, toolArgs)}");
                                }
                                else
                                {
                                    // 查询类工具只保留工具名和关键参数
                                    historyText.AppendLine($"  - 工具: {toolName} {FormatToolArgumentsForHistory(toolName, toolArgs)}");
                                }
                            }
                        }
                    }
                    else if (role == "tool")
                    {
                        // 根据工具名和内容长度智能截断
                        string toolName = msg["name"]?.ToString() ?? "未知工具";

                        // 重要工具保留更多信息
                        int maxLength = IsImportantToolForHistory(toolName) ? 1500 : 500;

                        if (content.Length <= maxLength)
                        {
                            historyText.AppendLine($"[工具执行结果 - {toolName}]:");
                            historyText.AppendLine(content);
                        }
                        else
                        {
                            // 尝试智能提取关键信息
                            string summary = ExtractToolResultSummary(toolName, content, maxLength);
                            historyText.AppendLine($"[工具执行结果 - {toolName}]:");
                            historyText.AppendLine(summary);
                            historyText.AppendLine($"(原始结果长度: {content.Length} 字符，已提取关键信息)");
                        }
                    }
                    historyText.AppendLine();
                }

                // 动态计算压缩摘要的 max_tokens（根据上下文窗口大小）
                // 规则：按上下文窗口的 20%~30% 取值（与compress-context提示词一致），并限制范围避免过小/过大
                int compressMaxTokens = 2000; // 默认值（无 context_length 时）
                if (contextLengthTokens > 0)
                {
                    // 小窗口（≤8k）按30%，中窗口（8k-32k）按25%，大窗口（>32k）按20%
                    double ratio = contextLengthTokens <= 8 * 1024 ? 0.30
                                 : contextLengthTokens <= 32 * 1024 ? 0.25
                                 : 0.20;
                    compressMaxTokens = (int)(contextLengthTokens * ratio);
                    // 兜底：压缩摘要不能太短（中文文档更容易丢信息），也不能无限膨胀导致成本/失败
                    compressMaxTokens = Math.Max(1000, Math.Min(20000, compressMaxTokens));
                }
                Debug.WriteLine($"压缩摘要 max_tokens: {compressMaxTokens}（上下文窗口: {contextLengthTokens}, 比例: {(compressMaxTokens * 100.0 / Math.Max(1, contextLengthTokens)):F1}%）");

                // 构建压缩请求
                var compressMessages = new JArray
                {
                    new JObject
                    {
                        ["role"] = "system",
                        ["content"] = compressPrompt
                    },
                    new JObject
                    {
                        ["role"] = "user",
                        ["content"] = historyText.ToString()
                    }
                };

                var compressRequest = new JObject
                {
                    ["model"] = modelName,
                    ["messages"] = compressMessages,
                    ["stream"] = false, // 压缩使用非流式，快速获取结果
                    ["temperature"] = 0.3, // 低温度确保压缩稳定
                    ["max_tokens"] = compressMaxTokens // 动态计算的压缩长度
                };

                Debug.WriteLine("开始调用AI压缩对话历史...");

                // 同步等待压缩结果
                StringBuilder compressedContent = new StringBuilder();
                var compressCts = new CancellationTokenSource();
                compressCts.CancelAfter(TimeSpan.FromSeconds(30)); // 30秒超时

                await OpenAIUtils.OpenAIApiClientAsync(apiUrl, apiKey, compressRequest.ToString(), compressCts.Token,
                    content => compressedContent.Append(content));

                string compressed = compressedContent.ToString();

                if (!string.IsNullOrEmpty(compressed))
                {
                    Debug.WriteLine($"压缩完成，原始记录数: {oldMessages.Count}, 压缩后长度: {compressed.Length}");

                    // 替换历史记录：用压缩摘要 + 最近2轮
                    _conversationHistory.Clear();

                    // 添加压缩后的摘要（作为助手消息）
                    _conversationHistory.Add(new JObject
                    {
                        ["role"] = "assistant",
                        ["content"] = $"【历史对话摘要】\n{compressed}"
                    });

                    // 添加最近的对话
                    foreach (var msg in recentMessages)
                    {
                        _conversationHistory.Add(msg);
                    }

                    Debug.WriteLine($"对话历史已压缩，当前记录数: {_conversationHistory.Count}");
                }
                else
                {
                    Debug.WriteLine("压缩结果为空，保持原历史记录");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"压缩对话历史失败: {ex.Message}");
                // 失败时不影响主流程，保持原历史记录
            }
        }

        /// <summary>
        /// 判断是否是重要工具（需要在压缩历史中保留完整信息）
        /// </summary>
        private bool IsImportantToolForHistory(string toolName)
        {
            // 编辑类工具：这些工具会修改文档，必须保留完整操作记录
            var editingTools = new HashSet<string>
            {
                "formatted_insert_content",    // 格式化插入内容
                "modify_text_style",            // 修改文字样式
                "insert_text_at_cursor",        // 光标处插入文本
                "find_and_replace"              // 查找替换
            };

            // 重要查询类工具：这些查询结果会影响后续操作
            var importantQueryTools = new HashSet<string>
            {
                "get_heading_content",          // 获取标题内容
                "get_selected_text",            // 获取选中文本
                "check_insert_position"         // 检查插入位置
            };

            return editingTools.Contains(toolName) || importantQueryTools.Contains(toolName);
        }

        /// <summary>
        /// 格式化工具参数为易读的历史记录格式
        /// </summary>
        private string FormatToolArgumentsForHistory(string toolName, string argsJson)
        {
            try
            {
                var args = JObject.Parse(argsJson);

                switch (toolName)
                {
                    case "formatted_insert_content":
                        {
                            string targetHeading = args["target_heading"]?.ToString() ?? "未指定";
                            string position = args["position"]?.ToString() ?? "end";
                            string content = args["content"]?.ToString() ?? "";
                            int contentLength = content.Length;

                            // 提取内容主题（前100字符）
                            string contentPreview = contentLength > 0
                                ? content.Substring(0, Math.Min(100, contentLength)).Replace("\n", " ").Replace("\r", "")
                                : "空内容";

                            return $"目标标题=\"{targetHeading}\", 插入位置={position}, 内容长度={contentLength}字符, 内容预览=\"{contentPreview}...\"";
                        }

                    case "modify_text_style":
                        {
                            string findText = args["find_text"]?.ToString() ?? "";
                            var styleChanges = args["style_changes"];

                            string styleDesc = "未指定";
                            if (styleChanges != null)
                            {
                                var styles = new List<string>();
                                if (styleChanges["font_name"] != null) styles.Add($"字体={styleChanges["font_name"]}");
                                if (styleChanges["font_size"] != null) styles.Add($"字号={styleChanges["font_size"]}");
                                if (styleChanges["font_color"] != null) styles.Add($"颜色={styleChanges["font_color"]}");
                                if (styleChanges["bold"] != null) styles.Add($"加粗={styleChanges["bold"]}");
                                if (styleChanges["italic"] != null) styles.Add($"斜体={styleChanges["italic"]}");
                                styleDesc = string.Join(", ", styles);
                            }

                            return $"查找文本=\"{findText}\", 样式修改=[{styleDesc}]";
                        }

                    case "insert_text_at_cursor":
                        {
                            string text = args["text"]?.ToString() ?? "";
                            return $"插入文本=\"{text.Substring(0, Math.Min(100, text.Length))}...\"";
                        }

                    case "find_and_replace":
                        {
                            string find = args["find_text"]?.ToString() ?? "";
                            string replace = args["replace_text"]?.ToString() ?? "";
                            return $"查找=\"{find}\", 替换为=\"{replace}\"";
                        }

                    case "get_heading_content":
                        {
                            string heading = args["heading_text"]?.ToString() ?? "";
                            return $"标题=\"{heading}\"";
                        }

                    case "get_document_headings":
                        {
                            int page = args["page"]?.ToObject<int>() ?? 1;
                            int pageSize = args["page_size"]?.ToObject<int>() ?? 50;
                            return $"分页={page}, 每页={pageSize}条";
                        }

                    case "get_selected_text":
                    case "get_document_statistics":
                    case "get_document_structure":
                        // 这些工具没有关键参数，只返回工具名即可
                        return "";

                    default:
                        // 其他工具：返回精简的JSON（不超过150字符）
                        string jsonStr = args.ToString(Formatting.None);
                        return jsonStr.Length > 150
                            ? jsonStr.Substring(0, 150) + "..."
                            : jsonStr;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"格式化工具参数失败: {ex.Message}");
                return argsJson.Length > 100
                    ? argsJson.Substring(0, 100) + "..."
                    : argsJson;
            }
        }

        /// <summary>
        /// 从工具执行结果中提取摘要信息
        /// </summary>
        private string ExtractToolResultSummary(string toolName, string content, int maxLength)
        {
            try
            {
                // 尝试解析为JSON
                var result = JObject.Parse(content);
                var summary = new StringBuilder();

                // 提取通用字段
                if (result["success"] != null)
                {
                    bool success = result["success"].ToObject<bool>();
                    summary.AppendLine($"执行状态: {(success ? "✅ 成功" : "❌ 失败")}");
                }

                if (result["message"] != null)
                {
                    summary.AppendLine($"消息: {result["message"]}");
                }

                // 根据工具类型提取特定字段
                switch (toolName)
                {
                    case "formatted_insert_content":
                        if (result["target_heading"] != null)
                            summary.AppendLine($"目标标题: {result["target_heading"]}");
                        if (result["position"] != null)
                            summary.AppendLine($"插入位置: {result["position"]}");
                        if (result["inserted_content_length"] != null)
                            summary.AppendLine($"插入内容长度: {result["inserted_content_length"]} 字符");

                        // 保留内容预览
                        string previewContent = result["preview_content"]?.ToString()
                                             ?? result["original_content"]?.ToString()
                                             ?? result["content"]?.ToString();
                        if (!string.IsNullOrEmpty(previewContent))
                        {
                            string preview = previewContent.Substring(0, Math.Min(300, previewContent.Length))
                                .Replace("\n", " ").Replace("\r", "");
                            summary.AppendLine($"内容预览: {preview}...");
                        }
                        break;

                    case "modify_text_style":
                        if (result["find_text"] != null)
                            summary.AppendLine($"目标文本: {result["find_text"]}");
                        if (result["style_parameters"] != null)
                            summary.AppendLine($"样式参数: {result["style_parameters"]}");
                        if (result["preview_styles"] != null)
                            summary.AppendLine($"样式预览: {result["preview_styles"]}");
                        break;

                    case "get_heading_content":
                        if (result["heading_text"] != null)
                            summary.AppendLine($"标题: {result["heading_text"]}");
                        if (result["content_length"] != null)
                            summary.AppendLine($"内容长度: {result["content_length"]} 字符");

                        // 保留部分内容
                        string headingContent = result["content"]?.ToString();
                        if (!string.IsNullOrEmpty(headingContent))
                        {
                            string contentPreview = headingContent.Substring(0, Math.Min(400, headingContent.Length));
                            summary.AppendLine($"内容摘要: {contentPreview}...");
                        }
                        break;

                    case "get_document_headings":
                        if (result["total_count"] != null)
                            summary.AppendLine($"标题总数: {result["total_count"]}");
                        if (result["page"] != null && result["page_size"] != null)
                            summary.AppendLine($"当前页: 第{result["page"]}页, 每页{result["page_size"]}条");

                        // 列出前几个标题
                        var headings = result["headings"] as JArray;
                        if (headings != null && headings.Count > 0)
                        {
                            summary.AppendLine("标题列表（前5个）:");
                            int count = Math.Min(5, headings.Count);
                            for (int i = 0; i < count; i++)
                            {
                                var h = headings[i];
                                summary.AppendLine($"  - [{h["level"]}] {h["text"]}");
                            }
                        }
                        break;

                    case "get_document_statistics":
                        if (result["page_count"] != null)
                            summary.AppendLine($"页数: {result["page_count"]}");
                        if (result["word_count"] != null)
                            summary.AppendLine($"字数: {result["word_count"]}");
                        if (result["paragraph_count"] != null)
                            summary.AppendLine($"段落数: {result["paragraph_count"]}");
                        break;

                    default:
                        // 其他工具：保留完整JSON（如果不超过限制）
                        if (content.Length <= maxLength)
                            return content;
                        break;
                }

                return summary.ToString().TrimEnd();
            }
            catch
            {
                // 如果不是JSON或解析失败，直接截断原文
                string truncated = content.Substring(0, Math.Min(maxLength, content.Length));
                if (content.Length > maxLength)
                    truncated += "\n...(内容已截断)";
                return truncated;
            }
        }

        // 处理Mermaid图片插入
        private void HandleMermaidImageInsertion(JObject jsonObject)
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleMermaidImageInsertion(jsonObject)));
                    return;
                }

                // 获取图片数据
                string imageData = jsonObject["imageData"]?.ToString();
                string svgData = jsonObject["svgData"]?.ToString();
                string mermaidCode = jsonObject["mermaidCode"]?.ToString();
                int width = jsonObject["width"]?.ToObject<int>() ?? 400;
                int height = jsonObject["height"]?.ToObject<int>() ?? 300;

                if (string.IsNullOrEmpty(imageData) && string.IsNullOrEmpty(svgData))
                {
                    Debug.WriteLine("图片数据为空（PNG与SVG都为空）");
                    MessageBox.Show("图片数据为空，无法插入");
                    return;
                }

                Debug.WriteLine($"准备插入Mermaid图片，尺寸: {width}x{height}");
                Debug.WriteLine($"Mermaid代码: {mermaidCode?.Substring(0, Math.Min(50, mermaidCode?.Length ?? 0))}...");

                // 优先插入 SVG（Word 2016+ 为矢量，缩放清晰）；失败再回退 PNG
                bool success = false;
                if (!string.IsNullOrEmpty(svgData))
                {
                    success = _markdownToWord.InsertMermaidImageFromSvg(svgData, mermaidCode, width, height);
                }
                if (!success && !string.IsNullOrEmpty(imageData))
                {
                    success = _markdownToWord.InsertMermaidImage(imageData, mermaidCode, width, height);
                }

                if (success)
                {
                    Debug.WriteLine("Mermaid图片插入成功");
                }
                else
                {
                    Debug.WriteLine("Mermaid图片插入失败");
                    MessageBox.Show("图片插入失败，请重试");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理Mermaid图片插入时出错: {ex.Message}");
                MessageBox.Show($"插入图片时出错: {ex.Message}");
            }
        }

        // 获取当前对话历史数量
        public int GetConversationHistoryCount()
        {
            return _conversationHistory.Count;
        }

        // 获取选中的模型配置
        private Model GetSelectedModel(int modelId)
        {
            try
            {
                if (modelId <= 0)
                    return null;

                return _modelService.GetModelById(modelId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取模型配置失败: {ex.Message}");
                return null;
            }
        }

        // 从Parameters JSON中提取model参数
        private string ExtractModelNameFromParameters(string parametersJson)
        {
            try
            {
                if (string.IsNullOrEmpty(parametersJson))
                    return "gpt-3.5-turbo"; // 默认模型

                var parameters = JObject.Parse(parametersJson);
                string modelName = parameters["model"]?.ToString();

                return !string.IsNullOrEmpty(modelName) ? modelName : "gpt-3.5-turbo";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"解析模型参数失败: {ex.Message}");
                return "gpt-3.5-turbo"; // 解析失败时使用默认模型
            }
        }

        // 获取AI参数：优先从模型配置中解析，其次从应用设置中获取
        private (double temperature, int? maxTokens, double topP) GetAIParameters(Model modelConfig, string chatMode)
        {
            try
            {
                // 首先尝试从模型配置的Parameters中解析
                if (modelConfig != null && !string.IsNullOrEmpty(modelConfig.Parameters))
                {
                    var parameters = JObject.Parse(modelConfig.Parameters);

                    var temperature = parameters["temperature"]?.ToObject<double>();
                    var maxTokens = parameters["max_tokens"]?.ToObject<int>();
                    var topP = parameters["top_p"]?.ToObject<double>();

                    // 如果模型配置中有完整的参数，使用它们
                    if (temperature.HasValue && topP.HasValue)
                    {
                        Debug.WriteLine($"使用模型配置参数: temperature={temperature.Value}, max_tokens={maxTokens?.ToString() ?? "null(自动)"}, top_p={topP.Value}");
                        return (temperature.Value, maxTokens, topP.Value);
                    }

                    // 如果只有部分参数，使用已有的参数并用应用设置补充缺失的
                    var (defaultTemp, defaultMaxTokens, defaultTopP) = _appSettingsService.GetAIParametersByMode(chatMode);

                    var finalTemp = temperature ?? defaultTemp;
                    var finalMaxTokens = maxTokens ?? defaultMaxTokens;
                    var finalTopP = topP ?? defaultTopP;

                    Debug.WriteLine($"使用混合参数(模型+应用设置): temperature={finalTemp}, max_tokens={finalMaxTokens?.ToString() ?? "null(自动)"}, top_p={finalTopP}");
                    return (finalTemp, finalMaxTokens, finalTopP);
                }

                // 如果模型配置中没有参数，使用应用设置中的默认值
                var appSettings = _appSettingsService.GetAIParametersByMode(chatMode);
                Debug.WriteLine($"使用应用设置参数: temperature={appSettings.temperature}, max_tokens={appSettings.maxTokens?.ToString() ?? "null(自动)"}, top_p={appSettings.topP}");
                return appSettings;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"解析AI参数时出错: {ex.Message}，使用应用设置默认值");

                // 出错时使用应用设置中的默认值
                var fallbackSettings = _appSettingsService.GetAIParametersByMode(chatMode);
                return fallbackSettings;
            }
        }

        // 根据聊天模式获取系统提示词
        private string GetSystemPromptByMode(string chatMode, string[] enabledTools = null, List<JObject> contexts = null)
        {
            string basePrompt = PromptService.GetPrompt(chatMode);

            // 构建上下文内容
            string contextContent = "";

            // 如果有已加载的Word文档内容，优先添加到上下文（只在系统提示词中出现一次，节省Token）
            if (!string.IsNullOrEmpty(_loadedDocumentContent) && !_loadedDocumentContent.StartsWith("错误"))
            {
                contextContent += "### 当前Word文档完整内容\n";
                contextContent += _loadedDocumentContent + "\n\n";
                Debug.WriteLine($"已将Word文档内容添加到系统提示词，字符数: {_loadedDocumentContent.Length}");
            }
            if (contexts != null && contexts.Count > 0)
            {
                int contextIndex = 1;
                foreach (var context in contexts)
                {
                    string contextType = context["type"]?.ToString();

                    if (contextType == "document")
                    {
                        // 文档类型的上下文
                        int documentId = context["id"]?.ToObject<int>() ?? 0;
                        string fileName = context["name"]?.ToString() ?? "未知文档";

                        // 获取文档的所有标题内容
                        var headings = _documentService.GetDocumentHeadings(documentId);
                        if (headings.Any())
                        {
                            contextContent += $"### {contextIndex}. 文档：{fileName}\n";
                            foreach (var heading in headings.OrderBy(h => h.OrderIndex))
                            {
                                if (!string.IsNullOrWhiteSpace(heading.Content))
                                {
                                    contextContent += $"#### {heading.HeadingText}\n{heading.Content}\n\n";
                                }
                            }
                        }
                    }
                    else if (contextType == "heading")
                    {
                        // 标题类型的上下文
                        int documentId = context["documentId"]?.ToObject<int>() ?? 0;
                        int headingId = context["headingId"]?.ToObject<int>() ?? 0;
                        string headingText = context["name"]?.ToString() ?? "未知标题";
                        string documentName = context["documentName"]?.ToString() ?? "未知文档";

                        // 获取特定标题的内容
                        var headings = _documentService.GetDocumentHeadings(documentId);
                        var targetHeading = headings.FirstOrDefault(h => h.Id == headingId);

                        if (targetHeading != null && !string.IsNullOrWhiteSpace(targetHeading.Content))
                        {
                            contextContent += $"### {contextIndex}. 标题：{headingText} (来自文档：{documentName})\n";
                            contextContent += $"{targetHeading.Content}\n\n";
                        }
                    }

                    contextIndex++;
                }
            }

            // 替换占位符
            basePrompt = basePrompt.Replace("${{docs}}", contextContent);

            // 当未选择任何上下文时，给出显式提示，便于模型理解与日志排查
            if (string.IsNullOrWhiteSpace(contextContent))
            {
                basePrompt += "\n\n（提示：当前无参考文档，将基于模型自身知识回答）";
            }

            switch (chatMode)
            {
                case "chat-agent":
                    // 为智能体模式动态添加工具信息
                    if (enabledTools != null && enabledTools.Length > 0)
                    {
                        basePrompt += "\n\n6. 使用以下用户启用的Word操作工具：\n";
                        foreach (string tool in enabledTools)
                        {
                            string toolDesc = GetToolDescription(tool);
                            if (!string.IsNullOrEmpty(toolDesc))
                            {
                                basePrompt += $"   - {tool}: {toolDesc}\n";
                            }
                        }
                        basePrompt += $"当前共启用了 {enabledTools.Length} 个工具。";
                    }
                    else
                    {
                        // 默认显示所有工具
                        basePrompt += "\n\n6. 使用以下Word操作工具：\n" +
                                      "   - check_insert_position: 检查插入位置和获取上下文\n" +
                                      "   - get_selected_text: 获取选中的文本\n" +
                                      "   - formatted_insert_content: 插入格式化内容\n" +
                                      "   - modify_text_style: 修改文本样式\n" +
                                      "   - get_document_statistics: 获取文档统计信息\n" +
                                      "   - get_document_images: 获取文档中的所有图片信息\n" +
                                      "   - get_document_formulas: 获取数学公式的位置和数量统计\n" +
                                      "   - get_document_tables: 获取文档中的表格信息\n" +
                                      "   - get_document_headings: 获取文档的标题列表（整体结构）\n" +
                                      "   - get_heading_content: 高效获取指定标题下的所有内容\n";
                    }
                    break;

                case "chat":
                default:
                    // Chat模式直接使用基础提示词，不需要额外处理
                    break;
            }

            // 输出最终系统提示词到调试日志，便于排查与核对（包含已替换的上下文内容）
            try
            {
                Debug.WriteLine("=== 最终系统提示词（已替换）开始 ===");
                Debug.WriteLine(basePrompt);
                Debug.WriteLine("=== 最终系统提示词（已替换）结束 ===");
            }
            catch { }

            return basePrompt;
        }

        // 复制内容到Word
        private void CopyContentToWord(string htmlContent)
        {
            try
            {
                if (string.IsNullOrEmpty(htmlContent))
                {
                    Debug.WriteLine("复制到Word的内容为空");
                    return;
                }

                // 获取公式信息
                JObject jsonObject = JObject.Parse(htmlContent);
                List<JObject> formulas = null;

                // 如果是从JavaScript发送的JSON，解析公式列表
                if (jsonObject != null && jsonObject["formulas"] != null)
                {
                    formulas = jsonObject["formulas"].ToObject<List<JObject>>();
                    htmlContent = jsonObject["content"].ToString();
                }

                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => CopyContentToWord(htmlContent, formulas)));
                    return;
                }

                CopyContentToWord(htmlContent, formulas);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"解析复制内容时出错: {ex.Message}");
                // 尝试使用原始HTML直接复制
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => CopyContentToWord(htmlContent, null)));
                }
                else
                {
                    CopyContentToWord(htmlContent, null);
                }
            }
        }

        // 重载方法，支持公式处理
        private void CopyContentToWord(string htmlContent, List<JObject> formulas)
        {
            try
            {
                // 使用MarkdownToWord工具类插入HTML内容
                _markdownToWord.InsertHtmlContent(htmlContent, formulas);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制到Word时出错: {ex.Message}");
                Debug.WriteLine($"复制到Word时出错详细信息: {ex}");
            }
        }

        // 处理顺序插入请求
        private void HandleSequenceInsertion(JObject jsonObject)
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleSequenceInsertion(jsonObject)));
                    return;
                }

                // 获取插入项数组
                var items = jsonObject["items"]?.ToObject<List<JObject>>();
                if (items == null || items.Count == 0)
                {
                    Debug.WriteLine("插入项为空");
                    return;
                }

                Debug.WriteLine($"开始处理 {items.Count} 个插入项");

                // 获取Word应用程序和选择区域，确保在文档末尾开始插入
                dynamic wordApp = _markdownToWord.GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法获取Word应用程序");
                    MessageBox.Show("无法连接到Word应用程序，请确保Word已打开");
                    return;
                }

                dynamic doc = _markdownToWord.GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    MessageBox.Show("无法获取活动的Word文档");
                    return;
                }

                dynamic selection = wordApp.Selection;

                // 保持当前光标位置，不移动到文档末尾
                Debug.WriteLine($"保持当前光标位置: {selection.Start}");

                // 如果当前在行中间，添加换行开始新行
                try
                {
                    var currentPara = selection.Paragraphs[1];
                    if (currentPara.Range.Text.Trim().Length > 0 && selection.Start > currentPara.Range.Start)
                    {
                        Debug.WriteLine("在行中间，添加换行开始新行");
                        selection.TypeText("\r\n");
                    }
                }
                catch (Exception posEx)
                {
                    Debug.WriteLine($"检查插入位置时出错: {posEx.Message}");
                    // 如果检查失败，确保至少有一个换行
                    try
                    {
                        if (selection.Start > 0)
                        {
                            selection.TypeText("\r\n");
                        }
                    }
                    catch { }
                }

                foreach (var item in items)
                {
                    string itemType = item["type"]?.ToString();
                    string content = item["content"]?.ToString();

                    Debug.WriteLine($"=== 处理项: {itemType} ===");

                    // 对于复杂类型的内容，需要特殊处理
                    if (itemType == "mermaidImage")
                    {
                        try
                        {
                            // 处理Mermaid图片（PNG格式）
                            Debug.WriteLine("处理Mermaid图片（PNG格式）");

                            // 获取内容对象
                            var mermaidContent = item["content"] as JObject;
                            if (mermaidContent != null)
                            {
                                string imageData = mermaidContent["imageData"]?.ToString();
                                string svgData = mermaidContent["svgData"]?.ToString();
                                string mermaidCode = mermaidContent["mermaidCode"]?.ToString();
                                int width = mermaidContent["width"]?.ToObject<int>() ?? 400;
                                int height = mermaidContent["height"]?.ToObject<int>() ?? 300;

                                Debug.WriteLine($"Mermaid图片数据长度: {imageData?.Length ?? 0}");
                                Debug.WriteLine($"Mermaid图片尺寸: {width}x{height}");

                                if (!string.IsNullOrEmpty(imageData) || !string.IsNullOrEmpty(svgData))
                                {
                                    // 优先插入 SVG（矢量，缩放清晰）；失败再回退 PNG
                                    bool success = false;
                                    if (!string.IsNullOrEmpty(svgData))
                                    {
                                        success = _markdownToWord.InsertMermaidImageFromSvg(svgData, mermaidCode, width, height);
                                    }
                                    if (!success && !string.IsNullOrEmpty(imageData))
                                    {
                                        success = _markdownToWord.InsertMermaidImage(imageData, mermaidCode, width, height);
                                    }
                                    if (success)
                                    {
                                        Debug.WriteLine("Mermaid PNG图片插入成功");

                                        // 在流程图后添加一个换行，确保与后续内容分隔
                                        try
                                        {
                                            Debug.WriteLine("在流程图后添加换行");
                                            selection.TypeText("\r\n");
                                        }
                                        catch (Exception lineEx)
                                        {
                                            Debug.WriteLine($"添加换行时出错: {lineEx.Message}");
                                        }
                                    }
                                    else
                                    {
                                        Debug.WriteLine("Mermaid图片插入失败（SVG/PNG都失败），回退到代码块");
                                        // 回退到代码块
                                        if (!string.IsNullOrEmpty(mermaidCode))
                                        {
                                            _markdownToWord.InsertCode(mermaidCode);

                                            // 代码块后也添加换行
                                            try
                                            {
                                                Debug.WriteLine("在代码块后添加换行");
                                                selection.TypeText("\r\n");
                                            }
                                            catch (Exception lineEx)
                                            {
                                                Debug.WriteLine($"添加换行时出错: {lineEx.Message}");
                                            }
                                        }
                                    }
                                    continue;
                                }
                                else
                                {
                                    Debug.WriteLine("Mermaid图片数据为空（PNG与SVG都为空），回退到代码块");
                                    // 回退到代码块
                                    if (!string.IsNullOrEmpty(mermaidCode))
                                    {
                                        _markdownToWord.InsertCode(mermaidCode);

                                        // 代码块后也添加换行
                                        try
                                        {
                                            Debug.WriteLine("在代码块后添加换行");
                                            selection.TypeText("\r\n");
                                        }
                                        catch (Exception lineEx)
                                        {
                                            Debug.WriteLine($"添加换行时出错: {lineEx.Message}");
                                        }
                                    }
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"处理Mermaid图片时出错: {ex.Message}");
                            // 继续处理其他项目
                        }
                    }
                    else if (itemType == "html")
                    {
                        // 插入HTML内容，保持格式
                        if (!string.IsNullOrEmpty(content))
                        {
                            Debug.WriteLine("调用InsertHtmlContentDirect");
                            _markdownToWord.InsertHtmlContentDirect(content);
                        }
                        else
                        {
                            Debug.WriteLine("HTML内容为空，跳过");
                        }
                    }
                    else if (itemType == "text")
                    {
                        // 插入文本
                        if (!string.IsNullOrEmpty(content))
                        {
                            Debug.WriteLine("调用InsertText");
                            _markdownToWord.InsertText(content);
                        }
                        else
                        {
                            Debug.WriteLine("文本内容为空，跳过");
                        }
                    }
                    else if (itemType == "formula")
                    {
                        // 插入公式
                        if (!string.IsNullOrEmpty(content))
                        {
                            Debug.WriteLine($"调用InsertEquation，公式: {content}");
                            _markdownToWord.InsertEquation(content);
                        }
                        else
                        {
                            Debug.WriteLine("公式内容为空，跳过");
                        }
                    }
                    else if (itemType == "code")
                    {
                        // 插入代码
                        if (!string.IsNullOrEmpty(content))
                        {
                            Debug.WriteLine($"调用InsertCode，代码: {content.Substring(0, Math.Min(50, content.Length))}...");
                            _markdownToWord.InsertCode(content);
                        }
                        else
                        {
                            Debug.WriteLine("代码内容为空，跳过");
                        }
                    }
                    else if (itemType == "mermaid")
                    {
                        // 插入Mermaid图表（作为代码块处理）
                        if (!string.IsNullOrEmpty(content))
                        {
                            Debug.WriteLine($"调用InsertCode处理Mermaid图表: {content.Substring(0, Math.Min(50, content.Length))}...");
                            _markdownToWord.InsertCode(content);

                            // 在Mermaid代码块后添加换行
                            try
                            {
                                Debug.WriteLine("在Mermaid代码块后添加换行");
                                selection.TypeText("\r\n");
                            }
                            catch (Exception lineEx)
                            {
                                Debug.WriteLine($"添加换行时出错: {lineEx.Message}");
                            }
                        }
                        else
                        {
                            Debug.WriteLine("Mermaid图表内容为空，跳过");
                        }
                    }
                    else if (itemType == "mermaidImagePending")
                    {
                        // 处理待处理的Mermaid图片（应该已经在JavaScript端处理完成，不应该出现在这里）
                        Debug.WriteLine("收到mermaidImagePending类型，这不应该发生，回退到代码块处理");

                        // 尝试获取代码并作为代码块插入
                        try
                        {
                            var pendingContent = item["content"] as JObject;
                            if (pendingContent != null)
                            {
                                string mermaidCode = pendingContent["mermaidCode"]?.ToString();
                                if (!string.IsNullOrEmpty(mermaidCode))
                                {
                                    _markdownToWord.InsertCode(mermaidCode);

                                    // 在Mermaid代码块后添加换行
                                    try
                                    {
                                        Debug.WriteLine("在Mermaid代码块后添加换行");
                                        selection.TypeText("\r\n");
                                    }
                                    catch (Exception lineEx)
                                    {
                                        Debug.WriteLine($"添加换行时出错: {lineEx.Message}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"处理mermaidImagePending时出错: {ex.Message}");
                        }
                    }
                    else if (itemType == "table")
                    {
                        // 插入表格
                        if (item["content"] != null)
                        {
                            Debug.WriteLine($"调用InsertTable，表格数据: {item["content"]}");
                            _markdownToWord.InsertTable(item["content"]);
                        }
                        else
                        {
                            Debug.WriteLine("表格内容为空，跳过");
                        }
                    }
                    else if (itemType == "linebreak")
                    {
                        // 插入换行 - 现在减少不必要的换行
                        Debug.WriteLine("跳过换行插入（已优化）");
                        // _markdownToWord.InsertLineBreak();
                    }
                    else
                    {
                        Debug.WriteLine($"未知的插入项类型: {itemType}");
                    }

                    Debug.WriteLine($"=== 完成处理项: {itemType} ===");
                }

                Debug.WriteLine("顺序插入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理顺序插入时出错: {ex.Message}");
                Debug.WriteLine($"错误堆栈: {ex.StackTrace}");
                MessageBox.Show($"插入内容到Word时出错: {ex.Message}");
            }
        }

        // 向WebView发送Markdown内容
        public void SendMarkdownToWebView(string userMessage, string markdown)
        {
            // 确保在UI线程执行
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SendMarkdownToWebView(userMessage, markdown)));
                return;
            }

            try
            {
                if (!_webViewInitialized || webView21 == null || webView21.CoreWebView2 == null)
                {
                    Debug.WriteLine("WebView2尚未初始化，无法发送消息");
                    return;
                }

                // 创建消息对象
                var messageData = new
                {
                    type = "markdownContent",
                    message = userMessage,
                    markdown = markdown
                };

                // 使用Newtonsoft.Json序列化JSON
                string json = JsonConvert.SerializeObject(messageData);

                // 发送到WebView
                webView21.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"向WebView发送消息时出错: {ex.Message}");
                // 在严重错误情况下尝试重新初始化WebView
                if (ex.Message.Contains("E_NOINTERFACE") || ex.Message.Contains("无法将类型"))
                {
                    MessageBox.Show("WebView2发生错误，请重新启动应用");
                }
            }
        }

        // 向WebView发送状态更新
        public void SendStatusToWebView(string status)
        {
            // 确保在UI线程执行
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SendStatusToWebView(status)));
                return;
            }

            try
            {
                if (!_webViewInitialized || webView21 == null || webView21.CoreWebView2 == null)
                    return;

                var messageData = new
                {
                    type = "status",
                    message = status
                };

                // 使用Newtonsoft.Json序列化JSON
                string json = JsonConvert.SerializeObject(messageData);
                webView21.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送状态更新时出错: {ex.Message}");
            }
        }

        // 辅助方法：将TeX公式转换为Word可用的线性格式
        private string ConvertTeXToLinearFormat(string texFormula)
        {
            try
            {
                // 移除可能的外部包装符号
                texFormula = texFormula.Trim();
                if (texFormula.StartsWith("$$") && texFormula.EndsWith("$$"))
                {
                    texFormula = texFormula.Substring(2, texFormula.Length - 4).Trim();
                }
                else if (texFormula.StartsWith("$") && texFormula.EndsWith("$"))
                {
                    texFormula = texFormula.Substring(1, texFormula.Length - 2).Trim();
                }

                // 特殊公式直接处理
                // 检查是否是二次方程公式
                if ((texFormula.Contains("-b") || texFormula.Contains("\\frac{-b")) &&
                    texFormula.Contains("\\sqrt{b^2-4ac}") &&
                    texFormula.Contains("2a"))
                {
                    // 二次方程式特殊处理
                    return "(-b±√(b^2-4ac))/(2a)";
                }

                // 替换常见的TeX命令为Word线性格式
                Dictionary<string, string> replacements = new Dictionary<string, string>
                {
                    // 分数
                    {@"\\frac{(.*?)}{(.*?)}", @"($1)/($2)"},
                    {@"\\dfrac{(.*?)}{(.*?)}", @"($1)/($2)"}, 
                    
                    // 上标和下标
                    {@"_{(.*?)}", @"_($1)"},
                    {@"\^{(.*?)}", @"^($1)"},
                    {@"_([a-zA-Z0-9])", @"_$1"},
                    {@"\^([a-zA-Z0-9])", @"^$1"},
                    
                    // 希腊字母（小写）
                    {@"\\alpha", @"α"},
                    {@"\\beta", @"β"},
                    {@"\\gamma", @"γ"},
                    {@"\\delta", @"δ"},
                    {@"\\epsilon", @"ε"},
                    {@"\\varepsilon", @"ε"},
                    {@"\\zeta", @"ζ"},
                    {@"\\eta", @"η"},
                    {@"\\theta", @"θ"},
                    {@"\\vartheta", @"ϑ"},
                    {@"\\iota", @"ι"},
                    {@"\\kappa", @"κ"},
                    {@"\\lambda", @"λ"},
                    {@"\\mu", @"μ"},
                    {@"\\nu", @"ν"},
                    {@"\\xi", @"ξ"},
                    {@"\\pi", @"π"},
                    {@"\\varpi", @"ϖ"},
                    {@"\\rho", @"ρ"},
                    {@"\\varrho", @"ϱ"},
                    {@"\\sigma", @"σ"},
                    {@"\\varsigma", @"ς"},
                    {@"\\tau", @"τ"},
                    {@"\\upsilon", @"υ"},
                    {@"\\phi", @"φ"},
                    {@"\\varphi", @"φ"},
                    {@"\\chi", @"χ"},
                    {@"\\psi", @"ψ"},
                    {@"\\omega", @"ω"},
                    
                    // 希腊字母（大写）
                    {@"\\Gamma", @"Γ"},
                    {@"\\Delta", @"Δ"},
                    {@"\\Theta", @"Θ"},
                    {@"\\Lambda", @"Λ"},
                    {@"\\Xi", @"Ξ"},
                    {@"\\Pi", @"Π"},
                    {@"\\Sigma", @"Σ"},
                    {@"\\Upsilon", @"Υ"},
                    {@"\\Phi", @"Φ"},
                    {@"\\Psi", @"Ψ"},
                    {@"\\Omega", @"Ω"},
                    
                    // 根号和乘方
                    {@"\\sqrt{(.*?)}", @"√($1)"},
                    {@"\\sqrt\[(.*?)\]{(.*?)}", @"root($1)($2)"},
                    
                    // 其他常见符号
                    {@"\\pm", @"±"},
                    {@"\\mp", @"∓"},
                    {@"\\infty", @"∞"},
                    {@"\\propto", @"∝"},
                    {@"\\partial", @"∂"},
                    {@"\\simeq", @"≃"},
                    {@"\\equiv", @"≡"},
                    {@"\\sim", @"∼"},
                    {@"\\cdot", @"·"},
                    {@"\\div", @"÷"},
                    {@"\\times", @"×"},
                    {@"\\leq", @"≤"},
                    {@"\\geq", @"≥"},
                    {@"\\le", @"≤"},
                    {@"\\ge", @"≥"},
                    {@"\\ll", @"≪"},
                    {@"\\gg", @"≫"},
                    {@"\\ne", @"≠"},
                    {@"\\neq", @"≠"},
                    {@"\\approx", @"≈"},
                    {@"\\perp", @"⊥"},
                    {@"\\parallel", @"∥"},
                    {@"\\mid", @"|"},
                    {@"\\bullet", @"•"},
                    {@"\\cap", @"∩"},
                    {@"\\cup", @"∪"},
                    {@"\\subset", @"⊂"},
                    {@"\\supset", @"⊃"},
                    {@"\\subseteq", @"⊆"},
                    {@"\\supseteq", @"⊇"},
                    {@"\\in", @"∈"},
                    {@"\\ni", @"∋"},
                    {@"\\notin", @"∉"},
                    
                    // 箭头
                    {@"\\leftarrow", @"←"},
                    {@"\\rightarrow", @"→"},
                    {@"\\Leftarrow", @"⇐"},
                    {@"\\Rightarrow", @"⇒"},
                    {@"\\leftrightarrow", @"↔"},
                    {@"\\Leftrightarrow", @"⇔"},
                    {@"\\mapsto", @"↦"},
                    
                    // 积分、求和、乘积
                    {@"\\int", @"∫"},
                    {@"\\iint", @"∬"},
                    {@"\\iiint", @"∭"},
                    {@"\\oint", @"∮"},
                    {@"\\sum", @"∑"},
                    {@"\\prod", @"∏"},
                    {@"\\coprod", @"∐"},
                    
                    // 括号处理
                    {@"\\left\(", @"("},
                    {@"\\right\)", @")"},
                    {@"\\left\[", @"["},
                    {@"\\right\]", @"]"},
                    {@"\\left\{", @"{"},
                    {@"\\right\}", @"}"},
                    {@"\\left\|", @"|"},
                    {@"\\right\|", @"|"},
                    
                    // 常见函数
                    {@"\\sin", @"sin"},
                    {@"\\cos", @"cos"},
                    {@"\\tan", @"tan"},
                    {@"\\cot", @"cot"},
                    {@"\\sec", @"sec"},
                    {@"\\csc", @"csc"},
                    {@"\\arcsin", @"arcsin"},
                    {@"\\arccos", @"arccos"},
                    {@"\\arctan", @"arctan"},
                    {@"\\sinh", @"sinh"},
                    {@"\\cosh", @"cosh"},
                    {@"\\tanh", @"tanh"},
                    {@"\\coth", @"coth"},
                    {@"\\log", @"log"},
                    {@"\\ln", @"ln"},
                    {@"\\exp", @"exp"},
                    {@"\\lim", @"lim"},
                    {@"\\max", @"max"},
                    {@"\\min", @"min"},
                    
                    // 二次方程公式
                    {@"-b\s*\\pm\s*\\sqrt{b\^2-4ac}\\over 2a", @"(-b±√(b^2-4ac))/(2a)"},
                    {@"\\frac{-b\s*\\pm\s*\\sqrt{b\^2-4ac}}{2a}", @"(-b±√(b^2-4ac))/(2a)"},
                    {@"\\frac{-b\s*\\pm\s*\\sqrt{b\^2\s*-\s*4ac}}{2a}", @"(-b±√(b^2-4ac))/(2a)"}
                };

                string result = texFormula;

                // 先处理嵌套括号的问题
                // 统一处理括号
                result = result.Replace("\\left(", "(")
                              .Replace("\\right)", ")")
                              .Replace("\\left[", "[")
                              .Replace("\\right]", "]")
                              .Replace("\\left{", "{")
                              .Replace("\\right}", "}")
                              .Replace("\\{", "{")
                              .Replace("\\}", "}");

                // 应用正则表达式替换
                foreach (var replacement in replacements)
                {
                    result = Regex.Replace(result, replacement.Key, replacement.Value);
                }

                // 最终的清理和多重处理
                // 1. 重复处理可能未被第一轮处理的复杂表达式
                if (result.Contains("frac"))
                {
                    Debug.WriteLine("检测到未处理完的分数表达式，进行再次处理");
                    foreach (var replacement in replacements)
                    {
                        if (replacement.Key.Contains("frac"))
                        {
                            result = Regex.Replace(result, replacement.Key, replacement.Value);
                        }
                    }
                }

                // 2. 处理不必要的括号
                result = result.Replace("((", "(")
                              .Replace("))", ")")
                              .Replace(")(", ")·(");

                // 3. 特殊处理二次方程公式
                if (result.Contains("-b") && result.Contains("sqrt") &&
                    (result.Contains("b^2-4ac") || result.Contains("b^2 - 4ac")))
                {
                    Debug.WriteLine("再次检测到二次方程公式，强制使用标准格式");
                    result = "(-b±√(b^2-4ac))/(2a)";
                }

                Debug.WriteLine($"转换公式结果: {result}");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"公式转换出错: {ex.Message}");
                // 返回原始公式
                return texFormula;
            }
        }

        // 向WebView发送欢迎消息的Markdown内容
        private void SendWelcomeMessageToWebView()
        {
            // 确保在UI线程执行
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SendWelcomeMessageToWebView()));
                return;
            }

            try
            {
                if (!_webViewInitialized || webView21 == null || webView21.CoreWebView2 == null)
                    return;

                // 获取当前聊天模式
                string chatMode = "chat"; // 默认模式

                // 获取欢迎消息的Markdown内容
                string welcomeMarkdown = GetWelcomeMessageMarkdown(chatMode);

                // 构建消息对象
                var messageData = new
                {
                    type = "setWelcomeMessage",
                    content = welcomeMarkdown,
                    format = "markdown" // 指明内容格式为Markdown
                };

                // 使用Newtonsoft.Json序列化JSON
                string json = JsonConvert.SerializeObject(messageData);

                // 发送到WebView
                webView21.CoreWebView2.PostWebMessageAsJson(json);

                Debug.WriteLine("欢迎消息(Markdown格式)已发送到前端");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送欢迎消息时出错: {ex.Message}");
            }
        }

        // 获取欢迎消息的Markdown内容
        public string GetWelcomeMessageMarkdown(string chatMode = "chat")
        {
            // 从PromptService获取欢迎消息
            string welcomeMessage = PromptService.GetPrompt("welcome");

            // 如果是智能体模式，可以添加额外的内容
            if (chatMode == "chat-agent")
            {
                welcomeMessage += "\n\n我还可以作为智能代理帮助您完成更复杂的任务，如文档规划、内容组织和多步骤操作。";
            }

            return welcomeMessage;
        }

        // 向WebView发送模型列表
        private void SendModelListToWebView()
        {
            // 确保在UI线程执行
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SendModelListToWebView()));
                return;
            }

            try
            {
                if (!_webViewInitialized || webView21 == null || webView21.CoreWebView2 == null)
                    return;

                // 获取所有对话模型（modelType = 1）
                var allModels = _modelService.GetAllModels();
                var chatModels = allModels.Where(m => m.modelType == 1).ToList();

                // 构建模型选择数据
                var modelOptions = chatModels.Select(model => new
                {
                    id = model.Id,
                    name = model.NickName,
                    template = model.Template?.TemplateName ?? "未知",
                    baseUrl = model.BaseUrl,
                    modelName = ExtractModelNameFromParameters(model.Parameters),
                    enableTools = model.EnableTools, // 包含工具调用启用状态
                    enableMulti = model.EnableMulti  // 0:禁用,1:多模态启用
                }).ToList();

                // 构建消息对象
                var messageData = new
                {
                    type = "modelList",
                    models = modelOptions
                };

                // 使用Newtonsoft.Json序列化JSON
                string json = JsonConvert.SerializeObject(messageData);

                // 发送到WebView
                webView21.CoreWebView2.PostWebMessageAsJson(json);

                Debug.WriteLine($"模型列表已发送到前端，共 {modelOptions.Count} 个对话模型");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送模型列表时出错: {ex.Message}");
            }
        }

        // 处理应用预览操作请求
        private async void HandleApplyPreviewedAction(JObject data)
        {
            string actionType = ""; // 提升到方法级别作用域
            string previewId = ""; // 提升到方法级别作用域
            try
            {
                Debug.WriteLine("开始处理应用预览操作...");

                actionType = data["action_type"]?.ToString();
                JObject parameters = data["parameters"] as JObject;
                previewId = data["preview_id"]?.ToString(); // 获取预览ID

                if (string.IsNullOrEmpty(actionType) || parameters == null)
                {
                    Debug.WriteLine("应用预览操作：参数无效");
                    return;
                }

                Debug.WriteLine($"处理预览操作: {actionType}, 预览ID: {previewId}");

                // 将参数转换为Dictionary
                var paramDict = parameters.ToObject<Dictionary<string, object>>();
                if (paramDict != null)
                {
                    // 设置为实际执行模式
                    paramDict["preview_only"] = false;

                    string result = "";
                    switch (actionType)
                    {
                        case "insert_content":
                            result = await TaskAsync.Run(() => WordTools.ExecuteFormattedInsertDirectly(paramDict));
                            break;
                        case "modify_style":
                            // 样式修改
                            result = await TaskAsync.Run(() => WordTools.ExecuteStyleModificationDirectly(paramDict));
                            break;

                        default:
                            Debug.WriteLine($"未知的操作类型: {actionType}");
                            return;
                    }

                    // 解析结果并发送到前端
                    Debug.WriteLine($"尝试解析结果JSON: {result}");

                    // 检查结果是否为有效的JSON格式
                    if (string.IsNullOrEmpty(result))
                    {
                        Debug.WriteLine("错误：返回结果为空");
                        var errorData = new
                        {
                            type = "actionApplied",
                            success = false,
                            action_type = actionType,
                            message = "操作返回空结果"
                        };

                        if (InvokeRequired)
                        {
                            Invoke(new Action(() => { SendActionResultToWebView(errorData, false); }));
                        }
                        else
                        {
                            SendActionResultToWebView(errorData, false);
                        }
                        return;
                    }

                    // 检查是否以有效的JSON字符开始
                    string trimmedResult = result.Trim();
                    if (!trimmedResult.StartsWith("{") && !trimmedResult.StartsWith("["))
                    {
                        Debug.WriteLine($"错误：返回结果不是有效的JSON格式，内容: {result}");
                        var errorData = new
                        {
                            type = "actionApplied",
                            success = false,
                            action_type = actionType,
                            message = $"操作返回格式错误: {result}"
                        };

                        if (InvokeRequired)
                        {
                            Invoke(new Action(() => { SendActionResultToWebView(errorData, false); }));
                        }
                        else
                        {
                            SendActionResultToWebView(errorData, false);
                        }
                        return;
                    }

                    var resultData = JsonConvert.DeserializeObject<dynamic>(result);
                    bool success = resultData.success == true;

                    var responseData = new
                    {
                        type = "actionApplied",
                        success = success,
                        action_type = actionType,
                        preview_id = previewId, // 包含预览ID
                        message = resultData.message?.ToString() ?? (success ? "操作成功应用" : "操作应用失败")
                    };

                    // 确保在UI线程执行WebView操作
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => {
                            SendActionResultToWebView(responseData, success);
                            // 继续与AI对话，告知操作结果
                            ContinueConversationAfterUserAction(actionType, success, resultData.message?.ToString());
                        }));
                    }
                    else
                    {
                        SendActionResultToWebView(responseData, success);
                        // 继续与AI对话，告知操作结果
                        ContinueConversationAfterUserAction(actionType, success, resultData.message?.ToString());
                    }
                }
            }
            catch (JsonReaderException jsonEx)
            {
                Debug.WriteLine($"JSON解析异常: {jsonEx.Message}");
                Debug.WriteLine($"异常详细信息: {jsonEx}");

                // 发送JSON解析错误信息到前端
                var errorData = new
                {
                    type = "actionApplied",
                    success = false,
                    action_type = actionType,
                    preview_id = previewId,
                    message = $"JSON解析失败: {jsonEx.Message}"
                };

                // 确保在UI线程执行WebView操作
                if (InvokeRequired)
                {
                    Invoke(new Action(() => {
                        SendActionResultToWebView(errorData, false);
                    }));
                }
                else
                {
                    SendActionResultToWebView(errorData, false);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理应用预览操作时出错: {ex.Message}");
                Debug.WriteLine($"异常详细信息: {ex}");

                // 发送错误信息到前端
                var errorData = new
                {
                    type = "actionApplied",
                    success = false,
                    action_type = actionType,
                    preview_id = previewId,
                    message = $"应用操作失败: {ex.Message}"
                };

                // 确保在UI线程执行WebView操作
                if (InvokeRequired)
                {
                    Invoke(new Action(() => {
                        SendActionResultToWebView(errorData, false);
                    }));
                }
                else
                {
                    SendActionResultToWebView(errorData, false);
                }
            }
        }

        // 处理来自OpenAI工具的预览事件
        private void HandleToolPreviewFromOpenAI(JObject previewData)
        {
            try
            {
                Debug.WriteLine("收到工具预览事件");

                // 确保在UI线程执行
                if (InvokeRequired)
                {
                    Invoke(new Action(() => HandleToolPreviewFromOpenAI(previewData)));
                    return;
                }

                // 发送预览数据到前端
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var messageData = new
                    {
                        type = "toolPreview",
                        success = previewData["success"]?.ToObject<bool>() ?? false,
                        preview_mode = previewData["preview_mode"]?.ToObject<bool>() ?? false,
                        action_type = previewData["action_type"]?.ToString() ?? "",
                        target_heading = previewData["target_heading"]?.ToString() ?? "",
                        format_type = previewData["format_type"]?.ToString() ?? "",
                        indent_level = previewData["indent_level"]?.ToObject<int>() ?? 0,
                        add_spacing = previewData["add_spacing"]?.ToObject<bool>() ?? true,
                        original_content = previewData["original_content"]?.ToString() ?? "",
                        preview_content = previewData["preview_content"]?.ToString() ?? "",
                        text_to_find = previewData["text_to_find"]?.ToString() ?? "",
                        preview_styles = previewData["preview_styles"]?.ToObject<string[]>() ?? new string[0],
                        style_parameters = previewData["style_parameters"] ?? new JObject(),
                        message = previewData["message"]?.ToString() ?? ""
                    };

                    string json = JsonConvert.SerializeObject(messageData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine("预览数据已发送到前端");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理工具预览事件时出错: {ex.Message}");
            }
        }

        // 处理来自OpenAI工具的进度事件
        private void HandleToolProgressFromOpenAI(string progressMessage)
        {
            try
            {
                Debug.WriteLine($"收到工具进度事件: {progressMessage}");

                // 确保在UI线程执行
                if (InvokeRequired)
                {
                    Invoke(new Action(() => HandleToolProgressFromOpenAI(progressMessage)));
                    return;
                }

                // 发送进度数据到前端
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var messageData = new
                    {
                        type = "toolProgress",
                        message = progressMessage,
                        timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
                    };

                    string json = JsonConvert.SerializeObject(messageData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine("工具进度已发送到前端");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理工具进度事件时出错: {ex.Message}");
            }
        }

        // 处理来自 OpenAI 的 usage 事件（token 用量）
        private void HandleUsageFromOpenAI(JObject usageData)
        {
            try
            {
                // 确保在UI线程执行
                if (InvokeRequired)
                {
                    Invoke(new Action(() => HandleUsageFromOpenAI(usageData)));
                    return;
                }

                // 记录到 VS 输出，便于观察每次生成结束的 usage 字段
                try
                {
                    int? promptTokens = usageData["prompt_tokens"]?.ToObject<int?>();
                    int? completionTokens = usageData["completion_tokens"]?.ToObject<int?>();
                    int? totalTokens = usageData["total_tokens"]?.ToObject<int?>();
                    int? maxTokens = usageData["max_tokens"]?.ToObject<int?>();
                    int? usedForMax = usageData["used_for_max_tokens"]?.ToObject<int?>();
                    int? remaining = usageData["remaining_tokens"]?.ToObject<int?>();
                    bool estimated = usageData["estimated"]?.ToObject<bool?>() ?? false;
                    bool suppressed = System.Threading.Volatile.Read(ref _suppressUsageToFrontendCount) > 0;

                    Debug.WriteLine($"[usage] prompt={promptTokens?.ToString() ?? "null"}, completion={completionTokens?.ToString() ?? "null"}, total={totalTokens?.ToString() ?? "null"}, max={maxTokens?.ToString() ?? "null"}, used_for_max={usedForMax?.ToString() ?? "null"}, remaining={remaining?.ToString() ?? "null"}, estimated={estimated}, suppressed_to_ui={suppressed}");
                }
                catch { }

                // 保存最近一次 usage
                _lastUsagePayload = usageData;

                // 压缩等内部请求期间不更新前端 token 组件，避免百分比跳变
                if (System.Threading.Volatile.Read(ref _suppressUsageToFrontendCount) > 0)
                {
                    return;
                }

                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    // 计算上下文长度（tokens）
                    int? contextLengthTokens = null;
                    if (_currentModelConfig?.ContextLength.HasValue == true && _currentModelConfig.ContextLength.Value > 0)
                    {
                        contextLengthTokens = _currentModelConfig.ContextLength.Value * 1024;
                    }

                    var messageData = new
                    {
                        type = "usage",
                        prompt_tokens = usageData["prompt_tokens"]?.ToObject<int?>(),
                        completion_tokens = usageData["completion_tokens"]?.ToObject<int?>(),
                        total_tokens = usageData["total_tokens"]?.ToObject<int?>(),
                        max_tokens = usageData["max_tokens"]?.ToObject<int?>(),
                        context_length = contextLengthTokens,
                        used_for_max_tokens = usageData["used_for_max_tokens"]?.ToObject<int?>(),
                        remaining_tokens = usageData["remaining_tokens"]?.ToObject<int?>(),
                        estimated = usageData["estimated"]?.ToObject<bool?>() ?? false,
                        timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
                    };

                    string json = JsonConvert.SerializeObject(messageData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理usage事件时出错: {ex.Message}");
            }
        }

        // 发送操作结果到WebView（UI线程辅助方法）
        private void SendActionResultToWebView(object responseData, bool success)
        {
            try
            {
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    string json = JsonConvert.SerializeObject(responseData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);
                    Debug.WriteLine($"预览操作应用结果已发送到前端: {success}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送操作结果到WebView时出错: {ex.Message}");
            }
        }

        // 处理获取文档列表请求
        private void HandleGetDocuments()
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleGetDocuments()));
                    return;
                }

                // 获取所有文档
                var documents = _documentService.GetAllDocuments();

                // 构建文档列表数据
                var documentList = documents.Select(doc => new
                {
                    id = doc.Id,
                    fileName = doc.FileName,
                    fileType = doc.FileType,
                    fileSize = doc.FileSize,
                    uploadTime = doc.UploadTime.ToString("yyyy-MM-dd HH:mm"),
                    headingCount = _documentService.GetDocumentHeadingCount(doc.Id)
                }).ToList();

                // 发送到前端
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var messageData = new
                    {
                        type = "documentList",
                        documents = documentList
                    };

                    string json = JsonConvert.SerializeObject(messageData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine($"文档列表已发送到前端，共 {documentList.Count} 个文档");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档列表时出错: {ex.Message}");
                SendDocumentErrorToWebView("获取文档列表失败: " + ex.Message);
            }
        }

        // 处理获取文档内容请求
        private void HandleGetDocumentContent(JObject data)
        {
            try
            {
                // 确保在UI线程执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleGetDocumentContent(data)));
                    return;
                }

                int documentId = data["documentId"]?.ToObject<int>() ?? 0;
                int headingId = data["headingId"]?.ToObject<int>() ?? 0;

                if (documentId <= 0)
                {
                    SendDocumentErrorToWebView("无效的文档ID");
                    return;
                }

                if (headingId > 0)
                {
                    // 获取特定标题的内容
                    var headings = _documentService.GetDocumentHeadings(documentId);
                    var targetHeading = headings.FirstOrDefault(h => h.Id == headingId);

                    if (targetHeading != null)
                    {
                        var messageData = new
                        {
                            type = "documentContent",
                            documentId = documentId,
                            headingId = headingId,
                            headingText = targetHeading.HeadingText,
                            content = targetHeading.Content,
                            headingLevel = targetHeading.HeadingLevel
                        };

                        string json = JsonConvert.SerializeObject(messageData);
                        webView21.CoreWebView2.PostWebMessageAsJson(json);

                        Debug.WriteLine($"标题内容已发送到前端: {targetHeading.HeadingText}");
                    }
                    else
                    {
                        SendDocumentErrorToWebView("未找到指定的标题");
                    }
                }
                else
                {
                    // 获取整个文档的标题列表
                    var headings = _documentService.GetDocumentHeadings(documentId);
                    var document = _documentService.GetDocumentById(documentId);

                    if (document != null)
                    {
                        var headingList = headings.Select(h => new
                        {
                            id = h.Id,
                            text = h.HeadingText,
                            level = h.HeadingLevel,
                            parentId = h.ParentHeadingId,
                            orderIndex = h.OrderIndex,
                            contentLength = h.Content?.Length ?? 0
                        }).ToList();

                        var messageData = new
                        {
                            type = "documentHeadingList",
                            documentId = documentId,
                            documentName = document.FileName,
                            headings = headingList
                        };

                        string json = JsonConvert.SerializeObject(messageData);
                        webView21.CoreWebView2.PostWebMessageAsJson(json);

                        Debug.WriteLine($"文档标题列表已发送到前端: {document.FileName}，共 {headingList.Count} 个标题");
                    }
                    else
                    {
                        SendDocumentErrorToWebView("未找到指定的文档");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取文档内容时出错: {ex.Message}");
                SendDocumentErrorToWebView("获取文档内容失败: " + ex.Message);
            }
        }

        // 发送文档错误信息到WebView
        private void SendDocumentErrorToWebView(string errorMessage)
        {
            try
            {
                if (_webViewInitialized && webView21?.CoreWebView2 != null)
                {
                    var errorData = new
                    {
                        type = "documentError",
                        message = errorMessage
                    };

                    string json = JsonConvert.SerializeObject(errorData);
                    webView21.CoreWebView2.PostWebMessageAsJson(json);

                    Debug.WriteLine("文档错误信息已发送到前端");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"发送文档错误信息失败: {ex.Message}");
            }
        }

        // 处理Agent配置更新请求
        private void HandleUpdateAgentConfig(JObject data)
        {
            try
            {
                Debug.WriteLine("开始处理Agent配置更新...");

                var config = data["config"] as JObject;
                if (config == null)
                {
                    Debug.WriteLine("Agent配置数据为空");
                    return;
                }

                // 获取配置值
                bool enableRangeNormalization = config["enableRangeNormalization"]?.ToObject<bool>() ?? true;
                string defaultInsertPosition = config["defaultInsertPosition"]?.ToString() ?? "end";
                bool showToolCalls = config["showToolCalls"]?.ToObject<bool>() ?? true;
                bool showDebugInfo = config["showDebugInfo"]?.ToObject<bool>() ?? false;

                Debug.WriteLine($"Agent配置更新:");
                Debug.WriteLine($"- 启用范围清理: {enableRangeNormalization}");
                Debug.WriteLine($"- 默认插入位置: {defaultInsertPosition}");
                Debug.WriteLine($"- 显示工具调用信息: {showToolCalls}");
                Debug.WriteLine($"- 显示调试详细信息: {showDebugInfo}");

                // 更新WordTools的配置
                WordTools.UpdateAgentConfig(enableRangeNormalization, defaultInsertPosition, showToolCalls, showDebugInfo);

                Debug.WriteLine("Agent配置更新完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理Agent配置更新时出错: {ex.Message}");
            }
        }

        // 继续与AI对话，告知用户操作结果
        private async void ContinueConversationAfterUserAction(string actionType, bool success, string message)
        {
            try
            {
                Debug.WriteLine($"继续与AI对话 - 操作类型: {actionType}, 成功: {success}, 消息: {message}");

                // 在任何对话构建之前读取开关：未开启则直接返回，避免额外token
                bool enablePostFeedbackEarly = false; // 默认不自动发送反馈
                try
                {
                    var cfgJson = await webView21.ExecuteScriptAsync("(function(){try{var s=localStorage.getItem('agentConfig');return s?JSON.parse(s):null;}catch(e){return null}})();");
                    if (!string.IsNullOrEmpty(cfgJson) && cfgJson != "null")
                    {
                        var cfg = JObject.Parse(cfgJson);
                        enablePostFeedbackEarly = cfg["enablePostFeedback"]?.ToObject<bool>() ?? false;
                    }
                }
                catch { enablePostFeedbackEarly = false; }

                if (!enablePostFeedbackEarly)
                {
                    Debug.WriteLine("已关闭自动发送插入结果反馈：不追加AI回复");
                    return;
                }

                // 构建用户反馈消息
                string userFeedback = success
                    ? $"我已接受并成功应用了{GetActionTypeDisplayName(actionType)}操作。{message}"
                    : $"我接受了{GetActionTypeDisplayName(actionType)}操作，但执行失败：{message}";

                // 添加用户反馈到对话历史
                AddUserMessageToHistory(userFeedback);

                // 获取当前使用的模型配置（使用最后一次对话的配置）
                var modelConfig = GetLastUsedModelConfig();
                if (modelConfig == null)
                {
                    Debug.WriteLine("无法获取模型配置，跳过后续对话");
                    return;
                }

                string modelName = ExtractModelNameFromParameters(modelConfig.Parameters);
                string apiUrl = modelConfig.BaseUrl;
                string apiKey = modelConfig.ApiKey;

                // 获取系统提示词（使用agent模式）
                string systemPrompt = GetSystemPromptByMode("chat-agent", null, null);

                // 构建消息数组
                JArray messages = new JArray();
                messages.Add(new JObject
                {
                    ["role"] = "system",
                    ["content"] = systemPrompt
                });

                // 添加对话历史
                foreach (var historyMessage in _conversationHistory)
                {
                    messages.Add(historyMessage);
                }

                // 获取AI参数
                var (temperature, maxTokens, topP) = GetAIParameters(modelConfig, "chat-agent");

                // 构建请求
                JObject requestBody = new JObject
                {
                    ["model"] = modelName,
                    ["messages"] = messages,
                    ["stream"] = true,
                    ["temperature"] = temperature,
                    ["top_p"] = topP
                };

                // max_tokens为可选参数，仅在有值时添加
                if (maxTokens.HasValue)
                {
                    requestBody["max_tokens"] = maxTokens.Value;
                }

                string json = requestBody.ToString();

                // 创建新的CancellationTokenSource
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();

                await webView21.ExecuteScriptAsync("startAppendingToExistingMessage()");

                // 收集生成的内容
                StringBuilder responseContent = new StringBuilder();

                // 调用API生成回复
                await OpenAIUtils.OpenAIApiClientAsync(apiUrl, apiKey, json, _cancellationTokenSource.Token, content =>
                {
                    responseContent.Append(content);

                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(async () => {
                            try
                            {
                                await webView21.ExecuteScriptAsync($"appendToExistingMessage(`{content.Replace("`", "\\`")}`)");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"向前端发送内容时出错: {ex.Message}");
                            }
                        }));
                    }
                    else
                    {
                        webView21.ExecuteScriptAsync($"appendToExistingMessage(`{content.Replace("`", "\\`")}`)").ConfigureAwait(false);
                    }
                }, null); // 不传递对话历史，因为这是简单的反馈回复

                // 将助手回复添加到历史记录
                string finalResponse = responseContent.ToString();
                if (!string.IsNullOrEmpty(finalResponse))
                {
                    AddAssistantMessageToHistory(finalResponse);
                }

                // 通知前端生成完成
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(async () => {
                        try
                        {
                            await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                            Debug.WriteLine("用户操作反馈对话完成");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"通知前端生成完成时出错: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                    Debug.WriteLine("用户操作反馈对话完成");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"继续对话时出错: {ex.Message}");

                if (InvokeRequired)
                {
                    BeginInvoke(new Action(async () => {
                        try
                        {
                            await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                        }
                        catch (Exception innerEx)
                        {
                            Debug.WriteLine($"通知前端完成时出错: {innerEx.Message}");
                        }
                    }));
                }
                else
                {
                    await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                }
            }
        }

        // 获取操作类型的显示名称
        private string GetActionTypeDisplayName(string actionType)
        {
            switch (actionType)
            {
                case "insert_content":
                    return "内容插入";
                case "modify_style":
                    return "样式修改";
                default:
                    return "操作";
            }
        }

        // 获取最后使用的模型配置
        private Model GetLastUsedModelConfig()
        {
            try
            {
                // 优先使用保存的当前模型配置
                if (_currentModelConfig != null)
                {
                    return _currentModelConfig;
                }

                // 如果没有保存的配置，获取第一个可用的对话模型
                var allModels = _modelService.GetAllModels();
                var chatModels = allModels.Where(m => m.modelType == 1).ToList();
                return chatModels.FirstOrDefault();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取模型配置失败: {ex.Message}");
                return null;
            }
        }

        // 处理拒绝预览操作请求
        private async void HandleRejectPreviewedAction(JObject data)
        {
            try
            {
                Debug.WriteLine("开始处理拒绝预览操作...");

                string actionType = data["action_type"]?.ToString() ?? "unknown";
                string previewId = data["preview_id"]?.ToString() ?? "";

                Debug.WriteLine($"拒绝预览操作: {actionType}, 预览ID: {previewId}");

                // 构建用户拒绝反馈消息
                string userFeedback = $"我拒绝了{GetActionTypeDisplayName(actionType)}操作。";

                // 继续与AI对话，告知用户拒绝了操作（遵循是否开启反馈的配置）
                bool enablePostFeedback = false; // 默认不自动发送反馈
                try
                {
                    var flagJson = await webView21.ExecuteScriptAsync("(function(){try{return localStorage.getItem('agentConfig')}catch(e){return null}})();");
                    if (!string.IsNullOrEmpty(flagJson) && flagJson != "null")
                    {
                        var cfg = JObject.Parse(flagJson);
                        enablePostFeedback = cfg["enablePostFeedback"]?.ToObject<bool>() ?? false;
                    }
                }
                catch { enablePostFeedback = false; }

                if (enablePostFeedback)
                {
                    await ContinueConversationAfterUserReject(actionType, userFeedback);
                }
                else
                {
                    Debug.WriteLine("已关闭自动发送插入结果反馈：拒绝操作不追加AI回复");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理拒绝预览操作时出错: {ex.Message}");
            }
        }

        // 继续与AI对话，告知用户拒绝操作
        private async Task ContinueConversationAfterUserReject(string actionType, string userFeedback)
        {
            try
            {
                Debug.WriteLine($"继续与AI对话 - 用户拒绝了操作: {actionType}");

                // 添加用户反馈到对话历史
                AddUserMessageToHistory(userFeedback);

                // 获取当前使用的模型配置
                var modelConfig = GetLastUsedModelConfig();
                if (modelConfig == null)
                {
                    Debug.WriteLine("无法获取模型配置，跳过后续对话");
                    return;
                }

                string modelName = ExtractModelNameFromParameters(modelConfig.Parameters);
                string apiUrl = modelConfig.BaseUrl;
                string apiKey = modelConfig.ApiKey;

                // 获取系统提示词（使用agent模式）
                string systemPrompt = GetSystemPromptByMode("chat-agent", null, null);

                // 构建消息数组
                JArray messages = new JArray();
                messages.Add(new JObject
                {
                    ["role"] = "system",
                    ["content"] = systemPrompt
                });

                // 添加对话历史
                foreach (var historyMessage in _conversationHistory)
                {
                    messages.Add(historyMessage);
                }

                // 获取AI参数
                var (temperature, maxTokens, topP) = GetAIParameters(modelConfig, "chat-agent");

                // 构建请求
                JObject requestBody = new JObject
                {
                    ["model"] = modelName,
                    ["messages"] = messages,
                    ["stream"] = true,
                    ["temperature"] = temperature,
                    ["top_p"] = topP
                };

                // max_tokens为可选参数，仅在有值时添加
                if (maxTokens.HasValue)
                {
                    requestBody["max_tokens"] = maxTokens.Value;
                }

                string json = requestBody.ToString();

                // 创建新的CancellationTokenSource
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();

                // 通知前端开始在现有消息中追加拒绝反馈回复
                await webView21.ExecuteScriptAsync("startAppendingToExistingMessage()");

                // 收集生成的内容
                StringBuilder responseContent = new StringBuilder();

                // 调用API生成回复
                await OpenAIUtils.OpenAIApiClientAsync(apiUrl, apiKey, json, _cancellationTokenSource.Token, content =>
                {
                    responseContent.Append(content);

                    if (InvokeRequired)
                    {
                        BeginInvoke(new Action(async () => {
                            try
                            {
                                await webView21.ExecuteScriptAsync($"appendToExistingMessage(`{content.Replace("`", "\\`")}`)");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"向前端发送内容时出错: {ex.Message}");
                            }
                        }));
                    }
                    else
                    {
                        webView21.ExecuteScriptAsync($"appendToExistingMessage(`{content.Replace("`", "\\`")}`)").ConfigureAwait(false);
                    }
                }, null);

                // 将助手回复添加到历史记录
                string finalResponse = responseContent.ToString();
                if (!string.IsNullOrEmpty(finalResponse))
                {
                    AddAssistantMessageToHistory(finalResponse);
                }

                // 通知前端生成完成
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(async () => {
                        try
                        {
                            await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                            Debug.WriteLine("用户拒绝操作反馈对话完成");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"通知前端生成完成时出错: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                    Debug.WriteLine("用户拒绝操作反馈对话完成");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"继续拒绝对话时出错: {ex.Message}");

                if (InvokeRequired)
                {
                    BeginInvoke(new Action(async () => {
                        try
                        {
                            await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                        }
                        catch (Exception innerEx)
                        {
                            Debug.WriteLine($"通知前端完成时出错: {innerEx.Message}");
                        }
                    }));
                }
                else
                {
                    await webView21.ExecuteScriptAsync("finishAppendingToExistingMessage()");
                }
            }
        }

        // 获取工具描述
        private string GetToolDescription(string toolName)
        {
            switch (toolName)
            {
                case "check_insert_position":
                    return "检查插入位置和获取上下文信息";
                case "get_selected_text":
                    return "获取选中的文本";
                case "formatted_insert_content":
                    return "插入格式化内容（支持在标题下或光标位置插入）";
                case "modify_text_style":
                    return "修改文字颜色、大小、间距等样式";
                case "get_document_statistics":
                    return "获取文档统计信息";
                case "get_document_images":
                    return "获取文档中的所有图片信息";
                case "get_document_formulas":
                    return "获取数学公式的位置和数量统计";
                case "get_document_tables":
                    return "获取文档中的表格信息";
                case "get_document_headings":
                    return "获取文档的标题列表（整体结构）";
                case "get_heading_content":
                    return "高效获取指定标题下的所有内容";
                default:
                    return "";
            }
        }
    }
}
