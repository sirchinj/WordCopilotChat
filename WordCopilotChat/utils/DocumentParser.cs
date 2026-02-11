using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using WordCopilotChat.models;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDocument = Microsoft.Office.Interop.Word.Document;

namespace WordCopilotChat.utils
{
    /// <summary>
    /// 文档解析工具类
    /// </summary>
    public static class DocumentParser
    {
        /// <summary>
        /// 解析Word文档（.docx, .doc）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="quickMode">快速模式，只提取标题不提取内容</param>
        /// <param name="progressCallback">进度回调函数</param>
        /// <returns>解析结果</returns>
        public static DocumentParseResult ParseWordDocument(string filePath, bool quickMode = false, Action<string> progressCallback = null)
        {
            var result = new DocumentParseResult();
            WordApp wordApp = null;
            WordDocument doc = null;

            try
            {
                Debug.WriteLine($"开始解析Word文档: {filePath}");
                progressCallback?.Invoke($"开始解析Word文档: {Path.GetFileName(filePath)}");

                // 创建Word应用程序
                Debug.WriteLine("正在创建Word应用程序实例...");
                progressCallback?.Invoke("正在创建Word应用程序实例...");
                wordApp = new WordApp();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone; // 禁用警告对话框

                Debug.WriteLine("正在打开文档...");
                progressCallback?.Invoke("正在打开文档...");
                doc = wordApp.Documents.Open(filePath, ReadOnly: true); // 以只读方式打开

                Debug.WriteLine($"文档已打开，共有 {doc.Paragraphs.Count} 个段落");
                progressCallback?.Invoke($"文档已打开，共有 {doc.Paragraphs.Count} 个段落");

                var headings = new List<DocumentHeading>();
                var headingStack = new Stack<DocumentHeading>(); // 用于跟踪父级标题
                int orderIndex = 0;
                int processedParagraphs = 0;

                Debug.WriteLine("开始扫描段落寻找标题...");
                progressCallback?.Invoke("开始扫描段落寻找标题...");
                foreach (Paragraph para in doc.Paragraphs)
                {
                    processedParagraphs++;

                    // 每处理100个段落输出一次进度
                    if (processedParagraphs % 100 == 0)
                    {
                        var progressMsg = $"已处理 {processedParagraphs}/{doc.Paragraphs.Count} 个段落，找到 {headings.Count} 个标题";
                        Debug.WriteLine(progressMsg);
                        progressCallback?.Invoke(progressMsg);
                    }

                    try
                    {
                        var style = para.get_Style();
                        string styleName = style.NameLocal;

                        // 检查是否为标题样式
                        if (IsHeadingStyle(styleName, out int level))
                        {
                            var headingText = para.Range.Text.Trim().Replace("\r", "");
                            if (!string.IsNullOrWhiteSpace(headingText))
                            {
                                Debug.WriteLine($"找到标题 (H{level}): {headingText}");
                                progressCallback?.Invoke($"找到标题 (H{level}): {headingText}");

                                // 找到父级标题
                                DocumentHeading parentHeading = null;
                                while (headingStack.Count > 0 && headingStack.Peek().HeadingLevel >= level)
                                {
                                    headingStack.Pop();
                                }
                                if (headingStack.Count > 0)
                                {
                                    parentHeading = headingStack.Peek();
                                }

                                var heading = new DocumentHeading
                                {
                                    HeadingText = headingText,
                                    HeadingLevel = level,
                                    ParentHeadingId = parentHeading?.Id,
                                    Content = "", // 内容稍后收集
                                    OrderIndex = orderIndex++
                                };

                                headings.Add(heading);
                                headingStack.Push(heading);
                            }
                        }
                    }
                    catch (Exception paraEx)
                    {
                        Debug.WriteLine($"处理段落 {processedParagraphs} 时出错: {paraEx.Message}");
                        // 继续处理下一个段落
                    }
                }

                Debug.WriteLine($"段落扫描完成，共找到 {headings.Count} 个标题");
                progressCallback?.Invoke($"段落扫描完成，共找到 {headings.Count} 个标题");

                // 标记是否使用了字体大小识别（字体大小识别方法内部已收集内容）
                bool usedFontSizeRecognition = false;

                // 如果没有找到标题样式，尝试通过字体大小识别标题
                if (headings.Count == 0)
                {
                    Debug.WriteLine("未找到标题样式，尝试通过字体大小识别标题...");
                    progressCallback?.Invoke("未找到标题样式，尝试通过字体大小识别标题...");

                    headings = TryParseLargerFontAsHeading(doc, progressCallback);
                    usedFontSizeRecognition = true; // 字体大小识别方法内部已收集内容

                    if (headings.Count > 0)
                    {
                        Debug.WriteLine($"通过字体大小识别到 {headings.Count} 个标题");
                        progressCallback?.Invoke($"通过字体大小识别到 {headings.Count} 个标题");
                    }
                    else
                    {
                        Debug.WriteLine("未能通过字体大小识别到标题");
                        progressCallback?.Invoke("未能通过字体大小识别到标题");
                    }
                }

                // 收集每个标题下的内容（除非是快速模式或已使用字体大小识别）
                if (headings.Count > 0 && !quickMode && !usedFontSizeRecognition)
                {
                    Debug.WriteLine("开始收集标题内容...");
                    progressCallback?.Invoke("开始收集标题内容...");
                    CollectWordHeadingContent(doc, headings, progressCallback);
                    Debug.WriteLine("标题内容收集完成");
                    progressCallback?.Invoke("标题内容收集完成");
                }
                else if (quickMode)
                {
                    Debug.WriteLine("快速模式：跳过内容收集");
                    progressCallback?.Invoke("快速模式：跳过内容收集");
                    // 在快速模式下，为所有标题设置空内容
                    foreach (var heading in headings)
                    {
                        heading.Content = "";
                    }
                }

                result.Headings = headings;
                result.Success = true;
                result.Message = $"Word文档解析成功，找到 {headings.Count} 个标题";

                Debug.WriteLine($"Word文档解析完成，共找到 {headings.Count} 个标题");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Word文档解析失败: {ex.Message}";
                Debug.WriteLine($"Word文档解析错误: {ex}");
                Debug.WriteLine($"错误详情: {ex.StackTrace}");
            }
            finally
            {
                // 清理资源
                try
                {
                    Debug.WriteLine("正在清理Word资源...");
                    if (doc != null)
                    {
                        doc.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    }
                    if (wordApp != null)
                    {
                        wordApp.Quit(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    }
                    Debug.WriteLine("Word资源清理完成");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"清理Word资源时出错: {ex.Message}");
                }
                finally
                {
                    // 强制垃圾回收
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                }
            }

            return result;
        }

        /// <summary>
        /// 解析Markdown文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>解析结果</returns>
        public static DocumentParseResult ParseMarkdownDocument(string filePath)
        {
            var result = new DocumentParseResult();

            try
            {
                string content = File.ReadAllText(filePath, Encoding.UTF8);
                var headings = new List<DocumentHeading>();
                var lines = content.Split('\n');

                var headingStack = new Stack<DocumentHeading>();
                int orderIndex = 0;
                var currentContent = new StringBuilder();
                DocumentHeading currentHeading = null;

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();

                    // 检查是否为标题行
                    var headingMatch = Regex.Match(trimmedLine, @"^(#{1,6})\s+(.+)$");
                    if (headingMatch.Success)
                    {
                        // 保存之前标题的内容
                        if (currentHeading != null)
                        {
                            currentHeading.Content = currentContent.ToString().Trim();
                        }

                        int level = headingMatch.Groups[1].Value.Length;
                        string headingText = headingMatch.Groups[2].Value.Trim();

                        // 找到父级标题
                        DocumentHeading parentHeading = null;
                        while (headingStack.Count > 0 && headingStack.Peek().HeadingLevel >= level)
                        {
                            headingStack.Pop();
                        }
                        if (headingStack.Count > 0)
                        {
                            parentHeading = headingStack.Peek();
                        }

                        currentHeading = new DocumentHeading
                        {
                            HeadingText = headingText,
                            HeadingLevel = level,
                            ParentHeadingId = parentHeading?.Id,
                            Content = "",
                            OrderIndex = orderIndex++
                        };

                        headings.Add(currentHeading);
                        headingStack.Push(currentHeading);
                        currentContent.Clear();
                    }
                    else
                    {
                        // 收集内容
                        if (currentHeading != null && !string.IsNullOrWhiteSpace(trimmedLine))
                        {
                            currentContent.AppendLine(line);
                        }
                    }
                }

                // 保存最后一个标题的内容
                if (currentHeading != null)
                {
                    currentHeading.Content = currentContent.ToString().Trim();
                }

                result.Headings = headings;
                result.Success = true;
                result.Message = "Markdown文档解析成功";

                Debug.WriteLine($"Markdown文档解析完成，共找到 {headings.Count} 个标题");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Markdown文档解析失败: {ex.Message}";
                Debug.WriteLine($"Markdown文档解析错误: {ex}");
            }

            return result;
        }

        /// <summary>
        /// 尝试通过字体大小识别标题（当文档没有使用标题样式时的备选方案）
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="progressCallback">进度回调</param>
        /// <returns>识别到的标题列表</returns>
        private static List<DocumentHeading> TryParseLargerFontAsHeading(WordDocument doc, Action<string> progressCallback = null)
        {
            var headings = new List<DocumentHeading>();

            try
            {
                // 第一遍扫描：收集所有段落的字体大小信息
                var fontSizeStats = new Dictionary<float, int>(); // 字体大小 -> 出现次数
                var paragraphInfos = new List<ParagraphFontInfo>();

                Debug.WriteLine("第一遍扫描：收集字体大小信息...");
                progressCallback?.Invoke("正在分析文档字体大小分布...");

                int paraIndex = 0;
                foreach (Paragraph para in doc.Paragraphs)
                {
                    paraIndex++;
                    try
                    {
                        string text = para.Range.Text.Trim().Replace("\r", "").Replace("\a", "");
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        // 获取段落的字体大小
                        float fontSize = 0;
                        try
                        {
                            // 尝试获取段落的字体大小
                            var range = para.Range;
                            var font = range.Font;

                            // Font.Size 可能返回 9999999 表示混合大小，取第一个字符的大小
                            if (font.Size > 0 && font.Size < 1000)
                            {
                                fontSize = font.Size;
                            }
                            else
                            {
                                // 混合字体大小，取第一个字符
                                if (range.Characters.Count > 0)
                                {
                                    var firstCharFont = range.Characters[1].Font;
                                    if (firstCharFont.Size > 0 && firstCharFont.Size < 1000)
                                    {
                                        fontSize = firstCharFont.Size;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            fontSize = 0;
                        }

                        if (fontSize > 0)
                        {
                            // 四舍五入到0.5磅
                            fontSize = (float)Math.Round(fontSize * 2) / 2;

                            // 统计字体大小出现次数（按文本长度加权）
                            int weight = Math.Min(text.Length, 100); // 限制权重，避免长段落过度影响
                            if (fontSizeStats.ContainsKey(fontSize))
                            {
                                fontSizeStats[fontSize] += weight;
                            }
                            else
                            {
                                fontSizeStats[fontSize] = weight;
                            }

                            paragraphInfos.Add(new ParagraphFontInfo
                            {
                                Index = paraIndex,
                                Text = text,
                                FontSize = fontSize,
                                IsBold = false // 稍后检查
                            });

                            // 检查是否加粗
                            try
                            {
                                var lastInfo = paragraphInfos[paragraphInfos.Count - 1];
                                lastInfo.IsBold = para.Range.Font.Bold != 0;
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"分析段落 {paraIndex} 字体时出错: {ex.Message}");
                    }
                }

                if (paragraphInfos.Count == 0 || fontSizeStats.Count == 0)
                {
                    Debug.WriteLine("未能获取有效的字体信息");
                    progressCallback?.Invoke("未能获取有效的字体信息");
                    return headings;
                }

                // 确定正文字体大小（出现频率最高的）
                float bodyFontSize = fontSizeStats.OrderByDescending(x => x.Value).First().Key;
                Debug.WriteLine($"检测到正文字体大小: {bodyFontSize}磅");
                progressCallback?.Invoke($"检测到正文字体大小: {bodyFontSize}磅");

                // 找出所有比正文大的字体大小，按大小降序排列
                var largerFontSizes = fontSizeStats.Keys
                    .Where(s => s > bodyFontSize + 1.5f) // 至少比正文大1.5磅才算标题
                    .OrderByDescending(s => s)
                    .Take(6) // 最多支持6级标题
                    .ToList();

                if (largerFontSizes.Count == 0)
                {
                    Debug.WriteLine("未找到比正文更大的字体");
                    progressCallback?.Invoke("未找到比正文更大的字体，无法识别标题");
                    return headings;
                }

                Debug.WriteLine($"找到 {largerFontSizes.Count} 种较大字体: {string.Join(", ", largerFontSizes.Select(s => s + "磅"))}");
                progressCallback?.Invoke($"找到 {largerFontSizes.Count} 种较大字体可能为标题");

                // 建立字体大小到标题级别的映射
                var fontSizeToLevel = new Dictionary<float, int>();
                for (int i = 0; i < largerFontSizes.Count; i++)
                {
                    fontSizeToLevel[largerFontSizes[i]] = i + 1; // H1, H2, H3...
                }

                // 第二遍扫描：识别标题
                Debug.WriteLine("第二遍扫描：识别标题...");
                progressCallback?.Invoke("正在识别标题...");

                var headingStack = new Stack<DocumentHeading>();
                int orderIndex = 0;

                foreach (var paraInfo in paragraphInfos)
                {
                    // 检查是否为标题候选
                    if (fontSizeToLevel.ContainsKey(paraInfo.FontSize))
                    {
                        int level = fontSizeToLevel[paraInfo.FontSize];
                        string headingText = paraInfo.Text;

                        // 标题文本不应太长（超过100字符的可能不是标题）
                        if (headingText.Length > 100) continue;

                        // 标题文本不应为纯数字或特殊符号
                        if (string.IsNullOrWhiteSpace(headingText.Trim('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', '、', ' ')))
                            continue;

                        Debug.WriteLine($"识别到标题 (H{level}, {paraInfo.FontSize}磅): {headingText}");

                        // 找到父级标题
                        DocumentHeading parentHeading = null;
                        while (headingStack.Count > 0 && headingStack.Peek().HeadingLevel >= level)
                        {
                            headingStack.Pop();
                        }
                        if (headingStack.Count > 0)
                        {
                            parentHeading = headingStack.Peek();
                        }

                        var heading = new DocumentHeading
                        {
                            HeadingText = headingText,
                            HeadingLevel = level,
                            ParentHeadingId = parentHeading?.Id,
                            Content = "",
                            OrderIndex = orderIndex++
                        };

                        headings.Add(heading);
                        headingStack.Push(heading);
                    }
                }

                // 收集标题内容
                if (headings.Count > 0)
                {
                    Debug.WriteLine("开始收集标题内容...");
                    progressCallback?.Invoke("开始收集标题内容...");
                    CollectHeadingContentByFontSize(paragraphInfos, headings, fontSizeToLevel, progressCallback);

                    // 提取表格（字体识别模式下，需要根据段落索引估算标题范围来关联表格）
                    if (doc.Tables.Count > 0)
                    {
                        Debug.WriteLine($"文档包含 {doc.Tables.Count} 个表格，开始提取...");
                        progressCallback?.Invoke($"检测到 {doc.Tables.Count} 个表格，正在提取...");

                        ExtractTablesForFontSizeMode(doc, headings, paragraphInfos, fontSizeToLevel, progressCallback);
                    }
                }

                Debug.WriteLine($"字体大小识别完成，共找到 {headings.Count} 个标题");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"通过字体大小识别标题时出错: {ex.Message}");
                progressCallback?.Invoke($"字体识别出错: {ex.Message}");
            }

            return headings;
        }

        /// <summary>
        /// 段落字体信息
        /// </summary>
        private class ParagraphFontInfo
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public float FontSize { get; set; }
            public bool IsBold { get; set; }
        }

        /// <summary>
        /// 根据字体大小收集标题内容（包括表格，按文档顺序）
        /// </summary>
        private static void CollectHeadingContentByFontSize(
            List<ParagraphFontInfo> paragraphInfos,
            List<DocumentHeading> headings,
            Dictionary<float, int> fontSizeToLevel,
            Action<string> progressCallback = null)
        {
            try
            {
                // 为每个标题找到对应的段落索引范围
                var headingIndices = new List<int>();
                int searchFrom = 0;

                foreach (var heading in headings)
                {
                    for (int i = searchFrom; i < paragraphInfos.Count; i++)
                    {
                        if (paragraphInfos[i].Text == heading.HeadingText &&
                            fontSizeToLevel.ContainsKey(paragraphInfos[i].FontSize))
                        {
                            headingIndices.Add(i);
                            searchFrom = i + 1;
                            break;
                        }
                    }
                }

                // 收集每个标题下的内容（仅段落文本，暂不含表格）
                for (int h = 0; h < headings.Count; h++)
                {
                    var currentHeading = headings[h];
                    int startIdx = headingIndices[h] + 1;
                    int endIdx = h + 1 < headingIndices.Count ? headingIndices[h + 1] : paragraphInfos.Count;

                    var content = new StringBuilder();
                    for (int i = startIdx; i < endIdx; i++)
                    {
                        var paraInfo = paragraphInfos[i];
                        // 只收集正文内容（不是标题的段落）
                        if (!fontSizeToLevel.ContainsKey(paraInfo.FontSize))
                        {
                            content.AppendLine(paraInfo.Text);
                        }
                        else
                        {
                            // 遇到同级或更高级标题时停止
                            int paraLevel = fontSizeToLevel[paraInfo.FontSize];
                            if (paraLevel <= currentHeading.HeadingLevel)
                            {
                                break;
                            }
                        }
                    }

                    currentHeading.Content = content.ToString().Trim();
                    Debug.WriteLine($"标题 '{currentHeading.HeadingText}' 收集到 {currentHeading.Content.Length} 字符内容（不含表格）");
                }

                progressCallback?.Invoke("段落内容收集完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"收集标题内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在字体识别模式下提取表格并关联到标题（按正确位置插入）
        /// </summary>
        private static void ExtractTablesForFontSizeMode(
            WordDocument doc,
            List<DocumentHeading> headings,
            List<ParagraphFontInfo> paragraphInfos,
            Dictionary<float, int> fontSizeToLevel,
            Action<string> progressCallback = null)
        {
            try
            {
                if (headings.Count == 0 || doc.Tables.Count == 0)
                    return;

                // 收集所有内容元素（段落和表格）及其位置
                var contentElements = new List<(int rangeStart, string type, string text, DocumentHeading owner)>();

                // 1. 收集段落信息（带Range位置）
                int paraIndex = 0;
                foreach (Paragraph para in doc.Paragraphs)
                {
                    paraIndex++;
                    try
                    {
                        string paraText = para.Range.Text.Trim().Replace("\r", "").Replace("\a", "");
                        if (string.IsNullOrWhiteSpace(paraText)) continue;

                        int rangeStart = para.Range.Start;

                        // 查找对应的 ParagraphFontInfo
                        var paraInfo = paragraphInfos.FirstOrDefault(p => p.Index == paraIndex);
                        if (paraInfo == null) continue;

                        // 跳过标题段落
                        if (fontSizeToLevel.ContainsKey(paraInfo.FontSize)) continue;

                        contentElements.Add((rangeStart, "paragraph", paraText, null));
                    }
                    catch { }
                }

                // 2. 收集表格信息（带Range位置）
                foreach (Table table in doc.Tables)
                {
                    try
                    {
                        int rangeStart = table.Range.Start;
                        string tableMarkdown = ExtractTableAsMarkdown(table);
                        if (!string.IsNullOrWhiteSpace(tableMarkdown))
                        {
                            contentElements.Add((rangeStart, "table", tableMarkdown, null));
                        }
                    }
                    catch { }
                }

                // 3. 按 Range 位置排序
                contentElements = contentElements.OrderBy(e => e.rangeStart).ToList();

                // 4. 获取每个标题的 Range 范围
                var headingRanges = new List<(DocumentHeading heading, int start, int end)>();
                paraIndex = 0;
                foreach (Paragraph para in doc.Paragraphs)
                {
                    paraIndex++;
                    try
                    {
                        string paraText = para.Range.Text.Trim().Replace("\r", "");
                        var paraInfo = paragraphInfos.FirstOrDefault(p => p.Index == paraIndex && p.Text == paraText);

                        if (paraInfo != null && fontSizeToLevel.ContainsKey(paraInfo.FontSize))
                        {
                            // 这是一个标题
                            var heading = headings.FirstOrDefault(h => h.HeadingText == paraText);
                            if (heading != null)
                            {
                                int headingStart = para.Range.End;
                                headingRanges.Add((heading, headingStart, int.MaxValue));
                            }
                        }
                    }
                    catch { }
                }

                // 设置每个标题的结束位置（下一个同级或更高级标题的开始位置）
                for (int i = 0; i < headingRanges.Count; i++)
                {
                    var (heading, start, _) = headingRanges[i];
                    int end = doc.Content.End;

                    // 查找下一个同级或更高级标题
                    for (int j = i + 1; j < headingRanges.Count; j++)
                    {
                        if (headingRanges[j].heading.HeadingLevel <= heading.HeadingLevel)
                        {
                            end = headingRanges[j].start;
                            break;
                        }
                    }

                    headingRanges[i] = (heading, start, end);
                }

                // 5. 为每个标题重新构建内容（按位置顺序包含段落和表格）
                int extractedTables = 0;
                foreach (var (heading, rangeStart, rangeEnd) in headingRanges)
                {
                    var contentBuilder = new StringBuilder();

                    foreach (var (elemStart, elemType, elemText, _) in contentElements)
                    {
                        // 只包含在标题范围内的元素
                        if (elemStart >= rangeStart && elemStart < rangeEnd)
                        {
                            if (contentBuilder.Length > 0)
                            {
                                contentBuilder.AppendLine();
                            }

                            if (elemType == "table")
                            {
                                contentBuilder.AppendLine(); // 表格前空一行
                                extractedTables++;
                            }

                            contentBuilder.AppendLine(elemText);
                        }
                    }

                    heading.Content = contentBuilder.ToString().Trim();
                    Debug.WriteLine($"标题 '{heading.HeadingText}' 最终内容长度: {heading.Content.Length} 字符");
                }

                Debug.WriteLine($"成功提取 {extractedTables} 个表格（按位置插入）");
                progressCallback?.Invoke($"成功提取 {extractedTables} 个表格");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"字体识别模式下提取表格失败: {ex.Message}");
                progressCallback?.Invoke($"提取表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查样式是否为标题样式
        /// </summary>
        private static bool IsHeadingStyle(string styleName, out int level)
        {
            level = 0;
            if (string.IsNullOrEmpty(styleName)) return false;

            // 中文标题样式
            if (styleName.Contains("标题"))
            {
                var match = Regex.Match(styleName, @"标题\s*(\d+)");
                if (match.Success && int.TryParse(match.Groups[1].Value, out level))
                {
                    return level >= 1 && level <= 6;
                }
            }

            // 英文标题样式
            if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(styleName, @"Heading\s*(\d+)", RegexOptions.IgnoreCase);
                if (match.Success && int.TryParse(match.Groups[1].Value, out level))
                {
                    return level >= 1 && level <= 6;
                }
            }

            return false;
        }

        /// <summary>
        /// 收集Word文档中每个标题下的内容（包括段落和表格，按文档顺序）
        /// </summary>
        private static void CollectWordHeadingContent(WordDocument doc, List<DocumentHeading> headings, Action<string> progressCallback = null)
        {
            try
            {
                Debug.WriteLine($"开始为 {headings.Count} 个标题收集内容...");

                // 1. 收集所有段落信息（带Range位置）
                var paragraphElements = new List<(int rangeStart, string text)>();
                var headingRanges = new List<(DocumentHeading heading, int start, int end)>();

                progressCallback?.Invoke("正在扫描文档结构...");

                foreach (Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        string paraText = para.Range.Text.Trim().Replace("\r", "").Replace("\a", "");
                        int rangeStart = para.Range.Start;
                        int rangeEnd = para.Range.End;

                        // 检查是否是标题
                        var style = para.get_Style();
                        if (IsHeadingStyle(style.NameLocal, out int level))
                        {
                            // 找到对应的标题对象
                            var heading = headings.FirstOrDefault(h => h.HeadingText == paraText && h.HeadingLevel == level);
                            if (heading != null)
                            {
                                headingRanges.Add((heading, rangeEnd, doc.Content.End));
                            }
                        }
                        else if (!string.IsNullOrWhiteSpace(paraText))
                        {
                            // 普通段落
                            paragraphElements.Add((rangeStart, paraText));
                        }
                    }
                    catch (Exception paraEx)
                    {
                        Debug.WriteLine($"处理段落时出错: {paraEx.Message}");
                    }
                }

                // 2. 设置每个标题的结束位置
                for (int i = 0; i < headingRanges.Count; i++)
                {
                    var (heading, start, _) = headingRanges[i];
                    int end = doc.Content.End;

                    // 查找下一个同级或更高级标题
                    for (int j = i + 1; j < headingRanges.Count; j++)
                    {
                        if (headingRanges[j].heading.HeadingLevel <= heading.HeadingLevel)
                        {
                            end = headingRanges[j].start;
                            break;
                        }
                    }

                    headingRanges[i] = (heading, start, end);
                    Debug.WriteLine($"标题范围: '{heading.HeadingText}' -> {start} - {end}");
                }

                // 3. 收集表格信息（带Range位置）
                var tableElements = new List<(int rangeStart, string markdown)>();
                if (doc.Tables.Count > 0)
                {
                    Debug.WriteLine($"文档包含 {doc.Tables.Count} 个表格，开始提取...");
                    progressCallback?.Invoke($"检测到 {doc.Tables.Count} 个表格，正在提取...");

                    int tableIndex = 0;
                    foreach (Table table in doc.Tables)
                    {
                        tableIndex++;
                        try
                        {
                            int tableStart = table.Range.Start;
                            string tableMarkdown = ExtractTableAsMarkdown(table);
                            Debug.WriteLine($"表格 {tableIndex}: Range {tableStart}, 行数: {table.Rows.Count}, 列数: {table.Columns.Count}");

                            if (!string.IsNullOrWhiteSpace(tableMarkdown))
                            {
                                tableElements.Add((tableStart, tableMarkdown));
                            }
                        }
                        catch (Exception tableEx)
                        {
                            Debug.WriteLine($"提取表格 {tableIndex} 时出错: {tableEx.Message}");
                        }
                    }

                    Debug.WriteLine($"表格提取完成，共 {tableElements.Count} 个");
                }

                // 4. 合并所有内容元素并按位置排序
                var allElements = new List<(int rangeStart, string type, string content)>();

                foreach (var (rangeStart, text) in paragraphElements)
                {
                    allElements.Add((rangeStart, "paragraph", text));
                }

                foreach (var (rangeStart, markdown) in tableElements)
                {
                    allElements.Add((rangeStart, "table", markdown));
                }

                // 按位置排序
                allElements = allElements.OrderBy(e => e.rangeStart).ToList();
                Debug.WriteLine($"共有 {allElements.Count} 个内容元素（段落+表格）");

                // 5. 为每个标题构建内容（按位置顺序）
                progressCallback?.Invoke("正在组合标题内容...");
                int extractedTableCount = 0;

                foreach (var (heading, rangeStart, rangeEnd) in headingRanges)
                {
                    var contentBuilder = new StringBuilder();

                    foreach (var (elemStart, elemType, elemContent) in allElements)
                    {
                        // 只包含在标题范围内的元素
                        if (elemStart >= rangeStart && elemStart < rangeEnd)
                        {
                            if (contentBuilder.Length > 0)
                            {
                                contentBuilder.AppendLine();
                            }

                            if (elemType == "table")
                            {
                                extractedTableCount++;
                                progressCallback?.Invoke($"表格 {extractedTableCount} 已关联到标题: {heading.HeadingText}");
                            }

                            contentBuilder.AppendLine(elemContent);
                        }
                    }

                    heading.Content = contentBuilder.ToString().Trim();
                }

                // 输出最终统计
                progressCallback?.Invoke($"表格提取完成：成功提取 {extractedTableCount} 个");
                for (int i = 0; i < headings.Count; i++)
                {
                    var heading = headings[i];
                    Debug.WriteLine($"标题 '{heading.HeadingText}' 最终内容长度: {heading.Content?.Length ?? 0} 字符");
                }

                Debug.WriteLine("所有标题内容收集完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"收集Word标题内容时出错: {ex.Message}");
                Debug.WriteLine($"错误详情: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 将Word表格提取为Markdown格式文本
        /// </summary>
        private static string ExtractTableAsMarkdown(Table table)
        {
            try
            {
                if (table.Rows.Count == 0 || table.Columns.Count == 0)
                {
                    return string.Empty;
                }

                var sb = new StringBuilder();
                sb.AppendLine("\n[表格]");

                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;

                // 提取表格内容
                for (int r = 1; r <= rowCount; r++)
                {
                    var rowData = new List<string>();

                    try
                    {
                        var row = table.Rows[r];

                        // 提取每个单元格
                        for (int c = 1; c <= colCount; c++)
                        {
                            try
                            {
                                var cell = row.Cells[c];
                                string cellText = cell.Range.Text;

                                // 清理单元格文本（移除表格结束符等特殊字符）
                                cellText = cellText.Replace("\r", " ").Replace("\a", "").Replace("\x07", "").Trim();

                                rowData.Add(cellText);
                            }
                            catch
                            {
                                rowData.Add("");
                            }
                        }

                        // 输出行数据（使用 | 分隔）
                        sb.AppendLine("| " + string.Join(" | ", rowData) + " |");

                        // 第一行后添加分隔线（Markdown表格格式）
                        if (r == 1)
                        {
                            var separator = new string[colCount];
                            for (int i = 0; i < colCount; i++)
                            {
                                separator[i] = "---";
                            }
                            sb.AppendLine("| " + string.Join(" | ", separator) + " |");
                        }
                    }
                    catch (Exception rowEx)
                    {
                        Debug.WriteLine($"提取表格第 {r} 行时出错: {rowEx.Message}");
                        continue;
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"表格转Markdown失败: {ex.Message}");
                return "[表格提取失败]";
            }
        }

        /// <summary>
        /// 验证文件格式
        /// </summary>
        public static bool IsSupportedFileType(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return false;

            string extension = Path.GetExtension(filePath).ToLower();
            return extension == ".docx" || extension == ".doc" || extension == ".md";
        }

        /// <summary>
        /// 获取文件类型
        /// </summary>
        public static string GetFileType(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return "";

            return Path.GetExtension(filePath).ToLower().TrimStart('.');
        }
    }

    /// <summary>
    /// 文档解析结果
    /// </summary>
    public class DocumentParseResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public List<DocumentHeading> Headings { get; set; } = new List<DocumentHeading>();
    }
} 