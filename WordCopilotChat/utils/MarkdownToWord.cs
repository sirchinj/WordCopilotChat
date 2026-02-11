using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace WordCopilotChat.utils
{
    /// <summary>
    /// 提供Markdown内容转换到Word文档的功能
    /// </summary>
    public class MarkdownToWord
    {
        /// <summary>
        /// 简单的HTML编码方法
        /// </summary>
        /// <param name="text">要编码的文本</param>
        /// <returns>编码后的HTML文本</returns>
        private string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
                
            return text.Replace("&", "&amp;")
                      .Replace("<", "&lt;")
                      .Replace(">", "&gt;")
                      .Replace("\"", "&quot;")
                      .Replace("'", "&#39;");
        }
        /// <summary>
        /// 获取或创建Word应用程序实例
        /// </summary>
        /// <returns>Word应用程序对象</returns>
        public dynamic GetWordApplication()
        {
            try
            {
                // 尝试获取已打开的Word实例
                dynamic wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                Debug.WriteLine("成功连接到已打开的Word实例");
                return wordApp;
            }
            catch
            {
                try
                {
                    // 如果没有打开的实例，创建一个新的Word实例
                    Type wordType = Type.GetTypeFromProgID("Word.Application");
                    dynamic wordApp = Activator.CreateInstance(wordType);
                    wordApp.Visible = true;
                    Debug.WriteLine("已创建新的Word实例");
                    return wordApp;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"无法启动或连接到Word: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// 获取活动文档，如果没有则创建一个新文档
        /// </summary>
        /// <param name="wordApp">Word应用程序对象</param>
        /// <returns>活动文档对象</returns>
        public dynamic GetActiveDocument(dynamic wordApp)
        {
            dynamic doc = wordApp.ActiveDocument;
            if (doc == null)
            {
                // 如果没有活动文档，则创建一个新文档
                doc = wordApp.Documents.Add();
                Debug.WriteLine("已创建新的Word文档");
            }
            return doc;
        }

        /// <summary>
        /// 将HTML内容插入到Word文档
        /// </summary>
        /// <param name="htmlContent">HTML内容</param>
        /// <param name="formulas">公式列表</param>
        public void InsertHtmlContent(string htmlContent, List<JObject> formulas = null)
        {
            try
            {
                Debug.WriteLine("开始复制到Word过程...");
                
                // 获取Word应用程序实例
                dynamic wordApp = GetWordApplication();
                dynamic doc = GetActiveDocument(wordApp);
                dynamic selection = wordApp.Selection;
                
                Debug.WriteLine($"初始光标位置: {selection.Start}");
                
                // 创建临时HTML文件
                string tempFile = Path.Combine(Path.GetTempPath(), $"WordCopilot_{Guid.NewGuid()}.html");
                try
                {
                    // 添加HTML基本标签
                    string completeHtml = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{ font-family: Calibri, Arial, sans-serif; }}
        pre {{ background-color: #f0f0f0; padding: 8px; border-radius: 4px; }}
        code {{ font-family: Consolas, 'Courier New', monospace; }}
        .formula-placeholder {{ background-color: #f8f9fa; padding: 5px; text-align: center; border: 1px solid #dee2e6; margin: 10px 0; }}
        table {{ border-collapse: collapse; width: 100%; border: 1px solid #000; }}
        th, td {{ border: 1px solid #000; padding: 6px; }}
        th {{ background-color: #f2f2f2; }}
    </style>
</head>
<body>
    {htmlContent}
</body>
</html>";

                    // 写入临时文件
                    File.WriteAllText(tempFile, completeHtml, Encoding.UTF8);

                    // 记录光标位置
                    int cursorStart = selection.Start;
                    Debug.WriteLine($"插入内容前的光标位置: {cursorStart}");
                    
                    // 插入HTML文件
                    selection.InsertFile(tempFile);
                    Debug.WriteLine($"HTML内容插入后的光标位置: {selection.Start}");
                    
                    // 处理插入后的表格样式
                    try
                    {
                        // 从当前插入点向上查找所有表格并应用样式
                        dynamic range = doc.Range(cursorStart, selection.Start);
                        dynamic tables = range.Tables;
                        
                        if (tables != null && tables.Count > 0)
                        {
                            Debug.WriteLine($"找到{tables.Count}个表格，应用边框样式");
                            
                            for (int i = 1; i <= tables.Count; i++)
                            {
                                dynamic table = tables[i];
                                // 应用网格线样式 (wdLineStyleSingle = 1)
                                table.Borders.InsideLineStyle = 1;
                                table.Borders.OutsideLineStyle = 1;
                                
                                // 设置边框宽度 (wdLineWidth025pt = 2，1/4磅)
                                table.Borders.InsideLineWidth = 2;
                                table.Borders.OutsideLineWidth = 2;
                                
                                // 自动调整表格
                                try { table.AutoFitBehavior(1); } catch { } // wdAutoFitContent = 1
                            }
                        }
                    }
                    catch (Exception exTable)
                    {
                        Debug.WriteLine($"应用表格样式时出错: {exTable.Message}");
                        // 继续处理，不中断流程
                    }
                    
                    // 处理公式
                    if (formulas != null && formulas.Count > 0)
                    {
                        Debug.WriteLine($"开始处理{formulas.Count}个公式");
                        
                        // 从头遍历文档内容
                        try 
                        {
                            // 确保将光标移动到文档开始
                            selection.HomeKey(6, 0); // wdStory = 6, wdMove = 0
                            
                            // 逐个处理公式
                            for (int i = 0; i < formulas.Count; i++)
                            {
                                string searchText = $"[公式{i+1}]";
                                Debug.WriteLine($"尝试查找公式占位符: {searchText}");
                                
                                // 简单地使用selection.Find直接查找和替换
                                selection.Find.ClearFormatting();
                                selection.Find.Text = searchText;
                                
                                bool found = selection.Find.Execute();
                                if (found)
                                {
                                    Debug.WriteLine($"成功找到占位符: {searchText}");
                                    
                                    // 获取公式信息
                                    string formula = formulas[i]["formula"].ToString();
                                    bool isDisplayMode = formulas[i]["isDisplayMode"].ToObject<bool>();
                                    
                                    // 删除占位符
                                    selection.Text = "";
                                    
                                    // 确保选区位置更新
                                    selection.Collapse(0); // 0 = wdCollapseStart
                                    
                                    // 关闭屏幕更新
                                    try
                                    {
                                        dynamic app = selection.Application;
                                        app.ScreenUpdating = false;
                                        
                                        // 记录公式插入前的位置
                                        int formulaStartPos = selection.Start;
                                        
                                        // 在当前位置插入公式
                                        bool success = InsertFormula(selection, formula, isDisplayMode);
                                        
                                        // 如果插入失败，确保光标位置正确
                                        if (!success)
                                        {
                                            selection.SetRange(formulaStartPos, selection.Start);
                                            selection.Collapse(0);
                                        }
                                        
                                        // 恢复屏幕更新
                                        app.ScreenUpdating = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"处理公式时屏幕更新控制失败: {ex.Message}");
                                        // 仍然尝试插入公式
                                        InsertFormula(selection, formula, isDisplayMode);
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine($"未找到占位符: {searchText}");
                                }
                            }
                            
                            Debug.WriteLine("所有公式处理完成");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"处理公式时出错: {ex.Message}");
                            MessageBox.Show($"处理公式时出错: {ex.Message}");
                        }
                    }

                    Debug.WriteLine("成功复制内容到Word");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"复制到Word失败: {ex.Message}");
                    Debug.WriteLine($"复制到Word失败详细信息: {ex}");
                }
                finally
                {
                    // 清理临时文件
                    try
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                    catch { /* 忽略清理错误 */ }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制到Word时出错: {ex.Message}");
                Debug.WriteLine($"复制到Word时出错详细信息: {ex}");
            }
        }

        /// <summary>
        /// 直接插入HTML内容到Word文档，保持格式
        /// </summary>
        /// <param name="htmlContent">HTML内容</param>
        public void InsertHtmlContentDirect(string htmlContent)
        {
            try
            {
                Debug.WriteLine($"开始插入HTML内容: {htmlContent.Substring(0, Math.Min(100, htmlContent.Length))}...");
                
                // 获取Word应用程序实例
                dynamic wordApp = GetWordApplication();
                dynamic doc = GetActiveDocument(wordApp);
                dynamic selection = wordApp.Selection;
                
                // 创建临时HTML文件
                string tempFile = Path.Combine(Path.GetTempPath(), $"WordCopilot_{Guid.NewGuid()}.html");
                try
                {
                    // 添加HTML基本标签和样式
                    string completeHtml = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{ 
            font-family: Calibri, Arial, sans-serif; 
            font-size: 11pt;
            line-height: 1.15;
            margin: 0;
            padding: 0;
        }}
        p {{ 
            margin: 0 0 8pt 0; 
            padding: 0;
        }}
        h1, h2, h3, h4, h5, h6 {{ 
            margin: 12pt 0 6pt 0;
            padding: 0;
            font-weight: bold;
        }}
        h1 {{ font-size: 16pt; }}
        h2 {{ font-size: 14pt; }}
        h3 {{ font-size: 12pt; }}
        ul, ol {{ 
            margin: 6pt 0 6pt 18pt; 
            padding: 0;
        }}
        li {{ 
            margin: 0 0 3pt 0; 
            padding: 0;
        }}
        strong, b {{ font-weight: bold; }}
        em, i {{ font-style: italic; }}
        code {{ 
            font-family: Consolas, 'Courier New', monospace; 
            background-color: #f0f0f0;
            padding: 1pt 3pt;
            border-radius: 3pt;
        }}
        pre {{ 
            background-color: #f0f0f0; 
            padding: 8pt; 
            border-radius: 4pt;
            font-family: Consolas, 'Courier New', monospace;
            margin: 6pt 0;
        }}
        table {{ 
            border-collapse: collapse; 
            width: 100%; 
            border: 1pt solid #000;
            margin: 6pt 0;
        }}
        th, td {{ 
            border: 1pt solid #000; 
            padding: 4pt 6pt;
            text-align: left;
        }}
        th {{ 
            background-color: #f2f2f2; 
            font-weight: bold;
        }}
        blockquote {{
            margin: 6pt 0 6pt 18pt;
            padding: 0 0 0 12pt;
            border-left: 3pt solid #ccc;
            font-style: italic;
        }}
    </style>
</head>
<body>
    {htmlContent}
</body>
</html>";

                    // 写入临时文件
                    File.WriteAllText(tempFile, completeHtml, Encoding.UTF8);

                    // 记录光标位置
                    int cursorStart = selection.Start;
                    Debug.WriteLine($"插入HTML内容前的光标位置: {cursorStart}");
                    
                    // 强制退出任何可能的公式编辑状态
                    try
                    {
                        if (selection.OMaths.Count > 0)
                        {
                            selection.MoveRight(1, 1);
                            selection.Collapse(0);
                        }
                        selection.ClearFormatting();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"退出公式状态时出错: {ex.Message}");
                    }
                    
                    // 插入HTML文件
                    selection.InsertFile(tempFile);
                    Debug.WriteLine($"HTML内容插入后的光标位置: {selection.Start}");
                    
                    // 处理插入后的表格样式
                    try
                    {
                        // 从当前插入点向上查找所有表格并应用样式
                        dynamic range = doc.Range(cursorStart, selection.Start);
                        dynamic tables = range.Tables;
                        
                        if (tables != null && tables.Count > 0)
                        {
                            Debug.WriteLine($"找到{tables.Count}个表格，应用边框样式");
                            
                            for (int i = 1; i <= tables.Count; i++)
                            {
                                dynamic table = tables[i];
                                // 应用网格线样式 (wdLineStyleSingle = 1)
                                table.Borders.InsideLineStyle = 1;
                                table.Borders.OutsideLineStyle = 1;
                                
                                // 设置边框宽度 (wdLineWidth025pt = 2，1/4磅)
                                table.Borders.InsideLineWidth = 2;
                                table.Borders.OutsideLineWidth = 2;
                                
                                // 自动调整表格
                                try { table.AutoFitBehavior(1); } catch { } // wdAutoFitContent = 1
                            }
                        }
                    }
                    catch (Exception exTable)
                    {
                        Debug.WriteLine($"应用表格样式时出错: {exTable.Message}");
                        // 继续处理，不中断流程
                    }
                    
                    // 确保光标位置正确
                    selection.Collapse(0); // 折叠到末尾
                    
                    Debug.WriteLine("HTML内容插入成功");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"插入HTML内容失败: {ex.Message}");
                    
                    // 回退：插入纯文本
                    try
                    {
                        // 移除HTML标签，只保留文本内容
                        string textContent = Regex.Replace(htmlContent, "<[^>]*>", "");
                        // 手动解码常见的HTML实体
                        textContent = textContent.Replace("&lt;", "<")
                                                .Replace("&gt;", ">")
                                                .Replace("&amp;", "&")
                                                .Replace("&quot;", "\"")
                                                .Replace("&#39;", "'")
                                                .Replace("&nbsp;", " ");
                        
                        if (!string.IsNullOrWhiteSpace(textContent))
                        {
                            selection.TypeText(textContent);
                            Debug.WriteLine("已回退为纯文本插入");
                        }
                    }
                    catch (Exception fallbackEx)
                    {
                        Debug.WriteLine($"回退插入也失败: {fallbackEx.Message}");
                    }
                }
                finally
                {
                    // 清理临时文件
                    try
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                        }
                    }
                    catch { /* 忽略清理错误 */ }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入HTML内容时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 将选中的文本内容插入到Word文档
        /// </summary>
        /// <param name="content">选中的文本内容</param>
        /// <param name="hasFormula">是否包含公式</param>
        public void InsertSelectedText(string content, bool hasFormula)
        {
            try
            {
                Debug.WriteLine($"开始插入选中内容: {content.Substring(0, Math.Min(50, content.Length))}...");
                
                // 获取Word应用程序实例
                dynamic wordApp = GetWordApplication();
                dynamic doc = GetActiveDocument(wordApp);
                dynamic selection = wordApp.Selection;
                
                // 保存原始位置和屏幕更新状态
                int startPos = selection.Start;
                bool originalScreenUpdating = wordApp.ScreenUpdating;
                
                try
                {
                    // 关闭屏幕更新
                    wordApp.ScreenUpdating = false;
                    
                    // 处理内容
                    if (hasFormula)
                    {
                        // 提取公式文本
                        string formula = content.Trim();
                        
                        // 是否为行内公式 $...$
                        if (formula.StartsWith("$") && formula.EndsWith("$") && !formula.StartsWith("$$"))
                        {
                            formula = formula.Substring(1, formula.Length - 2);
                            InsertFormula(selection, formula, false);
                        }
                        // 是否为行间公式 $$...$$
                        else if (formula.StartsWith("$$") && formula.EndsWith("$$"))
                        {
                            formula = formula.Substring(2, formula.Length - 4);
                            InsertFormula(selection, formula, true);
                        }
                        // 无标记公式，默认行内
                        else
                        {
                            InsertFormula(selection, formula, false);
                        }
                    }
                    else
                    {
                        // 普通文本直接插入
                        selection.TypeText(content);
                    }
                    
                    Debug.WriteLine("选中内容插入成功");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"选中内容插入过程出错: {ex.Message}");
                    
                    // 清理可能部分插入的内容
                    try
                    {
                        selection.SetRange(startPos, selection.Start);
                        selection.Delete();
                        selection.SetRange(startPos, startPos);
                        
                        // 回退为普通文本插入
                        selection.TypeText(content);
                    }
                    catch { }
                    
                    throw;
                }
                finally
                {
                    // 确保恢复屏幕更新
                    wordApp.ScreenUpdating = originalScreenUpdating;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入选中内容失败: {ex.Message}");
                MessageBox.Show($"插入到Word失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用简化方法插入公式到Word文档
        /// </summary>
        /// <param name="selection">Word选择区域</param>
        /// <param name="formulaText">公式文本</param>
        /// <param name="isDisplayMode">是否为显示模式</param>
        /// <returns>插入是否成功</returns>
        private bool InsertFormulaCleaner(dynamic selection, string formulaText, bool isDisplayMode)
        {
            Debug.WriteLine($"开始使用简化方法插入公式: {formulaText}, 显示模式: {isDisplayMode}");
            
            // 获取应用程序对象
            dynamic app = selection.Application;
            
            // 保存原始屏幕更新状态并暂停
            bool originalScreenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;
            
            // 记录原始位置
            int startPos = selection.Start;
            Debug.WriteLine($"公式插入起始位置: {startPos}");
            
            try
            {
                // 段落设置
                if (isDisplayMode)
                {
                    // 确保段落是居中的
                    selection.ParagraphFormat.Alignment = 1; // 居中对齐(wdAlignParagraphCenter)
                }
                
                // 使用最直接的方法插入公式
                try
                {
                    // 使用TypeText直接输入公式文本
                    selection.TypeText(formulaText);
                    Debug.WriteLine($"文本插入后位置: {selection.Start}");
                    
                    // 计算实际插入的文本长度
                    int endPos = selection.Start;
                    int actualLength = endPos - startPos;
                    Debug.WriteLine($"实际插入长度: {actualLength}, 预期长度: {formulaText.Length}");
                    
                    // 使用更安全的方式选择刚插入的文本
                    // 直接使用起始位置和结束位置创建范围，然后选择
                    dynamic doc = selection.Document;
                    dynamic range = doc.Range(startPos, endPos);
                    
                    // 验证范围内容是否正确
                    string rangeText = range.Text;
                    Debug.WriteLine($"选择的范围文本: '{rangeText}', 原始公式: '{formulaText}'");
                    
                    // 如果范围文本不匹配，说明有问题，直接返回失败
                    if (rangeText != formulaText)
                    {
                        Debug.WriteLine($"警告：范围文本不匹配！期望: '{formulaText}', 实际: '{rangeText}'");
                        // 不删除内容，保留已插入的文本
                        return false;
                    }
                    
                    // 选择范围
                    range.Select();
                    Debug.WriteLine($"已选择范围: {range.Start} 到 {range.End}");
                    
                    // 将选中的文本转换为公式
                    dynamic omathObj = selection.OMaths.Add(selection.Range);
                    
                    // 构建公式
                    omathObj.BuildUp();
                    Debug.WriteLine("公式构建完成");
                    
                    // 设置显示模式
                    if (isDisplayMode)
                    {
                        try
                        {
                            omathObj.DisplayType = 1; // wdOMathDisplay
                            Debug.WriteLine("设置为显示模式");
                            
                            // 添加新段落并恢复左对齐
                            selection.TypeParagraph();
                            selection.ParagraphFormat.Alignment = 0; // 左对齐(wdAlignParagraphLeft)
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"设置为显示模式失败: {ex.Message}");
                        }
                    }
                    
                    // 确保选区折叠到末尾
                    selection.Collapse(0);
                    
                    Debug.WriteLine("公式插入成功");
                    return true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"公式插入失败: {ex.Message}");
                    
                    // 检查是否有内容被插入
                    int currentPos = selection.Start;
                    if (currentPos > startPos)
                    {
                        Debug.WriteLine($"检测到部分内容已插入 (从 {startPos} 到 {currentPos})");
                        
                        // 检查插入的内容是否完整
                        try
                        {
                            dynamic doc = selection.Document;
                            dynamic range = doc.Range(startPos, currentPos);
                            string insertedText = range.Text;
                            Debug.WriteLine($"已插入的文本: '{insertedText}'");
                            
                            // 如果插入的文本等于原始公式文本，说明文本插入是成功的，只是公式转换失败
                            if (insertedText == formulaText)
                            {
                                Debug.WriteLine("文本插入完整，公式转换失败但保留文本");
                                return false; // 返回失败但不清理文本
                            }
                        }
                        catch (Exception checkEx)
                        {
                            Debug.WriteLine($"检查插入内容时出错: {checkEx.Message}");
                        }
                        
                        // 如果插入不完整，才进行清理
                        try
                        {
                            Debug.WriteLine("检测到不完整插入，进行清理");
                            selection.SetRange(startPos, currentPos);
                            selection.Delete();
                            selection.SetRange(startPos, startPos);
                            Debug.WriteLine("已清理失败的公式内容");
                        }
                        catch (Exception cleanupEx)
                        {
                            Debug.WriteLine($"清理失败的公式对象时出错: {cleanupEx.Message}");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("没有内容被插入，无需清理");
                    }
                    
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"公式插入过程出现未预期错误: {ex.Message}");
                
                // 只进行必要的清理
                try
                {
                    int currentPos = selection.Start;
                    if (currentPos != startPos)
                    {
                        Debug.WriteLine($"发生未预期错误，位置从 {startPos} 变为 {currentPos}");
                        // 只在位置确实改变时才清理
                        selection.SetRange(startPos, currentPos);
                        selection.Delete();
                        selection.SetRange(startPos, startPos);
                        Debug.WriteLine("已清理未预期错误产生的内容");
                    }
                }
                catch (Exception cleanupEx)
                {
                    Debug.WriteLine($"清理时出错: {cleanupEx.Message}");
                }
                
                return false;
            }
            finally
            {
                // 恢复屏幕更新状态
                app.ScreenUpdating = originalScreenUpdating;
            }
        }

        /// <summary>
        /// 向Word文档中插入数学公式
        /// </summary>
        /// <param name="selection">Word选择区域</param>
        /// <param name="formula">公式文本</param>
        /// <param name="isDisplayMode">是否是显示模式公式</param>
        /// <returns>插入是否成功</returns>
        public bool InsertFormula(dynamic selection, string formula, bool isDisplayMode)
        {
            Debug.WriteLine($"开始插入公式: {formula}, 显示模式: {isDisplayMode}");
            
            try
            {
                // 检查是否是对齐环境公式
                if (formula.Contains("\\begin{align") || formula.Contains("\\begin{aligned"))
                {
                    return InsertAlignedEquation(selection, formula);
                }
                
                // 预处理公式以移除不兼容的命令和环境
                string processedFormula = PreprocessSpecialFormulas(formula);
                
                // 使用简化的公式插入方法
                return InsertFormulaCleaner(selection, processedFormula, isDisplayMode);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"公式插入失败: {ex.Message}");
                
                // 不再自动插入回退文本，让上层方法处理
                // 这样可以避免重复插入内容
                return false;
            }
        }

        /// <summary>
        /// 插入对齐格式的公式到Word文档
        /// </summary>
        /// <param name="selection">Word选择区域</param>
        /// <param name="alignFormula">对齐公式文本</param>
        /// <returns>插入是否成功</returns>
        private bool InsertAlignedEquation(dynamic selection, string alignFormula)
        {
            try
            {
                Debug.WriteLine("开始处理对齐公式");
                
                // 预处理公式，移除对齐环境标记
                string processedFormula = PreprocessSpecialFormulas(alignFormula);
                Debug.WriteLine($"预处理后的对齐公式: {processedFormula}");
                
                // 获取应用程序对象
                dynamic app = selection.Application;
                
                // 保存原始屏幕更新状态并暂停
                bool originalScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;
                
                try
                {
                    // 记录原始位置
                    int startPos = selection.Start;
                    
                    // 是否为多行等式
                    if (processedFormula.Contains(";"))
                    {
                        string[] lines = processedFormula.Split(';');
                        bool success = true;
                        
                        // 处理多行等式
                        for (int i = 0; i < lines.Length; i++)
                        {
                            string line = lines[i].Trim();
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            
                            Debug.WriteLine($"处理第{i+1}行等式: {line}");
                            
                            // 清除当前选区位置
                            int lineStartPos = selection.Start;
                            
                            // 居中对齐
                            selection.ParagraphFormat.Alignment = 1; // 居中对齐
                            
                            try
                            {
                                // 直接输入公式文本
                                selection.TypeText(line);
                                
                                // 使用更安全的方式选择刚插入的文本
                                int lineEndPos = selection.Start;
                                dynamic doc = selection.Document;
                                dynamic range = doc.Range(lineStartPos, lineEndPos);
                                
                                // 验证范围内容
                                string rangeText = range.Text;
                                Debug.WriteLine($"第{i+1}行范围文本: '{rangeText}', 原始行: '{line}'");
                                
                                if (rangeText == line)
                                {
                                    // 选择范围并转换为公式
                                    range.Select();
                                    
                                    // 将选中的文本转换为公式
                                    dynamic omath = selection.OMaths.Add(selection.Range);
                                    omath.BuildUp();
                                    
                                    Debug.WriteLine($"第{i+1}行公式转换成功");
                                }
                                else
                                {
                                    Debug.WriteLine($"第{i+1}行文本不匹配，保留为文本");
                                }
                                
                                // 添加换行(除最后一行)
                                if (i < lines.Length - 1)
                                {
                                    selection.TypeParagraph();
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"第{i+1}行等式插入失败: {ex.Message}");
                                
                                // 清理此行可能的部分插入内容
                                selection.SetRange(lineStartPos, selection.Start);
                                selection.Delete();
                                selection.SetRange(lineStartPos, lineStartPos);
                                
                                // 回退为文本
                                selection.TypeText(line);
                                if (i < lines.Length - 1)
                                {
                                    selection.TypeParagraph();
                                }
                                
                                success = false;
                            }
                        }
                        
                        // 最后添加段落并恢复左对齐
                        selection.TypeParagraph();
                        selection.ParagraphFormat.Alignment = 0; // 左对齐
                        
                        return success;
                    }
                    else
                    {
                        // 单行等式处理
                        selection.ParagraphFormat.Alignment = 1; // 居中对齐
                        
                        try
                        {
                            // 直接输入公式文本
                            selection.TypeText(processedFormula);
                            
                            // 使用更安全的方式选择刚插入的文本
                            int endPos = selection.Start;
                            dynamic doc = selection.Document;
                            dynamic range = doc.Range(startPos, endPos);
                            
                            // 验证范围内容
                            string rangeText = range.Text;
                            Debug.WriteLine($"单行公式范围文本: '{rangeText}', 原始公式: '{processedFormula}'");
                            
                            if (rangeText == processedFormula)
                            {
                                // 选择范围并转换为公式
                                range.Select();
                                
                                // 将选中的文本转换为公式
                                dynamic omath = selection.OMaths.Add(selection.Range);
                                omath.BuildUp();
                                
                                // 尝试设置为显示模式
                                try { omath.DisplayType = 1; } catch { }
                                
                                Debug.WriteLine("单行公式转换成功");
                            }
                            else
                            {
                                Debug.WriteLine("单行公式文本不匹配，保留为文本");
                            }
                            
                            // 结束处理
                            selection.TypeParagraph();
                            selection.ParagraphFormat.Alignment = 0; // 左对齐
                            
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"单行对齐等式插入失败: {ex.Message}");
                            
                            // 清理失败内容
                            selection.SetRange(startPos, selection.Start);
                            selection.Delete();
                            selection.SetRange(startPos, startPos);
                            
                            // 回退为文本
                            selection.TypeText(processedFormula);
                            selection.TypeParagraph();
                            selection.ParagraphFormat.Alignment = 0; // 左对齐
                            
                            return false;
                        }
                    }
                }
                finally
                {
                    // 恢复屏幕更新
                    app.ScreenUpdating = originalScreenUpdating;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"对齐公式处理失败: {ex.Message}");
                
                // 插入原始文本作为回退
                try 
                {
                    selection.TypeText(alignFormula);
                    selection.TypeParagraph();
                } 
                catch { }
                
                return false;
            }
        }

        /// <summary>
        /// 预处理特殊公式，使其更适合Word公式编辑器
        /// </summary>
        /// <param name="formula">原始公式文本</param>
        /// <returns>处理后的公式文本</returns>
        private string PreprocessSpecialFormulas(string formula)
        {
            try
            {
                // 增加调试信息，检查公式内容和特殊字符
                Debug.WriteLine($"预处理公式原始内容: {formula}");
                Debug.WriteLine($"公式中是否包含align开始标记: {formula.Contains("\\begin{align}")}");
                Debug.WriteLine($"公式中是否包含align结束标记: {formula.Contains("\\end{align}")}");
                
                // 检查公式是否包含特殊字符的编码问题
                if (formula.Contains("\\begin"))
                {
                    Debug.WriteLine("检测到\\begin标记");
                    // 尝试查找具体的align标记
                    int beginIndex = formula.IndexOf("\\begin");
                    int endIndex = Math.Min(beginIndex + 20, formula.Length);
                    Debug.WriteLine($"\\begin附近内容: {formula.Substring(beginIndex, endIndex - beginIndex)}");
                }
                
                // 检测并预处理常见的特殊公式模式
                
                // 二次方程公式 - 单独处理最常见的二次方程式
                if ((formula.Contains("\\frac{-b") || formula.Contains("-b")) && 
                    (formula.Contains("\\sqrt{b^2-4ac}") || formula.Contains("b^2-4ac")) && 
                    formula.Contains("2a"))
                {
                    Debug.WriteLine("检测到二次方程公式，使用优化处理");
                    // 返回二次方程公式的优化线性格式，确保Word能正确识别
                    return "(-b±√(b^2-4ac))/(2a)";
                }
                
                // 强化对齐环境检测 - 使用更灵活的检测方式
                if ((formula.Contains("\\begin{align}") && formula.Contains("\\end{align}")) ||
                    (formula.Contains("align") && formula.Contains("&") && formula.Contains("\\\\")))
                {
                    Debug.WriteLine("检测到带对齐的公式环境，直接处理为多个独立公式");
                    
                    // 移除align环境标记 - 处理可能的空格和编码问题
                    string processedFormula = formula
                        .Replace("\\begin{align}", "")
                        .Replace("\\begin{align*}", "")
                        .Replace("\\end{align}", "")
                        .Replace("\\end{align*}", "")
                        .Trim();
                    
                    Debug.WriteLine($"移除align标记后: {processedFormula}");
                    
                    // 将换行符替换为分号，方便后续分割
                    processedFormula = processedFormula.Replace("\\\\", ";");
                    Debug.WriteLine($"替换换行符后: {processedFormula}");
                    
                    // 处理对齐标记
                    processedFormula = processedFormula.Replace("&", "=");
                    Debug.WriteLine($"替换对齐符号后: {processedFormula}");
                    
                    // 移除多余的空格和标记，简化输出
                    processedFormula = processedFormula.Replace("  ", " ").Trim();
                    
                    Debug.WriteLine($"对齐公式处理最终结果: {processedFormula}");
                    
                    // 如果这是一个简单的等式a = b + c, 就直接返回
                    if (processedFormula.Contains("=") && !processedFormula.Contains(";"))
                    {
                        return processedFormula;
                    }
                    
                    // 对于复杂的多行等式，创建一个线性格式
                    if (processedFormula.Contains(";"))
                    {
                        string[] lines = processedFormula.Split(';');
                        return string.Join(" ; ", lines.Select(line => line.Trim()));
                    }
                    
                    // 如果处理后还是没检测到预期格式，则返回处理后的结果
                    return processedFormula;
                }
                
                // 处理矩阵
                if (formula.Contains("\\begin{matrix}") || formula.Contains("\\begin{bmatrix}") || 
                    formula.Contains("\\begin{pmatrix}") || formula.Contains("\\begin{vmatrix}"))
                {
                    Debug.WriteLine("检测到矩阵环境");
                    formula = formula.Replace("\\begin{matrix}", "(")
                                   .Replace("\\end{matrix}", ")")
                                   .Replace("\\begin{bmatrix}", "[")
                                   .Replace("\\end{bmatrix}", "]")
                                   .Replace("\\begin{pmatrix}", "(")
                                   .Replace("\\end{pmatrix}", ")")
                                   .Replace("\\begin{vmatrix}", "|")
                                   .Replace("\\end{vmatrix}", "|")
                                   .Replace("\\\\", " ; ")
                                   .Replace("&", " , ");
                }
                
                return formula;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"预处理公式时出错: {ex.Message}");
                return formula;
            }
        }
        
        /// <summary>
        /// 插入纯文本内容 - 优化版本
        /// </summary>
        /// <param name="text">要插入的文本</param>
        public void InsertText(string text)
        {
            try
            {
                dynamic wordApp = GetWordApplication();
                dynamic selection = wordApp.Selection;
                
                if (!string.IsNullOrEmpty(text))
                {
                    Debug.WriteLine($"=== 开始插入文本 ===");
                    Debug.WriteLine($"文本内容: {text.Substring(0, Math.Min(100, text.Length))}...");
                    Debug.WriteLine($"插入前位置: {selection.Start}");
                    
                    // 简化的公式状态退出
                    try
                    {
                        if (selection.OMaths.Count > 0)
                        {
                            Debug.WriteLine("检测到公式编辑状态，退出");
                                    selection.MoveRight(1, 1);
                                    selection.Collapse(0);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"退出公式状态时出错: {ex.Message}");
                    }
                    
                    // 基本格式设置（不过度重置）
                    try
                    {
                        // 只设置必要的格式属性
                        selection.Font.Name = "Calibri";
                        selection.Font.Size = 11;
                        selection.ParagraphFormat.Alignment = 0; // 左对齐
                        
                        Debug.WriteLine("已设置基本格式");
                    }
                    catch (Exception formatEx)
                    {
                        Debug.WriteLine($"设置格式失败: {formatEx.Message}");
                    }
                    
                    // 直接插入文本
                    selection.TypeText(text);
                    Debug.WriteLine($"文本插入完成，新位置: {selection.Start}");
                    
                    Debug.WriteLine("=== 文本插入流程完成 ===");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入文本时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 插入数学公式 - 优化版本
        /// </summary>
        /// <param name="formula">LaTeX格式的公式</param>
        public void InsertEquation(string formula)
        {
            try
            {
                dynamic wordApp = GetWordApplication();
                dynamic selection = wordApp.Selection;
                
                // 记录插入前的位置
                int startPos = selection.Start;
                Debug.WriteLine($"公式插入开始位置: {startPos}");
                Debug.WriteLine($"原始公式: {formula}");
                
                // 简化的公式状态清理
                try
                {
                    if (selection.OMaths.Count > 0)
                    {
                        Debug.WriteLine("检测到当前在公式编辑状态，退出");
                        selection.MoveRight(1, 1);
                        selection.Collapse(0);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"清除公式状态时出错: {ex.Message}");
                }
                
                // 只在必要时添加换行
                if (selection.Start > 0)
                {
                    // 检查当前行是否有内容
                    try
                    {
                    var currentPara = selection.Paragraphs[1];
                        string currentText = currentPara.Range.Text.Trim();
                        if (currentText.Length > 0 && !currentText.EndsWith("\r"))
                    {
                            selection.TypeText("\r");
                        }
                    }
                    catch
                    {
                        // 如果无法检查段落，就添加一个换行
                        selection.TypeText("\r");
                    }
                }
                
                // 记录公式插入的实际开始位置
                int formulaStartPos = selection.Start;
                Debug.WriteLine($"公式实际插入位置: {formulaStartPos}");
                
                // 居中对齐公式
                selection.ParagraphFormat.Alignment = 1; // wdAlignParagraphCenter = 1
                
                // 预处理公式
                string processedFormula = formula;
                try
                {
                    processedFormula = PreprocessSpecialFormulas(formula);
                    Debug.WriteLine($"预处理后公式: {processedFormula}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"公式预处理失败: {ex.Message}");
                    processedFormula = formula; // 使用原始公式
                }
                
                // 公式插入尝试
                bool success = false;
                bool hasTextInserted = false;
                
                try
                {
                    success = InsertFormula(selection, processedFormula, true);
                    
                    // 检查位置是否变化
                    int currentPos = selection.Start;
                    hasTextInserted = (currentPos > formulaStartPos);
                    
                    if (success)
                    {
                        Debug.WriteLine("公式插入成功");
                    }
                    else
                    {
                        Debug.WriteLine($"公式插入失败，位置变化: {hasTextInserted}");
                        
                        // 如果没有任何内容被插入，插入原始公式文本
                        if (!hasTextInserted)
                        {
                            Debug.WriteLine("插入原始公式文本作为回退");
                                selection.TypeText(formula);
                                hasTextInserted = true;
                            }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"公式插入过程出错: {ex.Message}");
                    
                    // 异常处理：确保有内容被插入
                        if (!hasTextInserted)
                        {
                            try
                            {
                                selection.TypeText(formula);
                                hasTextInserted = true;
                            }
                        catch (Exception fallbackEx)
                            {
                            Debug.WriteLine($"回退插入失败: {fallbackEx.Message}");
                    }
                }
                }
                
                // 简化的格式重置
                try
                {
                    // 添加一个换行以分隔公式和后续内容
                    selection.TypeText("\r");
                    
                    // 重置为正常文本格式
                    selection.ParagraphFormat.Alignment = 0; // 左对齐
                    selection.Font.Name = "Calibri";
                    selection.Font.Size = 11;
                    
                    Debug.WriteLine("格式重置完成");
                }
                catch (Exception resetEx)
                {
                    Debug.WriteLine($"重置格式时出错: {resetEx.Message}");
                }
                
                if (success)
                {
                    Debug.WriteLine($"公式插入和格式化完成: {formula}");
                }
                else if (hasTextInserted)
                {
                    Debug.WriteLine($"公式以文本形式插入: {formula}");
                }
                else
                {
                    Debug.WriteLine($"警告：公式插入失败: {formula}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入公式时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 插入代码块 - 优化版本，与单独插入效果一致
        /// </summary>
        /// <param name="code">代码内容</param>
        public void InsertCode(string code)
        {
            try
            {
                Debug.WriteLine($"=== 开始插入代码块 ===");
                Debug.WriteLine($"代码内容: {code.Substring(0, Math.Min(100, code.Length))}...");
                
                dynamic wordApp = GetWordApplication();
                dynamic selection = wordApp.Selection;
                
                if (!string.IsNullOrEmpty(code))
                {
                    // 记录插入前的位置
                    int startPos = selection.Start;
                    Debug.WriteLine($"插入前位置: {startPos}");
                    
                    // 简化的公式状态退出
                    try
                    {
                        if (selection.OMaths.Count > 0)
                        {
                            Debug.WriteLine("检测到公式编辑状态，退出");
                            selection.MoveRight(1, 1);
                            selection.Collapse(0);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"退出公式状态时出错: {ex.Message}");
                    }
                    
                    // 确保在新行开始
                    if (selection.Start > 0)
                    {
                        try
                    {
                        var currentPara = selection.Paragraphs[1];
                            string currentText = currentPara.Range.Text.Trim();
                            if (currentText.Length > 0 && !currentText.EndsWith("\r"))
                        {
                                selection.TypeText("\r");
                            }
                        }
                        catch
                        {
                            // 如果无法检查段落，就添加一个换行
                            selection.TypeText("\r");
                        }
                    }
                    
                    // 使用HTML方式插入代码块，保持与单独插入一致的效果
                    string codeHTML = $"<pre style=\"background-color: #f0f0f0; padding: 8pt; border-radius: 4pt; font-family: Consolas, 'Courier New', monospace; margin: 6pt 0;\"><code>{HtmlEncode(code)}</code></pre>";
                    
                    // 创建临时HTML文件
                    string tempFile = Path.Combine(Path.GetTempPath(), $"WordCopilot_Code_{Guid.NewGuid()}.html");
                    try
                    {
                        // 创建完整的HTML文档
                        string completeHtml = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{ 
            font-family: Calibri, Arial, sans-serif; 
            font-size: 11pt;
            line-height: 1.15;
            margin: 0;
            padding: 0;
        }}
        pre {{ 
            background-color: #f0f0f0; 
            padding: 8pt; 
            border-radius: 4pt;
            font-family: Consolas, 'Courier New', monospace;
            margin: 6pt 0;
            font-size: 10pt;
            border: 1pt solid #d0d0d0;
            overflow-x: auto;
        }}
        code {{ 
            font-family: Consolas, 'Courier New', monospace;
            font-size: 10pt;
            white-space: pre-wrap;
            word-wrap: break-word;
        }}
    </style>
</head>
<body>
    {codeHTML}
</body>
</html>";

                        // 写入临时文件
                        File.WriteAllText(tempFile, completeHtml, Encoding.UTF8);
                        
                        // 记录插入前的位置
                        int insertStartPos = selection.Start;
                        Debug.WriteLine($"HTML代码块插入前位置: {insertStartPos}");
                        
                        // 插入HTML文件
                        selection.InsertFile(tempFile);
                        Debug.WriteLine($"HTML代码块插入后位置: {selection.Start}");
                        
                        // 添加一个换行，确保后续内容正确分隔
                        selection.TypeText("\r");
                        
                        // 重置为正常文本格式
                        selection.Font.Name = "Calibri";
                        selection.Font.Size = 11;
                        selection.ParagraphFormat.Alignment = 0; // 左对齐
                        selection.ParagraphFormat.LeftIndent = 0;
                        selection.ParagraphFormat.RightIndent = 0;
                        selection.ParagraphFormat.SpaceAfter = 0;
                        selection.ParagraphFormat.SpaceBefore = 0;
                        selection.Shading.BackgroundPatternColor = 16777215; // 白色背景
                        
                        Debug.WriteLine("代码块HTML插入成功");
                    }
                    catch (Exception htmlEx)
                    {
                        Debug.WriteLine($"HTML代码块插入失败: {htmlEx.Message}");
                        Debug.WriteLine("回退到传统代码块插入方式");
                        
                        // 回退到传统方式
                        try
                        {
                    // 设置代码块格式
                    selection.Font.Name = "Consolas";
                    selection.Font.Size = 10;
                    selection.Font.Bold = false;
                    selection.Font.Italic = false;
                    
                    // 设置段落格式
                    selection.ParagraphFormat.Alignment = 0; // 左对齐
                    selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1); // 缩进1厘米
                    selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(1); // 右缩进1厘米
                    selection.ParagraphFormat.SpaceAfter = 6; // 段后间距
                    selection.ParagraphFormat.SpaceBefore = 6; // 段前间距
                    
                    // 设置背景色
                    selection.Shading.BackgroundPatternColor = 15658734; // 浅灰色背景
                    
                    // 插入代码内容
                    selection.TypeText(code);
                    
                    // 重置格式
                    selection.Font.Name = "Calibri";
                    selection.Font.Size = 11;
                    selection.ParagraphFormat.LeftIndent = 0;
                    selection.ParagraphFormat.RightIndent = 0;
                    selection.ParagraphFormat.SpaceAfter = 0;
                    selection.ParagraphFormat.SpaceBefore = 0;
                    selection.Shading.BackgroundPatternColor = 16777215; // 白色背景
                    
                            Debug.WriteLine("传统代码块插入成功");
                        }
                        catch (Exception fallbackEx)
                        {
                            Debug.WriteLine($"传统代码块插入也失败: {fallbackEx.Message}");
                            throw;
                        }
                    }
                    finally
                    {
                        // 清理临时文件
                        try
                        {
                            if (File.Exists(tempFile))
                            {
                                File.Delete(tempFile);
                            }
                        }
                        catch { /* 忽略清理错误 */ }
                    }
                    
                    Debug.WriteLine("=== 代码块插入流程完成 ===");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入代码时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 插入换行符
        /// </summary>
        public void InsertLineBreak()
        {
            try
            {
                dynamic wordApp = GetWordApplication();
                dynamic selection = wordApp.Selection;
                
                // 强制退出任何可能的公式编辑状态
                try
                {
                    // 如果当前在公式编辑器中，先退出
                    if (selection.OMaths.Count > 0)
                    {
                        // 移动到公式外
                        selection.MoveRight(1, 1);
                        selection.Collapse(0);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"在换行时退出公式状态出错: {ex.Message}");
                }
                
                // 插入段落换行
                selection.TypeText("\r");
                
                // 重置格式，确保下一行是正常格式
                try
                {
                    selection.ClearFormatting();
                    selection.Font.Name = "Calibri";
                    selection.Font.Size = 11;
                    selection.Font.Bold = false;
                    selection.Font.Italic = false;
                    selection.ParagraphFormat.Alignment = 0; // 左对齐
                    selection.ParagraphFormat.LeftIndent = 0;
                    selection.ParagraphFormat.RightIndent = 0;
                    selection.Shading.BackgroundPatternColor = 16777215; // 白色背景
                }
                catch (Exception formatEx)
                {
                    Debug.WriteLine($"重置换行格式时出错: {formatEx.Message}");
                }
                
                Debug.WriteLine("插入换行符并重置格式");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入换行符时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="tableData">表格数据（JSON格式字符串或JObject）</param>
        public void InsertTable(object tableData)
        {
            try
            {
                dynamic wordApp = GetWordApplication();
                dynamic selection = wordApp.Selection;
                
                // 解析表格数据
                JObject tableJson = null;
                if (tableData is string jsonStr)
                {
                    tableJson = JObject.Parse(jsonStr);
                }
                else if (tableData is JObject jobj)
                {
                    tableJson = jobj;
                }
                else
                {
                    Debug.WriteLine("无效的表格数据格式");
                    return;
                }
                
                // 获取表格信息
                var headers = tableJson["headers"]?.ToObject<string[]>();
                var rows = tableJson["rows"]?.ToObject<string[][]>();
                
                if (headers == null || rows == null)
                {
                    Debug.WriteLine("表格数据缺少必要的头部或行信息");
                    return;
                }
                
                int numCols = headers.Length;
                int numRows = rows.Length + 1; // +1 for header row
                
                Debug.WriteLine($"开始插入表格: {numRows} 行 x {numCols} 列");
                
                // 强制退出任何可能的公式编辑状态
                try
                {
                    if (selection.OMaths.Count > 0)
                    {
                        selection.MoveRight(1, 1);
                        selection.Collapse(0);
                    }
                    selection.ClearFormatting();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"退出公式状态时出错: {ex.Message}");
                }
                
                // 确保在新行开始
                if (selection.Start > 0)
                {
                    var currentPara = selection.Paragraphs[1];
                    if (currentPara.Range.Text.Trim().Length > 0)
                    {
                        selection.TypeText("\r");
                    }
                }
                
                // 创建表格
                dynamic range = selection.Range;
                dynamic table = range.Tables.Add(range, numRows, numCols);
                
                // 设置表格样式
                table.Borders.InsideLineStyle = 1; // wdLineStyleSingle
                table.Borders.OutsideLineStyle = 1;
                table.Borders.InsideLineWidth = 2; // wdLineWidth025pt
                table.Borders.OutsideLineWidth = 2;
                
                // 设置表头
                for (int col = 1; col <= numCols; col++)
                {
                    var cell = table.Cell(1, col);
                    cell.Range.Text = headers[col - 1];
                    
                    // 设置表头格式
                    cell.Range.Font.Bold = true;
                    cell.Shading.BackgroundPatternColor = 15658734; // 浅灰色背景
                    cell.Range.ParagraphFormat.Alignment = 1; // 居中对齐
                }
                
                // 填充数据行
                for (int row = 0; row < rows.Length; row++)
                {
                    for (int col = 1; col <= numCols && col - 1 < rows[row].Length; col++)
                    {
                        var cell = table.Cell(row + 2, col); // +2 because row 1 is header
                        cell.Range.Text = rows[row][col - 1];
                        
                        // 设置数据单元格格式
                        cell.Range.Font.Bold = false;
                        cell.Range.ParagraphFormat.Alignment = 0; // 左对齐
                    }
                }
                
                // 自动调整表格
                try
                {
                    table.AutoFitBehavior(1); // wdAutoFitContent
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"自动调整表格时出错: {ex.Message}");
                }
                
                // 移动光标到表格后面
                selection.SetRange(table.Range.End + 1, table.Range.End + 1);
                selection.TypeText("\r"); // 添加段落分隔
                
                // 重置格式
                selection.ClearFormatting();
                selection.Font.Name = "Calibri";
                selection.Font.Size = 11;
                selection.Font.Bold = false;
                selection.ParagraphFormat.Alignment = 0; // 左对齐
                selection.Shading.BackgroundPatternColor = 16777215; // 白色背景
                
                Debug.WriteLine($"表格插入成功: {numRows} 行 x {numCols} 列");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入表格时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 插入Mermaid图片到Word文档
        /// </summary>
        /// <param name="imageData">PNG格式的base64图片数据</param>
        /// <param name="mermaidCode">Mermaid源代码（用于备注）</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        /// <returns>是否插入成功</returns>
        public bool InsertMermaidImage(string imageData, string mermaidCode, int width, int height)
        {
            try
            {
                Debug.WriteLine("=== 开始插入Mermaid图片 ===");
                
                // 获取Word应用程序和选择区域
                dynamic wordApp = GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法获取Word应用程序");
                    return false;
                }

                dynamic doc = GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    return false;
                }

                dynamic selection = wordApp.Selection;
                
                // 保持当前光标位置，不移动到文档末尾
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
                        selection.TypeText("\r\n");
                    }
                    catch { }
                }
                
                // 解析base64图片数据
                if (imageData.StartsWith("data:image/png;base64,"))
                {
                    imageData = imageData.Substring("data:image/png;base64,".Length);
                }
                
                // 将base64转换为字节数组
                byte[] imageBytes = Convert.FromBase64String(imageData);
                
                // 创建临时文件
                string tempImagePath = Path.GetTempFileName();
                string tempImageFile = Path.ChangeExtension(tempImagePath, ".png");
                
                try
                {
                    // 写入图片文件
                    File.WriteAllBytes(tempImageFile, imageBytes);
                    Debug.WriteLine($"临时图片文件已创建: {tempImageFile}");
                    
                    // 插入图片到Word
                    dynamic inlineShape = selection.InlineShapes.AddPicture(
                        tempImageFile,
                        false, // LinkToFile
                        true   // SaveWithDocument
                    );
                    
                // 设置图片尺寸：默认“适配页面可用宽度”，避免侧边栏预览较小导致插入后看不清
                // width/height 作为比例参考即可（不再直接按像素->磅）
                double aspect = 0.75;
                if (width > 0 && height > 0)
                {
                    aspect = (double)height / (double)width;
                }

                double availableWidth = 400;  // 兜底
                double availableHeight = 600; // 兜底
                try
                {
                    availableWidth = (double)doc.PageSetup.PageWidth - (double)doc.PageSetup.LeftMargin - (double)doc.PageSetup.RightMargin;
                    availableHeight = (double)doc.PageSetup.PageHeight - (double)doc.PageSetup.TopMargin - (double)doc.PageSetup.BottomMargin;
                }
                catch (Exception psEx)
                {
                    Debug.WriteLine($"获取页面可用宽高失败，使用兜底值: {psEx.Message}");
                }

                // 目标：尽量铺满页面宽度；高度过高则按高度上限回缩
                double targetWidth = Math.Max(240, availableWidth * 0.95);
                double maxHeight = Math.Max(240, availableHeight * 0.70);
                double targetHeight = targetWidth * aspect;
                if (targetHeight > maxHeight)
                {
                    targetHeight = maxHeight;
                    targetWidth = targetHeight / aspect;
                }

                try
                {
                    inlineShape.LockAspectRatio = true;
                }
                catch { }

                inlineShape.Width = targetWidth;
                inlineShape.Height = targetHeight;

                // 居中显示（不影响文档其它段落）
                try
                {
                    inlineShape.Range.ParagraphFormat.Alignment = 1; // wdAlignParagraphCenter
                }
                catch { }

                Debug.WriteLine($"图片插入成功，尺寸: {targetWidth}x{targetHeight}磅（适配页面宽度）");
                    
                    // 图片插入后移动光标到图片后面
                    selection.MoveRight(1, 1);
                    
                    Debug.WriteLine("=== Mermaid图片插入完成 ===");
                    return true;
                }
                finally
                {
                    // 清理临时文件
                    try
                    {
                        if (File.Exists(tempImageFile))
                        {
                            File.Delete(tempImageFile);
                            Debug.WriteLine($"临时文件已删除: {tempImageFile}");
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        Debug.WriteLine($"清理临时文件时出错: {cleanupEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"=== 插入Mermaid图片时出错 ===");
                Debug.WriteLine($"错误信息: {ex.Message}");
                Debug.WriteLine($"错误堆栈: {ex.StackTrace}");
                return false;
            }
        }
        
        /// <summary>
        /// 从SVG数据插入Mermaid图片到Word文档
        /// </summary>
        /// <param name="svgData">SVG数据</param>
        /// <param name="mermaidCode">Mermaid源代码（用于备注）</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        /// <returns>是否插入成功</returns>
        public bool InsertMermaidImageFromSvg(string svgData, string mermaidCode, int width, int height)
        {
            try
            {
                Debug.WriteLine("=== 开始从SVG插入Mermaid图片 ===");
                
                // 获取Word应用程序和选择区域
                dynamic wordApp = GetWordApplication();
                if (wordApp == null)
                {
                    Debug.WriteLine("无法获取Word应用程序");
                    return false;
                }

                dynamic doc = GetActiveDocument(wordApp);
                if (doc == null)
                {
                    Debug.WriteLine("无法获取活动文档");
                    return false;
                }

                dynamic selection = wordApp.Selection;
                
                // 保持当前光标位置，不移动到文档末尾
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
                        selection.TypeText("\r\n");
                    }
                    catch { }
                }
                
                // 使用简化的方法：直接将SVG转换为临时文件并插入
                // 创建临时SVG文件
                string tempSvgPath = Path.GetTempFileName();
                string tempSvgFile = Path.ChangeExtension(tempSvgPath, ".svg");
                
                try
                {
                    // 写入SVG文件
                    File.WriteAllText(tempSvgFile, svgData, Encoding.UTF8);
                    Debug.WriteLine($"临时SVG文件已创建: {tempSvgFile}");
                    
                    // 尝试直接插入SVG（Word 2016+支持）
                    try
                    {
                        dynamic inlineShape = selection.InlineShapes.AddPicture(
                            tempSvgFile,
                            false, // LinkToFile
                            true   // SaveWithDocument
                        );
                        
                        // 设置图片尺寸：同 PNG，默认适配页面可用宽度（矢量可无损缩放）
                        double aspect = 0.75;
                        if (width > 0 && height > 0)
                        {
                            aspect = (double)height / (double)width;
                        }

                        double availableWidth = 400;  // 兜底
                        double availableHeight = 600; // 兜底
                        try
                        {
                            availableWidth = (double)doc.PageSetup.PageWidth - (double)doc.PageSetup.LeftMargin - (double)doc.PageSetup.RightMargin;
                            availableHeight = (double)doc.PageSetup.PageHeight - (double)doc.PageSetup.TopMargin - (double)doc.PageSetup.BottomMargin;
                        }
                        catch (Exception psEx)
                        {
                            Debug.WriteLine($"获取页面可用宽高失败，使用兜底值: {psEx.Message}");
                        }

                        double targetWidth = Math.Max(240, availableWidth * 0.95);
                        double maxHeight = Math.Max(240, availableHeight * 0.70);
                        double targetHeight = targetWidth * aspect;
                        if (targetHeight > maxHeight)
                        {
                            targetHeight = maxHeight;
                            targetWidth = targetHeight / aspect;
                        }

                        try
                        {
                            inlineShape.LockAspectRatio = true;
                        }
                        catch { }

                        inlineShape.Width = targetWidth;
                        inlineShape.Height = targetHeight;

                        try
                        {
                            inlineShape.Range.ParagraphFormat.Alignment = 1; // wdAlignParagraphCenter
                        }
                        catch { }

                        Debug.WriteLine($"SVG图片插入成功，尺寸: {targetWidth}x{targetHeight}磅（适配页面宽度）");
                        
                        // 图片插入后移动光标到图片后面
                        selection.MoveRight(1, 1);
                        
                        Debug.WriteLine("=== SVG Mermaid图片插入完成 ===");
                        return true;
                    }
                    catch (Exception svgEx)
                    {
                        Debug.WriteLine($"直接插入SVG失败，尝试其他方法: {svgEx.Message}");
                        
                        // 如果SVG插入失败，尝试转换为HTML并插入
                        string htmlContent = $@"
                        <div style='text-align: center; margin: 10px 0;'>
                            {svgData}
                        </div>";
                        
                        // 创建临时HTML文件
                        string tempHtmlPath = Path.GetTempFileName();
                        string tempHtmlFile = Path.ChangeExtension(tempHtmlPath, ".html");
                        
                        try
                        {
                            string fullHtml = $@"
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <meta charset='utf-8'>
                                <style>
                                    body {{ margin: 0; padding: 10px; font-family: Arial, sans-serif; }}
                                    svg {{ max-width: 100%; height: auto; }}
                                </style>
                            </head>
                            <body>
                                {htmlContent}
                            </body>
                            </html>";
                            
                            File.WriteAllText(tempHtmlFile, fullHtml, Encoding.UTF8);
                            
                            // 插入HTML文件
                            selection.InsertFile(tempHtmlFile);
                            
                            Debug.WriteLine("通过HTML方式插入SVG成功");
                            return true;
                        }
                        catch (Exception htmlEx)
                        {
                            Debug.WriteLine($"HTML方式插入也失败: {htmlEx.Message}");
                            return false;
                        }
                        finally
                        {
                            try
                            {
                                if (File.Exists(tempHtmlFile))
                                {
                                    File.Delete(tempHtmlFile);
                                }
                            }
                            catch { }
                        }
                    }
                }
                finally
                {
                    // 清理临时SVG文件
                    try
                    {
                        if (File.Exists(tempSvgFile))
                        {
                            File.Delete(tempSvgFile);
                            Debug.WriteLine($"临时SVG文件已删除: {tempSvgFile}");
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        Debug.WriteLine($"清理临时SVG文件时出错: {cleanupEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"=== 从SVG插入Mermaid图片时出错 ===");
                Debug.WriteLine($"错误信息: {ex.Message}");
                Debug.WriteLine($"错误堆栈: {ex.StackTrace}");
                return false;
            }
        }
    }
}

