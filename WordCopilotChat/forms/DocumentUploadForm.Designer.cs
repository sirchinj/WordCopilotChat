namespace WordCopilotChat
{
    partial class DocumentUploadForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBoxFileSelect = new System.Windows.Forms.GroupBox();
            this.labelFileSize = new System.Windows.Forms.Label();
            this.labelSelectedFile = new System.Windows.Forms.Label();
            this.buttonSelectFile = new System.Windows.Forms.Button();
            this.labelInstruction = new System.Windows.Forms.Label();
            this.groupBoxParse = new System.Windows.Forms.GroupBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelParseResult = new System.Windows.Forms.Label();
            this.buttonParse = new System.Windows.Forms.Button();
            this.labelParseInfo = new System.Windows.Forms.Label();
            this.textBoxParseLog = new System.Windows.Forms.TextBox();
            this.groupBoxHeadings = new System.Windows.Forms.GroupBox();
            this.listViewHeadings = new System.Windows.Forms.ListView();
            this.columnHeaderHeading = new System.Windows.Forms.ColumnHeader();
            this.columnHeaderLevel = new System.Windows.Forms.ColumnHeader();
            this.columnHeaderContentLength = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStripHeadings = new System.Windows.Forms.ContextMenuStrip();
            this.toolStripMenuItemViewContent = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemRemove = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItemSelectAll = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemUnselectAll = new System.Windows.Forms.ToolStripMenuItem();
            this.buttonUpload = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBoxFileSelect.SuspendLayout();
            this.groupBoxParse.SuspendLayout();
            this.groupBoxHeadings.SuspendLayout();
            this.contextMenuStripHeadings.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxFileSelect
            // 
            this.groupBoxFileSelect.Controls.Add(this.labelFileSize);
            this.groupBoxFileSelect.Controls.Add(this.labelSelectedFile);
            this.groupBoxFileSelect.Controls.Add(this.buttonSelectFile);
            this.groupBoxFileSelect.Controls.Add(this.labelInstruction);
            this.groupBoxFileSelect.Location = new System.Drawing.Point(12, 12);
            this.groupBoxFileSelect.Name = "groupBoxFileSelect";
            this.groupBoxFileSelect.Size = new System.Drawing.Size(660, 100);
            this.groupBoxFileSelect.TabIndex = 0;
            this.groupBoxFileSelect.TabStop = false;
            this.groupBoxFileSelect.Text = "1. 选择文档文件";
            // 
            // labelFileSize
            // 
            this.labelFileSize.AutoSize = true;
            this.labelFileSize.ForeColor = System.Drawing.Color.Gray;
            this.labelFileSize.Location = new System.Drawing.Point(15, 70);
            this.labelFileSize.Name = "labelFileSize";
            this.labelFileSize.Size = new System.Drawing.Size(0, 12);
            this.labelFileSize.TabIndex = 3;
            // 
            // labelSelectedFile
            // 
            this.labelSelectedFile.AutoSize = true;
            this.labelSelectedFile.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelSelectedFile.Location = new System.Drawing.Point(15, 50);
            this.labelSelectedFile.Name = "labelSelectedFile";
            this.labelSelectedFile.Size = new System.Drawing.Size(70, 12);
            this.labelSelectedFile.TabIndex = 2;
            this.labelSelectedFile.Text = "未选择文件";
            // 
            // buttonSelectFile
            // 
            this.buttonSelectFile.Location = new System.Drawing.Point(570, 25);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(80, 30);
            this.buttonSelectFile.TabIndex = 1;
            this.buttonSelectFile.Text = "选择文件";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // labelInstruction
            // 
            this.labelInstruction.AutoSize = true;
            this.labelInstruction.Location = new System.Drawing.Point(15, 25);
            this.labelInstruction.Name = "labelInstruction";
            this.labelInstruction.Size = new System.Drawing.Size(329, 12);
            this.labelInstruction.TabIndex = 0;
            this.labelInstruction.Text = "请选择要上传的文档文件，支持格式：Word (.docx, .doc) 和 Markdown (.md)";
            // 
            // groupBoxParse
            // 
            this.groupBoxParse.Controls.Add(this.textBoxParseLog);
            this.groupBoxParse.Controls.Add(this.progressBar);
            this.groupBoxParse.Controls.Add(this.labelParseResult);
            this.groupBoxParse.Controls.Add(this.buttonParse);
            this.groupBoxParse.Controls.Add(this.labelParseInfo);
            this.groupBoxParse.Location = new System.Drawing.Point(12, 128);
            this.groupBoxParse.Name = "groupBoxParse";
            this.groupBoxParse.Size = new System.Drawing.Size(660, 160);
            this.groupBoxParse.TabIndex = 1;
            this.groupBoxParse.TabStop = false;
            this.groupBoxParse.Text = "2. 解析文档结构";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 130);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(540, 20);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.TabIndex = 3;
            this.progressBar.Visible = false;
            // 
            // textBoxParseLog
            // 
            this.textBoxParseLog.Location = new System.Drawing.Point(15, 70);
            this.textBoxParseLog.Multiline = true;
            this.textBoxParseLog.Name = "textBoxParseLog";
            this.textBoxParseLog.ReadOnly = true;
            this.textBoxParseLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxParseLog.Size = new System.Drawing.Size(540, 55);
            this.textBoxParseLog.TabIndex = 4;
            this.textBoxParseLog.Visible = false;
            // 
            // labelParseResult
            // 
            this.labelParseResult.AutoSize = true;
            this.labelParseResult.Location = new System.Drawing.Point(15, 50);
            this.labelParseResult.Name = "labelParseResult";
            this.labelParseResult.Size = new System.Drawing.Size(0, 12);
            this.labelParseResult.TabIndex = 2;
            // 
            // buttonParse
            // 
            this.buttonParse.Enabled = false;
            this.buttonParse.Location = new System.Drawing.Point(570, 25);
            this.buttonParse.Name = "buttonParse";
            this.buttonParse.Size = new System.Drawing.Size(80, 30);
            this.buttonParse.TabIndex = 1;
            this.buttonParse.Text = "解析文档";
            this.buttonParse.UseVisualStyleBackColor = true;
            this.buttonParse.Click += new System.EventHandler(this.buttonParse_Click);
            // 
            // labelParseInfo
            // 
            this.labelParseInfo.AutoSize = true;
            this.labelParseInfo.Location = new System.Drawing.Point(15, 25);
            this.labelParseInfo.Name = "labelParseInfo";
            this.labelParseInfo.Size = new System.Drawing.Size(365, 12);
            this.labelParseInfo.TabIndex = 0;
            this.labelParseInfo.Text = "点击解析文档，系统将自动提取文档中的标题和内容（支持标题样式和大字体识别）。";
            // 
            // groupBoxHeadings
            // 
            this.groupBoxHeadings.Controls.Add(this.listViewHeadings);
            this.groupBoxHeadings.Location = new System.Drawing.Point(12, 304);
            this.groupBoxHeadings.Name = "groupBoxHeadings";
            this.groupBoxHeadings.Size = new System.Drawing.Size(660, 280);
            this.groupBoxHeadings.TabIndex = 2;
            this.groupBoxHeadings.TabStop = false;
            this.groupBoxHeadings.Text = "3. 预览文档结构（双击查看内容，右键可剔除标题）";
            // 
            // listViewHeadings
            // 
            this.listViewHeadings.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderHeading,
            this.columnHeaderLevel,
            this.columnHeaderContentLength});
            this.listViewHeadings.ContextMenuStrip = this.contextMenuStripHeadings;
            this.listViewHeadings.FullRowSelect = true;
            this.listViewHeadings.GridLines = true;
            this.listViewHeadings.Location = new System.Drawing.Point(15, 25);
            this.listViewHeadings.Name = "listViewHeadings";
            this.listViewHeadings.Size = new System.Drawing.Size(635, 240);
            this.listViewHeadings.TabIndex = 0;
            this.listViewHeadings.UseCompatibleStateImageBehavior = false;
            this.listViewHeadings.View = System.Windows.Forms.View.Details;
            this.listViewHeadings.DoubleClick += new System.EventHandler(this.listViewHeadings_DoubleClick);
            // 
            // columnHeaderHeading
            // 
            this.columnHeaderHeading.Text = "标题";
            this.columnHeaderHeading.Width = 400;
            // 
            // columnHeaderLevel
            // 
            this.columnHeaderLevel.Text = "级别";
            this.columnHeaderLevel.Width = 60;
            // 
            // columnHeaderContentLength
            // 
            this.columnHeaderContentLength.Text = "内容长度";
            this.columnHeaderContentLength.Width = 80;
            // 
            // buttonUpload
            // 
            this.buttonUpload.Enabled = false;
            this.buttonUpload.Location = new System.Drawing.Point(512, 600);
            this.buttonUpload.Name = "buttonUpload";
            this.buttonUpload.Size = new System.Drawing.Size(80, 30);
            this.buttonUpload.TabIndex = 3;
            this.buttonUpload.Text = "上传文档";
            this.buttonUpload.UseVisualStyleBackColor = true;
            this.buttonUpload.Click += new System.EventHandler(this.buttonUpload_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(598, 600);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(80, 30);
            this.buttonCancel.TabIndex = 4;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // contextMenuStripHeadings
            // 
            this.contextMenuStripHeadings.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemViewContent,
            this.toolStripMenuItemRemove,
            this.toolStripSeparator1,
            this.toolStripMenuItemSelectAll,
            this.toolStripMenuItemUnselectAll});
            this.contextMenuStripHeadings.Name = "contextMenuStripHeadings";
            this.contextMenuStripHeadings.Size = new System.Drawing.Size(149, 98);
            // 
            // toolStripMenuItemViewContent
            // 
            this.toolStripMenuItemViewContent.Name = "toolStripMenuItemViewContent";
            this.toolStripMenuItemViewContent.Size = new System.Drawing.Size(148, 22);
            this.toolStripMenuItemViewContent.Text = "查看内容";
            this.toolStripMenuItemViewContent.Click += new System.EventHandler(this.toolStripMenuItemViewContent_Click);
            // 
            // toolStripMenuItemRemove
            // 
            this.toolStripMenuItemRemove.Name = "toolStripMenuItemRemove";
            this.toolStripMenuItemRemove.Size = new System.Drawing.Size(148, 22);
            this.toolStripMenuItemRemove.Text = "剔除选中标题";
            this.toolStripMenuItemRemove.Click += new System.EventHandler(this.toolStripMenuItemRemove_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(145, 6);
            // 
            // toolStripMenuItemSelectAll
            // 
            this.toolStripMenuItemSelectAll.Name = "toolStripMenuItemSelectAll";
            this.toolStripMenuItemSelectAll.Size = new System.Drawing.Size(148, 22);
            this.toolStripMenuItemSelectAll.Text = "选择全部";
            this.toolStripMenuItemSelectAll.Click += new System.EventHandler(this.toolStripMenuItemSelectAll_Click);
            // 
            // toolStripMenuItemUnselectAll
            // 
            this.toolStripMenuItemUnselectAll.Name = "toolStripMenuItemUnselectAll";
            this.toolStripMenuItemUnselectAll.Size = new System.Drawing.Size(148, 22);
            this.toolStripMenuItemUnselectAll.Text = "取消全部选择";
            this.toolStripMenuItemUnselectAll.Click += new System.EventHandler(this.toolStripMenuItemUnselectAll_Click);
            // 
            // DocumentUploadForm
            // 
            this.AcceptButton = this.buttonUpload;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(684, 642);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonUpload);
            this.Controls.Add(this.groupBoxHeadings);
            this.Controls.Add(this.groupBoxParse);
            this.Controls.Add(this.groupBoxFileSelect);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DocumentUploadForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "上传文档";
            this.Load += new System.EventHandler(this.DocumentUploadForm_Load);
            this.groupBoxFileSelect.ResumeLayout(false);
            this.groupBoxFileSelect.PerformLayout();
            this.groupBoxParse.ResumeLayout(false);
            this.groupBoxParse.PerformLayout();
            this.groupBoxHeadings.ResumeLayout(false);
            this.contextMenuStripHeadings.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxFileSelect;
        private System.Windows.Forms.Label labelFileSize;
        private System.Windows.Forms.Label labelSelectedFile;
        private System.Windows.Forms.Button buttonSelectFile;
        private System.Windows.Forms.Label labelInstruction;
        private System.Windows.Forms.GroupBox groupBoxParse;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label labelParseResult;
        private System.Windows.Forms.Button buttonParse;
        private System.Windows.Forms.Label labelParseInfo;
        private System.Windows.Forms.GroupBox groupBoxHeadings;
        private System.Windows.Forms.ListView listViewHeadings;
        private System.Windows.Forms.ColumnHeader columnHeaderHeading;
        private System.Windows.Forms.ColumnHeader columnHeaderLevel;
        private System.Windows.Forms.ColumnHeader columnHeaderContentLength;
        private System.Windows.Forms.Button buttonUpload;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.TextBox textBoxParseLog;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripHeadings;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemViewContent;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemRemove;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemSelectAll;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemUnselectAll;
    }
} 