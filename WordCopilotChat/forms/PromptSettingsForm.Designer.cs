namespace WordCopilotChat
{
    partial class PromptSettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageChat;
        private System.Windows.Forms.TabPage tabPageAgent;
        private System.Windows.Forms.TabPage tabPageWelcome;
        private System.Windows.Forms.TabPage tabPageCompress;
        private System.Windows.Forms.TextBox txtChatPrompt;
        private System.Windows.Forms.TextBox txtAgentPrompt;
        private System.Windows.Forms.TextBox txtWelcomePrompt;
        private System.Windows.Forms.TextBox txtCompressPrompt;
        private System.Windows.Forms.Label lblChatPrompt;
        private System.Windows.Forms.Label lblAgentPrompt;
        private System.Windows.Forms.Label lblWelcomePrompt;
        private System.Windows.Forms.Label lblCompressPrompt;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnPreviewChat;
        private System.Windows.Forms.Button btnPreviewAgent;
        private System.Windows.Forms.Button btnPreviewWelcome;
        private System.Windows.Forms.Button btnPreviewCompress;
        private System.Windows.Forms.Label lblChatDesc;
        private System.Windows.Forms.Label lblAgentDesc;
        private System.Windows.Forms.Label lblWelcomeDesc;
        private System.Windows.Forms.Label lblCompressDesc;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;

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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageChat = new System.Windows.Forms.TabPage();
            this.btnPreviewChat = new System.Windows.Forms.Button();
            this.txtChatPrompt = new System.Windows.Forms.TextBox();
            this.lblChatDesc = new System.Windows.Forms.Label();
            this.lblChatPrompt = new System.Windows.Forms.Label();
            this.tabPageAgent = new System.Windows.Forms.TabPage();
            this.btnPreviewAgent = new System.Windows.Forms.Button();
            this.txtAgentPrompt = new System.Windows.Forms.TextBox();
            this.lblAgentDesc = new System.Windows.Forms.Label();
            this.lblAgentPrompt = new System.Windows.Forms.Label();
            this.tabPageWelcome = new System.Windows.Forms.TabPage();
            this.btnPreviewWelcome = new System.Windows.Forms.Button();
            this.txtWelcomePrompt = new System.Windows.Forms.TextBox();
            this.lblWelcomeDesc = new System.Windows.Forms.Label();
            this.lblWelcomePrompt = new System.Windows.Forms.Label();
            this.tabPageCompress = new System.Windows.Forms.TabPage();
            this.btnPreviewCompress = new System.Windows.Forms.Button();
            this.txtCompressPrompt = new System.Windows.Forms.TextBox();
            this.lblCompressDesc = new System.Windows.Forms.Label();
            this.lblCompressPrompt = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPageChat.SuspendLayout();
            this.tabPageAgent.SuspendLayout();
            this.tabPageWelcome.SuspendLayout();
            this.tabPageCompress.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPageChat);
            this.tabControl1.Controls.Add(this.tabPageAgent);
            this.tabControl1.Controls.Add(this.tabPageWelcome);
            this.tabControl1.Controls.Add(this.tabPageCompress);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(760, 500);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPageChat
            // 
            this.tabPageChat.Controls.Add(this.btnPreviewChat);
            this.tabPageChat.Controls.Add(this.txtChatPrompt);
            this.tabPageChat.Controls.Add(this.lblChatDesc);
            this.tabPageChat.Controls.Add(this.lblChatPrompt);
            this.tabPageChat.Location = new System.Drawing.Point(4, 22);
            this.tabPageChat.Name = "tabPageChat";
            this.tabPageChat.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageChat.Size = new System.Drawing.Size(752, 474);
            this.tabPageChat.TabIndex = 0;
            this.tabPageChat.Text = "智能问答模式";
            this.tabPageChat.UseVisualStyleBackColor = true;
            // 
            // btnPreviewChat
            // 
            this.btnPreviewChat.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPreviewChat.Location = new System.Drawing.Point(671, 445);
            this.btnPreviewChat.Name = "btnPreviewChat";
            this.btnPreviewChat.Size = new System.Drawing.Size(75, 23);
            this.btnPreviewChat.TabIndex = 3;
            this.btnPreviewChat.Text = "预览";
            this.btnPreviewChat.UseVisualStyleBackColor = true;
            this.btnPreviewChat.Click += new System.EventHandler(this.BtnPreviewChat_Click);
            // 
            // txtChatPrompt
            // 
            this.txtChatPrompt.AcceptsReturn = true;
            this.txtChatPrompt.AcceptsTab = true;
            this.txtChatPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtChatPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtChatPrompt.Location = new System.Drawing.Point(6, 60);
            this.txtChatPrompt.Multiline = true;
            this.txtChatPrompt.Name = "txtChatPrompt";
            this.txtChatPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtChatPrompt.Size = new System.Drawing.Size(740, 379);
            this.txtChatPrompt.TabIndex = 2;
            this.txtChatPrompt.WordWrap = true;
            this.txtChatPrompt.TextChanged += new System.EventHandler(this.OnTextChanged);
            // 
            // lblChatDesc
            // 
            this.lblChatDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblChatDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblChatDesc.Location = new System.Drawing.Point(6, 25);
            this.lblChatDesc.Name = "lblChatDesc";
            this.lblChatDesc.Size = new System.Drawing.Size(740, 32);
            this.lblChatDesc.TabIndex = 1;
            this.lblChatDesc.Text = "智能问答模式用于普通对话，主要处理文档内容生成、数学公式、表格等基础功能。该模式下AI不会主动调用Word工具。";
            // 
            // lblChatPrompt
            // 
            this.lblChatPrompt.AutoSize = true;
            this.lblChatPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblChatPrompt.Location = new System.Drawing.Point(6, 6);
            this.lblChatPrompt.Name = "lblChatPrompt";
            this.lblChatPrompt.Size = new System.Drawing.Size(92, 17);
            this.lblChatPrompt.TabIndex = 0;
            this.lblChatPrompt.Text = "智能问答模式提示词";
            // 
            // tabPageAgent
            // 
            this.tabPageAgent.Controls.Add(this.btnPreviewAgent);
            this.tabPageAgent.Controls.Add(this.txtAgentPrompt);
            this.tabPageAgent.Controls.Add(this.lblAgentDesc);
            this.tabPageAgent.Controls.Add(this.lblAgentPrompt);
            this.tabPageAgent.Location = new System.Drawing.Point(4, 22);
            this.tabPageAgent.Name = "tabPageAgent";
            this.tabPageAgent.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageAgent.Size = new System.Drawing.Size(752, 474);
            this.tabPageAgent.TabIndex = 1;
            this.tabPageAgent.Text = "智能体模式";
            this.tabPageAgent.UseVisualStyleBackColor = true;
            // 
            // btnPreviewAgent
            // 
            this.btnPreviewAgent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPreviewAgent.Location = new System.Drawing.Point(671, 445);
            this.btnPreviewAgent.Name = "btnPreviewAgent";
            this.btnPreviewAgent.Size = new System.Drawing.Size(75, 23);
            this.btnPreviewAgent.TabIndex = 3;
            this.btnPreviewAgent.Text = "预览";
            this.btnPreviewAgent.UseVisualStyleBackColor = true;
            this.btnPreviewAgent.Click += new System.EventHandler(this.BtnPreviewAgent_Click);
            // 
            // txtAgentPrompt
            // 
            this.txtAgentPrompt.AcceptsReturn = true;
            this.txtAgentPrompt.AcceptsTab = true;
            this.txtAgentPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtAgentPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtAgentPrompt.Location = new System.Drawing.Point(6, 60);
            this.txtAgentPrompt.Multiline = true;
            this.txtAgentPrompt.Name = "txtAgentPrompt";
            this.txtAgentPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtAgentPrompt.Size = new System.Drawing.Size(740, 379);
            this.txtAgentPrompt.TabIndex = 2;
            this.txtAgentPrompt.WordWrap = true;
            this.txtAgentPrompt.TextChanged += new System.EventHandler(this.OnTextChanged);
            // 
            // lblAgentDesc
            // 
            this.lblAgentDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblAgentDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblAgentDesc.Location = new System.Drawing.Point(6, 25);
            this.lblAgentDesc.Name = "lblAgentDesc";
            this.lblAgentDesc.Size = new System.Drawing.Size(740, 32);
            this.lblAgentDesc.TabIndex = 1;
            this.lblAgentDesc.Text = "智能体模式具备智能代理能力，可以主动调用Word工具来完成复杂的文档操作任务，如插入内容、修改样式、获取文档信息等。";
            // 
            // lblAgentPrompt
            // 
            this.lblAgentPrompt.AutoSize = true;
            this.lblAgentPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblAgentPrompt.Location = new System.Drawing.Point(6, 6);
            this.lblAgentPrompt.Name = "lblAgentPrompt";
            this.lblAgentPrompt.Size = new System.Drawing.Size(128, 17);
            this.lblAgentPrompt.TabIndex = 0;
            this.lblAgentPrompt.Text = "智能体模式提示词";
            // 
            // tabPageWelcome
            // 
            this.tabPageWelcome.Controls.Add(this.btnPreviewWelcome);
            this.tabPageWelcome.Controls.Add(this.txtWelcomePrompt);
            this.tabPageWelcome.Controls.Add(this.lblWelcomeDesc);
            this.tabPageWelcome.Controls.Add(this.lblWelcomePrompt);
            this.tabPageWelcome.Location = new System.Drawing.Point(4, 22);
            this.tabPageWelcome.Name = "tabPageWelcome";
            this.tabPageWelcome.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageWelcome.Size = new System.Drawing.Size(752, 474);
            this.tabPageWelcome.TabIndex = 2;
            this.tabPageWelcome.Text = "欢迎页";
            this.tabPageWelcome.UseVisualStyleBackColor = true;
            // 
            // btnPreviewWelcome
            // 
            this.btnPreviewWelcome.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPreviewWelcome.Location = new System.Drawing.Point(671, 445);
            this.btnPreviewWelcome.Name = "btnPreviewWelcome";
            this.btnPreviewWelcome.Size = new System.Drawing.Size(75, 23);
            this.btnPreviewWelcome.TabIndex = 3;
            this.btnPreviewWelcome.Text = "预览";
            this.btnPreviewWelcome.UseVisualStyleBackColor = true;
            this.btnPreviewWelcome.Click += new System.EventHandler(this.BtnPreviewWelcome_Click);
            // 
            // txtWelcomePrompt
            // 
            this.txtWelcomePrompt.AcceptsReturn = true;
            this.txtWelcomePrompt.AcceptsTab = true;
            this.txtWelcomePrompt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtWelcomePrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWelcomePrompt.Location = new System.Drawing.Point(6, 60);
            this.txtWelcomePrompt.Multiline = true;
            this.txtWelcomePrompt.Name = "txtWelcomePrompt";
            this.txtWelcomePrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtWelcomePrompt.Size = new System.Drawing.Size(740, 379);
            this.txtWelcomePrompt.TabIndex = 2;
            this.txtWelcomePrompt.WordWrap = true;
            this.txtWelcomePrompt.TextChanged += new System.EventHandler(this.OnTextChanged);
            // 
            // lblWelcomeDesc
            // 
            this.lblWelcomeDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblWelcomeDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblWelcomeDesc.Location = new System.Drawing.Point(6, 25);
            this.lblWelcomeDesc.Name = "lblWelcomeDesc";
            this.lblWelcomeDesc.Size = new System.Drawing.Size(740, 32);
            this.lblWelcomeDesc.TabIndex = 1;
            this.lblWelcomeDesc.Text = "欢迎页在用户首次打开聊天界面时显示，用于介绍助手功能并提供使用示例。支持Markdown格式，包括数学公式、表格、代码块等。";
            // 
            // lblWelcomePrompt
            // 
            this.lblWelcomePrompt.AutoSize = true;
            this.lblWelcomePrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblWelcomePrompt.Location = new System.Drawing.Point(6, 6);
            this.lblWelcomePrompt.Name = "lblWelcomePrompt";
            this.lblWelcomePrompt.Size = new System.Drawing.Size(68, 17);
            this.lblWelcomePrompt.TabIndex = 0;
            this.lblWelcomePrompt.Text = "欢迎页内容";
            // 
            // tabPageCompress
            // 
            this.tabPageCompress.Controls.Add(this.btnPreviewCompress);
            this.tabPageCompress.Controls.Add(this.txtCompressPrompt);
            this.tabPageCompress.Controls.Add(this.lblCompressDesc);
            this.tabPageCompress.Controls.Add(this.lblCompressPrompt);
            this.tabPageCompress.Location = new System.Drawing.Point(4, 22);
            this.tabPageCompress.Name = "tabPageCompress";
            this.tabPageCompress.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageCompress.Size = new System.Drawing.Size(752, 474);
            this.tabPageCompress.TabIndex = 3;
            this.tabPageCompress.Text = "上下文压缩";
            this.tabPageCompress.UseVisualStyleBackColor = true;
            // 
            // btnPreviewCompress
            // 
            this.btnPreviewCompress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPreviewCompress.Location = new System.Drawing.Point(671, 445);
            this.btnPreviewCompress.Name = "btnPreviewCompress";
            this.btnPreviewCompress.Size = new System.Drawing.Size(75, 23);
            this.btnPreviewCompress.TabIndex = 3;
            this.btnPreviewCompress.Text = "预览";
            this.btnPreviewCompress.UseVisualStyleBackColor = true;
            this.btnPreviewCompress.Click += new System.EventHandler(this.BtnPreviewCompress_Click);
            // 
            // txtCompressPrompt
            // 
            this.txtCompressPrompt.AcceptsReturn = true;
            this.txtCompressPrompt.AcceptsTab = true;
            this.txtCompressPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCompressPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompressPrompt.Location = new System.Drawing.Point(6, 60);
            this.txtCompressPrompt.Multiline = true;
            this.txtCompressPrompt.Name = "txtCompressPrompt";
            this.txtCompressPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtCompressPrompt.Size = new System.Drawing.Size(740, 379);
            this.txtCompressPrompt.TabIndex = 2;
            this.txtCompressPrompt.WordWrap = true;
            this.txtCompressPrompt.TextChanged += new System.EventHandler(this.OnTextChanged);
            // 
            // lblCompressDesc
            // 
            this.lblCompressDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblCompressDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblCompressDesc.Location = new System.Drawing.Point(6, 25);
            this.lblCompressDesc.Name = "lblCompressDesc";
            this.lblCompressDesc.Size = new System.Drawing.Size(740, 32);
            this.lblCompressDesc.TabIndex = 1;
            this.lblCompressDesc.Text = "当对话历史Token占用过高时，系统会自动调用此提示词对历史对话进行压缩总结，以节省Token并保持上下文连贯性。";
            // 
            // lblCompressPrompt
            // 
            this.lblCompressPrompt.AutoSize = true;
            this.lblCompressPrompt.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCompressPrompt.Location = new System.Drawing.Point(6, 6);
            this.lblCompressPrompt.Name = "lblCompressPrompt";
            this.lblCompressPrompt.Size = new System.Drawing.Size(116, 17);
            this.lblCompressPrompt.TabIndex = 0;
            this.lblCompressPrompt.Text = "上下文压缩提示词";
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(537, 530);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 30);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnReset
            // 
            this.btnReset.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnReset.Location = new System.Drawing.Point(618, 530);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 30);
            this.btnReset.TabIndex = 2;
            this.btnReset.Text = "重置";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.BtnReset_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(699, 530);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Location = new System.Drawing.Point(456, 530);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 30);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "确定";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Location = new System.Drawing.Point(375, 530);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 30);
            this.btnExport.TabIndex = 5;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnImport.Location = new System.Drawing.Point(294, 530);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 30);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "导入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // PromptSettingsForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(784, 572);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(600, 500);
            this.Name = "PromptSettingsForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "提示词设置";
            this.tabControl1.ResumeLayout(false);
            this.tabPageChat.ResumeLayout(false);
            this.tabPageChat.PerformLayout();
            this.tabPageAgent.ResumeLayout(false);
            this.tabPageAgent.PerformLayout();
            this.tabPageWelcome.ResumeLayout(false);
            this.tabPageWelcome.PerformLayout();
            this.tabPageCompress.ResumeLayout(false);
            this.tabPageCompress.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
    }
} 