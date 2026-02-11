namespace WordCopilotChat
{
    partial class AddModelForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtNickName;
        private System.Windows.Forms.TextBox txtBaseUrl;
        private System.Windows.Forms.TextBox txtApiKey;
        private System.Windows.Forms.ComboBox cmbTemplate;
        private System.Windows.Forms.RadioButton rbChat;
        private System.Windows.Forms.RadioButton rbEmbedding;
        private System.Windows.Forms.CheckBox chkMultiModal;
        private System.Windows.Forms.CheckBox chkTools;
        private System.Windows.Forms.CheckBox chkThink;
        private System.Windows.Forms.TextBox txtParameters;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblNickName;
        private System.Windows.Forms.Label lblBaseUrl;
        private System.Windows.Forms.Label lblApiKey;
        private System.Windows.Forms.Label lblTemplate;
        private System.Windows.Forms.Label lblTemplateDesc;
        private System.Windows.Forms.Label lblModelType;
        private System.Windows.Forms.Label lblModelSupport;
        private System.Windows.Forms.Label lblContextLength;
        private System.Windows.Forms.NumericUpDown numContextLength;
        private System.Windows.Forms.Label lblContextLengthDesc;
        private System.Windows.Forms.Label lblParameters;
        private System.Windows.Forms.Label lblParamDesc;

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
            this.txtNickName = new System.Windows.Forms.TextBox();
            this.txtBaseUrl = new System.Windows.Forms.TextBox();
            this.txtApiKey = new System.Windows.Forms.TextBox();
            this.cmbTemplate = new System.Windows.Forms.ComboBox();
            this.rbChat = new System.Windows.Forms.RadioButton();
            this.rbEmbedding = new System.Windows.Forms.RadioButton();
            this.chkMultiModal = new System.Windows.Forms.CheckBox();
            this.chkTools = new System.Windows.Forms.CheckBox();
            this.chkThink = new System.Windows.Forms.CheckBox();
            this.txtParameters = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblNickName = new System.Windows.Forms.Label();
            this.lblBaseUrl = new System.Windows.Forms.Label();
            this.lblApiKey = new System.Windows.Forms.Label();
            this.lblTemplate = new System.Windows.Forms.Label();
            this.lblTemplateDesc = new System.Windows.Forms.Label();
            this.lblModelType = new System.Windows.Forms.Label();
            this.lblModelSupport = new System.Windows.Forms.Label();
            this.lblContextLength = new System.Windows.Forms.Label();
            this.numContextLength = new System.Windows.Forms.NumericUpDown();
            this.lblContextLengthDesc = new System.Windows.Forms.Label();
            this.lblParameters = new System.Windows.Forms.Label();
            this.lblParamDesc = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numContextLength)).BeginInit();
            this.SuspendLayout();
            // 
            // txtNickName
            // 
            this.txtNickName.Location = new System.Drawing.Point(110, 20);
            this.txtNickName.Name = "txtNickName";
            this.txtNickName.Size = new System.Drawing.Size(350, 21);
            this.txtNickName.TabIndex = 0;
            // 
            // txtBaseUrl
            // 
            this.txtBaseUrl.Location = new System.Drawing.Point(110, 55);
            this.txtBaseUrl.Name = "txtBaseUrl";
            this.txtBaseUrl.Size = new System.Drawing.Size(350, 21);
            this.txtBaseUrl.TabIndex = 1;
            // 
            // txtApiKey
            // 
            this.txtApiKey.Location = new System.Drawing.Point(110, 90);
            this.txtApiKey.Name = "txtApiKey";
            this.txtApiKey.PasswordChar = '*';
            this.txtApiKey.Size = new System.Drawing.Size(350, 21);
            this.txtApiKey.TabIndex = 2;
            this.txtApiKey.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtApiKey_MouseDown);
            this.txtApiKey.MouseLeave += new System.EventHandler(this.txtApiKey_MouseLeave);
            this.txtApiKey.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtApiKey_MouseUp);
            // 
            // cmbTemplate
            // 
            this.cmbTemplate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTemplate.FormattingEnabled = true;
            this.cmbTemplate.Location = new System.Drawing.Point(110, 125);
            this.cmbTemplate.Name = "cmbTemplate";
            this.cmbTemplate.Size = new System.Drawing.Size(350, 20);
            this.cmbTemplate.TabIndex = 3;
            // 
            // rbChat
            // 
            this.rbChat.AutoSize = true;
            this.rbChat.Checked = true;
            this.rbChat.Location = new System.Drawing.Point(110, 180);
            this.rbChat.Name = "rbChat";
            this.rbChat.Size = new System.Drawing.Size(119, 16);
            this.rbChat.TabIndex = 4;
            this.rbChat.TabStop = true;
            this.rbChat.Text = "对话模型（Chat）";
            this.rbChat.UseVisualStyleBackColor = true;
            // 
            // rbEmbedding
            // 
            this.rbEmbedding.AutoSize = true;
            this.rbEmbedding.Enabled = false;
            this.rbEmbedding.Location = new System.Drawing.Point(270, 180);
            this.rbEmbedding.Name = "rbEmbedding";
            this.rbEmbedding.Size = new System.Drawing.Size(161, 16);
            this.rbEmbedding.TabIndex = 5;
            this.rbEmbedding.Text = "词嵌入模型（Embedding）";
            this.rbEmbedding.UseVisualStyleBackColor = true;
            // 
            // chkMultiModal
            // 
            this.chkMultiModal.AutoSize = true;
            this.chkMultiModal.Location = new System.Drawing.Point(110, 215);
            this.chkMultiModal.Name = "chkMultiModal";
            this.chkMultiModal.Size = new System.Drawing.Size(60, 16);
            this.chkMultiModal.TabIndex = 6;
            this.chkMultiModal.Text = "多模态";
            this.chkMultiModal.UseVisualStyleBackColor = true;
            // 
            // chkTools
            // 
            this.chkTools.AutoSize = true;
            this.chkTools.Location = new System.Drawing.Point(200, 215);
            this.chkTools.Name = "chkTools";
            this.chkTools.Size = new System.Drawing.Size(72, 16);
            this.chkTools.TabIndex = 7;
            this.chkTools.Text = "工具调用";
            this.chkTools.UseVisualStyleBackColor = true;
            // 
            // chkThink
            // 
            this.chkThink.AutoSize = true;
            this.chkThink.Enabled = false;
            this.chkThink.Location = new System.Drawing.Point(290, 215);
            this.chkThink.Name = "chkThink";
            this.chkThink.Size = new System.Drawing.Size(72, 16);
            this.chkThink.TabIndex = 8;
            this.chkThink.Text = "深度思考";
            this.chkThink.UseVisualStyleBackColor = true;
            this.chkThink.Visible = false;
            // 
            // txtParameters
            // 
            this.txtParameters.Location = new System.Drawing.Point(110, 280);
            this.txtParameters.Multiline = true;
            this.txtParameters.Name = "txtParameters";
            this.txtParameters.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtParameters.Size = new System.Drawing.Size(350, 120);
            this.txtParameters.TabIndex = 9;
            this.txtParameters.Text = "{\r\n  \"model\": \"gpt-3.5-turbo\"\r\n}";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(200, 450);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(80, 35);
            this.btnSave.TabIndex = 10;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(300, 450);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 35);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // lblNickName
            // 
            this.lblNickName.AutoSize = true;
            this.lblNickName.Location = new System.Drawing.Point(20, 23);
            this.lblNickName.Name = "lblNickName";
            this.lblNickName.Size = new System.Drawing.Size(35, 12);
            this.lblNickName.TabIndex = 12;
            this.lblNickName.Text = "名称:";
            // 
            // lblBaseUrl
            // 
            this.lblBaseUrl.AutoSize = true;
            this.lblBaseUrl.Location = new System.Drawing.Point(20, 58);
            this.lblBaseUrl.Name = "lblBaseUrl";
            this.lblBaseUrl.Size = new System.Drawing.Size(59, 12);
            this.lblBaseUrl.TabIndex = 13;
            this.lblBaseUrl.Text = "接口地址:";
            // 
            // lblApiKey
            // 
            this.lblApiKey.AutoSize = true;
            this.lblApiKey.Location = new System.Drawing.Point(20, 93);
            this.lblApiKey.Name = "lblApiKey";
            this.lblApiKey.Size = new System.Drawing.Size(53, 12);
            this.lblApiKey.TabIndex = 14;
            this.lblApiKey.Text = "API KEY:";
            // 
            // lblTemplate
            // 
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Location = new System.Drawing.Point(20, 128);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new System.Drawing.Size(59, 12);
            this.lblTemplate.TabIndex = 15;
            this.lblTemplate.Text = "请求模板:";
            // 
            // lblTemplateDesc
            // 
            this.lblTemplateDesc.AutoSize = true;
            this.lblTemplateDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblTemplateDesc.Location = new System.Drawing.Point(110, 155);
            this.lblTemplateDesc.Name = "lblTemplateDesc";
            this.lblTemplateDesc.Size = new System.Drawing.Size(377, 12);
            this.lblTemplateDesc.TabIndex = 16;
            this.lblTemplateDesc.Text = "所有主流服务商和推理工具均支持Openai接口规范，所以只定义一个！";
            // 
            // lblModelType
            // 
            this.lblModelType.AutoSize = true;
            this.lblModelType.Location = new System.Drawing.Point(20, 182);
            this.lblModelType.Name = "lblModelType";
            this.lblModelType.Size = new System.Drawing.Size(59, 12);
            this.lblModelType.TabIndex = 17;
            this.lblModelType.Text = "模型类型:";
            // 
            // lblModelSupport
            // 
            this.lblModelSupport.AutoSize = true;
            this.lblModelSupport.Location = new System.Drawing.Point(20, 217);
            this.lblModelSupport.Name = "lblModelSupport";
            this.lblModelSupport.Size = new System.Drawing.Size(59, 12);
            this.lblModelSupport.TabIndex = 18;
            this.lblModelSupport.Text = "模型支持:";
            // 
            // lblContextLength
            // 
            this.lblContextLength.AutoSize = true;
            this.lblContextLength.Location = new System.Drawing.Point(20, 248);
            this.lblContextLength.Name = "lblContextLength";
            this.lblContextLength.Size = new System.Drawing.Size(71, 12);
            this.lblContextLength.TabIndex = 21;
            this.lblContextLength.Text = "上下文长度:";
            // 
            // numContextLength
            // 
            this.numContextLength.Location = new System.Drawing.Point(110, 245);
            this.numContextLength.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numContextLength.Name = "numContextLength";
            this.numContextLength.Size = new System.Drawing.Size(100, 21);
            this.numContextLength.TabIndex = 22;
            this.numContextLength.Value = new decimal(new int[] {
            128,
            0,
            0,
            0});
            // 
            // lblContextLengthDesc
            // 
            this.lblContextLengthDesc.AutoSize = true;
            this.lblContextLengthDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblContextLengthDesc.Location = new System.Drawing.Point(220, 248);
            this.lblContextLengthDesc.Name = "lblContextLengthDesc";
            this.lblContextLengthDesc.Size = new System.Drawing.Size(281, 12);
            this.lblContextLengthDesc.TabIndex = 23;
            this.lblContextLengthDesc.Text = "单位：k（例如128表示128k tokens，0表示不限制）";
            // 
            // lblParameters
            // 
            this.lblParameters.AutoSize = true;
            this.lblParameters.Location = new System.Drawing.Point(20, 283);
            this.lblParameters.Name = "lblParameters";
            this.lblParameters.Size = new System.Drawing.Size(59, 12);
            this.lblParameters.TabIndex = 19;
            this.lblParameters.Text = "请求参数:";
            // 
            // lblParamDesc
            // 
            this.lblParamDesc.AutoSize = true;
            this.lblParamDesc.ForeColor = System.Drawing.Color.Gray;
            this.lblParamDesc.Location = new System.Drawing.Point(110, 410);
            this.lblParamDesc.Name = "lblParamDesc";
            this.lblParamDesc.Size = new System.Drawing.Size(209, 12);
            this.lblParamDesc.TabIndex = 20;
            this.lblParamDesc.Text = "messages(input) 和 stream 不用填写";
            // 
            // AddModelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 510);
            this.Controls.Add(this.lblContextLengthDesc);
            this.Controls.Add(this.numContextLength);
            this.Controls.Add(this.lblContextLength);
            this.Controls.Add(this.lblParamDesc);
            this.Controls.Add(this.lblParameters);
            this.Controls.Add(this.lblModelSupport);
            this.Controls.Add(this.lblModelType);
            this.Controls.Add(this.lblTemplateDesc);
            this.Controls.Add(this.lblTemplate);
            this.Controls.Add(this.lblApiKey);
            this.Controls.Add(this.lblBaseUrl);
            this.Controls.Add(this.lblNickName);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtParameters);
            this.Controls.Add(this.chkThink);
            this.Controls.Add(this.chkTools);
            this.Controls.Add(this.chkMultiModal);
            this.Controls.Add(this.rbEmbedding);
            this.Controls.Add(this.rbChat);
            this.Controls.Add(this.cmbTemplate);
            this.Controls.Add(this.txtApiKey);
            this.Controls.Add(this.txtBaseUrl);
            this.Controls.Add(this.txtNickName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddModelForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "模型编辑";
            ((System.ComponentModel.ISupportInitialize)(this.numContextLength)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
} 