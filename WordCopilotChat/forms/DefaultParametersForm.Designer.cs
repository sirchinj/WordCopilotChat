namespace WordCopilotChat
{
    partial class DefaultParametersForm
    {
        private System.ComponentModel.IContainer components = null;
        
        // 控件声明
        private System.Windows.Forms.GroupBox grpDefault;
        private System.Windows.Forms.Label lblDefaultTemp;
        private System.Windows.Forms.NumericUpDown nudDefaultTemp;
        private System.Windows.Forms.Label lblDefaultMaxTokens;
        private System.Windows.Forms.NumericUpDown nudDefaultMaxTokens;
        private System.Windows.Forms.Label lblDefaultTopP;
        private System.Windows.Forms.NumericUpDown nudDefaultTopP;
        
        private System.Windows.Forms.GroupBox grpChat;
        private System.Windows.Forms.Label lblChatTemp;
        private System.Windows.Forms.NumericUpDown nudChatTemp;
        private System.Windows.Forms.Label lblChatMaxTokens;
        private System.Windows.Forms.NumericUpDown nudChatMaxTokens;
        private System.Windows.Forms.Label lblChatTopP;
        private System.Windows.Forms.NumericUpDown nudChatTopP;
        
        private System.Windows.Forms.GroupBox grpAgent;
        private System.Windows.Forms.Label lblAgentTemp;
        private System.Windows.Forms.NumericUpDown nudAgentTemp;
        private System.Windows.Forms.Label lblAgentMaxTokens;
        private System.Windows.Forms.NumericUpDown nudAgentMaxTokens;
        private System.Windows.Forms.Label lblAgentTopP;
        private System.Windows.Forms.NumericUpDown nudAgentTopP;
        
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnReset;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.grpDefault = new System.Windows.Forms.GroupBox();
            this.lblDefaultTemp = new System.Windows.Forms.Label();
            this.nudDefaultTemp = new System.Windows.Forms.NumericUpDown();
            this.lblDefaultMaxTokens = new System.Windows.Forms.Label();
            this.nudDefaultMaxTokens = new System.Windows.Forms.NumericUpDown();
            this.lblDefaultTopP = new System.Windows.Forms.Label();
            this.nudDefaultTopP = new System.Windows.Forms.NumericUpDown();
            this.grpChat = new System.Windows.Forms.GroupBox();
            this.lblChatTemp = new System.Windows.Forms.Label();
            this.nudChatTemp = new System.Windows.Forms.NumericUpDown();
            this.lblChatMaxTokens = new System.Windows.Forms.Label();
            this.nudChatMaxTokens = new System.Windows.Forms.NumericUpDown();
            this.lblChatTopP = new System.Windows.Forms.Label();
            this.nudChatTopP = new System.Windows.Forms.NumericUpDown();
            this.grpAgent = new System.Windows.Forms.GroupBox();
            this.lblAgentTemp = new System.Windows.Forms.Label();
            this.nudAgentTemp = new System.Windows.Forms.NumericUpDown();
            this.lblAgentMaxTokens = new System.Windows.Forms.Label();
            this.nudAgentMaxTokens = new System.Windows.Forms.NumericUpDown();
            this.lblAgentTopP = new System.Windows.Forms.Label();
            this.nudAgentTopP = new System.Windows.Forms.NumericUpDown();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.grpDefault.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultTemp)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultMaxTokens)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultTopP)).BeginInit();
            this.grpChat.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudChatTemp)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudChatMaxTokens)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudChatTopP)).BeginInit();
            this.grpAgent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentTemp)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentMaxTokens)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentTopP)).BeginInit();
            this.SuspendLayout();
            // 
            // grpDefault
            // 
            this.grpDefault.Controls.Add(this.lblDefaultTemp);
            this.grpDefault.Controls.Add(this.nudDefaultTemp);
            this.grpDefault.Controls.Add(this.lblDefaultMaxTokens);
            this.grpDefault.Controls.Add(this.nudDefaultMaxTokens);
            this.grpDefault.Controls.Add(this.lblDefaultTopP);
            this.grpDefault.Controls.Add(this.nudDefaultTopP);
            this.grpDefault.Location = new System.Drawing.Point(12, 12);
            this.grpDefault.Name = "grpDefault";
            this.grpDefault.Size = new System.Drawing.Size(440, 120);
            this.grpDefault.TabIndex = 0;
            this.grpDefault.TabStop = false;
            this.grpDefault.Text = "通用默认参数";
            // 
            // lblDefaultTemp
            // 
            this.lblDefaultTemp.Location = new System.Drawing.Point(15, 25);
            this.lblDefaultTemp.Name = "lblDefaultTemp";
            this.lblDefaultTemp.Size = new System.Drawing.Size(150, 20);
            this.lblDefaultTemp.TabIndex = 0;
            this.lblDefaultTemp.Text = "Temperature (0.0-2.0):";
            this.lblDefaultTemp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudDefaultTemp
            // 
            this.nudDefaultTemp.DecimalPlaces = 1;
            this.nudDefaultTemp.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudDefaultTemp.Location = new System.Drawing.Point(170, 23);
            this.nudDefaultTemp.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nudDefaultTemp.Name = "nudDefaultTemp";
            this.nudDefaultTemp.Size = new System.Drawing.Size(80, 21);
            this.nudDefaultTemp.TabIndex = 1;
            // 
            // lblDefaultMaxTokens
            // 
            this.lblDefaultMaxTokens.Location = new System.Drawing.Point(15, 55);
            this.lblDefaultMaxTokens.Name = "lblDefaultMaxTokens";
            this.lblDefaultMaxTokens.Size = new System.Drawing.Size(150, 20);
            this.lblDefaultMaxTokens.TabIndex = 2;
            this.lblDefaultMaxTokens.Text = "Max Tokens (最小1):";
            this.lblDefaultMaxTokens.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudDefaultMaxTokens
            // 
            this.nudDefaultMaxTokens.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nudDefaultMaxTokens.Location = new System.Drawing.Point(170, 53);
            this.nudDefaultMaxTokens.Maximum = new decimal(new int[] {
            2147483647,
            0,
            0,
            0});
            this.nudDefaultMaxTokens.Name = "nudDefaultMaxTokens";
            this.nudDefaultMaxTokens.Size = new System.Drawing.Size(120, 21);
            this.nudDefaultMaxTokens.TabIndex = 3;
            this.nudDefaultMaxTokens.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblDefaultTopP
            // 
            this.lblDefaultTopP.Location = new System.Drawing.Point(15, 85);
            this.lblDefaultTopP.Name = "lblDefaultTopP";
            this.lblDefaultTopP.Size = new System.Drawing.Size(150, 20);
            this.lblDefaultTopP.TabIndex = 4;
            this.lblDefaultTopP.Text = "Top-P (0.0-1.0):";
            this.lblDefaultTopP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudDefaultTopP
            // 
            this.nudDefaultTopP.DecimalPlaces = 1;
            this.nudDefaultTopP.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudDefaultTopP.Location = new System.Drawing.Point(170, 83);
            this.nudDefaultTopP.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudDefaultTopP.Name = "nudDefaultTopP";
            this.nudDefaultTopP.Size = new System.Drawing.Size(80, 21);
            this.nudDefaultTopP.TabIndex = 5;
            // 
            // grpChat
            // 
            this.grpChat.Controls.Add(this.lblChatTemp);
            this.grpChat.Controls.Add(this.nudChatTemp);
            this.grpChat.Controls.Add(this.lblChatMaxTokens);
            this.grpChat.Controls.Add(this.nudChatMaxTokens);
            this.grpChat.Controls.Add(this.lblChatTopP);
            this.grpChat.Controls.Add(this.nudChatTopP);
            this.grpChat.Location = new System.Drawing.Point(12, 142);
            this.grpChat.Name = "grpChat";
            this.grpChat.Size = new System.Drawing.Size(440, 120);
            this.grpChat.TabIndex = 1;
            this.grpChat.TabStop = false;
            this.grpChat.Text = "Chat模式参数";
            // 
            // lblChatTemp
            // 
            this.lblChatTemp.Location = new System.Drawing.Point(15, 25);
            this.lblChatTemp.Name = "lblChatTemp";
            this.lblChatTemp.Size = new System.Drawing.Size(150, 20);
            this.lblChatTemp.TabIndex = 0;
            this.lblChatTemp.Text = "Temperature (0.0-2.0):";
            this.lblChatTemp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudChatTemp
            // 
            this.nudChatTemp.DecimalPlaces = 1;
            this.nudChatTemp.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudChatTemp.Location = new System.Drawing.Point(170, 23);
            this.nudChatTemp.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nudChatTemp.Name = "nudChatTemp";
            this.nudChatTemp.Size = new System.Drawing.Size(80, 21);
            this.nudChatTemp.TabIndex = 1;
            // 
            // lblChatMaxTokens
            // 
            this.lblChatMaxTokens.Location = new System.Drawing.Point(15, 55);
            this.lblChatMaxTokens.Name = "lblChatMaxTokens";
            this.lblChatMaxTokens.Size = new System.Drawing.Size(150, 20);
            this.lblChatMaxTokens.TabIndex = 2;
            this.lblChatMaxTokens.Text = "Max Tokens (最小1):";
            this.lblChatMaxTokens.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudChatMaxTokens
            // 
            this.nudChatMaxTokens.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nudChatMaxTokens.Location = new System.Drawing.Point(170, 53);
            this.nudChatMaxTokens.Maximum = new decimal(new int[] {
            2147483647,
            0,
            0,
            0});
            this.nudChatMaxTokens.Name = "nudChatMaxTokens";
            this.nudChatMaxTokens.Size = new System.Drawing.Size(120, 21);
            this.nudChatMaxTokens.TabIndex = 3;
            this.nudChatMaxTokens.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblChatTopP
            // 
            this.lblChatTopP.Location = new System.Drawing.Point(15, 85);
            this.lblChatTopP.Name = "lblChatTopP";
            this.lblChatTopP.Size = new System.Drawing.Size(150, 20);
            this.lblChatTopP.TabIndex = 4;
            this.lblChatTopP.Text = "Top-P (0.0-1.0):";
            this.lblChatTopP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudChatTopP
            // 
            this.nudChatTopP.DecimalPlaces = 1;
            this.nudChatTopP.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudChatTopP.Location = new System.Drawing.Point(170, 83);
            this.nudChatTopP.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudChatTopP.Name = "nudChatTopP";
            this.nudChatTopP.Size = new System.Drawing.Size(80, 21);
            this.nudChatTopP.TabIndex = 5;
            // 
            // grpAgent
            // 
            this.grpAgent.Controls.Add(this.lblAgentTemp);
            this.grpAgent.Controls.Add(this.nudAgentTemp);
            this.grpAgent.Controls.Add(this.lblAgentMaxTokens);
            this.grpAgent.Controls.Add(this.nudAgentMaxTokens);
            this.grpAgent.Controls.Add(this.lblAgentTopP);
            this.grpAgent.Controls.Add(this.nudAgentTopP);
            this.grpAgent.Location = new System.Drawing.Point(12, 272);
            this.grpAgent.Name = "grpAgent";
            this.grpAgent.Size = new System.Drawing.Size(440, 120);
            this.grpAgent.TabIndex = 2;
            this.grpAgent.TabStop = false;
            this.grpAgent.Text = "Agent模式参数";
            // 
            // lblAgentTemp
            // 
            this.lblAgentTemp.Location = new System.Drawing.Point(15, 25);
            this.lblAgentTemp.Name = "lblAgentTemp";
            this.lblAgentTemp.Size = new System.Drawing.Size(150, 20);
            this.lblAgentTemp.TabIndex = 0;
            this.lblAgentTemp.Text = "Temperature (0.0-2.0):";
            this.lblAgentTemp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudAgentTemp
            // 
            this.nudAgentTemp.DecimalPlaces = 1;
            this.nudAgentTemp.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudAgentTemp.Location = new System.Drawing.Point(170, 23);
            this.nudAgentTemp.Maximum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nudAgentTemp.Name = "nudAgentTemp";
            this.nudAgentTemp.Size = new System.Drawing.Size(80, 21);
            this.nudAgentTemp.TabIndex = 1;
            // 
            // lblAgentMaxTokens
            // 
            this.lblAgentMaxTokens.Location = new System.Drawing.Point(15, 55);
            this.lblAgentMaxTokens.Name = "lblAgentMaxTokens";
            this.lblAgentMaxTokens.Size = new System.Drawing.Size(150, 20);
            this.lblAgentMaxTokens.TabIndex = 2;
            this.lblAgentMaxTokens.Text = "Max Tokens (最小1):";
            this.lblAgentMaxTokens.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudAgentMaxTokens
            // 
            this.nudAgentMaxTokens.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nudAgentMaxTokens.Location = new System.Drawing.Point(170, 53);
            this.nudAgentMaxTokens.Maximum = new decimal(new int[] {
            2147483647,
            0,
            0,
            0});
            this.nudAgentMaxTokens.Name = "nudAgentMaxTokens";
            this.nudAgentMaxTokens.Size = new System.Drawing.Size(120, 21);
            this.nudAgentMaxTokens.TabIndex = 3;
            this.nudAgentMaxTokens.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblAgentTopP
            // 
            this.lblAgentTopP.Location = new System.Drawing.Point(15, 85);
            this.lblAgentTopP.Name = "lblAgentTopP";
            this.lblAgentTopP.Size = new System.Drawing.Size(150, 20);
            this.lblAgentTopP.TabIndex = 4;
            this.lblAgentTopP.Text = "Top-P (0.0-1.0):";
            this.lblAgentTopP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudAgentTopP
            // 
            this.nudAgentTopP.DecimalPlaces = 1;
            this.nudAgentTopP.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudAgentTopP.Location = new System.Drawing.Point(170, 83);
            this.nudAgentTopP.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudAgentTopP.Name = "nudAgentTopP";
            this.nudAgentTopP.Size = new System.Drawing.Size(80, 21);
            this.nudAgentTopP.TabIndex = 5;
            // 
            // btnSave
            // 
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(217, 410);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(298, 410);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(12, 410);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(80, 25);
            this.btnReset.TabIndex = 5;
            this.btnReset.Text = "重置默认";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.BtnReset_Click);
            // 
            // DefaultParametersForm
            // 
            this.ClientSize = new System.Drawing.Size(480, 520);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.grpAgent);
            this.Controls.Add(this.grpChat);
            this.Controls.Add(this.grpDefault);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DefaultParametersForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "默认AI参数设置";
            this.grpDefault.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultTemp)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultMaxTokens)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDefaultTopP)).EndInit();
            this.grpChat.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudChatTemp)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudChatMaxTokens)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudChatTopP)).EndInit();
            this.grpAgent.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentTemp)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentMaxTokens)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAgentTopP)).EndInit();
            this.ResumeLayout(false);

        }
    }
}

