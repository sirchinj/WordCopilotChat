using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordCopilotChat.models;
using WordCopilotChat.services;

namespace WordCopilotChat
{
    public partial class AddModelForm : Form
    {
        private RequestTemplateService _templateService;
        private ModelService _modelService;
        private Model _currentModel;
        private bool _isEditMode;

        public AddModelForm(Model model = null)
        {
            InitializeComponent();
            _templateService = new RequestTemplateService();
            _modelService = new ModelService();
            _currentModel = model;
            _isEditMode = model != null;
            
            LoadTemplates();
            
            if (_isEditMode)
            {
                LoadModelData();
                this.Text = "编辑模型";
            }
            else
            {
                this.Text = "新增模型";
                SetDefaultValues();
            }
        }



        private void LoadTemplates()
        {
            try
            {
                var templates = _templateService.GetAllTemplates();
                cmbTemplate.DisplayMember = "TemplateName";
                cmbTemplate.ValueMember = "Id";
                cmbTemplate.DataSource = templates;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载模板失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadModelData()
        {
            if (_currentModel == null) return;

            txtNickName.Text = _currentModel.NickName;
            txtBaseUrl.Text = _currentModel.BaseUrl;
            txtApiKey.Text = _currentModel.ApiKey;
            txtParameters.Text = _currentModel.Parameters;

            if (_currentModel.TemplateId > 0)
            {
                cmbTemplate.SelectedValue = _currentModel.TemplateId;
            }

            rbChat.Checked = _currentModel.modelType == 1;
            rbEmbedding.Checked = _currentModel.modelType == 2;

            chkMultiModal.Checked = _currentModel.EnableMulti == 1;
            chkTools.Checked = _currentModel.EnableTools == 1;
            chkThink.Checked = _currentModel.EnableThink == 1;
            
            // 加载上下文长度（默认128k）
            numContextLength.Value = _currentModel.ContextLength ?? 128;
        }

        private void SetDefaultValues()
        {
            txtParameters.Text = "{\n  \"model\": \"gpt-3.5-turbo\"\n}";
            rbChat.Checked = true;
            numContextLength.Value = 128; // 默认128k tokens
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ValidateInput())
                return;

            try
            {
                var model = CreateModelFromInput();
                bool success;

                if (_isEditMode)
                {
                    model.Id = _currentModel.Id;
                    success = _modelService.UpdateModel(model);
                }
                else
                {
                    success = _modelService.AddModel(model);
                }

                if (success)
                {
                    MessageBox.Show("保存成功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("保存失败!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtNickName.Text))
            {
                MessageBox.Show("请输入模型名称!", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtNickName.Focus();
                return false;
            }

            if (cmbTemplate.SelectedValue == null)
            {
                MessageBox.Show("请选择请求模板!", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbTemplate.Focus();
                return false;
            }

            return true;
        }

        private Model CreateModelFromInput()
        {
            // 获取上下文长度，如果为0则设置为null（表示不限制）
            int contextLengthValue = (int)numContextLength.Value;
            int? contextLength = contextLengthValue > 0 ? (int?)contextLengthValue : null;
            
            return new Model
            {
                NickName = txtNickName.Text.Trim(),
                BaseUrl = txtBaseUrl.Text.Trim(),
                ApiKey = txtApiKey.Text.Trim(),
                Parameters = txtParameters.Text.Trim(),
                TemplateId = (int)cmbTemplate.SelectedValue,
                modelType = rbChat.Checked ? 1 : 2,
                EnableMulti = chkMultiModal.Checked ? 1 : 0,
                EnableTools = chkTools.Checked ? 1 : 0,
                EnableThink = chkThink.Checked ? 1 : 0,
                ContextLength = contextLength
            };
        }

        private void txtApiKey_MouseDown(object sender, MouseEventArgs e)
        {
            // 按住时显示明文
            txtApiKey.PasswordChar = '\0';
        }

        private void txtApiKey_MouseUp(object sender, MouseEventArgs e)
        {
            // 松开恢复掩码
            txtApiKey.PasswordChar = '*';
        }

        private void txtApiKey_MouseLeave(object sender, EventArgs e)
        {
            // 失去焦点时确保恢复掩码
            txtApiKey.PasswordChar = '*';
        }
    }
} 