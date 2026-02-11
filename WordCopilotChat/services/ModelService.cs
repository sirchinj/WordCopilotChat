using System;
using System.Collections.Generic;
using System.Linq;
using WordCopilotChat.db;
using WordCopilotChat.models;

namespace WordCopilotChat.services
{
    public class ModelService
    {
        private readonly IFreeSql _freeSql;

        public ModelService()
        {
            _freeSql = FreeSqlDB.Sqlite;
        }

        /// <summary>
        /// 初始化Model表结构
        /// </summary>
        public void InitializeModelTable()
        {
            try
            {
                // 确保表存在
                _freeSql.CodeFirst.SyncStructure<Model>();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化Model表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取所有模型（带模板信息）
        /// </summary>
        public List<Model> GetAllModels()
        {
            return _freeSql.Select<Model, RequestTemplate>()
                .LeftJoin((m, t) => m.TemplateId == t.Id)
                .ToList((m, t) => new Model
                {
                    Id = m.Id,
                    NickName = m.NickName,
                    TemplateId = m.TemplateId,
                    Template = t,
                    ApiKey = m.ApiKey,
                    BaseUrl = m.BaseUrl,
                    Parameters = m.Parameters,
                    EnableMulti = m.EnableMulti,
                    EnableTools = m.EnableTools,
                    EnableThink = m.EnableThink,
                    modelType = m.modelType,
                    ContextLength = m.ContextLength
                });
        }

        /// <summary>
        /// 根据ID获取模型
        /// </summary>
        public Model GetModelById(int id)
        {
            return _freeSql.Select<Model, RequestTemplate>()
                .LeftJoin((m, t) => m.TemplateId == t.Id)
                .Where((m, t) => m.Id == id)
                .ToOne((m, t) => new Model
                {
                    Id = m.Id,
                    NickName = m.NickName,
                    TemplateId = m.TemplateId,
                    Template = t,
                    ApiKey = m.ApiKey,
                    BaseUrl = m.BaseUrl,
                    Parameters = m.Parameters,
                    EnableMulti = m.EnableMulti,
                    EnableTools = m.EnableTools,
                    EnableThink = m.EnableThink,
                    modelType = m.modelType,
                    ContextLength = m.ContextLength
                });
        }

        /// <summary>
        /// 添加新模型
        /// </summary>
        public bool AddModel(Model model)
        {
            try
            {
                if (model == null || string.IsNullOrWhiteSpace(model.NickName))
                    return false;

                return _freeSql.Insert(model).ExecuteAffrows() > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"添加模型失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 更新模型
        /// </summary>
        public bool UpdateModel(Model model)
        {
            try
            {
                if (model == null || model.Id <= 0 || string.IsNullOrWhiteSpace(model.NickName))
                    return false;

                return _freeSql.Update<Model>()
                    .Set(m => m.NickName, model.NickName)
                    .Set(m => m.TemplateId, model.TemplateId)
                    .Set(m => m.ApiKey, model.ApiKey)
                    .Set(m => m.BaseUrl, model.BaseUrl)
                    .Set(m => m.Parameters, model.Parameters)
                    .Set(m => m.EnableMulti, model.EnableMulti)
                    .Set(m => m.EnableTools, model.EnableTools)
                    .Set(m => m.EnableThink, model.EnableThink)
                    .Set(m => m.modelType, model.modelType)
                    .Set(m => m.ContextLength, model.ContextLength)
                    .Where(m => m.Id == model.Id)
                    .ExecuteAffrows() > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新模型失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 删除模型
        /// </summary>
        public bool DeleteModel(int id)
        {
            try
            {
                return _freeSql.Delete<Model>().Where(x => x.Id == id).ExecuteAffrows() > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"删除模型失败: {ex.Message}");
                return false;
            }
        }
    }
} 