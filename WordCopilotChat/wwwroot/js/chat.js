// WebView2与C#通信相关变量和方法
let hostObject = null;

// 全局变量
let isGenerating = false;
let currentContent = '';
let messageIdCounter = 0;
let isAtBottom = true;
let currentChatMode = 'chat'; // 当前聊天模式
let selectedModelId = 0; // 当前选中的模型ID
let availableModels = []; // 可用的模型列表

// 上下文相关变量
let selectedContexts = []; // 已选择的上下文列表
let contextBar = null; // 上下文显示栏
let contextItems = null; // 上下文项目容器
let contextSelector = null; // 上下文选择器
let contextSelectorContent = null; // 上下文选择器内容
let availableDocuments = []; // 可用的文档列表
let currentDocumentHeadings = []; // 当前文档的标题列表

// 图片相关变量
let selectedImages = []; // 已选择的图片列表 [{id, name, base64, size}]
let imageBar = null; // 图片显示栏
let imageItems = null; // 图片项目容器
let imageIdCounter = 0; // 图片ID计数器

// 预览操作相关变量
let currentPreviewedAction = null;
let currentMessageId = null; // 当前正在生成的消息ID
let isPreviewPending = false; // 是否有待处理的预览
let toolProgressHostMessage = null; // 工具进度与预览共享的宿主消息
// 预览决策日志（仅记录支持预览的工具：formatted_insert_content / modify_text_style）
let previewDecisionLogs = []; // { previewId, actionType, toolName, decision: 'accepted'|'rejected', timestamp }

function mapActionTypeToToolName(actionType) {
    if (actionType === 'insert_content') return 'formatted_insert_content';
    if (actionType === 'modify_style') return 'modify_text_style';
    return actionType || '';
}

function recordPreviewDecision(actionType, decision, previewId) {
    const toolName = mapActionTypeToToolName(actionType);
    if (!toolName) return;
    const now = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    // 去重：以 previewId 为键更新
    const idx = previewDecisionLogs.findIndex(x => x.previewId === previewId);
    const entry = { previewId, actionType, toolName, decision, timestamp: now };
    if (idx >= 0) {
        previewDecisionLogs[idx] = entry;
    } else {
        previewDecisionLogs.push(entry);
    }
    console.log('记录预览决策:', entry);
}

// 近况的操作决策状态（跨轮对话使用，用于前端工具白名单收窄）
let operationDecisionState = {
    // toolName: 'accepted' | 'rejected'
};

// 同步决策状态
function updateOperationDecisionState(toolName, decision) {
    if (!toolName) return;
    operationDecisionState[toolName] = decision;
}

// 改写 recordPreviewDecision -> 同步状态
const _recordPreviewDecision_orig = recordPreviewDecision;
recordPreviewDecision = function (actionType, decision, previewId) {
    const toolName = mapActionTypeToToolName(actionType);
    _recordPreviewDecision_orig(actionType, decision, previewId);
    if (toolName) {
        updateOperationDecisionState(toolName, decision);
    }
};



// 现代化预览管理器
class PreviewManager {
    constructor() {
        this.previews = new Map(); // 存储所有预览
        this.counter = 0;
    }

    // 创建新预览
    createPreview(data, actionData) {
        const id = `preview_${++this.counter}`;
        const preview = {
            id: id,
            data: data,
            actionData: actionData,
            status: 'pending', // pending, accepted, rejected, applied
            timestamp: Date.now(),
            element: null
        };
        this.previews.set(id, preview);
        return preview;
    }

    // 更新预览状态
    updateStatus(id, status, message = null) {
        const preview = this.previews.get(id);
        if (preview) {
            preview.status = status;
            if (message) preview.message = message;
            this.updatePreviewUI(preview);

            // 状态更新后，检查是否需要更新批量操作按钮
            setTimeout(() => {
                showFloatingBatchActions();
            }, 100);
        }
    }

    // 获取待处理的预览
    getPendingPreviews() {
        return Array.from(this.previews.values()).filter(p => p.status === 'pending');
    }

    // 批量接受所有待处理预览（顺序处理，避免数据丢失）
    async acceptAll() {
        const pending = this.getPendingPreviews();
        if (pending.length === 0) return;

        console.log(`🔄 开始批量接受 ${pending.length} 个预览（顺序处理）`);

        for (let i = 0; i < pending.length; i++) {
            const preview = pending[i];
            console.log(`📝 处理第 ${i + 1}/${pending.length} 个预览: ${preview.id}`);

            // 记录用户决策为“接受”
            try {
                recordPreviewDecision(preview.actionData.action_type, 'accepted', preview.id);
            } catch (e) {
                console.warn('记录批量接受决策失败:', e);
            }

            this.updateStatus(preview.id, 'applying');

            // 发送应用请求
            this.applyPreview(preview);

            // 添加延迟，避免并发冲突（除了最后一个）
            if (i < pending.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 500)); // 500ms延迟
                console.log(`⏳ 等待500ms后处理下一个预览...`);
            }
        }

        console.log(`✅ 批量接受完成，共处理 ${pending.length} 个预览`);
    }

    // 批量拒绝所有待处理预览
    rejectAll() {
        const pending = this.getPendingPreviews();
        if (pending.length === 0) return;

        console.log(`批量拒绝 ${pending.length} 个预览`);
        pending.forEach(preview => {
            // 记录用户决策为“拒绝”
            try {
                recordPreviewDecision(preview.actionData.action_type, 'rejected', preview.id);
            } catch (e) {
                console.warn('记录批量拒绝决策失败:', e);
            }
            this.updateStatus(preview.id, 'rejected', '已拒绝');
        });

        // 清除预览待处理状态
        isPreviewPending = false;
    }

    // 从DOM收集插入设置（避免事件未触发导致的不同步）
    collectInsertSettings(previewElement, actionData) {
        try {
            if (!actionData) return actionData;
            actionData.parameters = actionData.parameters || {};
            if (previewElement) {
                const trimEl = previewElement.querySelector('input.insert-setting[data-setting-checkbox="trim_spaces"], input.ms-setting[data-setting-checkbox="trim_spaces"]');
                if (trimEl) actionData.parameters.trim_spaces = !!trimEl.checked;
                const posEl = previewElement.querySelector('select.insert-setting[data-setting-select="insert_position"], select.ms-setting[data-setting-select="insert_position"]');
                if (posEl && posEl.value) actionData.parameters.insert_position = posEl.value;
            }
            return actionData;
        } catch (_) {
            return actionData;
        }
    }

    // 应用预览（增强日志）
    applyPreview(preview) {
        // 从 DOM 元素读取最新的 actionData（包含用户在设置中修改的参数）
        let finalActionData = preview.actionData;
        if (preview.element) {
            const dataStr = preview.element.getAttribute('data-action-data');
            if (dataStr) {
                try {
                    finalActionData = JSON.parse(dataStr);
                    console.log('✅ 已从 DOM 读取最新的操作参数');
                } catch (e) {
                    console.warn('解析 data-action-data 失败，使用原始参数:', e);
                }
            }
            // 再次从DOM控件收集插入设置，确保与UI一致
            finalActionData = this.collectInsertSettings(preview.element, finalActionData);
        }

        console.log(`📤 发送预览应用请求: ${preview.id}`, {
            action_type: finalActionData.action_type,
            parameters: finalActionData.parameters
        });

        const messageData = {
            type: 'applyPreviewedAction',
            action_type: finalActionData.action_type,
            parameters: finalActionData.parameters,
            preview_id: preview.id // 添加预览ID用于追踪
        };

        sendMessageToCSharp(messageData);

        console.log(`✅ 预览应用请求已发送: ${preview.id}`);
    }

    // 更新预览UI
    updatePreviewUI(preview) {
        if (!preview.element) return;

        const element = preview.element;
        const header = element.querySelector('.tool-preview-title');
        const actions = element.querySelector('.tool-preview-actions');
        const statusIndicator = element.querySelector('.preview-status');

        // 更新状态指示器
        if (statusIndicator) {
            statusIndicator.remove();
        }

        let statusBadge = '';
        let headerIcon = '';
        let statusColor = '';

        switch (preview.status) {
            case 'pending':
                headerIcon = '⏳';
                statusBadge = '<span class="preview-status pending">待处理</span>';
                statusColor = 'border-yellow-200 bg-yellow-50';
                break;
            case 'applying':
                headerIcon = '🔄';
                statusBadge = '<span class="preview-status applying">应用中...</span>';
                statusColor = 'border-blue-200 bg-blue-50';
                break;
            case 'applied':
                headerIcon = '✅';
                statusBadge = '<span class="preview-status applied">已应用</span>';
                statusColor = 'border-green-200 bg-green-50';
                break;
            case 'rejected':
                headerIcon = '❌';
                statusBadge = '<span class="preview-status rejected">已拒绝</span>';
                statusColor = 'border-red-200 bg-red-50';
                break;
        }

        // 更新头部
        if (header) {
            const title = header.querySelector('span:last-child');
            if (title) {
                const iconSpan = header.querySelector('.icon');
                if (iconSpan) iconSpan.textContent = headerIcon;
                title.innerHTML += statusBadge;
            }
        }

        // 更新元素样式
        element.className = element.className.replace(/border-\w+-200|bg-\w+-50|pending|applying|applied|rejected/g, '');
        element.classList.add(preview.status);

        // 更新操作按钮
        if (actions) {
            if (preview.status === 'pending') {
                // 保持操作按钮
            } else if (preview.status === 'applied') {
                // 已应用：显示小字状态，不再显示按钮
                actions.innerHTML = `
                    <div class="preview-result applied">
                        <span class="icon">✅</span>
                        <span>已接受</span>
                    </div>
                `;
            } else if (preview.status === 'rejected') {
                // 已拒绝：显示小字状态，不再显示按钮
                actions.innerHTML = `
                    <div class="preview-result rejected">
                        <span class="icon">❌</span>
                        <span>已拒绝</span>
                    </div>
                `;
            } else {
                // 其他状态（如 applying）：显示状态信息
                actions.innerHTML = `
                    <div class="preview-result ${preview.status}">
                        <span class="icon">${headerIcon}</span>
                        <span>${preview.message || this.getStatusText(preview.status)}</span>
                    </div>
                `;
            }
        }
    }

    getStatusText(status) {
        switch (status) {
            case 'applied': return '操作已成功应用';
            case 'rejected': return '操作已被拒绝';
            case 'applying': return '正在应用操作...';
            default: return '等待处理';
        }
    }

    // 移除预览
    removePreview(id) {
        const preview = this.previews.get(id);
        if (preview) {
            console.log(`移除预览: ${id}`);
            this.previews.delete(id);

            // 移除DOM元素
            if (preview.element) {
                preview.element.remove();
            }

            return true;
        }
        return false;
    }
}

// 全局预览管理器实例
const previewManager = new PreviewManager();
let enabledTools = {}; // 启用的工具列表
let defaultTools = [
    'check_insert_position',
    'get_selected_text',
    'formatted_insert_content',
    'modify_text_style',
    'get_document_statistics',
    'get_document_images',
    'get_document_formulas',
    'get_document_tables',
    'get_document_headings',
    'get_heading_content'
]; // 默认工具列表

// DOM 元素引用
const chatMessages = document.getElementById('chat-messages');
const messageInput = document.getElementById('message-input');
const messageHighlighter = document.getElementById('message-highlighter');
const sendBtn = document.getElementById('send-button');
const typingIndicator = document.getElementById('typing-indicator');
const charCount = document.getElementById('char-count');
const chatModeSelect = document.getElementById('chat-mode');
const modelSelect = document.getElementById('model-select');
const toolsSettingsBtn = document.getElementById('tools-settings-btn');
const toolsSettingsModal = document.getElementById('tools-settings-modal');
const quickSelector = document.getElementById('quick-selector');
const quickSelectorContent = document.getElementById('quick-selector-content');
const previewActionPanel = document.getElementById('preview-action-panel');
const previewBody = document.getElementById('preview-body');

// 初始化
document.addEventListener('DOMContentLoaded', function () {
    initializeChat();
    setupEventListeners();
    autoResizeInput();
    initializeContextMenu();
    initializeContextElements();
    initializeImageElements();

    // 初始化智能体模式锁定状态
    initializeAgentModeLock();

    // 添加键盘快捷键
    document.addEventListener('keydown', function (e) {
        // ESC键停止生成
        if (e.key === 'Escape' && isGenerating) {
            e.preventDefault();
            stopGeneration();
        }

        // Ctrl+Enter 发送消息（原有功能保持）
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && messageInput) {
            e.preventDefault();
            sendMessage();
        }
    });

    // 等待MathJax加载完成后处理欢迎消息
    if (typeof MathJax !== 'undefined') {
        MathJax.Hub.Queue(function () {
            processWelcomeMessage();
        });
    } else {
        // 如果MathJax还没加载，延迟处理
        setTimeout(function () {
            if (typeof MathJax !== 'undefined') {
                MathJax.Hub.Queue(function () {
                    processWelcomeMessage();
                });
            } else {
                // MathJax可能没有加载，直接处理
                processWelcomeMessage();
            }
        }, 1000);
    }

    // 设置输入框的快捷键监听
    setupInputShortcuts();
});

// 处理欢迎消息的函数
function processWelcomeMessage() {
    const welcomeMessage = document.querySelector('.message.assistant-message');
    if (welcomeMessage) {
        // 等待一小段时间，确保DOM完全加载
        setTimeout(() => {
            // 强制MathJax重新处理欢迎消息
            if (typeof MathJax !== 'undefined') {
                const markdownContent = welcomeMessage.querySelector('.markdown-content');

                if (markdownContent && !markdownContent.hasAttribute('data-welcome-processed')) {
                    markdownContent.setAttribute('data-welcome-processed', 'true');

                    // 首先处理代码块
                    processRenderedContent(markdownContent);

                    // 然后处理MathJax
                    MathJax.Hub.Queue(["Typeset", MathJax.Hub, markdownContent]);

                    // 在MathJax完成后添加公式工具栏
                    MathJax.Hub.Queue(function () {
                        markdownContent.querySelectorAll('script[type^="math/tex"]').forEach((script, index) => {
                            // 检查是否已经被处理过
                            if (!script.parentElement.classList.contains('equation-container') &&
                                !script.hasAttribute('data-processed')) {

                                const formula = script.textContent.trim();
                                if (formula) {
                                    const isDisplayMode = script.type.includes('mode=display');

                                    const container = document.createElement('div');
                                    container.className = 'equation-container';

                                    const toolbar = document.createElement('div');
                                    toolbar.className = 'math-toolbar';
                                    toolbar.innerHTML = `
                                        <p>手动插入：alt+=</p>
                                        <div>
                                            <button class="copy-math-button" onclick="copyMath('${btoa(formula)}')">复制公式</button>
                                            <button class="copy-to-word-button" onclick="insertMathToWord('${btoa(formula)}')">插入到Word</button>
                                        </div>
                                    `;

                                    // 检查script的父节点是否存在
                                    if (script.parentNode) {
                                        script.parentNode.insertBefore(container, script);
                                        container.appendChild(toolbar);
                                        container.appendChild(script);

                                        // 标记为已处理
                                        script.setAttribute('data-processed', 'true');
                                    }
                                }
                            }
                        });
                    });
                }
            } else {
                const markdownContent = welcomeMessage.querySelector('.markdown-content');
                if (markdownContent && !markdownContent.hasAttribute('data-welcome-processed')) {
                    markdownContent.setAttribute('data-welcome-processed', 'true');
                    processRenderedContent(markdownContent);
                }
            }
        }, 500); // 增加延迟到500ms，确保DOM稳定
    }
}

// 自定义弹窗函数
function showCustomAlert(message, callback) {
    const existingAlert = document.querySelector('.custom-alert-container');
    if (existingAlert) {
        existingAlert.remove();
    }

    const alertContainer = document.createElement('div');
    alertContainer.className = 'custom-alert-container';

    const alertContent = document.createElement('div');
    alertContent.className = 'custom-alert-content';

    alertContent.innerHTML = `
        <p>${message}</p>
        <div class="alert-buttons">
            <button class="alert-confirm-btn">确定</button>
        </div>
    `;

    alertContainer.appendChild(alertContent);
    document.body.appendChild(alertContainer);

    const confirmBtn = alertContent.querySelector('.alert-confirm-btn');
    confirmBtn.addEventListener('click', () => {
        alertContainer.remove();
        if (callback) {
            callback();
        }
    });

    alertContainer.addEventListener('click', (e) => {
        if (e.target === alertContainer) {
            alertContainer.remove();
        }
    });
}

// 初始化自定义右键菜单
function initializeContextMenu() {
    document.addEventListener('contextmenu', function (e) {
        // 只在聊天消息区域启用自定义右键菜单
        if (e.target.closest('.chat-messages')) {
            e.preventDefault();

            const selection = window.getSelection();
            const selectedText = selection.toString().trim();

            if (selectedText) {
                showContextMenu(e.pageX, e.pageY, selectedText);
            }
        }
    });

    document.addEventListener('click', function () {
        hideContextMenu();
    });

    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            hideContextMenu();
        }
    });
}

// 显示右键菜单
function showContextMenu(x, y, selectedText) {
    hideContextMenu();

    const contextMenu = document.createElement('div');
    contextMenu.className = 'custom-context-menu';
    contextMenu.style.cssText = `
        position: fixed;
        left: ${x}px;
        top: ${y}px;
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        z-index: 1000;
        min-width: 120px;
    `;

    contextMenu.innerHTML = `
        <div class="context-menu-item" style="padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #f3f4f6;" data-action="copy">
            复制选中内容
                    </div>
        <div class="context-menu-item" style="padding: 8px 12px; cursor: pointer;" data-action="insert">
            插入到Word
                </div>
    `;

    document.body.appendChild(contextMenu);

    // 添加悬停效果
    contextMenu.querySelectorAll('.context-menu-item').forEach(item => {
        item.addEventListener('mouseenter', () => {
            item.style.backgroundColor = '#f3f4f6';
        });
        item.addEventListener('mouseleave', () => {
            item.style.backgroundColor = '';
        });

        item.addEventListener('click', () => {
            const action = item.getAttribute('data-action');
            if (action === 'copy') {
                copyToClipboard(selectedText);
                showCustomAlert('已复制选中内容');
            } else if (action === 'insert') {
                insertToWord(selectedText);
            }
            hideContextMenu();
        });
    });
}

// 隐藏右键菜单
function hideContextMenu() {
    const contextMenu = document.querySelector('.custom-context-menu');
    if (contextMenu) {
        contextMenu.remove();
    }
}

// 初始化通信
if (window.chrome && window.chrome.webview) {
    console.log('WebView2环境已检测到');

    window.chrome.webview.addEventListener('message', function (event) {
        console.log('收到来自C#的消息:', event.data);
        handleMessageFromCSharp(event.data);
    });

    window.chrome.webview.postMessage({
        type: 'ready',
        message: 'JavaScript环境已准备就绪'
    });
} else {
    console.log('未检测到WebView2环境');
}

// 处理从C#接收的消息
function handleMessageFromCSharp(data) {
    try {
        if (data.type === 'startGenerating') {
            startGeneratingOutline();
        } else if (data.type === 'appendContent') {
            appendOutlineContent(data.content);
        } else if (data.type === 'finishGenerating') {
            finishGeneratingOutline();
        } else if (data.type === 'showError') {
            showError(data.message);
        } else if (data.type === 'setWelcomeMessage') {
            // 处理设置欢迎消息的请求
            const welcomeContent = document.querySelector('#welcome-message .markdown-content');
            if (welcomeContent) {
                // 检查内容格式，如果是markdown则需要渲染
                if (data.format === 'markdown') {
                    welcomeContent.innerHTML = renderMarkdown(data.content);
                } else {
                    welcomeContent.innerHTML = data.content;
                }
                console.log('欢迎消息已更新');

                // 重新处理欢迎消息
                processWelcomeMessage();
            } else {
                console.error('未找到欢迎消息元素');
            }
        } else if (data.type === 'documentHeadings') {
            // 处理文档标题数据
            const receiveTime = performance.now();
            console.log(`⏱️ 收到文档标题数据, 接收时间: ${receiveTime.toFixed(2)}ms`, data);

            // 只有在仍在获取状态时才处理数据
            if (isFetchingHeadings) {
                if (data.cancelled) {
                    // 处理取消状态
                    showSelectorError(data.message || '获取标题已取消');
                } else if (data.error) {
                    showSelectorError(data.error);
                } else {
                    const processStartTime = performance.now();
                    console.log(`⏱️ 开始处理标题数据, 处理开始时间: ${processStartTime.toFixed(2)}ms`);
                    showHeadingsInSelector(
                        data.headings,
                        data.page || 0,
                        data.append || false,
                        data.hasMore || false,
                        data.total || 0
                    );
                    console.log(`⏱️ 标题数据处理完成, 总耗时: ${(performance.now() - processStartTime).toFixed(2)}ms`);
                }
            } else {
                console.log('⏱️ 标题获取已取消，忽略返回的数据');
            }
        } else if (data.type === 'generateMermaidPNG') {
            // 处理生成Mermaid PNG的请求
            handleGenerateMermaidPNG(data.containerIndex, data.mermaidCode);
        } else if (data.type === 'toolPreview') {
            // 处理工具预览结果
            handleToolPreview(data);
        } else if (data.type === 'actionApplied') {
            // 处理操作应用结果
            handleActionApplied(data);
        } else if (data.type === 'modelList') {
            // 处理模型列表数据
            handleModelList(data.models);
        } else if (data.type === 'documentList') {
            // 处理文档列表数据
            showDocumentsInSelector(data.documents);
        } else if (data.type === 'documentHeadingList') {
            // 处理文档标题列表数据
            showDocumentHeadingsInSelector(data.documentId, data.documentName, data.headings);
        } else if (data.type === 'documentContent') {
            // 处理单个标题内容（如果需要的话）
            console.log('收到文档内容:', data);
        } else if (data.type === 'documentError') {
            // 处理文档错误
            console.error('文档错误:', data.message);
            if (contextSelectorContent) {
                contextSelectorContent.innerHTML = `<div class="context-selector-empty">错误: ${data.message}</div>`;
            }
        } else if (data.type === 'toolProgress') {
            // 处理工具调用进度
            handleToolProgress(data);
        }
    } catch (error) {
        console.error('处理C#消息时出错:', error);
    }
}

// 初始化聊天界面
function initializeChat() {
    if (typingIndicator) {
        typingIndicator.style.display = 'none';
    }

    scrollToBottom();

    if (messageInput) {
        messageInput.focus();
    }
}

// 设置事件监听器
function setupEventListeners() {
    if (sendBtn) {
        sendBtn.addEventListener('click', sendMessage);
    }

    if (messageInput) {
        messageInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
            // Backspace 整块删除 @标签
            if (e.key === 'Backspace') {
                if (tryDeleteWholeMention(true)) {
                    e.preventDefault();
                }
            }
        });

        messageInput.addEventListener('input', function () {
            updateCharacterCount();
            autoResizeInput();
            updateInputHighlights();
        });

        // 初始化一次高亮
        setTimeout(updateInputHighlights, 0);

        // 使用 ResizeObserver 同步高亮层的尺寸和位置，避免换行错位
        try {
            const ro = new ResizeObserver(() => syncHighlighterRect());
            ro.observe(messageInput);
            window.addEventListener('resize', syncHighlighterRect);
            messageInput.addEventListener('scroll', syncHighlighterScroll);
            syncHighlighterRect();
        } catch (e) {
            // 低版本不支持则使用定时器降级
            setInterval(() => { syncHighlighterRect(); syncHighlighterScroll(); }, 200);
        }
    }

    if (chatMessages) {
        chatMessages.addEventListener('scroll', handleScroll);
    }

    // 模型选择器事件
    if (modelSelect) {
        modelSelect.addEventListener('change', function () {
            selectedModelId = parseInt(this.value) || 0;
            console.log('选择模型ID:', selectedModelId);
        });
    }
}

// 检测光标是否紧贴某个 @标签之后，并整块删除
function tryDeleteWholeMention(expandToSpace) {
    if (!messageInput) return false;
    if (messageInput.selectionStart !== messageInput.selectionEnd) return false; // 有选区时按默认行为
    const text = messageInput.value;
    const cursor = messageInput.selectionStart;
    // 退格整块删除也要兼容带零宽分隔符的新mention
    const re = /@[\s\S]*?\u200b|@[^\s@]+/g;
    let m;
    while ((m = re.exec(text)) !== null) {
        const start = m.index;
        const end = start + m[0].length;
        if (end === cursor || end + 1 === cursor) { // 允许尾部后跟一个空格
            const before = text.slice(0, start);
            // 同时吞掉标签后紧随的一个空格，避免需要按两次退格
            const after = text.slice(end + (expandToSpace && text[end] === ' ' ? 1 : 0));
            messageInput.value = before + after;
            const newPos = before.length;
            messageInput.setSelectionRange(newPos, newPos);
            updateCharacterCount();
            autoResizeInput();
            updateInputHighlights();
            return true;
        }
    }
    return false;
}

// 渲染输入框内的 @标题 高亮为小标签
function updateInputHighlights() {
    if (!messageHighlighter || !messageInput) return;
    const raw = messageInput.value;
    if (!raw) {
        messageHighlighter.innerHTML = '';
        return;
    }
    // 将文本中的 @mention 识别为一个整体：优先匹配带零宽分隔符（\u200b）的新格式
    const tokens = [];
    let i = 0;
    const re = /@[\s\S]*?\u200b|@[^\s@]+/g;
    let m;
    while ((m = re.exec(raw)) !== null) {
        const start = m.index;
        const end = start + m[0].length;
        if (start > i) tokens.push({ t: 'text', v: raw.slice(i, start) });
        tokens.push({ t: 'mention', v: m[0] });
        i = end;
    }
    if (i < raw.length) tokens.push({ t: 'text', v: raw.slice(i) });

    // 拼装HTML - 使用纯 inline 布局确保换行行为与 textarea 完全一致
    let html = '';
    for (const token of tokens) {
        if (token.t === 'mention') {
            // 直接使用原始文本（包含NBSP和零宽字符），确保宽度和换行与textarea一致
            // 视觉上用背景色标记这是一个mention
            const rawText = escapeHtml(token.v);
            // 移除零宽字符用于显示，但保留NBSP以保持宽度
            const displayText = rawText.replace(/\u200b/g, '');
            html += `<span class="mention-tag">${displayText}</span>`;
        } else {
            // 普通文本，保持原样
            const textHtml = escapeHtml(token.v);
            html += textHtml;
        }
    }
    messageHighlighter.innerHTML = html;
    // 同步容器位置与尺寸
    syncHighlighterRect();
}

// 将高亮层定位到与 textarea 可视文本同一位置
function syncHighlighterRect() {
    if (!messageHighlighter || !messageInput) return;
    const rect = messageInput.getBoundingClientRect();
    const parentRect = messageInput.offsetParent ? messageInput.offsetParent.getBoundingClientRect() : { left: 0, top: 0 };
    const style = window.getComputedStyle(messageInput);
    const padLeft = parseFloat(style.paddingLeft || '0');
    const padTop = parseFloat(style.paddingTop || '0');
    const padRight = parseFloat(style.paddingRight || '0');
    const padBottom = parseFloat(style.paddingBottom || '0');
    messageHighlighter.style.left = (messageInput.offsetLeft + padLeft) + 'px';
    messageHighlighter.style.top = (messageInput.offsetTop + padTop) + 'px';
    messageHighlighter.style.width = (messageInput.clientWidth - padLeft - padRight) + 'px';
    messageHighlighter.style.height = (messageInput.clientHeight - padTop - padBottom) + 'px';
    // 同步滚动偏移
    syncHighlighterScroll();
}

function syncHighlighterScroll() {
    if (!messageHighlighter || !messageInput) return;
    messageHighlighter.style.transform = `translate(${-messageInput.scrollLeft}px, ${-messageInput.scrollTop}px)`;
}

// 发送消息
function sendMessage() {
    if (isGenerating) return; // 如果正在生成，则返回

    const message = messageInput.value.trim();
    if (!message) return;

    // 获取当前选择的模式
    currentChatMode = chatModeSelect ? chatModeSelect.value : 'chat';

    // 在发送消息前再次检查：如果是Agent模式，确保模型支持工具调用
    if (currentChatMode === 'chat-agent') {
        const currentModel = getSelectedModelInfo();
        if (currentModel && currentModel.enableTools !== 1) {
            console.log('检测到模型不支持工具调用，阻止发送并自动切换模式');
            // 自动切换回智能问答模式
            switchToChatMode(currentModel.name);
            return; // 阻止发送消息
        }
    }

    console.log('发送消息:', message, '模式:', currentChatMode);

    // 添加用户消息到界面（UI显示原始消息，不包含操作记录）
    addUserMessage(message);

    // 清空输入框
    messageInput.value = '';
    updateCharacterCount();
    autoResizeInput();
    updateInputHighlights();

    // 发送到C#
    if (window.chrome && window.chrome.webview) {
        // 在发送前，若存在待处理预览，按要求默认“拒绝”，并记录到决策日志
        try {
            const pending = previewManager.getPendingPreviews();
            if (pending && pending.length > 0) {
                console.log(`发现 ${pending.length} 个待处理预览，按默认规则标记为“拒绝”并记录`);
                previewManager.rejectAll(); // 内部也会调用 recordPreviewDecision
            }
        } catch (e) {
            console.warn('默认拒绝待处理预览失败:', e);
        }

        // 直接使用用户设置的工具列表，不进行前端收窄
        // AI模型会根据提示词中的【操作决策与工具选择规则】自主判断
        let enabledForMessage = getEnabledToolsList();
        console.log('启用的工具列表:', enabledForMessage);

        // 构造发送给模型的消息：在原始消息后追加"操作记录"块
        let messageToSend = message;
        if (currentChatMode === 'chat-agent' && previewDecisionLogs.length > 0) {
            const lines = previewDecisionLogs.map(x => `- ${x.toolName}: ${x.decision === 'accepted' ? '接受' : '拒绝'}`);
            const recordBlock = `\n\n[操作记录]\n${lines.join('\n')}`;
            messageToSend = message + recordBlock;
        }

        const messageData = {
            type: 'userMessage',
            message: messageToSend,
            mode: currentChatMode,
            selectedModelId: selectedModelId,
            enabledTools: enabledForMessage
        };

        // 如果是智能体模式，添加启用的工具列表
        if (currentChatMode === 'chat-agent') {
            console.log('=== 工具设置调试信息 ===');
            console.log('当前聊天模式:', currentChatMode);
            console.log('enabledTools对象(用户设置):', enabledTools);
            console.log('本轮启用的工具白名单:', messageData.enabledTools);
            console.log('白名单长度:', messageData.enabledTools.length);
            console.log('=== 调试信息结束 ===');
        }

        // 添加上下文内容
        if (selectedContexts && selectedContexts.length > 0) {
            messageData.contexts = selectedContexts;
            console.log('=== 上下文内容调试信息 ===');
            console.log('选中的上下文数量:', selectedContexts.length);
            console.log('上下文列表:', selectedContexts);
            console.log('=== 上下文调试信息结束 ===');
        }

        // 添加图片数据
        if (selectedImages && selectedImages.length > 0) {
            messageData.images = selectedImages.map(img => ({
                name: img.name,
                base64: img.base64,
                type: img.type,
                size: img.size
            }));
            console.log('=== 图片内容调试信息 ===');
            console.log('选中的图片数量:', selectedImages.length);
            console.log('图片总大小:', selectedImages.reduce((sum, img) => sum + img.size, 0));
            console.log('=== 图片调试信息结束 ===');

            // 发送后清空图片列表
            setTimeout(() => {
                clearAllImages();
            }, 100);
        }

        console.log('发送消息数据:', messageData);
        window.chrome.webview.postMessage(messageData);

        // 清空本轮的决策日志，避免重复累积到后续轮次
        previewDecisionLogs = [];
        // 在WebView2环境中，不在这里调用startGenerating，等待C#的响应
    } else {
        // 测试模式 - 直接开始生成响应
        setTimeout(() => {
            simulateResponse(message);
        }, 1000);
    }
}

// 添加用户消息
function addUserMessage(content) {
    const messageId = `user-message-${messageIdCounter++}`;

    // 保存当前的图片列表（用于点击预览）
    const currentImages = selectedImages.map(img => ({ ...img }));

    // 构建图片HTML（如果有图片）
    let imagesHtml = '';
    if (currentImages && currentImages.length > 0) {
        imagesHtml = '<div class="message-images">';
        currentImages.forEach((image, index) => {
            imagesHtml += `
                <div class="message-image-item" data-image-index="${index}" title="${escapeHtml(image.name)}">
                    <img src="${image.base64}" alt="${escapeHtml(image.name)}">
                </div>
            `;
        });
        imagesHtml += '</div>';
    }

    const messageHtml = `
        <div class="message user-message" id="${messageId}">
            <div class="message-avatar">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    <circle cx="12" cy="7" r="4" stroke="currentColor" stroke-width="2"/>
                </svg>
            </div>
            <div class="message-content">
                ${imagesHtml}
                <div class="message-text">
                    <div class="markdown-content">${escapeHtml(content)}</div>
                </div>
                <div class="message-actions">
                    <button class="action-btn copy-btn" onclick="copyMessage('${messageId}')" title="复制消息">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <rect x="9" y="9" width="13" height="13" rx="2" ry="2" stroke="currentColor" stroke-width="2" fill="none"/>
                            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" stroke="currentColor" stroke-width="2" fill="none"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>
    `;

    chatMessages.insertAdjacentHTML('beforeend', messageHtml);

    // 为消息中的图片添加点击预览事件
    if (currentImages && currentImages.length > 0) {
        const messageElement = document.getElementById(messageId);
        const imageItems = messageElement.querySelectorAll('.message-image-item');
        imageItems.forEach((item, index) => {
            item.addEventListener('click', function () {
                previewMessageImage(currentImages[index]);
            });
        });
    }

    scrollToBottom();
}

// 预览消息中的图片
function previewMessageImage(image) {
    const modal = document.getElementById('image-preview-modal');
    const modalImg = document.getElementById('image-preview-content');
    const caption = document.getElementById('image-preview-caption');

    if (modal && modalImg && caption) {
        modal.style.display = 'block';
        modalImg.src = image.base64;
        caption.textContent = image.name;
    }
}

// 添加助手消息
function addAssistantMessage(content = '') {
    const messageId = `assistant-message-${messageIdCounter++}`;
    const messageHtml = `
        <div class="message assistant-message" id="${messageId}">
            <div class="message-avatar">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/>
                    <path d="M8 14s1.5 2 4 2 4-2 4-2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                    <line x1="9" y1="9" x2="9.01" y2="9" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                    <line x1="15" y1="9" x2="15.01" y2="9" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                </svg>
                    </div>
            <div class="message-content">
                <div class="message-text">
                    <div class="markdown-content" id="${messageId}-content">${content}</div>
                </div>
                <div class="message-actions">
                    <button class="action-btn copy-btn" onclick="copyMessage('${messageId}')" title="复制消息">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <rect x="9" y="9" width="13" height="13" rx="2" ry="2" stroke="currentColor" stroke-width="2" fill="none"/>
                            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" stroke="currentColor" stroke-width="2" fill="none"/>
                        </svg>
                    </button>
                    <button class="action-btn insert-word-btn" onclick="insertMessageToWord('${messageId}')" title="插入到Word">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" stroke-width="2" fill="none"/>
                            <polyline points="14,2 14,8 20,8" stroke="currentColor" stroke-width="2" fill="none"/>
                            <line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" stroke-width="2"/>
                            <line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" stroke-width="2"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>
    `;

    chatMessages.insertAdjacentHTML('beforeend', messageHtml);
    scrollToBottom();

    // 返回DOM元素而不是ID
    const element = document.getElementById(messageId);
    return element;
}

// 开始生成响应
function startGenerating() {
    isGenerating = true;
    currentContent = '';
    // 清空工具卡片缓存，避免被后续完整渲染覆盖
    toolCardsCache = [];

    if (sendBtn) {
        // 改变发送按钮为停止按钮
        sendBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="6" y="6" width="12" height="12" rx="2" fill="currentColor"/>
            </svg>
        `;
        sendBtn.classList.add('stop-mode');
        sendBtn.title = '停止生成 (ESC)';
        sendBtn.onclick = stopGeneration;
    }

    // 显示简化的typing indicator
    if (typingIndicator) {
        typingIndicator.style.display = 'block';
    }

    scrollToBottom();
}

// 结束生成响应
function finishGenerating() {
    isGenerating = false;
    currentMessageId = null; // 清除当前消息ID

    // 如果有未处理的预览，清除预览状态
    if (isPreviewPending) {
        console.log('生成结束时清除未处理的预览状态');
        isPreviewPending = false;
    }
    // 释放共享宿主消息引用，避免后续请求误复用
    toolProgressHostMessage = null;

    if (sendBtn) {
        sendBtn.classList.remove('stop-mode');
        sendBtn.onclick = sendMessage; // 恢复发送功能
        // 恢复发送按钮原始图标
        sendBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <line x1="22" y1="2" x2="11" y2="13" stroke="currentColor" stroke-width="2"/>
                <polygon points="22,2 15,22 11,13 2,9 22,2" fill="currentColor"/>
            </svg>
        `;
        sendBtn.title = '发送消息 (Ctrl+Enter)';
    }

    if (typingIndicator) {
        typingIndicator.style.display = 'none';
    }

    // 再次渲染所有Mermaid图表，确保完全加载
    if (typeof mermaid !== 'undefined') {
        try {
            mermaid.init(undefined, document.querySelectorAll('.mermaid'));
        } catch (error) {
            console.error('Mermaid最终渲染错误:', error);
        }
    }

    // 模型处理完成后，检查是否需要显示批量操作按钮
    setTimeout(() => {
        checkAndShowBatchActions();
    }, 500);
}

// C#调用的函数
function startGeneratingOutline() {
    startGenerating();
    // 不再预先创建空的助手消息，让appendOutlineContent在需要时创建
    currentMessageId = null;
    return 'will_create_when_needed';
}

function appendOutlineContent(content) {
    // 检查是否还在生成状态，如果已经停止则忽略
    if (!isGenerating) {
        console.log('生成已停止，忽略新内容:', content.substring(0, 50) + '...');
        return;
    }

    // 如果有预览待处理，暂停内容追加
    if (isPreviewPending) {
        console.log('有预览待处理，暂停内容追加:', content.substring(0, 50) + '...');
        return;
    }

    // 允许空白内容（空格/换行）参与渲染，避免流式丢失空格/换行
    if (content === undefined || content === null) {
        console.log('内容为undefined/null，跳过');
        return;
    }

    currentContent += content;

    // 获取最后一个助手消息
    let lastMessage = chatMessages.querySelector('.assistant-message:last-child .markdown-content');

    // 如果尚未创建消息且目前仅有空白，先缓冲，等到有可视内容再创建
    if (!lastMessage) {
        if (currentContent.trim() === '') {
            return; // 仅缓冲空白，不创建消息
        }
        const messageElement = addAssistantMessage('');
        lastMessage = messageElement.querySelector('.markdown-content');
        // 更新currentMessageId
        if (messageElement && messageElement.id) {
            currentMessageId = messageElement.id;
            console.log('设置currentMessageId为:', currentMessageId);
        }
    }

    if (lastMessage) {
        // 先解析ReAct内容，再渲染Markdown
        const parsedContent = parseReActContent(currentContent);
        let renderedHTML = renderMarkdown(parsedContent);

        // 将工具卡片占位符替换为实际HTML（兼容<p>包裹/注释占位的情况）
        try {
            toolCardsCache.forEach(card => {
                // 新占位符（HTML注释，不会被Markdown改写）
                const placeholderComment = `<!--TOOL_CARD_${card.id}-->`;
                // 旧占位符（历史兼容：会被Markdown改写为<strong>...>）
                const placeholderLegacy = `__TOOL_CARD_${card.id}__`;
                const patterns = [
                    // 新占位符逻辑
                    `<p>${placeholderComment}</p>`,
                    placeholderComment,
                    // 旧占位符的几种可能形态
                    `<p>${placeholderLegacy}</p>`,
                    `<p>__TOOL_CARD_${card.id}__</p>`,
                    placeholderLegacy,
                    // 旧占位符被Markdown包装为<strong>
                    `<p><strong>TOOL_CARD_${card.id}</strong></p>`,
                    `<strong>TOOL_CARD_${card.id}</strong>`
                ];
                patterns.forEach(pattern => {
                    if (renderedHTML.indexOf(pattern) >= 0) {
                        renderedHTML = renderedHTML.replace(pattern, card.html);
                    }
                });
            });
        } catch (e) {
            console.warn('工具卡片占位符替换失败:', e);
        }

        lastMessage.innerHTML = renderedHTML;

        // 首次渲染过后，关闭卡片后续动画，避免每次流式刷新触发进入动画造成闪烁
        try {
            toolCardsCache.forEach(card => {
                if (!card.mounted) {
                    card.mounted = true;
                    // 为后续字符串替换的HTML加上 no-animate 标记
                    if (typeof card.html === 'string') {
                        if (card.html.indexOf('tool-call-card no-animate') < 0) {
                            card.html = card.html.replace('tool-call-card ', 'tool-call-card no-animate ');
                        }
                    }
                }
            });
        } catch (e) {
            console.warn('关闭卡片动画失败:', e);
        }

        // 重新渲染HTML后，清除处理标记，确保工具栏能被重新添加
        lastMessage.removeAttribute('data-processed');
        lastMessage.removeAttribute('data-mathjax-processed');

        // 立即处理渲染内容
        processRenderedContent(lastMessage);

        // 强制MathJax重新渲染整个消息内容
        if (typeof MathJax !== 'undefined') {
            // 先移除所有现有的MathJax元素，避免冲突
            const existingMathJax = lastMessage.querySelectorAll('.MathJax, .MathJax_Display, .MathJax_Preview, [id^="MathJax"]');
            existingMathJax.forEach(element => element.remove());

            // 强制MathJax重新扫描和渲染
            MathJax.Hub.Queue(["Typeset", MathJax.Hub, lastMessage]);

            // 渲染完成后确保公式工具栏正确添加
            MathJax.Hub.Queue(function () {
                setTimeout(() => {
                    const newScripts = lastMessage.querySelectorAll('script[type^="math/tex"]');
                    newScripts.forEach(script => {
                        if (!script.parentElement.classList.contains('equation-container') &&
                            !script.hasAttribute('data-processed')) {
                            addFormulaToolbar(script);
                        }
                    });
                }, 150); // 增加延迟确保MathJax完全完成
            });
        }
    }

    scrollToBottom();
}

function finishGeneratingOutline() {
    finishGenerating();

    // 重置工具进度宿主消息，为下次对话做准备
    toolProgressHostMessage = null;

    // 不再自动隐藏工具进度容器，让用户可以查看工具执行过程
    // hideToolProgressContainer();

    scrollToBottom();
}

// 开始在现有消息中追加内容（用于用户操作反馈）
function startAppendingToExistingMessage() {
    console.log('🔄 开始在现有消息中追加用户反馈回复');

    // 重置状态，准备追加内容
    startGenerating();

    // 获取包含预览卡片的消息（toolProgressHostMessage）
    let targetMessage = toolProgressHostMessage && document.body.contains(toolProgressHostMessage)
        ? toolProgressHostMessage
        : chatMessages.querySelector('.assistant-message:last-child');

    if (!targetMessage) {
        console.warn('未找到目标消息，创建新消息');
        targetMessage = addAssistantMessage('');
    }

    // 设置当前消息ID
    currentMessageId = targetMessage.id;
    console.log('✅ 将在现有消息中追加内容，消息ID:', currentMessageId);

    // 重置内容，准备追加AI反馈
    currentContent = '';

    return currentMessageId;
}

// 向现有消息追加内容
function appendToExistingMessage(content) {
    // 检查是否还在生成状态
    if (!isGenerating) {
        console.log('⚠️ 生成已停止，忽略新内容:', content.substring(0, 50) + '...');
        return;
    }

    // 允许空白内容参与渲染
    if (content === undefined || content === null) {
        console.log('⚠️ 内容为undefined/null，跳过');
        return;
    }

    console.log('📝 向现有消息追加内容:', content.substring(0, 30) + '...');
    currentContent += content;

    // 获取目标消息
    let targetMessage = null;
    if (currentMessageId) {
        targetMessage = document.getElementById(currentMessageId);
    }

    if (!targetMessage) {
        targetMessage = chatMessages.querySelector('.assistant-message:last-child');
    }

    if (targetMessage) {
        const messageContent = targetMessage.querySelector('.message-content .markdown-content');
        if (messageContent) {
            // 查找或创建反馈内容区域
            let feedbackArea = messageContent.querySelector('.user-feedback-area');
            if (!feedbackArea) {
                // 创建反馈区域
                feedbackArea = document.createElement('div');
                feedbackArea.className = 'user-feedback-area';
                feedbackArea.style.cssText = `
                    margin-top: 16px;
                    padding-top: 12px;
                    border-top: 1px solid #e5e7eb;
                    font-size: 14px;
                    line-height: 1.6;
                `;
                messageContent.appendChild(feedbackArea);
            }

            // 解析并渲染反馈内容
            const parsedContent = parseReActContent(currentContent);
            feedbackArea.innerHTML = renderMarkdown(parsedContent);

            // 重新处理渲染内容
            feedbackArea.removeAttribute('data-processed');
            feedbackArea.removeAttribute('data-mathjax-processed');
            processRenderedContent(feedbackArea);

            // MathJax渲染
            if (typeof MathJax !== 'undefined') {
                const existingMathJax = feedbackArea.querySelectorAll('.MathJax, .MathJax_Display, .MathJax_Preview, [id^="MathJax"]');
                existingMathJax.forEach(element => element.remove());

                MathJax.Hub.Queue(["Typeset", MathJax.Hub, feedbackArea]);
                MathJax.Hub.Queue(function () {
                    setTimeout(() => {
                        const newScripts = feedbackArea.querySelectorAll('script[type^="math/tex"]');
                        newScripts.forEach(script => {
                            if (!script.parentElement.classList.contains('equation-container') &&
                                !script.hasAttribute('data-processed')) {
                                addFormulaToolbar(script);
                            }
                        });
                    }, 150);
                });
            }
        }
    }

    scrollToBottom();
}

// 完成在现有消息中追加内容
function finishAppendingToExistingMessage() {
    console.log('🏁 完成在现有消息中追加内容');
    finishGenerating();
    scrollToBottom();
}

// 完成生成并隐藏工具进度（用于测试等特殊场景）
function finishGeneratingWithHideProgress() {
    finishGenerating();

    // 隐藏工具进度容器
    hideToolProgressContainer();

    scrollToBottom();
}

// 隐藏工具进度容器
function hideToolProgressContainer() {
    const lastAssistantMessage = chatMessages.querySelector('.assistant-message:last-child');
    if (lastAssistantMessage) {
        const progressContainer = lastAssistantMessage.querySelector('.tool-progress-container');
        if (progressContainer) {
            // 添加淡出动画
            progressContainer.style.transition = 'opacity 0.5s ease, transform 0.5s ease';
            progressContainer.style.opacity = '0';
            progressContainer.style.transform = 'translateY(-10px)';

            // 延迟移除元素
            setTimeout(() => {
                if (progressContainer.parentNode) {
                    progressContainer.parentNode.removeChild(progressContainer);
                }
            }, 500);
        }
    }
}



// 停止生成
function stopGeneration() {
    if (!isGenerating) return;

    try {
        console.log('用户请求停止生成');

        // 立即设置状态为停止，防止后续内容继续添加
        isGenerating = false;

        // 清除预览待处理状态
        if (isPreviewPending) {
            console.log('停止生成时清除预览状态');
            isPreviewPending = false;
        }

        // 通知C#停止生成
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage({
                type: 'stopGeneration',
                messageId: currentMessageId
            });
        }

        // 在当前消息后添加停止提示
        const lastMessage = chatMessages.querySelector('.assistant-message:last-child .markdown-content');
        if (lastMessage && currentContent.trim()) {
            // 如果已经有内容，添加停止标记
            lastMessage.innerHTML += '<div class="generation-stopped">⏹️ 生成已停止</div>';
        }

        console.log('生成已停止，UI状态已更新');
    } catch (error) {
        console.error('停止生成时出错:', error);
    } finally {
        // 无论如何都要更新UI状态
        finishGenerating();
    }
}

function showError(message) {
    finishGenerating();
    addAssistantMessage(`<div style="color: #dc2626; background: #fef2f2; padding: 12px; border-radius: 8px; border-left: 4px solid #dc2626;"><strong>错误:</strong> ${escapeHtml(message)}</div>`);
}

// Markdown渲染
function renderMarkdown(content) {
    // 预处理：移除模型特殊标记（如 DeepSeek 的 box 标记）
    let processedContent = content;

    // 移除 <|begin_of_box|> 和 |end_of_box|> 标记
    processedContent = processedContent.replace(/<\|begin_of_box\|>/g, '');
    processedContent = processedContent.replace(/<\|end_of_box\|>/g, '');

    // 移除其他可能的特殊标记
    processedContent = processedContent.replace(/<\|[^|]+\|>/g, '');

    // 预处理LaTeX公式 - 将$...$和$$...$$转换为MathJax script标签

    // 处理显示模式公式 $$...$$
    processedContent = processedContent.replace(/\$\$([\s\S]*?)\$\$/g, function (match, formula) {
        const cleanFormula = formula.trim();
        return `<script type="math/tex; mode=display">${cleanFormula}</script>`;
    });

    // 处理内联公式 $...$
    processedContent = processedContent.replace(/\$([^$\n]+?)\$/g, function (match, formula) {
        const cleanFormula = formula.trim();
        return `<script type="math/tex">${cleanFormula}</script>`;
    });

    if (typeof marked !== 'undefined') {
        marked.setOptions({
            breaks: true,
            gfm: true,
            highlight: function (code, lang) {
                // 如果是Mermaid代码块，不进行高亮处理
                if (lang === 'mermaid') {
                    return code;
                }
                if (typeof hljs !== 'undefined') {
                    const language = hljs.getLanguage(lang) ? lang : 'plaintext';
                    return hljs.highlight(code, { language }).value;
                }
                return code;
            }
        });
        return marked.parse(processedContent);
    }

    // 简单的Markdown处理
    return processedContent
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\*(.*?)\*/g, '<em>$1</em>')
        .replace(/`(.*?)`/g, '<code>$1</code>')
        .replace(/\n/g, '<br>');
}

// 处理渲染后的内容
function processRenderedContent(element) {
    // 检查是否已经处理过，避免重复处理
    if (element.hasAttribute('data-processed')) {
        return;
    }

    // 处理代码块
    element.querySelectorAll('pre code').forEach((block, index) => {
        // 跳过思考过程中的代码块和Mermaid代码块，避免高亮处理
        if (block.closest('.thinking-process') || block.classList.contains('language-mermaid')) {
            return;
        }

        const pre = block.parentElement;
        if (!pre.querySelector('.code-toolbar')) {
            const language = (block.className.match(/language-(\w+)/) || ['', 'text'])[1];
            const code = block.innerText;

            // 安全编码处理，避免特殊字符问题
            let encodedCode = '';
            try {
                encodedCode = btoa(unescape(encodeURIComponent(code)));
            } catch (e) {
                console.warn('代码编码失败，使用原始代码:', e);
                encodedCode = '';
            }

            const toolbar = document.createElement('div');
            toolbar.className = 'code-toolbar';
            toolbar.innerHTML = `
                <p>${language}</p>
                <div>
                    <button class="copy-code-button" onclick="copyCode('${encodedCode}')">复制代码</button>
                    <button class="copy-to-word-button" onclick="insertCodeToWord('${encodedCode}', '${language}')">插入到Word</button>
                </div>
            `;

            pre.insertBefore(toolbar, pre.firstChild);
        }
    });

    // 处理Mermaid图表
    processMermaidDiagrams(element);

    // 处理表格
    element.querySelectorAll('table').forEach((table, index) => {
        if (!table.parentElement.classList.contains('table-container')) {
            const container = document.createElement('div');
            container.className = 'table-container';

            const toolbar = document.createElement('div');
            toolbar.className = 'table-toolbar';
            toolbar.innerHTML = `
                <p>表格</p>
                <div>
                    <button class="copy-table-button" onclick="copyTable(this)">复制表格</button>
                    <button class="copy-to-word-button" onclick="insertTableToWord(this)">插入到Word</button>
                </div>
            `;

            table.parentNode.insertBefore(container, table);
            container.appendChild(toolbar);
            container.appendChild(table);
        }
    });

    // 处理MathJax - 改进时序控制
    if (typeof MathJax !== 'undefined' && !element.hasAttribute('data-mathjax-processed')) {
        element.setAttribute('data-mathjax-processed', 'true');

        // 先处理现有的公式script标签
        const existingScripts = element.querySelectorAll('script[type^="math/tex"]');
        existingScripts.forEach(script => {
            if (!script.parentElement.classList.contains('equation-container') &&
                !script.hasAttribute('data-processed')) {
                addFormulaToolbar(script);
            }
        });

        // 然后进行MathJax渲染
        MathJax.Hub.Queue(["Typeset", MathJax.Hub, element]);

        // MathJax完成后再次检查新生成的公式
        MathJax.Hub.Queue(function () {
            setTimeout(() => {
                const newScripts = element.querySelectorAll('script[type^="math/tex"]');
                newScripts.forEach(script => {
                    if (!script.parentElement.classList.contains('equation-container') &&
                        !script.hasAttribute('data-processed')) {
                        addFormulaToolbar(script);
                    }
                });
            }, 100);
        });
    }

    // 标记整个元素已处理
    element.setAttribute('data-processed', 'true');
}

// 处理Mermaid图表
function processMermaidDiagrams(element) {
    // 初始化 Mermaid
    if (typeof mermaid !== 'undefined') {
        mermaid.initialize({
            startOnLoad: true,
            theme: 'default',
            flowchart: {
                useMaxWidth: true,
                htmlLabels: true
            }
        });
    }

    // 处理Mermaid图表
    const mermaidDivs = element.querySelectorAll('pre code.language-mermaid');
    mermaidDivs.forEach((codeElement, index) => {
        // 获取Mermaid代码
        const mermaidCode = codeElement.textContent;

        // 创建一个新的div来放置渲染后的图表
        const mermaidDiv = document.createElement('div');
        mermaidDiv.className = 'mermaid-container';

        // 创建工具栏
        const toolbar = document.createElement('div');
        toolbar.className = 'mermaid-toolbar';
        toolbar.innerHTML = `<p>流程图</p>
            <div class="mermaid-buttons">
                <button class="copy-mermaid-button">复制代码</button>
                <button class="insert-mermaid-word-button">插入Word</button>
                <div class="download-dropdown">
                    <button class="download-mermaid-button">下载图片</button>
                    <div class="download-options">
                        <button class="download-svg">SVG格式</button>
                        <button class="download-png">PNG格式</button>
                    </div>
                </div>
            </div>`;

        // 创建Mermaid图表div
        const mermaidContent = document.createElement('div');
        mermaidContent.className = 'mermaid';
        mermaidContent.textContent = mermaidCode;

        // 保存原始代码到数据属性，以便后续提取时使用
        mermaidContent.setAttribute('data-mermaid-code', mermaidCode);

        // 添加工具栏和图表到容器
        mermaidDiv.appendChild(toolbar);
        mermaidDiv.appendChild(mermaidContent);

        // 替换原来的pre元素
        const preElement = codeElement.parentElement;
        preElement.parentElement.replaceChild(mermaidDiv, preElement);

        // 为每个图表添加唯一ID
        mermaidContent.id = `mermaid-diagram-${Date.now()}-${index}`;

        // 添加复制按钮事件
        const copyButton = toolbar.querySelector('.copy-mermaid-button');
        copyButton.addEventListener('click', () => {
            navigator.clipboard.writeText(mermaidCode).then(() => {
                const originalText = copyButton.textContent;
                copyButton.textContent = '已复制!';
                setTimeout(() => { copyButton.textContent = originalText; }, 1500);
            });
        });

        // 添加插入Word按钮事件
        const insertWordButton = toolbar.querySelector('.insert-mermaid-word-button');
        insertWordButton.addEventListener('click', () => {
            insertMermaidToWord(mermaidContent, mermaidCode);
        });

        // 添加下载按钮事件
        const downloadDropdown = toolbar.querySelector('.download-dropdown');
        const downloadButton = toolbar.querySelector('.download-mermaid-button');
        downloadButton.addEventListener('click', () => {
            downloadDropdown.classList.toggle('active');
        });

        // 点击其他地方关闭下拉菜单
        document.addEventListener('click', (event) => {
            if (!event.target.closest('.download-dropdown') && downloadDropdown.classList.contains('active')) {
                downloadDropdown.classList.remove('active');
            }
        });

        // 下载SVG格式
        const downloadSvgButton = toolbar.querySelector('.download-svg');
        downloadSvgButton.addEventListener('click', () => {
            // 等待渲染完成
            setTimeout(() => {
                const svgElement = mermaidContent.querySelector('svg');
                if (svgElement) {
                    // 克隆SVG以便修改
                    const svgClone = svgElement.cloneNode(true);

                    // 确保SVG有正确的命名空间
                    svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

                    // 获取SVG源代码
                    const svgData = new XMLSerializer().serializeToString(svgClone);

                    // 创建Blob对象
                    const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });

                    // 创建下载链接
                    const downloadLink = document.createElement('a');
                    downloadLink.href = URL.createObjectURL(svgBlob);
                    downloadLink.download = `mermaid-diagram-${Date.now()}.svg`;

                    // 模拟点击下载
                    document.body.appendChild(downloadLink);
                    downloadLink.click();
                    document.body.removeChild(downloadLink);

                    // 关闭下拉菜单
                    downloadDropdown.classList.remove('active');
                }
            }, 100);
        });

        // 下载PNG格式
        const downloadPngButton = toolbar.querySelector('.download-png');
        downloadPngButton.addEventListener('click', () => {
            // 等待渲染完成
            setTimeout(() => {
                const svgElement = mermaidContent.querySelector('svg');
                if (svgElement) {
                    // 获取SVG尺寸
                    const svgRect = svgElement.getBoundingClientRect();
                    const width = svgRect.width;
                    const height = svgRect.height;

                    // 克隆SVG以便修改
                    const svgClone = svgElement.cloneNode(true);

                    // 确保SVG有正确的命名空间
                    svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

                    // 获取SVG源代码
                    const svgData = new XMLSerializer().serializeToString(svgClone);

                    // 创建Canvas
                    const canvas = document.createElement('canvas');
                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');

                    // 填充白色背景
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, width, height);

                    // 创建Image对象
                    const img = new Image();
                    img.onload = function () {
                        // 在Canvas上绘制SVG
                        ctx.drawImage(img, 0, 0);

                        // 将Canvas转换为PNG
                        try {
                            const pngUrl = canvas.toDataURL('image/png');

                            // 创建下载链接
                            const downloadLink = document.createElement('a');
                            downloadLink.href = pngUrl;
                            downloadLink.download = `mermaid-diagram-${Date.now()}.png`;

                            // 模拟点击下载
                            document.body.appendChild(downloadLink);
                            downloadLink.click();
                            document.body.removeChild(downloadLink);
                        } catch (e) {
                            console.error('PNG转换失败:', e);
                            alert('PNG转换失败，请尝试下载SVG格式');
                        }

                        // 关闭下拉菜单
                        downloadDropdown.classList.remove('active');
                    };

                    // 设置Image源
                    img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
                }
            }, 100);
        });
    });

    // 渲染所有Mermaid图表
    if (typeof mermaid !== 'undefined') {
        try {
            // 在渲染前保存所有data-mermaid-code属性
            const mermaidDivsWithCode = element.querySelectorAll('.mermaid');
            const savedCodes = [];

            mermaidDivsWithCode.forEach((div, index) => {
                const code = div.getAttribute('data-mermaid-code');
                if (code) {
                    savedCodes.push({
                        index: index,
                        element: div,
                        code: code
                    });
                }
            });

            // 渲染Mermaid图表
            mermaid.init(undefined, element.querySelectorAll('.mermaid'));

            // 渲染后恢复data-mermaid-code属性
            setTimeout(() => {
                savedCodes.forEach(item => {
                    const renderedDiv = item.element;
                    if (renderedDiv && !renderedDiv.hasAttribute('data-mermaid-code')) {
                        renderedDiv.setAttribute('data-mermaid-code', item.code);
                        console.log(`恢复data-mermaid-code属性，索引: ${item.index}`);
                    }
                });

                // 确保所有新渲染的SVG元素都有对应的原始代码引用
                element.querySelectorAll('.mermaid svg').forEach((svg, index) => {
                    const parentDiv = svg.closest('.mermaid');
                    if (parentDiv) {
                        const code = parentDiv.getAttribute('data-mermaid-code');
                        if (code) {
                            // 在SVG元素上也保存一份原始代码引用
                            svg.setAttribute('data-original-code', code);
                            console.log(`在SVG上设置data-original-code属性，索引: ${index}`);
                        }
                    }
                });
            }, 200);

        } catch (error) {
            console.error('Mermaid渲染错误:', error);
        }
    }
}

// 处理生成Mermaid PNG的请求
function handleGenerateMermaidPNG(containerIndex, mermaidCode) {
    console.log(`收到生成Mermaid PNG请求: 容器索引 ${containerIndex}`);

    // 查找对应的Mermaid容器
    const mermaidContainers = document.querySelectorAll('.mermaid-container');
    let targetContainer = null;

    // 根据代码匹配找到对应容器
    for (let container of mermaidContainers) {
        const mermaidContent = container.querySelector('.mermaid');
        if (mermaidContent && mermaidContent.textContent.trim() === mermaidCode) {
            targetContainer = container;
            break;
        }
    }

    if (targetContainer) {
        const mermaidContent = targetContainer.querySelector('.mermaid');
        const svgElement = mermaidContent.querySelector('svg');

        if (svgElement) {
            console.log('找到SVG元素，开始生成PNG');

            // 使用现有的insertMermaidToWord函数逻辑生成PNG
            setTimeout(() => {
                // 获取SVG尺寸
                const svgRect = svgElement.getBoundingClientRect();
                const width = svgRect.width;
                const height = svgRect.height;

                // 克隆SVG以便修改
                const svgClone = svgElement.cloneNode(true);

                // 确保SVG有正确的命名空间
                svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

                // 获取SVG源代码
                const svgData = new XMLSerializer().serializeToString(svgClone);

                // 创建Canvas
                const canvas = document.createElement('canvas');
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext('2d');

                // 填充白色背景
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, width, height);

                // 创建Image对象
                const img = new Image();
                img.onload = function () {
                    // 在Canvas上绘制SVG
                    ctx.drawImage(img, 0, 0);

                    // 将Canvas转换为PNG
                    try {
                        const pngDataUrl = canvas.toDataURL('image/png');

                        // 发送PNG数据到C#端处理
                        if (window.chrome && window.chrome.webview) {
                            window.chrome.webview.postMessage({
                                type: 'insertMermaidImage',
                                imageData: pngDataUrl,
                                mermaidCode: mermaidCode,
                                width: width,
                                height: height
                            });

                            console.log('Mermaid PNG数据已发送到C#端');
                        } else {
                            console.error('WebView2环境不可用');
                        }
                    } catch (e) {
                        console.error('PNG转换失败:', e);
                    }
                };

                img.onerror = function () {
                    console.error('SVG加载失败');
                };

                // 设置Image源
                img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
            }, 100); // 小延迟确保渲染完成
        } else {
            console.error('未找到SVG元素');
        }
    } else {
        console.error(`未找到容器索引 ${containerIndex} 对应的Mermaid容器`);
    }
}

// 生成Mermaid图片数据（同步版本）
function generateMermaidImageData(svgElement, mermaidCode) {
    try {
        // 获取SVG尺寸
        const svgRect = svgElement.getBoundingClientRect();
        const width = svgRect.width;
        const height = svgRect.height;

        // 克隆SVG以便修改
        const svgClone = svgElement.cloneNode(true);

        // 确保SVG有正确的命名空间
        svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

        // 获取SVG源代码
        const svgData = new XMLSerializer().serializeToString(svgClone);

        // 创建Canvas
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');

        // 填充白色背景
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);

        // 创建Image对象（同步处理）
        const img = new Image();

        // 使用同步方式处理（注意：这可能会阻塞UI）
        const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(svgBlob);

        // 尝试同步绘制（这里实际上还是异步的，但我们可以返回一个包含所有必要信息的对象）
        return {
            svgData: svgData,
            width: width,
            height: height,
            mermaidCode: mermaidCode
        };
    } catch (error) {
        console.error('生成Mermaid图片数据失败:', error);
        return null;
    }
}

// 将SVG转换为PNG数据URL
async function generatePNGFromSVG(svgElement) {
    return new Promise((resolve, reject) => {
        try {
            console.log('开始生成PNG: 获取SVG元素尺寸');

            // 获取SVG尺寸
            const svgRect = svgElement.getBoundingClientRect();
            const width = svgRect.width || 400;
            const height = svgRect.height || 300;

            console.log(`SVG尺寸: ${width}x${height}`);

            // 克隆SVG以便修改
            const svgClone = svgElement.cloneNode(true);

            // 确保SVG有正确的命名空间
            svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

            // 获取SVG源代码
            const svgData = new XMLSerializer().serializeToString(svgClone);
            console.log(`SVG数据长度: ${svgData.length}`);

            // 创建Canvas
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');

            if (!ctx) {
                console.error('无法获取Canvas上下文');
                reject(new Error('无法获取Canvas上下文'));
                return;
            }

            // 填充白色背景
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, width, height);

            // 创建Image对象
            const img = new Image();

            // 设置超时处理
            const timeout = setTimeout(() => {
                console.error('SVG加载超时');
                reject(new Error('SVG加载超时'));
            }, 5000);

            img.onload = function () {
                clearTimeout(timeout);
                console.log('SVG图像已加载，开始绘制到Canvas');

                // 在Canvas上绘制SVG
                ctx.drawImage(img, 0, 0);

                // 将Canvas转换为PNG
                try {
                    const pngDataUrl = canvas.toDataURL('image/png');
                    console.log(`PNG数据URL长度: ${pngDataUrl.length}`);
                    resolve(pngDataUrl);
                } catch (e) {
                    console.error('PNG转换失败:', e);
                    reject(e);
                }
            };

            img.onerror = function (error) {
                clearTimeout(timeout);
                console.error('SVG加载失败:', error);

                // 尝试使用另一种方式加载SVG
                try {
                    console.log('尝试使用备用方法加载SVG');
                    // 使用Blob URL
                    const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
                    const url = URL.createObjectURL(svgBlob);

                    const backupImg = new Image();
                    backupImg.onload = function () {
                        ctx.drawImage(backupImg, 0, 0);
                        try {
                            const pngDataUrl = canvas.toDataURL('image/png');
                            URL.revokeObjectURL(url);
                            console.log('备用方法成功生成PNG');
                            resolve(pngDataUrl);
                        } catch (e) {
                            console.error('备用方法PNG转换失败:', e);
                            URL.revokeObjectURL(url);
                            reject(e);
                        }
                    };

                    backupImg.onerror = function (backupError) {
                        console.error('备用方法SVG加载失败:', backupError);
                        URL.revokeObjectURL(url);
                        reject(new Error('SVG加载失败 (备用方法)'));
                    };

                    backupImg.src = url;
                } catch (backupError) {
                    console.error('备用方法失败:', backupError);
                    reject(new Error('SVG加载失败'));
                }
            };

            // 设置Image源
            console.log('设置SVG数据源');
            try {
                const base64Data = btoa(unescape(encodeURIComponent(svgData)));
                img.src = 'data:image/svg+xml;base64,' + base64Data;
                console.log('SVG数据源设置完成');
            } catch (e) {
                console.error('设置SVG数据源失败:', e);
                clearTimeout(timeout);
                reject(e);
            }

        } catch (error) {
            console.error('生成PNG时出错:', error);
            reject(error);
        }
    });
}

// 插入Mermaid图表到Word
function insertMermaidToWord(mermaidContent, mermaidCode) {
    // 等待渲染完成
    setTimeout(() => {
        const svgElement = mermaidContent.querySelector('svg');
        if (svgElement) {
            // 获取SVG尺寸
            const svgRect = svgElement.getBoundingClientRect();
            const width = svgRect.width;
            const height = svgRect.height;

            // 克隆SVG以便修改
            const svgClone = svgElement.cloneNode(true);

            // 确保SVG有正确的命名空间
            svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

            // 获取SVG源代码
            const svgData = new XMLSerializer().serializeToString(svgClone);

            // 创建Canvas
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');

            // 填充白色背景
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, width, height);

            // 创建Image对象
            const img = new Image();
            img.onload = function () {
                // 在Canvas上绘制SVG
                ctx.drawImage(img, 0, 0);

                // 将Canvas转换为PNG
                try {
                    const pngDataUrl = canvas.toDataURL('image/png');

                    // 发送到C#端处理
                    if (window.chrome && window.chrome.webview) {
                        window.chrome.webview.postMessage({
                            type: 'insertMermaidImage',
                            imageData: pngDataUrl,
                            mermaidCode: mermaidCode,
                            width: width,
                            height: height
                        });

                        // 显示成功提示
                        showCustomAlert('Mermaid流程图已插入到Word文档');
                    } else {
                        showCustomAlert('WebView2环境不可用');
                    }
                } catch (e) {
                    console.error('PNG转换失败:', e);
                    showCustomAlert('图片转换失败，请重试');
                }
            };

            img.onerror = function () {
                console.error('SVG加载失败');
                showCustomAlert('图片加载失败，请重试');
            };

            // 设置Image源
            img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
        } else {
            console.error('未找到SVG元素');
            showCustomAlert('未找到流程图，请重试');
        }
    }, 200); // 增加延迟确保渲染完成
}

// 新增的公式工具栏添加函数
function addFormulaToolbar(script) {
    const formula = script.textContent.trim();
    if (!formula) return;

    // 安全编码处理
    let encodedFormula = '';
    try {
        encodedFormula = btoa(unescape(encodeURIComponent(formula)));
    } catch (e) {
        console.warn('公式编码失败:', e);
        encodedFormula = '';
    }

    const isDisplayMode = script.type.includes('mode=display');

    const container = document.createElement('div');
    container.className = 'equation-container';

    const toolbar = document.createElement('div');
    toolbar.className = 'math-toolbar';
    toolbar.innerHTML = `
        <p>手动插入：alt+=</p>
        <div>
            <button class="copy-math-button" onclick="copyMath('${encodedFormula}')">复制公式</button>
            <button class="copy-to-word-button" onclick="insertMathToWord('${encodedFormula}')">插入到Word</button>
        </div>
    `;

    // 检查script的父节点是否存在
    if (script.parentNode) {
        script.parentNode.insertBefore(container, script);
        container.appendChild(toolbar);
        container.appendChild(script);

        // 标记为已处理
        script.setAttribute('data-processed', 'true');
    }
}

// 复制消息
function copyMessage(messageId) {
    const message = document.getElementById(messageId);
    if (message) {
        const text = message.querySelector('.markdown-content').innerText;
        copyToClipboard(text);
        showCustomAlert('消息已复制到剪贴板');
    }
}

// 复制代码
function copyCode(encodedCode) {
    if (!encodedCode) {
        showCustomAlert('代码内容为空');
        return;
    }

    let code = '';
    try {
        code = decodeURIComponent(escape(atob(encodedCode)));
    } catch (e) {
        // 兼容旧的编码方式
        try {
            code = atob(encodedCode);
        } catch (e2) {
            console.error('代码解码失败:', e2);
            showCustomAlert('代码解码失败');
            return;
        }
    }

    copyToClipboard(code);
    showCustomAlert('代码已复制到剪贴板');
}

// 复制数学公式
function copyMath(encodedFormula) {
    if (!encodedFormula) {
        showCustomAlert('公式内容为空');
        return;
    }

    let formula = '';
    try {
        formula = decodeURIComponent(escape(atob(encodedFormula)));
    } catch (e) {
        // 兼容旧的编码方式
        try {
            formula = atob(encodedFormula);
        } catch (e2) {
            console.error('公式解码失败:', e2);
            showCustomAlert('公式解码失败');
            return;
        }
    }

    copyToClipboard(formula);
    showCustomAlert('公式已复制到剪贴板');
}

// 复制表格
function copyTable(button) {
    const table = button.closest('.table-container').querySelector('table');
    if (table) {
        let markdownTable = '';

        // 处理表头
        const headers = Array.from(table.querySelectorAll('thead th'));
        if (headers.length > 0) {
            markdownTable += '| ' + headers.map(th => th.innerText.trim()).join(' | ') + ' |\n';
            markdownTable += '| ' + headers.map(() => '---').join(' | ') + ' |\n';
        }

        // 处理表格内容
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cells = Array.from(row.querySelectorAll('td'));
            markdownTable += '| ' + cells.map(cell => cell.innerText.trim()).join(' | ') + ' |\n';
        });

        copyToClipboard(markdownTable);
        showCustomAlert('表格已复制到剪贴板（Markdown格式）');
    }
}

// 插入表格到Word
function insertTableToWord(button) {
    const table = button.closest('.table-container').querySelector('table');
    if (table) {
        const tableHTML = table.outerHTML;
        insertToWord(tableHTML);
    }
}

// 插入代码到Word
function insertCodeToWord(encodedCode, language) {
    if (!encodedCode) {
        showCustomAlert('代码内容为空');
        return;
    }

    let code = '';
    try {
        code = decodeURIComponent(escape(atob(encodedCode)));
    } catch (e) {
        // 兼容旧的编码方式
        try {
            code = atob(encodedCode);
        } catch (e2) {
            console.error('代码解码失败:', e2);
            showCustomAlert('代码解码失败');
            return;
        }
    }

    const codeHTML = `<pre><code class="language-${language}">${escapeHtml(code)}</code></pre>`;
    insertToWord(codeHTML);
}

// 插入数学公式到Word
function insertMathToWord(encodedFormula) {
    if (!encodedFormula) {
        showCustomAlert('公式内容为空');
        return;
    }

    let formula = '';
    try {
        formula = decodeURIComponent(escape(atob(encodedFormula)));
    } catch (e) {
        // 兼容旧的编码方式
        try {
            formula = atob(encodedFormula);
        } catch (e2) {
            console.error('公式解码失败:', e2);
            showCustomAlert('公式解码失败');
            return;
        }
    }

    // 发送公式到Word，让C#处理公式转换
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'copyToWord',
            content: formula,
            format: 'formula'
        });
        showCustomAlert('公式已插入到Word文档');
    } else {
        showCustomAlert('WebView2环境不可用');
    }
}

// 插入消息到Word
function insertMessageToWord(messageId) {
    const message = document.getElementById(messageId);
    if (message) {
        // 针对欢迎消息的特殊处理 - 动态提取内容
        if (messageId === 'welcome-message') {
            console.log('处理欢迎消息的插入');

            // 获取欢迎消息的原始内容
            const markdownContent = message.querySelector('.markdown-content');
            if (!markdownContent) {
                console.log('未找到markdown内容');
                return;
            }

            // 动态解析欢迎消息内容
            const insertItems = parseWelcomeMessageContent(markdownContent);

            // 调试输出解析结果
            console.log('===== 解析结果 =====');
            insertItems.forEach((item, index) => {
                if (item.type === 'text') {
                    console.log(`${index}. 文本: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
                } else if (item.type === 'formula') {
                    console.log(`${index}. 公式: ${item.content}`);
                } else if (item.type === 'table') {
                    console.log(`${index}. 表格: ${JSON.stringify(item.content.headers)}`);
                } else if (item.type === 'code') {
                    console.log(`${index}. 代码: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
                } else if (item.type === 'mermaid') {
                    console.log(`${index}. Mermaid图表: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
                } else if (item.type === 'linebreak') {
                    console.log(`${index}. 换行`);
                }
            });
            console.log('=====================');

            console.log('准备按顺序插入内容：', insertItems.length, '个项目');

            // 处理待生成PNG的Mermaid项目
            const processMermaidPendingItems = async (items) => {
                const processedItems = [];

                for (const item of items) {
                    if (item.type === 'mermaidImagePending') {
                        console.log(`开始生成Mermaid PNG: ${item.content.containerIndex}`);
                        try {
                            // 查找对应的Mermaid容器
                            const mermaidContainers = document.querySelectorAll('.mermaid-container');
                            let targetContainer = null;

                            // 通过索引查找容器
                            if (mermaidContainers[item.content.containerIndex]) {
                                targetContainer = mermaidContainers[item.content.containerIndex];
                                console.log(`通过索引找到Mermaid容器: ${item.content.containerIndex}`);
                            }

                            // 如果索引查找失败，尝试通过代码匹配
                            if (!targetContainer) {
                                for (let container of mermaidContainers) {
                                    const mermaidDiv = container.querySelector('.mermaid');
                                    if (mermaidDiv) {
                                        const containerCode = mermaidDiv.getAttribute('data-mermaid-code');
                                        if (containerCode && containerCode.trim() === item.content.mermaidCode.trim()) {
                                            targetContainer = container;
                                            console.log('通过data-mermaid-code属性找到Mermaid容器');
                                            break;
                                        }

                                        if (mermaidDiv.textContent.trim() === item.content.mermaidCode.trim()) {
                                            targetContainer = container;
                                            console.log('通过文本内容找到Mermaid容器');
                                            break;
                                        }
                                    }
                                }
                            }

                            // 如果仍未找到，尝试部分匹配
                            if (!targetContainer) {
                                for (let container of mermaidContainers) {
                                    const mermaidDiv = container.querySelector('.mermaid');
                                    if (mermaidDiv) {
                                        // 尝试去除空白后比较
                                        const cleanCode = item.content.mermaidCode.replace(/\s+/g, '');
                                        const cleanDivCode = mermaidDiv.textContent.replace(/\s+/g, '');
                                        const cleanAttrCode = mermaidDiv.getAttribute('data-mermaid-code');
                                        const cleanAttr = cleanAttrCode ? cleanAttrCode.replace(/\s+/g, '') : '';

                                        if (cleanCode === cleanDivCode || cleanCode === cleanAttr) {
                                            targetContainer = container;
                                            console.log('通过清理空白后比较找到Mermaid容器');
                                            break;
                                        }

                                        // 尝试部分匹配（如果代码长度超过50个字符）
                                        if (cleanCode.length > 50) {
                                            const codeStart = cleanCode.substring(0, 50);
                                            if ((cleanDivCode && cleanDivCode.includes(codeStart)) ||
                                                (cleanAttr && cleanAttr.includes(codeStart))) {
                                                targetContainer = container;
                                                console.log('通过部分内容匹配找到Mermaid容器');
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (targetContainer) {
                                const mermaidDiv = targetContainer.querySelector('.mermaid');
                                const svgElement = mermaidDiv ? mermaidDiv.querySelector('svg') : null;

                                if (svgElement) {
                                    // 生成PNG数据
                                    console.log('找到SVG元素，开始生成PNG...');
                                    const pngData = await generatePNGFromSVG(svgElement);

                                    if (pngData) {
                                        // 成功生成PNG，添加到处理后的项目中
                                        processedItems.push({
                                            type: 'mermaidImage',
                                            content: {
                                                imageData: pngData,
                                                mermaidCode: item.content.mermaidCode,
                                                width: item.content.svgWidth,
                                                height: item.content.svgHeight
                                            }
                                        });
                                        console.log(`Mermaid PNG生成成功: ${item.content.containerIndex}`);
                                        continue;
                                    } else {
                                        console.warn('PNG数据生成失败，回退到代码块');
                                    }
                                } else {
                                    console.warn('找到容器但未找到SVG元素，回退到代码块');
                                }
                            } else {
                                console.warn(`未找到匹配的Mermaid容器，回退到代码块: ${item.content.containerIndex}`);
                            }

                            // PNG生成失败，回退到代码块
                            processedItems.push({
                                type: 'mermaid',
                                content: item.content.mermaidCode
                            });

                        } catch (error) {
                            console.error(`处理Mermaid PNG时出错:`, error);
                            // 出错时回退到代码块
                            processedItems.push({
                                type: 'mermaid',
                                content: item.content.mermaidCode
                            });
                        }
                    } else {
                        // 非Mermaid项目直接添加
                        processedItems.push(item);
                    }
                }

                return processedItems;
            };

            if (window.chrome && window.chrome.webview) {
                // 检查是否有待处理的Mermaid PNG项目
                const hasPendingMermaid = insertItems.some(item => item.type === 'mermaidImagePending');

                if (hasPendingMermaid) {
                    console.log('检测到待生成PNG的Mermaid项目，开始异步处理...');

                    // 异步处理Mermaid PNG生成
                    processMermaidPendingItems(insertItems).then(processedItems => {
                        console.log(`PNG处理完成，最终发送 ${processedItems.length} 个项目`);

                        // 发送处理后的插入序列给C#
                        window.chrome.webview.postMessage({
                            type: 'insertSequence',
                            items: processedItems
                        });

                        showCustomAlert('欢迎内容已插入到Word文档');
                    }).catch(error => {
                        console.error('处理Mermaid PNG时出错:', error);

                        // 出错时发送原始序列（Mermaid会作为代码块处理）
                        const fallbackItems = insertItems.map(item => {
                            if (item.type === 'mermaidImagePending') {
                                return {
                                    type: 'mermaid',
                                    content: item.content.mermaidCode
                                };
                            }
                            return item;
                        });

                        window.chrome.webview.postMessage({
                            type: 'insertSequence',
                            items: fallbackItems
                        });

                        showCustomAlert('欢迎内容已插入到Word文档（部分Mermaid以代码形式插入）');
                    });
                } else {
                    // 没有待处理的Mermaid，直接发送
                    window.chrome.webview.postMessage({
                        type: 'insertSequence',
                        items: insertItems
                    });

                    showCustomAlert('欢迎内容已插入到Word文档');
                }
            } else {
                showCustomAlert('WebView2环境不可用');
            }
            return;
        }

        // ===== 修复：普通AI回复消息也使用智能解析 =====
        console.log('处理普通AI回复消息的插入');

        const markdownContent = message.querySelector('.markdown-content');
        if (!markdownContent) {
            console.log('未找到markdown内容');
            return;
        }

        // 使用相同的智能解析函数
        const insertItems = parseWelcomeMessageContent(markdownContent);

        // 调试输出解析结果
        console.log('===== AI回复解析结果 =====');
        insertItems.forEach((item, index) => {
            if (item.type === 'text') {
                console.log(`${index}. 文本: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
            } else if (item.type === 'formula') {
                console.log(`${index}. 公式: ${item.content}`);
            } else if (item.type === 'table') {
                console.log(`${index}. 表格: ${JSON.stringify(item.content.headers)}`);
            } else if (item.type === 'code') {
                console.log(`${index}. 代码: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
            } else if (item.type === 'mermaid') {
                console.log(`${index}. Mermaid图表: ${item.content.substring(0, 30)}${item.content.length > 30 ? '...' : ''}`);
            } else if (item.type === 'linebreak') {
                console.log(`${index}. 换行`);
            }
        });
        console.log('=============================');

        console.log('准备按顺序插入内容：', insertItems.length, '个项目');

        if (window.chrome && window.chrome.webview) {
            // 修改：对所有Mermaid图表尝试优先使用PNG格式
            const hasMermaid = insertItems.some(item => item.type === 'mermaid');

            if (hasMermaid) {
                console.log('检测到Mermaid图表，尝试转换为PNG格式...');

                // 转换Mermaid为PNG待处理项
                const itemsWithPendingMermaid = insertItems.map(item => {
                    if (item.type === 'mermaid') {
                        // 查找对应的已渲染Mermaid容器
                        const mermaidContainers = document.querySelectorAll('.mermaid-container');
                        let targetContainer = null;

                        // 改进：通过多种方式查找Mermaid容器
                        // 1. 首先尝试通过data-mermaid-code属性匹配
                        for (let container of mermaidContainers) {
                            const mermaidDiv = container.querySelector('.mermaid');
                            if (mermaidDiv) {
                                const containerCode = mermaidDiv.getAttribute('data-mermaid-code');
                                if (containerCode && containerCode.trim() === item.content.trim()) {
                                    targetContainer = container;
                                    console.log('找到匹配的Mermaid容器，通过data-mermaid-code属性');
                                    break;
                                }

                                // 如果没有data-mermaid-code属性，尝试比较文本内容
                                if (mermaidDiv.textContent.trim() === item.content.trim()) {
                                    targetContainer = container;
                                    console.log('找到匹配的Mermaid容器，通过文本内容');
                                    break;
                                }
                            }
                        }

                        // 2. 如果仍未找到，尝试遍历所有容器进行部分匹配
                        if (!targetContainer) {
                            for (let container of mermaidContainers) {
                                const mermaidDiv = container.querySelector('.mermaid');
                                if (mermaidDiv) {
                                    // 尝试去除空白后比较
                                    const cleanCode = item.content.replace(/\s+/g, '');
                                    const cleanDivCode = mermaidDiv.textContent.replace(/\s+/g, '');
                                    const cleanAttrCode = mermaidDiv.getAttribute('data-mermaid-code');
                                    const cleanAttr = cleanAttrCode ? cleanAttrCode.replace(/\s+/g, '') : '';

                                    if (cleanCode === cleanDivCode || cleanCode === cleanAttr) {
                                        targetContainer = container;
                                        console.log('找到匹配的Mermaid容器，通过清理空白后比较');
                                        break;
                                    }

                                    // 尝试部分匹配（如果代码长度超过50个字符）
                                    if (cleanCode.length > 50) {
                                        const codeStart = cleanCode.substring(0, 50);
                                        if ((cleanDivCode && cleanDivCode.includes(codeStart)) ||
                                            (cleanAttr && cleanAttr.includes(codeStart))) {
                                            targetContainer = container;
                                            console.log('找到匹配的Mermaid容器，通过部分内容匹配');
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        if (targetContainer) {
                            const mermaidDiv = targetContainer.querySelector('.mermaid');
                            const svgElement = mermaidDiv ? mermaidDiv.querySelector('svg') : null;

                            if (svgElement) {
                                // 获取SVG尺寸
                                const svgRect = svgElement.getBoundingClientRect();
                                const svgWidth = svgRect.width || 400;
                                const svgHeight = svgRect.height || 300;

                                // 标记为待处理PNG
                                return {
                                    type: 'mermaidImagePending',
                                    content: {
                                        containerIndex: Array.from(mermaidContainers).indexOf(targetContainer),
                                        mermaidCode: item.content,
                                        svgWidth: svgWidth,
                                        svgHeight: svgHeight
                                    }
                                };
                            }
                        }
                        // 如果没找到容器或SVG，保持原样
                        console.log('未找到匹配的Mermaid容器，将使用代码块格式');
                        return item;
                    }
                    return item;
                });

                // 异步处理所有待生成的PNG
                const processMermaidPendingItems = async (items) => {
                    const processedItems = [];

                    for (const item of items) {
                        if (item.type === 'mermaidImagePending') {
                            console.log(`开始生成Mermaid PNG: ${item.content.containerIndex}`);
                            try {
                                // 查找对应的Mermaid容器
                                const mermaidContainers = document.querySelectorAll('.mermaid-container');
                                let targetContainer = null;

                                // 通过索引查找容器
                                if (mermaidContainers[item.content.containerIndex]) {
                                    targetContainer = mermaidContainers[item.content.containerIndex];
                                } else {
                                    // 如果索引查找失败，尝试通过代码匹配
                                    for (let container of mermaidContainers) {
                                        const mermaidDiv = container.querySelector('.mermaid');
                                        if (mermaidDiv) {
                                            const containerCode = mermaidDiv.getAttribute('data-mermaid-code');
                                            if (containerCode && containerCode.trim() === item.content.mermaidCode.trim()) {
                                                targetContainer = container;
                                                break;
                                            }

                                            if (mermaidDiv.textContent.trim() === item.content.mermaidCode.trim()) {
                                                targetContainer = container;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (targetContainer) {
                                    const mermaidDiv = targetContainer.querySelector('.mermaid');
                                    const svgElement = mermaidDiv ? mermaidDiv.querySelector('svg') : null;

                                    if (svgElement) {
                                        // 生成PNG数据
                                        const pngData = await generatePNGFromSVG(svgElement);

                                        if (pngData) {
                                            // 成功生成PNG，添加到处理后的项目中
                                            processedItems.push({
                                                type: 'mermaidImage',
                                                content: {
                                                    imageData: pngData,
                                                    mermaidCode: item.content.mermaidCode,
                                                    width: item.content.svgWidth,
                                                    height: item.content.svgHeight
                                                }
                                            });
                                            console.log(`Mermaid PNG生成成功: ${item.content.containerIndex}`);
                                            continue;
                                        }
                                    }
                                }

                                // PNG生成失败，回退到代码块
                                console.warn(`Mermaid PNG生成失败，回退到代码块: ${item.content.containerIndex}`);
                                processedItems.push({
                                    type: 'mermaid',
                                    content: item.content.mermaidCode
                                });

                            } catch (error) {
                                console.error(`处理Mermaid PNG时出错:`, error);
                                // 出错时回退到代码块
                                processedItems.push({
                                    type: 'mermaid',
                                    content: item.content.mermaidCode
                                });
                            }
                        } else {
                            // 非Mermaid项目直接添加
                            processedItems.push(item);
                        }
                    }

                    return processedItems;
                };

                // 异步处理Mermaid PNG生成
                processMermaidPendingItems(itemsWithPendingMermaid).then(processedItems => {
                    console.log(`PNG处理完成，最终发送 ${processedItems.length} 个项目`);

                    // 发送处理后的插入序列给C#
                    window.chrome.webview.postMessage({
                        type: 'insertSequence',
                        items: processedItems
                    });

                    showCustomAlert('AI回复内容已插入到Word文档');
                }).catch(error => {
                    console.error('处理Mermaid PNG时出错:', error);

                    // 出错时发送原始序列（Mermaid会作为代码块处理）
                    const fallbackItems = itemsWithPendingMermaid.map(item => {
                        if (item.type === 'mermaidImagePending') {
                            return {
                                type: 'mermaid',
                                content: item.content.mermaidCode
                            };
                        }
                        return item;
                    });

                    window.chrome.webview.postMessage({
                        type: 'insertSequence',
                        items: fallbackItems
                    });

                    showCustomAlert('AI回复内容已插入到Word文档（部分Mermaid以代码形式插入）');
                });
            } else {
                // 没有Mermaid图表，直接发送
                window.chrome.webview.postMessage({
                    type: 'insertSequence',
                    items: insertItems
                });

                showCustomAlert('AI回复内容已插入到Word文档');
            }
        } else {
            showCustomAlert('WebView2环境不可用');
        }
    }
}

// 解析欢迎消息内容的新函数 - 修复版本，解决公式、表格、代码块为空的问题
function parseWelcomeMessageContent(markdownContent) {
    const insertItems = [];

    // 克隆内容以避免修改原始DOM
    const contentClone = markdownContent.cloneNode(true);

    // 预处理：提取所有公式，并标记其位置
    const formulas = [];

    // 处理所有的equation-container
    const equationContainers = contentClone.querySelectorAll('.equation-container');
    equationContainers.forEach(container => {
        const script = container.querySelector('script[type^="math/tex"]');
        if (script) {
            const formula = script.textContent.trim();
            if (formula) {
                formulas.push(formula);
                console.log('从equation-container提取公式:', formula);

                // 用特殊标记替换整个公式容器
                const marker = document.createElement('span');
                marker.className = 'formula-marker';
                marker.dataset.formulaIndex = formulas.length - 1;
                marker.textContent = `[FORMULA_${formulas.length - 1}]`; // 添加可见文本
                container.parentNode.replaceChild(marker, container);
            }
        }
    });

    // 处理剩余的独立script元素
    const independentScripts = contentClone.querySelectorAll('script[type^="math/tex"]');
    independentScripts.forEach(script => {
        if (script.parentNode && contentClone.contains(script)) {
            const formula = script.textContent.trim();
            if (formula) {
                formulas.push(formula);
                console.log('从独立script提取公式:', formula);

                const marker = document.createElement('span');
                marker.className = 'formula-marker';
                marker.dataset.formulaIndex = formulas.length - 1;
                marker.textContent = `[FORMULA_${formulas.length - 1}]`; // 添加可见文本
                script.parentNode.replaceChild(marker, script);
            }
        }
    });

    // 清除所有MathJax相关的元素
    const mathJaxElements = contentClone.querySelectorAll('.MathJax, .MathJax_Display, .MathJax_Preview, [id^="MathJax"], [class*="MathJax"]');
    console.log(`发现 ${mathJaxElements.length} 个MathJax元素，正在移除`);
    mathJaxElements.forEach(element => element.remove());

    const remainingMathScripts = contentClone.querySelectorAll('script[type^="math/"]');
    console.log(`发现 ${remainingMathScripts.length} 个剩余的数学脚本，正在移除`);
    remainingMathScripts.forEach(script => script.remove());

    // 预处理：提取所有表格
    const tables = [];
    const tableElements = contentClone.querySelectorAll('table');
    tableElements.forEach(table => {
        const tableData = extractTableData(table);
        if (tableData) {
            tables.push(tableData);
            console.log('提取表格数据:', tableData);

            const marker = document.createElement('span');
            marker.className = 'table-marker';
            marker.dataset.tableIndex = tables.length - 1;
            marker.textContent = `[TABLE_${tables.length - 1}]`; // 添加可见文本

            const container = table.closest('.table-container');
            if (container) {
                container.parentNode.replaceChild(marker, container);
            } else {
                table.parentNode.replaceChild(marker, table);
            }
        }
    });

    // 预处理：提取所有代码块和Mermaid图表
    const codeBlocks = [];
    const mermaidDiagrams = [];

    // 首先查找已经渲染的Mermaid容器
    const mermaidContainers = contentClone.querySelectorAll('.mermaid-container');
    mermaidContainers.forEach((container, index) => {
        const mermaidDiv = container.querySelector('.mermaid');
        if (mermaidDiv) {
            // 从数据属性中获取原始Mermaid代码
            const code = mermaidDiv.getAttribute('data-mermaid-code');
            if (code && code.trim()) {
                mermaidDiagrams.push({
                    code: code.trim(),
                    language: 'mermaid'
                });
                console.log('提取已渲染的Mermaid图表原始代码:', code.substring(0, 50) + '...');

                const marker = document.createElement('span');
                marker.className = 'mermaid-marker';
                marker.dataset.mermaidIndex = mermaidDiagrams.length - 1;
                marker.textContent = `[MERMAID_${mermaidDiagrams.length - 1}]`;
                container.parentNode.replaceChild(marker, container);
            } else {
                console.warn('无法从Mermaid容器中获取原始代码');
            }
        }
    });

    // 然后查找未渲染的代码块（包括可能的Mermaid代码块）
    const preElements = contentClone.querySelectorAll('pre');
    preElements.forEach(pre => {
        const codeElement = pre.querySelector('code');
        if (codeElement) {
            const code = codeElement.textContent.trim();
            if (code) {
                const language = (codeElement.className.match(/language-(\w+)/) || ['', 'text'])[1];

                // 检查是否是Mermaid图表
                if (language === 'mermaid') {
                    mermaidDiagrams.push({
                        code: code,
                        language: language
                    });
                    console.log('提取未渲染的Mermaid图表:', code.substring(0, 50) + '...');

                    const marker = document.createElement('span');
                    marker.className = 'mermaid-marker';
                    marker.dataset.mermaidIndex = mermaidDiagrams.length - 1;
                    marker.textContent = `[MERMAID_${mermaidDiagrams.length - 1}]`;
                    pre.parentNode.replaceChild(marker, pre);
                } else {
                    codeBlocks.push({
                        code: code,
                        language: language
                    });
                    console.log('提取代码块:', code.substring(0, 50) + '...');

                    const marker = document.createElement('span');
                    marker.className = 'code-marker';
                    marker.dataset.codeIndex = codeBlocks.length - 1;
                    marker.textContent = `[CODE_${codeBlocks.length - 1}]`;
                    pre.parentNode.replaceChild(marker, pre);
                }
            }
        }
    });

    console.log(`提取完成 - 公式: ${formulas.length}, 表格: ${tables.length}, 代码块: ${codeBlocks.length}, Mermaid图表: ${mermaidDiagrams.length}`);

    // 使用递归的方式处理所有节点
    function processNodeRecursively(node) {
        const items = [];

        if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
                items.push({
                    type: 'text',
                    content: text
                });
            }
        }
        else if (node.nodeType === Node.ELEMENT_NODE) {
            // 跳过工具栏和其他UI元素
            if (node.classList && (
                node.classList.contains('mermaid-toolbar') ||
                node.classList.contains('code-toolbar') ||
                node.classList.contains('table-toolbar') ||
                node.classList.contains('math-toolbar') ||
                node.classList.contains('equation-container')
            )) {
                return items; // 跳过工具栏
            }

            // 检查是否是特殊标记
            if (node.classList && node.classList.contains('formula-marker')) {
                const formulaIndex = parseInt(node.dataset.formulaIndex);
                if (!isNaN(formulaIndex) && formulaIndex < formulas.length) {
                    items.push({
                        type: 'formula',
                        content: formulas[formulaIndex]
                    });
                    console.log(`处理公式标记 ${formulaIndex}: ${formulas[formulaIndex]}`);
                }
            }
            else if (node.classList && node.classList.contains('table-marker')) {
                const tableIndex = parseInt(node.dataset.tableIndex);
                if (!isNaN(tableIndex) && tableIndex < tables.length) {
                    items.push({
                        type: 'table',
                        content: tables[tableIndex]
                    });
                    console.log(`处理表格标记 ${tableIndex}`);
                }
            }
            else if (node.classList && node.classList.contains('code-marker')) {
                const codeIndex = parseInt(node.dataset.codeIndex);
                if (!isNaN(codeIndex) && codeIndex < codeBlocks.length) {
                    items.push({
                        type: 'code',
                        content: codeBlocks[codeIndex].code
                    });
                    console.log(`处理代码标记 ${codeIndex}: ${codeBlocks[codeIndex].code.substring(0, 30)}...`);
                }
            }
            else if (node.classList && node.classList.contains('mermaid-marker')) {
                const mermaidIndex = parseInt(node.dataset.mermaidIndex);
                if (!isNaN(mermaidIndex) && mermaidIndex < mermaidDiagrams.length) {
                    // 查找对应的渲染后的Mermaid图表
                    const mermaidContainers = document.querySelectorAll('.mermaid-container');
                    let mermaidContainer = null;

                    // 修复：改进查找Mermaid容器的逻辑
                    // 1. 首先尝试通过代码内容匹配查找
                    const mermaidCode = mermaidDiagrams[mermaidIndex].code;
                    for (let container of mermaidContainers) {
                        const mermaidContent = container.querySelector('.mermaid');
                        if (mermaidContent) {
                            // 检查data-mermaid-code属性
                            const containerCode = mermaidContent.getAttribute('data-mermaid-code');
                            if (containerCode && containerCode.trim() === mermaidCode.trim()) {
                                mermaidContainer = container;
                                console.log(`找到匹配的Mermaid容器，通过data-mermaid-code属性`);
                                break;
                            }

                            // 如果没有data-mermaid-code属性，尝试比较文本内容
                            if (mermaidContent.textContent.trim() === mermaidCode.trim()) {
                                mermaidContainer = container;
                                console.log(`找到匹配的Mermaid容器，通过文本内容`);
                                break;
                            }
                        }
                    }

                    // 2. 如果仍未找到，尝试通过索引查找（不太可靠但作为备选）
                    if (!mermaidContainer && mermaidContainers.length > mermaidIndex) {
                        mermaidContainer = mermaidContainers[mermaidIndex];
                        console.log(`通过索引找到Mermaid容器: ${mermaidIndex}`);
                    }

                    // 3. 如果找到容器，尝试从DOM中查找SVG元素
                    if (mermaidContainer) {
                        const mermaidDiv = mermaidContainer.querySelector('.mermaid');
                        const svgElement = mermaidDiv ? mermaidDiv.querySelector('svg') : null;

                        if (svgElement) {
                            // 找到了渲染的SVG，同步生成PNG
                            try {
                                // 获取SVG的尺寸
                                const svgRect = svgElement.getBoundingClientRect();
                                const svgWidth = svgRect.width || 400;
                                const svgHeight = svgRect.height || 300;

                                // 标记为需要PNG，但在发送前异步生成
                                items.push({
                                    type: 'mermaidImagePending',
                                    content: {
                                        containerIndex: mermaidIndex,
                                        mermaidCode: mermaidDiagrams[mermaidIndex].code,
                                        svgWidth: svgWidth,
                                        svgHeight: svgHeight
                                    }
                                });
                                console.log(`处理Mermaid标记 ${mermaidIndex}: 标记为PNG待生成`);
                            } catch (error) {
                                console.warn(`Mermaid PNG准备失败，使用代码块:`, error);
                                items.push({
                                    type: 'mermaid',
                                    content: mermaidDiagrams[mermaidIndex].code
                                });
                            }
                        } else {
                            // 没有找到SVG，直接使用代码块格式
                            console.log(`找到Mermaid容器但未找到SVG，使用代码块格式: ${mermaidIndex}`);
                            items.push({
                                type: 'mermaid',
                                content: mermaidDiagrams[mermaidIndex].code
                            });
                        }
                    } else {
                        // 没有找到容器，直接使用代码块格式
                        console.log(`未找到Mermaid容器 ${mermaidIndex}，使用代码块格式`);
                        items.push({
                            type: 'mermaid',
                            content: mermaidDiagrams[mermaidIndex].code
                        });
                    }
                }
            }
            // 处理普通HTML元素
            else {
                const tagName = node.tagName ? node.tagName.toLowerCase() : '';

                // 对于段落、标题、列表等，保持HTML格式
                if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'blockquote'].includes(tagName)) {
                    // 检查是否包含特殊标记
                    const hasSpecialMarkers = node.querySelector('.formula-marker, .table-marker, .code-marker, .mermaid-marker');

                    if (hasSpecialMarkers) {
                        // 递归处理子节点
                        Array.from(node.childNodes).forEach(child => {
                            items.push(...processNodeRecursively(child));
                        });
                    } else {
                        // 没有特殊标记，保持HTML格式
                        const htmlContent = node.outerHTML;
                        if (htmlContent.trim()) {
                            items.push({
                                type: 'html',
                                content: htmlContent
                            });
                        }
                    }
                }
                // 对于内联元素，递归处理
                else if (['span', 'strong', 'em', 'b', 'i', 'a', 'code'].includes(tagName)) {
                    // 检查是否包含特殊标记
                    const hasSpecialMarkers = node.querySelector('.formula-marker, .table-marker, .code-marker, .mermaid-marker');

                    if (hasSpecialMarkers) {
                        // 递归处理子节点
                        Array.from(node.childNodes).forEach(child => {
                            items.push(...processNodeRecursively(child));
                        });
                    } else {
                        // 保持HTML格式
                        const htmlContent = node.outerHTML;
                        if (htmlContent.trim()) {
                            items.push({
                                type: 'html',
                                content: htmlContent
                            });
                        }
                    }
                }
                // 对于其他元素，递归处理子节点
                else {
                    Array.from(node.childNodes).forEach(child => {
                        items.push(...processNodeRecursively(child));
                    });
                }
            }
        }

        return items;
    }

    // 处理所有子节点
    Array.from(contentClone.childNodes).forEach(node => {
        insertItems.push(...processNodeRecursively(node));
    });

    // 合并相邻的相同类型项目
    const mergedItems = [];
    let lastItem = null;

    for (const item of insertItems) {
        if (lastItem && lastItem.type === 'text' && item.type === 'text') {
            // 合并相邻的文本项
            lastItem.content += ' ' + item.content;
        } else if (lastItem && lastItem.type === 'html' && item.type === 'html') {
            // 合并相邻的HTML项
            lastItem.content += item.content;
        } else {
            mergedItems.push(item);
            lastItem = item;
        }
    }

    console.log('解析完成，总项目数:', mergedItems.length);
    mergedItems.forEach((item, index) => {
        let preview = 'N/A';
        if (item.content) {
            if (typeof item.content === 'string') {
                preview = item.content.substring(0, 50);
            } else if (typeof item.content === 'object') {
                // 对于表格等对象类型，显示对象信息
                if (item.type === 'table') {
                    preview = `表格 ${item.content.rows?.length || 0} 行 ${item.content.headers?.length || 0} 列`;
                } else {
                    preview = JSON.stringify(item.content).substring(0, 50);
                }
            }
        }
        console.log(`${index}. ${item.type}: ${preview}...`);
    });

    return mergedItems;
}

// 提取表格数据的辅助函数
function extractTableData(tableElement) {
    const headers = [];
    const rows = [];

    // 提取表头
    const headerCells = tableElement.querySelectorAll('thead th, tr:first-child th');
    headerCells.forEach(cell => {
        headers.push(cell.textContent.trim());
    });

    // 提取数据行
    const dataRows = tableElement.querySelectorAll('tbody tr, tr:not(:first-child)');
    dataRows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('td, th');
        cells.forEach(cell => {
            rowData.push(cell.textContent.trim());
        });
        if (rowData.length > 0) {
            rows.push(rowData);
        }
    });

    // 如果没有明确的表头，使用第一行作为表头
    if (headers.length === 0 && rows.length > 0) {
        headers.push(...rows[0]);
        rows.shift();
    }

    return headers.length > 0 ? { headers, rows } : null;
}

// 插入到Word
function insertToWord(content) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'copyToWord',
            content: content,
            format: 'html'
        });
        showCustomAlert('内容已插入到Word文档');
    } else {
        showCustomAlert('WebView2环境不可用');
    }
}

// 复制到剪贴板
function copyToClipboard(text) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(() => {
            console.log('复制成功');
        }).catch(err => {
            console.error('复制失败:', err);
        });
    } else {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
    }
}

// HTML转义
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 自动调整输入框高度
function autoResizeInput() {
    if (messageInput) {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.min(messageInput.scrollHeight, 120) + 'px';
    }
}

// 更新字符计数
function updateCharacterCount() {
    if (charCount && messageInput) {
        const count = messageInput.value.length;
        charCount.textContent = count;

        if (count > 3500) {
            charCount.style.color = '#dc2626';
        } else if (count > 3000) {
            charCount.style.color = '#f59e0b';
        } else {
            charCount.style.color = '#6b7280';
        }
    }
}

// 滚动到底部
function scrollToBottom() {
    if (chatMessages) {
        chatMessages.scrollTop = chatMessages.scrollHeight;
        isAtBottom = true;
    }
}

// 处理滚动事件
function handleScroll() {
    if (chatMessages) {
        const isNearBottom = chatMessages.scrollTop + chatMessages.clientHeight >= chatMessages.scrollHeight - 10;
        isAtBottom = isNearBottom;
    }
}

// 键盘快捷键
document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
        e.preventDefault();
        if (messageInput) {
            messageInput.focus();
        }
    }

    if (e.key === 'Escape' && messageInput === document.activeElement) {
        messageInput.value = '';
        updateCharacterCount();
        autoResizeInput();
    }
});

// 模拟响应（测试用）
function simulateResponse(userMessage) {
    const responses = [
        `你说的"${userMessage}"很有趣！这是一个测试回复。\n\n这里有一些**粗体文字**和*斜体文字*。\n\n代码示例：\n\`\`\`javascript\nconsole.log("Hello World!");\n\`\`\`\n\n数学公式：$E = mc^2$\n\n表格示例：\n| 列1 | 列2 |\n|-----|-----|\n| 数据1 | 数据2 |\n| 数据3 | 数据4 |`,
        `关于"${userMessage}"，我可以提供以下信息：\n\n1. 这是第一点\n2. 这是第二点\n3. 这是第三点\n\n公式：$$\\frac{a}{b} = c$$`,
        `谢谢你输入"${userMessage}"！这是一个简单的回复。`
    ];

    // 根据模式添加不同的响应内容
    let response = responses[Math.floor(Math.random() * responses.length)];
    if (currentChatMode === 'chat-agent') {
        response += `\n\n**(智能体模式)** 我正在使用Agent功能为您提供更智能的回复。`;
    }

    // 直接使用完整的消息生成流程
    startGeneratingOutline();

    // 模拟流式响应
    let index = 0;
    const chunks = response.split(' ');

    const addChunk = () => {
        if (index < chunks.length) {
            const chunk = (index > 0 ? ' ' : '') + chunks[index];
            appendOutlineContent(chunk);
            index++;
            setTimeout(addChunk, 100);
        } else {
            finishGeneratingOutline();
        }
    };

    addChunk();
}

// 清空对话历史
function clearConversationHistory() {
    // 显示确认对话框
    showCustomAlert('确定要清空对话历史吗？这将删除所有聊天记录，操作不可撤销。', function () {
        // 用户确认后执行清空操作

        // 清空前端显示的消息（除了欢迎消息）
        const chatMessages = document.getElementById('chat-messages');
        const welcomeMessage = chatMessages.querySelector('.message.assistant-message');

        // 保存欢迎消息
        const welcomeHTML = welcomeMessage ? welcomeMessage.outerHTML : '';

        // 清空所有消息
        chatMessages.innerHTML = '';

        // 重新添加欢迎消息
        if (welcomeHTML) {
            chatMessages.innerHTML = welcomeHTML;
        }

        // 重置前端会话相关状态（防止“操作记录”残留到下一轮）
        try {
            // 1) 清空预览决策与操作记录
            previewDecisionLogs = [];
            operationDecisionState = {};

            // 2) 清空预览管理器与其DOM
            if (previewManager && previewManager.previews) {
                try {
                    // 移除预览DOM
                    document.querySelectorAll('.modern-inline-preview').forEach(el => el.remove());
                } catch (_) { }
                previewManager.previews.clear();
                previewManager.counter = 0;
            }
            isPreviewPending = false;
            hideFloatingBatchActions && hideFloatingBatchActions();

            console.log('已重置操作记录与预览状态（previewDecisionLogs/operationDecisionState/previewManager）');
        } catch (e) {
            console.warn('重置本地会话状态时出错:', e);
        }

        // 通知C#后端清空对话历史
        try {
            window.chrome.webview.postMessage({
                type: 'clearHistory',
                message: 'clear conversation history'
            });

            console.log('已发送清空对话历史请求到C#后端');
        } catch (error) {
            console.error('发送清空历史请求时出错:', error);
        }

        // 显示成功提示
        setTimeout(() => {
            console.log('对话历史已清空');
        }, 500);
    });
}

// ==================== 工具设置相关功能 ====================

// Agent配置选项的默认值
let agentConfig = {
    enableRangeNormalization: true,
    defaultInsertPosition: 'end',
    enablePostFeedback: false, // 默认不自动发送插入结果反馈
    showToolCalls: true,
    showDebugInfo: false
};

// 初始化工具设置
function initializeToolsSettings() {
    // 从localStorage加载保存的工具设置
    const savedTools = localStorage.getItem('enabledTools');
    const savedConfig = localStorage.getItem('agentConfig');

    console.log('=== 初始化工具设置调试 ===');
    console.log('localStorage中的工具设置:', savedTools);
    console.log('localStorage中的Agent配置:', savedConfig);

    if (savedTools) {
        try {
            enabledTools = JSON.parse(savedTools);
            console.log('成功加载保存的工具设置:', enabledTools);
        } catch (e) {
            console.error('解析保存的工具设置失败:', e);
            resetToDefaultTools();
        }
    } else {
        console.log('localStorage中没有工具设置，使用默认设置');
        resetToDefaultTools();
    }

    // 加载Agent配置
    if (savedConfig) {
        try {
            agentConfig = { ...agentConfig, ...JSON.parse(savedConfig) };
            console.log('成功加载保存的Agent配置:', agentConfig);
        } catch (e) {
            console.error('解析保存的Agent配置失败:', e);
        }
    }

    // 统一策略：强制关闭“自动发送插入结果反馈”（该功能已下线）
    if (agentConfig.enablePostFeedback !== false) {
        agentConfig.enablePostFeedback = false;
        try { localStorage.setItem('agentConfig', JSON.stringify(agentConfig)); } catch (e) { }
    }

    // 更新UI控件状态
    updateConfigUI();

    // 应用工具进度显示设置
    updateToolProgressVisibility();

    console.log('最终的enabledTools:', enabledTools);
    console.log('最终的agentConfig:', agentConfig);
    console.log('=== 初始化调试结束 ===');

    // 监听聊天模式变化
    if (chatModeSelect) {
        chatModeSelect.addEventListener('change', function () {
            // 检查是否尝试切换到智能体模式但未解锁
            if (this.value === 'chat-agent' && !isAgentModeUnlocked) {
                // 重置为智能问答模式
                this.value = 'chat';
                // 弹出解锁提示
                unlockAgentMode();
                return;
            }

            // 如果切换到智能体模式，先检查当前模型是否支持工具调用
            if (this.value === 'chat-agent') {
                const currentModel = getSelectedModelInfo();
                if (currentModel && currentModel.enableTools !== 1) {
                    // 模型不支持工具调用，阻止切换
                    console.log('模型不支持工具调用，阻止切换到Agent模式');
                    this.value = 'chat'; // 保持在智能问答模式
                    currentChatMode = 'chat';
                    showModelToolsWarning(currentModel.name, false);
                    return;
                }

                // 模型支持工具调用，允许切换
                currentChatMode = this.value;
                updateToolsSettingsVisibility();

                // 立即发送配置到后端
                notifyBackendConfigChange();
            } else {
                // 切换到其他模式（如智能问答）
                currentChatMode = this.value;
                updateToolsSettingsVisibility();
            }
        });
    }

    // 初始化工具设置按钮可见性
    updateToolsSettingsVisibility();
}

// 重置为默认工具设置
function resetToDefaultTools() {
    enabledTools = {};
    defaultTools.forEach(tool => {
        enabledTools[tool] = true;
    });
}

// 处理聊天模式切换
function onChatModeChange() {
    if (chatModeSelect) {
        currentChatMode = chatModeSelect.value;
        updateToolsSettingsVisibility();
    }
}

// 更新配置UI控件状态
function updateConfigUI() {
    // 更新范围清理开关
    const rangeNormalizationCheckbox = document.getElementById('config-enable-range-normalization');
    if (rangeNormalizationCheckbox) {
        rangeNormalizationCheckbox.checked = agentConfig.enableRangeNormalization;
    }

    // 更新默认插入位置选择
    const insertPositionSelect = document.getElementById('config-default-insert-position');
    const postFeedbackCheckbox = document.getElementById('config-enable-post-feedback');
    if (insertPositionSelect) {
        insertPositionSelect.value = agentConfig.defaultInsertPosition;
    }
    if (postFeedbackCheckbox) {
        postFeedbackCheckbox.checked = !!agentConfig.enablePostFeedback;
    }

    // 更新工具进度显示开关
    const toolCallsCheckbox = document.getElementById('config-show-tool-calls');
    if (toolCallsCheckbox) {
        toolCallsCheckbox.checked = agentConfig.showToolCalls;
    }

    const debugInfoCheckbox = document.getElementById('config-show-debug-info');
    if (debugInfoCheckbox) {
        debugInfoCheckbox.checked = agentConfig.showDebugInfo;
    }
}

// 更新工具设置按钮的可见性
function updateToolsSettingsVisibility() {
    if (toolsSettingsBtn) {
        if (currentChatMode === 'chat-agent') {
            toolsSettingsBtn.style.display = 'flex';
        } else {
            toolsSettingsBtn.style.display = 'none';
        }
    }
}

// 打开工具设置模态框
function openToolsSettings() {
    if (toolsSettingsModal) {
        // 更新工具复选框状态
        defaultTools.forEach(toolId => {
            const checkbox = document.getElementById(`tool-${toolId}`);
            if (checkbox) {
                checkbox.checked = enabledTools[toolId] || false;
            }
        });

        // 更新Agent配置UI状态
        updateConfigUI();
        // 隐藏已迁移到预览卡片的选项（清除空格/默认插入位置）
        try {
            const hideContainer = (el) => {
                if (!el) return;
                const container = el.closest('.settings-item') || el.parentElement?.parentElement || el.parentElement;
                if (container) container.style.display = 'none';
            };
            hideContainer(document.getElementById('config-enable-range-normalization'));
            hideContainer(document.getElementById('config-default-insert-position'));
        } catch (e) { console.warn('隐藏旧Agent选项失败:', e); }

        toolsSettingsModal.style.display = 'flex';
        document.body.style.overflow = 'hidden'; // 禁止背景滚动
    }
}

// 关闭工具设置模态框
function closeToolsSettings() {
    if (toolsSettingsModal) {
        toolsSettingsModal.style.display = 'none';
        document.body.style.overflow = 'auto'; // 恢复滚动
    }
}

// 保存工具设置
function saveToolsSettings() {
    // 获取所有工具复选框的状态
    defaultTools.forEach(toolId => {
        const checkbox = document.getElementById(`tool-${toolId}`);
        if (checkbox) {
            enabledTools[toolId] = checkbox.checked;
        }
    });

    // 获取Agent配置选项的状态
    const rangeNormalizationCheckbox = document.getElementById('config-enable-range-normalization');
    const insertPositionSelect = document.getElementById('config-default-insert-position');
    const postFeedbackCheckbox = document.getElementById('config-enable-post-feedback');
    const toolCallsCheckbox = document.getElementById('config-show-tool-calls');
    const debugInfoCheckbox = document.getElementById('config-show-debug-info');

    if (rangeNormalizationCheckbox) {
        agentConfig.enableRangeNormalization = rangeNormalizationCheckbox.checked;
    }
    if (insertPositionSelect) {
        agentConfig.defaultInsertPosition = insertPositionSelect.value;
    }
    if (postFeedbackCheckbox) {
        agentConfig.enablePostFeedback = postFeedbackCheckbox.checked;
    }
    if (toolCallsCheckbox) {
        agentConfig.showToolCalls = toolCallsCheckbox.checked;
    }
    if (debugInfoCheckbox) {
        agentConfig.showDebugInfo = debugInfoCheckbox.checked;
    }

    // 保存到localStorage
    try {
        localStorage.setItem('enabledTools', JSON.stringify(enabledTools));
        localStorage.setItem('agentConfig', JSON.stringify(agentConfig));

        console.log('工具设置已保存:', enabledTools);
        console.log('Agent配置已保存:', agentConfig);

        // 通知后端更新配置
        notifyBackendConfigChange();

        // 显示保存成功提示
        showToolsSettingsSaved();

        // 关闭模态框
        closeToolsSettings();
    } catch (e) {
        console.error('保存设置失败:', e);
        alert('保存设置失败，请重试');
    }
}

// 通知后端配置变更
function notifyBackendConfigChange() {
    try {
        const messageData = {
            type: 'updateAgentConfig',
            config: agentConfig
        };

        sendMessageToCSharp(messageData);
        console.log('已通知后端更新Agent配置:', agentConfig);

        // 立即应用工具进度显示设置
        updateToolProgressVisibility();
    } catch (e) {
        console.error('通知后端配置变更失败:', e);
    }
}

// 更新工具进度容器的可见性
function updateToolProgressVisibility() {
    const progressContainers = document.querySelectorAll('.tool-progress-container');
    progressContainers.forEach(container => {
        if (agentConfig.showToolCalls || agentConfig.showDebugInfo) {
            container.style.display = 'block';
            console.log('显示工具进度容器');
        } else {
            container.style.display = 'none';
            console.log('隐藏工具进度容器');
        }
    });
}

// 显示设置保存成功提示
function showToolsSettingsSaved() {
    // 创建临时提示元素
    const toast = document.createElement('div');
    toast.className = 'toast-message';
    toast.textContent = '工具设置已保存';
    toast.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background-color: #28a745;
        color: white;
        padding: 12px 20px;
        border-radius: 6px;
        font-size: 14px;
        z-index: 1001;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        animation: toastShow 0.3s ease;
    `;

    // 添加动画样式
    const style = document.createElement('style');
    style.textContent = `
        @keyframes toastShow {
            from { opacity: 0; transform: translateX(100%); }
            to { opacity: 1; transform: translateX(0); }
        }
        @keyframes toastHide {
            from { opacity: 1; transform: translateX(0); }
            to { opacity: 0; transform: translateX(100%); }
        }
    `;
    document.head.appendChild(style);

    document.body.appendChild(toast);

    // 3秒后移除提示
    setTimeout(() => {
        toast.style.animation = 'toastHide 0.3s ease forwards';
        setTimeout(() => {
            if (toast.parentNode) {
                toast.parentNode.removeChild(toast);
            }
            if (style.parentNode) {
                style.parentNode.removeChild(style);
            }
        }, 300);
    }, 3000);
}

// 全选工具
function selectAllTools() {
    defaultTools.forEach(toolId => {
        const checkbox = document.getElementById(`tool-${toolId}`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
}

// 全不选工具
function selectNoneTools() {
    defaultTools.forEach(toolId => {
        const checkbox = document.getElementById(`tool-${toolId}`);
        if (checkbox) {
            checkbox.checked = false;
        }
    });
}

// 恢复默认工具选择
function resetDefaultTools() {
    defaultTools.forEach(toolId => {
        const checkbox = document.getElementById(`tool-${toolId}`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
}

// 获取当前启用的工具列表
function getEnabledToolsList() {
    const enabledList = [];
    for (const [toolId, enabled] of Object.entries(enabledTools)) {
        if (enabled) {
            enabledList.push(toolId);
        }
    }
    console.log('当前启用的工具列表:', enabledList);
    console.log('enabledTools对象:', enabledTools);
    return enabledList;
}

// 点击模态框背景关闭
if (toolsSettingsModal) {
    toolsSettingsModal.addEventListener('click', function (e) {
        if (e.target === toolsSettingsModal) {
            closeToolsSettings();
        }
    });
}

// 在页面加载时初始化工具设置和模型选择
document.addEventListener('DOMContentLoaded', function () {
    setTimeout(() => {
        initializeToolsSettings();
        initializeModelSelector();

        // 延迟发送初始配置到后端，确保WebView2已经准备好
        setTimeout(() => {
            if (currentChatMode === 'chat-agent') {
                notifyBackendConfigChange();
                console.log('页面加载完成，已发送初始Agent配置到后端');
            }
        }, 1000);
    }, 100);
});

// 初始化模型选择器
function initializeModelSelector() {
    // 从localStorage加载保存的模型选择
    const savedModelId = localStorage.getItem('selectedModelId');
    if (savedModelId) {
        selectedModelId = parseInt(savedModelId) || 0;
        console.log('加载保存的模型ID:', selectedModelId);
    }

    // 监听模型选择变化，保存到localStorage
    if (modelSelect) {
        modelSelect.addEventListener('change', function () {
            selectedModelId = parseInt(this.value) || 0;
            localStorage.setItem('selectedModelId', selectedModelId.toString());
            console.log('模型选择已保存:', selectedModelId);

            // 检查新选择的模型是否支持工具调用
            checkCurrentModelSupportsTools();
        });
    }
}

// 调试功能：清除工具设置
window.clearToolsSettings = function () {
    localStorage.removeItem('enabledTools');
    console.log('已清除工具设置，请刷新页面');
};

// 调试功能：查看当前工具设置
window.showToolsSettings = function () {
    console.log('当前enabledTools:', enabledTools);
    console.log('localStorage中的设置:', localStorage.getItem('enabledTools'));
    console.log('当前启用的工具列表:', getEnabledToolsList());
};

// ==================== 快捷选择器功能 ====================

// 设置输入框快捷键
function setupInputShortcuts() {
    if (!messageInput) return;

    messageInput.addEventListener('input', function (e) {
        const text = e.target.value;
        const cursorPos = e.target.selectionStart;

        console.log('输入事件 - 文本:', text, '光标位置:', cursorPos);

        // 优先检查是否应该隐藏选择器（这样可以立即响应退格操作）
        if (quickSelector && quickSelector.style.display !== 'none') {
            const triggerIndex = findTriggerIndex(text, cursorPos);
            if (triggerIndex === -1) {
                console.log('未找到触发字符，隐藏选择器并停止获取标题');
                hideQuickSelector();
                // 如果正在获取标题，发送停止请求
                if (isFetchingHeadings) {
                    stopFetchingHeadings();
                }
                return; // 立即返回，不再处理其他逻辑
            }
        }

        // 检查上下文选择器是否应该隐藏
        if (contextSelector && contextSelector.style.display !== 'none') {
            const triggerIndex = findContextTriggerIndex(text, cursorPos);
            if (triggerIndex === -1) {
                console.log('未找到#触发字符，隐藏上下文选择器');
                hideContextSelector();
                return; // 立即返回，不再处理其他逻辑
            }
        }

        // 检查是否输入了 @、/ 或 # 触发快捷选择
        if (text.length > 0 && cursorPos > 0) {
            const charBefore = text[cursorPos - 1];

            if (charBefore === '@' || charBefore === '/' || charBefore === '#') {
                // @ 符号只在 agent 模式下工作
                if (charBefore === '@' && currentChatMode !== 'chat-agent') {
                    console.log('@ 符号只在 Agent 模式下可用');
                    return;
                }

                // 必须在空格后输入才触发（不允许在开头触发）
                const prevChar = cursorPos > 1 ? text[cursorPos - 2] : '';
                if (prevChar === ' ' && cursorPos > 1) {
                    const triggerTime = performance.now();
                    console.log(`⏱️ 触发快捷选择器, 触发时间: ${triggerTime.toFixed(2)}ms, 触发字符: ${charBefore}`);

                    if (charBefore === '#') {
                        // # 符号触发上下文选择器
                        showContextSelector();
                    } else {
                        // @ 和 / 符号触发Word标题选择器
                        showQuickSelector();
                    }
                }
            }
        }
    });

    messageInput.addEventListener('keydown', function (e) {
        // 处理快捷选择器（Word标题）
        if (quickSelector && quickSelector.style.display !== 'none') {
            if (e.key === 'Escape') {
                e.preventDefault();
                hideQuickSelector();
            } else if (e.key === 'ArrowDown') {
                e.preventDefault();
                navigateSelector('down');
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                navigateSelector('up');
            } else if (e.key === 'Enter') {
                e.preventDefault();
                selectCurrentItem();
            }
        }

        // 处理上下文选择器（文档/标题）
        if (contextSelector && contextSelector.style.display !== 'none') {
            if (e.key === 'Escape') {
                e.preventDefault();
                hideContextSelector();
            } else if (e.key === 'ArrowDown') {
                e.preventDefault();
                navigateContextSelector('down');
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                navigateContextSelector('up');
            } else if (e.key === 'Enter') {
                e.preventDefault();
                selectCurrentContextItem();
            }
        }
    });
}

// 查找触发字符的位置
function findTriggerIndex(text, cursorPos) {
    for (let i = cursorPos - 1; i >= 0; i--) {
        const char = text[i];
        if (char === '@' || char === '/') {
            // 必须在空格后（不允许在开头）
            const prevChar = i > 0 ? text[i - 1] : '';
            if (prevChar === ' ' && i > 0) {
                return i;
            }
        } else if (char === ' ') {
            break;
        }
    }
    return -1;
}

// 查找上下文触发字符（#）的位置
function findContextTriggerIndex(text, cursorPos) {
    for (let i = cursorPos - 1; i >= 0; i--) {
        const char = text[i];
        if (char === '#') {
            // 必须在空格后（不允许在开头）
            const prevChar = i > 0 ? text[i - 1] : '';
            if (prevChar === ' ' && i > 0) {
                return i;
            }
        } else if (char === ' ') {
            break;
        }
    }
    return -1;
}

// 获取用户的分页加载偏好设置
function getPagedLoadingPreference() {
    const saved = localStorage.getItem('headingPagedLoading');
    return saved === null ? true : saved === 'true'; // 默认启用分页
}

// 保存用户的分页加载偏好设置
function setPagedLoadingPreference(enabled) {
    localStorage.setItem('headingPagedLoading', enabled.toString());
}

// 生成设置区域的HTML
function generateSettingsHtml() {
    const pagedLoadingEnabled = getPagedLoadingPreference();
    return `
        <div class="selector-settings">
            <label class="selector-setting-item">
                <input type="checkbox" id="paged-loading-checkbox" ${pagedLoadingEnabled ? 'checked' : ''}>
                <span>分页加载标题（小文档可关闭以提升速度）</span>
            </label>
        </div>
    `;
}

// 绑定设置勾选框事件
function bindSettingsCheckboxEvent() {
    const checkbox = document.getElementById('paged-loading-checkbox');
    if (checkbox) {
        checkbox.addEventListener('change', (e) => {
            const enabled = e.target.checked;
            setPagedLoadingPreference(enabled);
            console.log(`分页加载设置已更改为: ${enabled ? '启用' : '禁用'}`);

            // 清空当前标题列表
            window.currentHeadings = [];
            window.currentPage = 0;
            window.hasMoreHeadings = false;

            // 重新加载标题，使用最新的设置HTML
            quickSelectorContent.innerHTML = generateSettingsHtml() + '<div class="selector-loading">正在重新获取文档标题...</div>';

            // 重新绑定事件（因为DOM被重建了）
            bindSettingsCheckboxEvent();

            // 根据新的设置加载标题
            fetchDocumentHeadings(0, false, !enabled);
        });
    }
}

// 显示快捷选择器
function showQuickSelector() {
    if (!quickSelector || !quickSelectorContent) return;

    const showStartTime = performance.now();
    console.log(`⏱️ showQuickSelector开始, 开始时间: ${showStartTime.toFixed(2)}ms`);

    // 显示选择器
    quickSelector.style.display = 'block';
    console.log(`⏱️ 选择器显示完成, 耗时: ${(performance.now() - showStartTime).toFixed(2)}ms`);

    // 获取用户的分页加载偏好
    const pagedLoadingEnabled = getPagedLoadingPreference();

    // 显示加载状态和分页设置
    quickSelectorContent.innerHTML = generateSettingsHtml() + '<div class="selector-loading">正在获取文档标题...</div>';

    // 绑定勾选框事件
    bindSettingsCheckboxEvent();

    console.log(`⏱️ 加载状态显示完成, 耗时: ${(performance.now() - showStartTime).toFixed(2)}ms`);

    // 请求标题数据
    console.log(`⏱️ 准备调用fetchDocumentHeadings, 耗时: ${(performance.now() - showStartTime).toFixed(2)}ms`);
    fetchDocumentHeadings(0, false, !pagedLoadingEnabled); // 根据用户设置决定是否一次性加载所有
}

// 隐藏快捷选择器
function hideQuickSelector() {
    if (quickSelector) {
        quickSelector.style.display = 'none';
    }
    // 重置选中项
    currentSelectedIndex = -1;

    // 如果正在获取标题，停止获取
    if (isFetchingHeadings) {
        stopFetchingHeadings();
    }
}

// 当前选中的项目索引
let currentSelectedIndex = -1;
// 是否正在获取标题
let isFetchingHeadings = false;

// 获取文档标题（支持分页和一次性加载所有）
function fetchDocumentHeadings(page = 0, append = false, loadAll = false) {
    try {
        // 设置获取状态
        isFetchingHeadings = true;
        const startTime = performance.now();
        console.log(`⏱️ 开始获取文档标题... 页码: ${page}, 追加: ${append}, 一次性加载: ${loadAll}, 开始时间: ${startTime.toFixed(2)}ms`);

        if (window.chrome && window.chrome.webview) {
            console.log(`⏱️ 向C#发送getDocumentHeadings请求, 耗时: ${(performance.now() - startTime).toFixed(2)}ms`);
            window.chrome.webview.postMessage({
                type: 'getDocumentHeadings',
                page: loadAll ? 0 : page,
                pageSize: loadAll ? 99999 : 10, // 如果一次性加载，设置一个很大的数字
                append: append
            });
        } else {
            // 测试数据 - 模拟分页
            setTimeout(() => {
                if (isFetchingHeadings) { // 检查是否仍在获取状态
                    const allTestData = [
                        { text: "《艺术概论》试题", level: 1, page: 1 },
                        { text: "1.人类的语言虽无处不谈论神话解释艺术的起源", level: 2, page: 1 },
                        { text: "2.《盲诗人》中所引，音高角度对工声中", level: 2, page: 1 },
                        { text: "3.认为所有的艺术都源于对自然界和社会现实的模仿", level: 2, page: 1 },
                        { text: "4.人的感性冲动和理性冲动必须通过游戏冲动", level: 2, page: 1 },
                        { text: "5.艺术源于游戏，提出这种艺术起源观点的美学家是", level: 2, page: 1 },
                        { text: "6.第二章 艺术的本质", level: 1, page: 2 },
                        { text: "7.艺术是社会生活的反映", level: 2, page: 2 },
                        { text: "8.艺术的审美特性", level: 2, page: 2 },
                        { text: "9.艺术的意识形态特性", level: 2, page: 2 },
                        { text: "10.第三章 艺术的功能", level: 1, page: 3 },
                        { text: "11.艺术的认识功能", level: 2, page: 3 },
                        { text: "12.艺术的教育功能", level: 2, page: 3 },
                        { text: "13.艺术的审美功能", level: 2, page: 3 }
                    ];

                    const pageSize = 10;
                    const startIndex = page * pageSize;
                    const endIndex = Math.min(startIndex + pageSize, allTestData.length);
                    const pagedData = allTestData.slice(startIndex, endIndex);
                    const hasMore = endIndex < allTestData.length;

                    showHeadingsInSelector(pagedData, page, append, hasMore, allTestData.length);
                }
            }, 500); // 模拟延迟
        }
    } catch (error) {
        console.error('获取文档标题失败:', error);
        showSelectorError('获取文档标题失败，请确保Word文档已打开');
        isFetchingHeadings = false;
    }
}

// 在选择器中显示标题列表（支持分页和懒加载）
function showHeadingsInSelector(headings, page = 0, append = false, hasMore = false, total = 0) {
    if (!quickSelectorContent) return;

    const functionStartTime = performance.now();
    console.log(`⏱️ showHeadingsInSelector开始, 函数开始时间: ${functionStartTime.toFixed(2)}ms`);

    // 重置获取状态
    isFetchingHeadings = false;
    console.log(`⏱️ 标题获取完成 - 页码: ${page}, 追加: ${append}, 还有更多: ${hasMore}, 耗时: ${(performance.now() - functionStartTime).toFixed(2)}ms`);

    if (!headings || headings.length === 0) {
        if (!append) {
            quickSelectorContent.innerHTML = generateSettingsHtml() + '<div class="selector-empty">文档中没有找到标题</div>';
            bindSettingsCheckboxEvent();
        }
        console.log(`⏱️ 标题为空，直接返回, 耗时: ${(performance.now() - functionStartTime).toFixed(2)}ms`);
        return;
    }

    const dataProcessStartTime = performance.now();
    console.log(`⏱️ 开始处理标题数据, 数据处理开始时间: ${dataProcessStartTime.toFixed(2)}ms`);

    // 存储或追加标题数据
    if (!append || !window.currentHeadings) {
        window.currentHeadings = [...headings];
        window.currentPage = 0;
        window.hasMoreHeadings = hasMore;
        window.totalHeadings = total;
    } else {
        window.currentHeadings = [...window.currentHeadings, ...headings];
        window.currentPage = page;
        window.hasMoreHeadings = hasMore;
    }

    console.log(`⏱️ 标题数据存储完成, 耗时: ${(performance.now() - dataProcessStartTime).toFixed(2)}ms`);

    // 构建层级化的标题结构
    const buildStartTime = performance.now();
    console.log(`⏱️ 开始构建层级化标题结构, 标题数量: ${window.currentHeadings.length}, 构建开始时间: ${buildStartTime.toFixed(2)}ms`);
    const hierarchicalHtml = buildHierarchicalHeadings(window.currentHeadings);
    console.log(`⏱️ 层级化标题结构构建完成, 耗时: ${(performance.now() - buildStartTime).toFixed(2)}ms`);

    // 添加加载更多按钮
    let loadMoreHtml = '';
    if (window.hasMoreHeadings) {
        loadMoreHtml = `
            <div class="load-more-container">
                <button class="load-more-btn" onclick="loadMoreHeadings()">
                    显示更多标题 (已显示 ${window.currentHeadings.length}/${window.totalHeadings})
                </button>
            </div>
        `;
    }

    const renderStartTime = performance.now();
    console.log(`⏱️ 开始渲染HTML到DOM, 渲染开始时间: ${renderStartTime.toFixed(2)}ms`);
    quickSelectorContent.innerHTML = generateSettingsHtml() + hierarchicalHtml + loadMoreHtml;
    console.log(`⏱️ HTML渲染到DOM完成, 耗时: ${(performance.now() - renderStartTime).toFixed(2)}ms`);

    // 重新绑定设置勾选框事件
    bindSettingsCheckboxEvent();

    // 重置选中项
    currentSelectedIndex = -1;

    console.log(`⏱️ showHeadingsInSelector函数完成, 总耗时: ${(performance.now() - functionStartTime).toFixed(2)}ms`);

    // 不再设置自动滚动加载，只支持手动点击加载更多
    // setupLazyLoading();
}

// 构建层级化的标题HTML
function buildHierarchicalHeadings(headings) {
    if (!headings || headings.length === 0) return '';

    const startTime = performance.now();
    console.log(`⏱️ buildHierarchicalHeadings开始, 标题数量: ${headings.length}, 开始时间: ${startTime.toFixed(2)}ms`);

    let html = '';

    // 预计算最小级别，避免重复计算
    const minLevelStartTime = performance.now();
    const minLevel = Math.min(...headings.map(h => h.level));
    console.log(`⏱️ 计算最小级别完成, 最小级别: ${minLevel}, 耗时: ${(performance.now() - minLevelStartTime).toFixed(2)}ms`);

    const foreachStartTime = performance.now();
    headings.forEach((heading, index) => {
        const itemStartTime = performance.now();
        const level = heading.level;
        const levelClass = `heading-level-${level <= 6 ? level : 'default'}`;

        // 计算缩进级别（相对于最小级别）
        const indentLevel = Math.max(0, level - minLevel);

        // 构建层级指示器
        let hierarchyIndicator = '';
        if (indentLevel > 0) {
            // 构建完整的层级线条
            let prefix = '';

            // 为每个层级构建正确的连接符
            for (let currentLevel = minLevel + 1; currentLevel <= level; currentLevel++) {
                const levelIndent = currentLevel - minLevel - 1;

                if (currentLevel === level) {
                    // 当前级别：检查是否是最后一个同级元素
                    let isLastInLevel = true;
                    for (let i = index + 1; i < headings.length; i++) {
                        if (headings[i].level < level) {
                            break;
                        } else if (headings[i].level === level) {
                            isLastInLevel = false;
                            break;
                        }
                    }
                    prefix += (isLastInLevel ? '└─ ' : '├─ ');
                } else {
                    // 父级别：检查该级别后面是否还有同级或更高级别的元素
                    let hasMoreInThisLevel = false;
                    for (let i = index + 1; i < headings.length; i++) {
                        if (headings[i].level < currentLevel) {
                            break;
                        } else if (headings[i].level === currentLevel) {
                            hasMoreInThisLevel = true;
                            break;
                        }
                    }
                    prefix += hasMoreInThisLevel ? '│  ' : '   ';
                }
            }

            hierarchyIndicator = prefix;
        }

        // 检查是否有子标题
        let hasChildren = false;
        let childrenCount = 0;
        for (let i = index + 1; i < headings.length; i++) {
            if (headings[i].level > level) {
                if (!hasChildren) hasChildren = true;
                if (headings[i].level === level + 1) {
                    childrenCount++;
                }
            } else if (headings[i].level <= level) {
                break;
            }
        }

        // 构建父级信息
        let parentInfo = '';
        if (indentLevel > 0) {
            // 构建完整的父级路径
            const parentPath = [];
            let searchLevel = level - 1;

            while (searchLevel >= 1) {
                let found = false;
                for (let i = index - 1; i >= 0; i--) {
                    if (headings[i].level === searchLevel) {
                        parentPath.unshift(headings[i].text);
                        found = true;
                        break;
                    } else if (headings[i].level < searchLevel) {
                        break;
                    }
                }
                if (!found) break;
                searchLevel--;
            }

            if (parentPath.length > 0) {
                parentInfo = `路径: ${parentPath.join(' → ')}`;
            }
        }

        // 清理标题文本用于显示（与插入时的清理逻辑保持一致）
        const displayText = cleanHeadingText(heading.text);

        // 构建标题项
        html += `
            <div class="heading-item hierarchical" data-index="${index}" onclick="selectHeading(${index})" style="padding-left: ${indentLevel * 20 + 8}px;">
                <div class="heading-hierarchy">
                    <span class="hierarchy-indicator">${hierarchyIndicator}</span>
                    <div class="heading-level ${levelClass}">${level}</div>
                </div>
                <div class="heading-content">
                    <div class="heading-text-container">
                        <div class="heading-text">${escapeHtml(displayText)}</div>
                        ${parentInfo ? `<div class="parent-info">${parentInfo}</div>` : ''}
                    </div>
                    <div class="heading-meta">
                        <span class="heading-page">p.${heading.page}</span>
                        ${hasChildren ? `<span class="has-children" title="包含 ${childrenCount} 个直接子标题">📁${childrenCount}</span>` : '<span class="no-children">📄</span>'}
                    </div>
                </div>
            </div>
        `;

        // 每10个项目打印一次进度（避免日志太多）
        if ((index + 1) % 10 === 0 || index === headings.length - 1) {
            console.log(`⏱️ 处理标题项进度: ${index + 1}/${headings.length}, 当前项耗时: ${(performance.now() - itemStartTime).toFixed(2)}ms`);
        }
    });

    console.log(`⏱️ forEach循环完成, 耗时: ${(performance.now() - foreachStartTime).toFixed(2)}ms`);
    console.log(`⏱️ buildHierarchicalHeadings完成, 总耗时: ${(performance.now() - startTime).toFixed(2)}ms`);

    return html;
}

// 停止获取文档标题
function stopFetchingHeadings() {
    if (isFetchingHeadings) {
        console.log('停止获取文档标题');
        isFetchingHeadings = false;

        // 发送停止请求到C#
        try {
            if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage({
                    type: 'stopGeneration'
                });
            }
        } catch (error) {
            console.error('发送停止请求失败:', error);
        }

        // 更新选择器显示
        if (quickSelectorContent) {
            quickSelectorContent.innerHTML = generateSettingsHtml() + '<div class="selector-cancelled">已取消获取标题</div>';
            bindSettingsCheckboxEvent();
        }
    }
}

// 加载更多标题
function loadMoreHeadings() {
    if (isFetchingHeadings || !window.hasMoreHeadings) return;

    const nextPage = (window.currentPage || 0) + 1;
    console.log(`加载更多标题 - 页码: ${nextPage}`);

    // 显示加载状态
    const loadMoreBtn = document.querySelector('.load-more-btn');
    if (loadMoreBtn) {
        loadMoreBtn.textContent = '正在加载...';
        loadMoreBtn.disabled = true;
    }

    fetchDocumentHeadings(nextPage, true);
}

// 设置懒加载滚动监听
function setupLazyLoading() {
    if (!quickSelectorContent) return;

    // 移除之前的监听器
    quickSelectorContent.removeEventListener('scroll', handleLazyScroll);

    // 添加新的监听器
    quickSelectorContent.addEventListener('scroll', handleLazyScroll);
}

// 处理懒加载滚动
function handleLazyScroll() {
    if (isFetchingHeadings || !window.hasMoreHeadings) return;

    const container = quickSelectorContent;
    const scrollTop = container.scrollTop;
    const scrollHeight = container.scrollHeight;
    const clientHeight = container.clientHeight;

    // 当滚动到底部附近时（距离底部50px以内）
    if (scrollTop + clientHeight >= scrollHeight - 50) {
        console.log('滚动到底部，触发懒加载');
        loadMoreHeadings();
    }
}

// 显示错误信息
function showSelectorError(message) {
    if (!quickSelectorContent) return;
    quickSelectorContent.innerHTML = generateSettingsHtml() + `<div class="selector-error">${escapeHtml(message)}</div>`;
    bindSettingsCheckboxEvent();
    // 重置获取状态
    isFetchingHeadings = false;
}

// 选择标题
function selectHeading(index) {
    if (!window.currentHeadings || !window.currentHeadings[index]) return;

    const heading = window.currentHeadings[index];
    const headings = window.currentHeadings;

    // 仅用于构建 UI 显示所需的父级信息（不再用于插入文本）
    let fullDescription = '';

    // 找到所有父级标题
    const parentChain = [];
    const currentLevel = heading.level;

    // 构建完整的父级链条
    let searchLevel = currentLevel - 1;
    while (searchLevel >= 1) {
        // 从当前位置往前查找指定级别的标题
        let found = false;
        for (let i = index - 1; i >= 0; i--) {
            if (headings[i].level === searchLevel) {
                parentChain.unshift(headings[i]);
                found = true;
                break;
            } else if (headings[i].level < searchLevel) {
                // 如果遇到更高级的标题，说明中间级别缺失，停止查找
                break;
            }
        }

        if (!found) break;
        searchLevel--;
    }

    // 清理标题文本：将制表符、换行符等特殊空白字符统一替换为普通空格
    // 这样可以避免光标位置计算错误
    const cleanedHeadingText = cleanHeadingText(heading.text);

    // 构建插入到输入框的文本：@标题（内部空格替换为不换行空格，整体不被换行拆开）+ 零宽分隔符 + 空格
    const nonBreakHeading = cleanedHeadingText.replace(/\s+/g, '\u00A0'); // NBSP
    const mentionText = `@${nonBreakHeading}\u200b `;
    console.log('选中标题(原始):', heading.text, '清理后:', cleanedHeadingText, '插入文本:', mentionText, '完整描述(仅用于日志):', fullDescription);

    insertHeadingToInput(mentionText);
    hideQuickSelector();
}

// 将标题插入到输入框
function insertHeadingToInput(headingText) {
    if (!messageInput) return;

    const text = messageInput.value;
    const cursorPos = messageInput.selectionStart;

    // 找到触发字符的位置
    const triggerIndex = findTriggerIndex(text, cursorPos);
    if (triggerIndex === -1) return;

    console.log('插入标题:', headingText, '触发位置:', triggerIndex);

    // 替换从触发字符开始到光标位置的文本
    const beforeTrigger = text.substring(0, triggerIndex);
    const afterCursor = text.substring(cursorPos);
    // 在标题后面添加一个空格，提升用户体验
    const newText = beforeTrigger + headingText + ' ' + afterCursor;

    messageInput.value = newText;

    // 设置新的光标位置（在标题和空格之后）
    const newCursorPos = triggerIndex + headingText.length + 1;
    messageInput.setSelectionRange(newCursorPos, newCursorPos);

    // 自动调整高度
    autoResizeInput();

    // 更新字符计数
    updateCharacterCount();
    // 刷新高亮
    updateInputHighlights();

    // 聚焦输入框
    messageInput.focus();

    console.log('标题插入完成，新文本:', newText);
}

// 键盘导航
function navigateSelector(direction) {
    const items = quickSelectorContent.querySelectorAll('.heading-item');
    if (items.length === 0) return;

    // 移除当前高亮
    if (currentSelectedIndex >= 0 && currentSelectedIndex < items.length) {
        items[currentSelectedIndex].classList.remove('selected');
    }

    // 计算新的选中索引
    if (direction === 'down') {
        currentSelectedIndex = currentSelectedIndex < items.length - 1 ? currentSelectedIndex + 1 : 0;
    } else if (direction === 'up') {
        currentSelectedIndex = currentSelectedIndex > 0 ? currentSelectedIndex - 1 : items.length - 1;
    }

    // 添加新的高亮
    if (currentSelectedIndex >= 0 && currentSelectedIndex < items.length) {
        items[currentSelectedIndex].classList.add('selected');
        // 滚动到可见区域
        items[currentSelectedIndex].scrollIntoView({ block: 'nearest' });
    }
}

// 选择当前高亮的项目
function selectCurrentItem() {
    if (currentSelectedIndex >= 0 && window.currentHeadings && window.currentHeadings[currentSelectedIndex]) {
        selectHeading(currentSelectedIndex);
    }
}

// 清理标题文本：将制表符、换行符等特殊空白字符统一替换为普通空格
// 这样可以避免光标位置计算错误，确保显示和插入的文本一致
function cleanHeadingText(text) {
    if (!text) return '';
    return text
        .replace(/\t/g, ' ')      // 制表符替换为空格
        .replace(/\r\n/g, ' ')    // Windows换行符
        .replace(/[\r\n]/g, ' ')  // 其他换行符
        .replace(/\s+/g, ' ')     // 多个连续空格合并为一个
        .trim();                   // 去除首尾空格
}

// HTML转义
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ==================== 预览操作功能 ====================

// 处理工具预览结果
function handleToolPreview(data) {
    console.log('收到工具预览数据:', data);
    console.log('preview_mode:', data.preview_mode);
    console.log('success:', data.success);

    if (data.preview_mode && data.success) {
        console.log('条件满足，准备显示现代化内联预览');

        // 设置预览待处理标志，暂停正常内容追加
        isPreviewPending = true;
        console.log('已设置预览待处理标志，暂停内容追加');

        // 准备操作数据
        let parameters = {};

        if (data.action_type === 'modify_style' && data.style_parameters && Object.keys(data.style_parameters).length > 0) {
            // 样式修改操作，使用style_parameters
            parameters = data.style_parameters;

            // 检查文本长度，防止COM异常
            if (parameters.text_to_find && parameters.text_to_find.length > 255) {
                console.warn('文本过长，截断为255字符:', parameters.text_to_find.length);
                parameters.text_to_find = parameters.text_to_find.substring(0, 255);
            }
        } else {
            // 插入内容操作，构建标准参数
            parameters = {
                target_heading: data.target_heading || "",
                content: data.original_content || "",
                format_type: data.format_type || "paragraph",
                indent_level: data.indent_level || 0,
                add_spacing: data.add_spacing !== undefined ? data.add_spacing : true,
                insert_position: data.insert_position || agentConfig.defaultInsertPosition
            };
        }

        // 创建操作数据
        const actionData = {
            action_type: data.action_type,
            parameters: parameters
        };

        // 使用预览管理器创建预览
        const preview = previewManager.createPreview(data, actionData);
        console.log('创建预览:', preview.id);

        // 生成现代化工具预览并添加到聊天消息中（始终追加到当前正在生成的对话消息，而不是新开一条）
        const previewHtml = generateModernToolPreview(data, preview.id);
        // 优先使用最后一条助手消息；没有则创建
        let messageElement = chatMessages.querySelector('.assistant-message:last-child') || addAssistantMessage('');
        toolProgressHostMessage = messageElement; // 记住宿主，便于后续继续往里追加
        const messageContent = messageElement.querySelector('.message-content .markdown-content');
        if (messageContent) {
            // 插入预览卡片（新的卡片模式会自动穿插在对话中）
            const container = document.createElement('div');
            container.innerHTML = previewHtml;
            messageContent.insertAdjacentElement('beforeend', container.firstElementChild);
        } else {
            messageElement = addAssistantMessage(previewHtml);
            toolProgressHostMessage = messageElement;
        }

        // 将DOM元素关联到预览对象
        preview.element = messageElement.querySelector('.tool-preview-container');

        // 绑定事件处理器
        bindToolPreviewEvents(preview.element, preview.id, actionData);

        // 渲染预览内容中的数学公式、Mermaid流程图等
        setTimeout(() => {
            renderPreviewContent(preview.element);
        }, 100); // 延迟渲染，确保DOM已完全插入

        // 不再立即显示批量操作按钮，等待模型处理完成后再显示
        console.log('现代化预览创建完成，等待模型处理完成后显示批量按钮');

        // 当插入/样式预览已经生成时，视为该工具本轮调用已结束，立即将工具卡片标记为“完成”
        try {
            if (data.action_type === 'insert_content' || data.action_type === 'modify_style') {
                const nowLabel = new Date().toLocaleTimeString('zh-CN', { hour12: false });
                // 优先按工具名更新，避免 currentToolCard 不存在导致无法更新
                const toolName = data.action_type === 'insert_content' ? 'formatted_insert_content' : 'modify_text_style';
                const ok = markToolCardCompletedByName(toolName, nowLabel);
                if (!ok && currentToolCard) {
                    updateToolCallCard('completed', '预览已生成', nowLabel);
                }
            }
        } catch (e) {
            console.warn('预览生成后更新工具卡片状态失败:', e);
        }

        // 预览创建完成，结束生成状态
        finishGeneratingOutline();
        console.log('预览创建完成，已结束AI思考状态');

        // 保留旧的兼容性
        currentPreviewedAction = actionData;
    } else {
        console.log('条件不满足，无法显示预览');
        console.log('原因: preview_mode =', data.preview_mode, ', success =', data.success);

        // 即使预览失败，也要结束生成状态
        finishGeneratingOutline();
        console.log('预览失败，已结束AI思考状态');
    }
}

// 处理操作应用结果
function handleActionApplied(data) {
    console.log('收到操作应用结果:', data);

    // 优先处理新的预览管理器
    const applyingPreviews = Array.from(previewManager.previews.values()).filter(p => p.status === 'applying');

    if (applyingPreviews.length > 0) {
        // 如果有预览ID，精确匹配；否则更新最早的应用中预览
        let targetPreview = null;

        if (data.preview_id) {
            targetPreview = previewManager.previews.get(data.preview_id);
            if (targetPreview && targetPreview.status !== 'applying') {
                console.warn(`预览 ${data.preview_id} 状态不是 applying，当前状态: ${targetPreview.status}`);
                targetPreview = null;
            }
        }

        // 如果没有找到精确匹配，使用最早的应用中预览（按时间戳排序）
        if (!targetPreview) {
            targetPreview = applyingPreviews.sort((a, b) => a.timestamp - b.timestamp)[0];
            console.log(`使用最早的应用中预览: ${targetPreview.id}`);
        }

        if (targetPreview) {
            console.log(`更新预览状态: ${targetPreview.id} -> ${data.success ? 'applied' : 'rejected'}`);

            if (data.success) {
                previewManager.updateStatus(targetPreview.id, 'applied', data.message || '操作成功应用');
            } else {
                previewManager.updateStatus(targetPreview.id, 'rejected', data.message || '操作应用失败');
            }
        }

        // 更新批量操作按钮
        updateBatchActionButtons();
        return;
    }

    // 兼容旧版预览的处理逻辑
    const activePreview = document.querySelector('.inline-preview:not([style*="display: none"])');

    if (activePreview) {
        // 处理内联预览的结果显示
        const header = activePreview.querySelector('.preview-title');
        const actions = activePreview.querySelector('.preview-actions');

        if (data.success) {
            // 成功状态
            if (header) {
                header.innerHTML = '<span class="icon">✅</span><span>操作成功</span><span class="preview-type-badge" style="background: #10b981;">已应用</span>';
            }
            if (actions) {
                actions.innerHTML = `
                    <div style="color: #059669; background: #ecfdf5; padding: 8px 12px; border-radius: 6px; font-size: 13px; width: 100%; text-align: center;">
                        <strong>✅ ${data.message}</strong>
                    </div>
                `;
            }

            // 添加成功样式
            activePreview.style.borderLeftColor = '#10b981';
            activePreview.style.background = '#f0fdf4';

            // 3秒后优雅淡出
            setTimeout(() => {
                activePreview.style.transition = 'all 0.5s ease';
                activePreview.style.opacity = '0';
                activePreview.style.transform = 'scale(0.95)';
                activePreview.style.maxHeight = '0';
                activePreview.style.margin = '0';
                activePreview.style.padding = '0';
                activePreview.style.overflow = 'hidden';

                setTimeout(() => {
                    const parentElement = activePreview.parentElement;
                    activePreview.remove();

                    // 清理可能的空白容器
                    if (parentElement && parentElement.classList.contains('assistant-message') &&
                        parentElement.textContent.trim() === '') {
                        parentElement.style.display = 'none';
                        setTimeout(() => parentElement.remove(), 100);
                    }
                }, 500);
            }, 2000);

        } else {
            // 失败状态
            if (header) {
                header.innerHTML = '<span class="icon">❌</span><span>操作失败</span><span class="preview-type-badge" style="background: #ef4444;">失败</span>';
            }
            if (actions) {
                actions.innerHTML = `
                    <div style="color: #dc2626; background: #fef2f2; padding: 8px 12px; border-radius: 6px; font-size: 13px; width: 100%; text-align: center;">
                        <strong>❌ ${data.message}</strong>
                    </div>
                `;
            }

            // 添加失败样式
            activePreview.style.borderLeftColor = '#ef4444';
            activePreview.style.background = '#fef2f2';

            // 5秒后优雅淡出
            setTimeout(() => {
                activePreview.style.transition = 'all 0.5s ease';
                activePreview.style.opacity = '0';
                activePreview.style.transform = 'scale(0.95)';
                activePreview.style.maxHeight = '0';
                activePreview.style.margin = '0';
                activePreview.style.padding = '0';
                activePreview.style.overflow = 'hidden';

                setTimeout(() => {
                    const parentElement = activePreview.parentElement;
                    activePreview.remove();

                    // 清理可能的空白容器
                    if (parentElement && parentElement.classList.contains('assistant-message') &&
                        parentElement.textContent.trim() === '') {
                        parentElement.style.display = 'none';
                        setTimeout(() => parentElement.remove(), 100);
                    }
                }, 500);
            }, 4000);
        }
    } else {
        // 没有找到内联预览，静默处理，不添加额外消息
        console.log('未找到对应的预览元素，操作结果已在预览管理器中处理');
        hidePreviewPanel();
    }

    // 清除当前预览操作
    currentPreviewedAction = null;
}

// 生成现代化内联预览HTML (类似Cursor/Cline)
function generateModernInlinePreview(data, previewId) {
    const icon = getPreviewIcon(data.action_type);
    const title = getPreviewTitle(data.action_type);
    const previewClass = getPreviewClass(data.action_type);

    // 生成预览内容
    let previewContent = '';
    switch (data.action_type) {
        case 'insert_content':
            previewContent = generateInsertPreviewContent(data);
            break;
        case 'modify_style':
            previewContent = generateStylePreviewContent(data);
            break;
        default:
            previewContent = generateGenericPreviewContent(data);
            break;
    }

    const previewHtml = `
        <div class="modern-inline-preview ${previewClass}" id="${previewId}" data-preview-id="${previewId}">
            <div class="preview-header">
                <div class="preview-title">
                    <span class="icon">${icon}</span>
                    <span>${title}</span>
                    <span class="preview-type-badge pending">${data.action_type.replace('_', ' ')}</span>
                </div>
                <div class="preview-metadata">
                    <span class="timestamp">${new Date().toLocaleTimeString()}</span>
                </div>
            </div>
            <div class="preview-content">
                ${previewContent}
            </div>
            <div class="preview-actions">
                <button class="preview-btn reject-btn" onclick="rejectModernPreview('${previewId}'); event.stopPropagation();">
                    <span class="icon">❌</span>
                    <span>拒绝</span>
                </button>
                <button class="preview-btn accept-btn" onclick="acceptModernPreview('${previewId}'); event.stopPropagation();">
                    <span class="icon">✅</span>
                    <span>接受</span>
                </button>
            </div>
        </div>
    `;

    return previewHtml;
}

// 更新批量操作按钮（已弃用，使用悬浮按钮代替）
function updateBatchActionButtons() {
    const pendingCount = previewManager.getPendingPreviews().length;
    console.log(`旧版批量操作按钮更新（已弃用） - 待处理预览数量: ${pendingCount}`);

    // 移除旧版批量操作按钮，现在使用悬浮按钮
    const batchContainer = document.querySelector('.batch-actions-container');
    if (batchContainer) {
        batchContainer.remove();
        console.log('移除旧版批量操作按钮');
    }

    return; // 不再创建旧版按钮

    // 查找现有的批量操作按钮
    // let batchContainer = document.querySelector('.batch-actions-container');

    if (pendingCount > 1 && !batchContainer) {
        // 创建批量操作按钮
        const batchActionsHtml = `
            <div class="batch-actions-container">
                <div class="batch-actions-card">
                    <div class="batch-header">
                        <span class="icon">📋</span>
                        <span>批量操作</span>
                        <span class="count-badge">${pendingCount} 个待处理</span>
                    </div>
                    <div class="batch-buttons">
                        <button class="batch-btn reject-all-btn" onclick="rejectAllPreviews()">
                            <span class="icon">❌</span>
                            <span>全部拒绝</span>
                        </button>
                        <button class="batch-btn accept-all-btn" onclick="acceptAllPreviews()">
                            <span class="icon">✅</span>
                            <span>全部接受</span>
                        </button>
                    </div>
                </div>
            </div>
        `;

        addAssistantMessage(batchActionsHtml);
    } else if (pendingCount <= 1 && batchContainer) {
        // 移除批量操作按钮，带有优雅的淡出动画
        batchContainer.style.transition = 'all 0.3s ease';
        batchContainer.style.opacity = '0';
        batchContainer.style.transform = 'translateY(-10px)';
        setTimeout(() => {
            if (batchContainer.parentElement) {
                batchContainer.remove();
            }
        }, 300);
    } else if (pendingCount > 1 && batchContainer) {
        // 更新计数
        const countBadge = batchContainer.querySelector('.count-badge');
        if (countBadge) {
            countBadge.textContent = `${pendingCount} 个待处理`;
        }
    }
}

// 预览辅助函数
function getPreviewIcon(actionType) {
    switch (actionType) {
        case 'insert_content': return '📝';
        case 'modify_style': return '🎨';
        case 'extract_content': return '📋';
        default: return '⚡';
    }
}

function getPreviewTitle(actionType) {
    switch (actionType) {
        case 'insert_content': return '插入内容预览';
        case 'modify_style': return '样式修改预览';
        case 'extract_content': return '内容提取预览';
        default: return '操作预览';
    }
}

function getPreviewClass(actionType) {
    switch (actionType) {
        case 'insert_content': return 'preview-insert';
        case 'modify_style': return 'preview-modify';
        case 'extract_content': return 'preview-extract';
        default: return 'preview-generic';
    }
}

function generateGenericPreviewContent(data) {
    return `
        <div class="preview-label">操作类型:</div>
        <div class="preview-value">${data.action_type}</div>
        <div class="preview-label">预览信息:</div>
        <div class="preview-value">${data.message || '无详细信息'}</div>
    `;
}

// 现代化预览操作函数
function acceptModernPreview(previewId) {
    console.log('接受现代化预览:', previewId);

    const preview = previewManager.previews.get(previewId);
    if (!preview) {
        console.error('未找到预览:', previewId);
        return;
    }

    // 更新状态为应用中
    previewManager.updateStatus(previewId, 'applying');

    // 发送应用请求
    previewManager.applyPreview(preview);

    // 清除预览待处理状态（如果没有其他待处理预览）
    if (previewManager.getPendingPreviews().length === 0) {
        isPreviewPending = false;
        hideFloatingBatchActions();
    } else {
        // 检查是否需要更新悬浮按钮
        const pendingCount = previewManager.getPendingPreviews().length;
        if (pendingCount >= 1) {
            showFloatingBatchActions();
        } else {
            hideFloatingBatchActions();
        }
    }
}

function rejectModernPreview(previewId) {
    console.log('拒绝现代化预览:', previewId);

    // 更新状态为已拒绝
    previewManager.updateStatus(previewId, 'rejected');

    // 清除预览待处理状态（如果没有其他待处理预览）
    if (previewManager.getPendingPreviews().length === 0) {
        isPreviewPending = false;
        hideFloatingBatchActions();
    } else {
        // 检查是否需要更新悬浮按钮
        const pendingCount = previewManager.getPendingPreviews().length;
        if (pendingCount >= 1) {
            showFloatingBatchActions();
        } else {
            hideFloatingBatchActions();
        }
    }
}

// 批量操作函数（异步处理）
async function acceptAllPreviews() {
    console.log('🚀 执行全部接受操作');

    // 显示处理状态
    const pendingCount = previewManager.getPendingPreviews().length;
    console.log(`准备批量处理 ${pendingCount} 个预览`);

    // 执行批量接受（异步顺序处理）
    await previewManager.acceptAll();

    isPreviewPending = false;
    updateBatchActionButtons();

    console.log('✅ 全部接受操作完成');
}

function rejectAllPreviews() {
    console.log('执行全部拒绝操作');
    previewManager.rejectAll();
    updateBatchActionButtons();
}

// 检查并显示批量操作按钮（在模型处理完成后）
function checkAndShowBatchActions() {
    const pendingCount = previewManager.getPendingPreviews().length;
    console.log(`模型处理完成，检查批量操作按钮 - 待处理预览数量: ${pendingCount}`);

    if (pendingCount >= 1) {
        console.log('有待处理预览，显示悬浮批量操作按钮');
        showFloatingBatchActions();
    } else {
        console.log('没有待处理预览，隐藏批量操作按钮');
        hideFloatingBatchActions();
    }
}

// 显示悬浮批量操作按钮
function showFloatingBatchActions() {
    const pendingCount = previewManager.getPendingPreviews().length;

    // 如果没有待处理的预览，隐藏悬浮按钮并返回
    if (pendingCount === 0) {
        hideFloatingBatchActions();
        console.log('没有待处理预览，隐藏悬浮批量操作按钮');
        return;
    }

    // 移除现有的悬浮按钮
    hideFloatingBatchActions();

    // 创建悬浮批量操作按钮
    const floatingBatch = document.createElement('div');
    floatingBatch.id = 'floating-batch-actions';
    floatingBatch.className = 'floating-batch-actions';
    floatingBatch.innerHTML = `
        <div class="floating-batch-card">
            <div class="floating-batch-header">
                <span class="batch-icon">📋</span>
                <span class="batch-text">批量操作</span>
                <span class="batch-count">${pendingCount}</span>
            </div>
            <div class="floating-batch-buttons">
                                 <button class="floating-btn reject-all-btn" onclick="handleFloatingRejectAll()" title="拒绝${pendingCount > 1 ? '所有' : '当前'}预览">
                     <span class="icon">❌</span>
                     <span>${pendingCount > 1 ? '全部拒绝' : '拒绝'}</span>
                 </button>
                 <button class="floating-btn accept-all-btn" onclick="handleFloatingAcceptAll()" title="接受${pendingCount > 1 ? '所有' : '当前'}预览">
                     <span class="icon">✅</span>
                     <span>${pendingCount > 1 ? '全部接受' : '接受'}</span>
                 </button>
            </div>
            <button class="close-floating-batch" onclick="hideFloatingBatchActions()" title="关闭">
                <span>✕</span>
            </button>
        </div>
    `;

    // 添加到页面
    document.body.appendChild(floatingBatch);

    // 添加出现动画
    setTimeout(() => {
        floatingBatch.classList.add('show');
    }, 10);

    console.log(`悬浮批量操作按钮已显示，待处理: ${pendingCount}`);
}

// 隐藏悬浮批量操作按钮
function hideFloatingBatchActions() {
    const existingFloating = document.getElementById('floating-batch-actions');
    if (existingFloating) {
        existingFloating.classList.add('hide');
        setTimeout(() => {
            if (existingFloating.parentElement) {
                existingFloating.remove();
            }
        }, 300);
        console.log('悬浮批量操作按钮已隐藏');
    }
}

// 悬浮按钮的处理函数（异步）
async function handleFloatingAcceptAll() {
    console.log('🎯 悬浮按钮：执行全部接受');

    // 显示处理中状态
    const floatingCard = document.querySelector('.floating-batch-card');
    if (floatingCard) {
        const originalContent = floatingCard.innerHTML;
        floatingCard.innerHTML = `
            <div style="text-align: center; padding: 20px; color: #059669;">
                <div style="font-size: 24px; margin-bottom: 8px;">🔄</div>
                <div style="font-weight: 600;">批量处理中...</div>
                <div style="font-size: 12px; opacity: 0.8; margin-top: 4px;">正在顺序应用预览</div>
            </div>
        `;

        // 执行批量操作
        await acceptAllPreviews();

        // 恢复原内容并隐藏
        setTimeout(() => {
            hideFloatingBatchActions();
        }, 1000);
    } else {
        await acceptAllPreviews();
        hideFloatingBatchActions();
    }
}

async function handleFloatingRejectAll() {
    console.log('🎯 悬浮按钮：执行全部拒绝');
    rejectAllPreviews();
    hideFloatingBatchActions();
}

// 测试所有修复的功能
function testAllFixes() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; border-radius: 16px; box-shadow: 0 8px 32px rgba(5,150,105,0.3);">
            <h2 style="margin: 0 0 16px 0;">🎉 完全修复验证</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">✅ 所有问题已修复：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.5;">
                    <li>🐛 <strong>空白回复修复</strong>：不再创建空的助手消息</li>
                    <li>⏰ <strong>批量按钮时机</strong>：模型处理完成后才显示</li>
                                         <li>📍 <strong>Cursor风格位置</strong>：输入框右上角，不占用聊天区域</li>
                     <li>🎨 <strong>简洁白色设计</strong>：去掉渐变，采用白色卡片设计</li>
                    <li>📱 <strong>响应式设计</strong>：移动端自适应</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 8px; margin-top: 12px; font-size: 13px; opacity: 0.9;">
                <strong>测试说明：</strong>将创建2个预览，模拟生成结束后悬浮按钮会出现在右侧
            </div>
        </div>
    `);

    setTimeout(() => {
        testModernPreviewSystem();
    }, 1000);
}

// 测试单个预览的悬浮按钮（新功能）
function testSinglePreview() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 16px; background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; border-radius: 12px; text-align: center;">
            <h3 style="margin: 0 0 8px 0;">🔧 单预览测试</h3>
            <p style="margin: 0; font-size: 14px;">现在1个预览也会显示悬浮按钮了！</p>
        </div>
    `);

    // 创建单个测试预览
    const preview = previewManager.createPreview({
        action_type: 'insert_content',
        message: '单个预览 - 插入MySQL和Redis对比内容'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: '## MySQL vs Redis对比\\n\\n1. **存储方式**：MySQL持久化存储，Redis内存存储\\n2. **数据结构**：MySQL关系型，Redis键值对\\n3. **使用场景**：MySQL适合复杂查询，Redis适合缓存',
            format_type: 'paragraph'
        }
    });

    // 模拟生成完成，触发悬浮按钮显示
    setTimeout(() => {
        checkAndShowBatchActions();
        addAssistantMessage(`
            <div style="padding: 14px; background: #f0fdf4; border: 1px solid #10b981; border-radius: 10px; margin: 16px 0;">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                    <span style="font-size: 16px;">🎯</span>
                    <strong style="color: #059669;">单预览悬浮按钮已显示！</strong>
                </div>
                <p style="margin: 0; font-size: 13px; color: #065f46; line-height: 1.4;">
                    📍 <strong>新功能</strong>：1个预览也会显示悬浮按钮<br>
                    🔘 <strong>按钮文案</strong>：自动调整为"接受"/"拒绝"<br>
                    ✨ <strong>用户体验</strong>：不需要等待多个预览才能看到批量操作
                </p>
            </div>
        `);
    }, 500);
}

// 快速测试悬浮按钮 - Cursor风格位置
function testFloatingButton() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 16px; background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); color: white; border-radius: 12px; text-align: center;">
            <h3 style="margin: 0 0 8px 0;">🎯 Cursor风格悬浮按钮</h3>
            <p style="margin: 0; font-size: 14px;">位置：输入框右上角 | 样式：简洁白色卡片</p>
        </div>
    `);

    // 创建测试预览
    const preview1 = previewManager.createPreview({
        action_type: 'insert_content',
        message: '测试预览1 - 插入内容'
    }, {
        action_type: 'insert_content',
        parameters: { content: 'MySQL和Redis对比的内容' }
    });

    const preview2 = previewManager.createPreview({
        action_type: 'modify_style',
        message: '测试预览2 - 修改样式'
    }, {
        action_type: 'modify_style',
        parameters: { text_to_find: '测试文本', font_bold: true }
    });

    // 直接显示悬浮按钮
    setTimeout(() => {
        showFloatingBatchActions();
        addAssistantMessage(`
            <div style="padding: 14px; background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px; margin: 16px 0; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                    <span style="font-size: 16px;">✅</span>
                    <strong style="color: #059669;">悬浮按钮已显示</strong>
                </div>
                <p style="margin: 0; font-size: 13px; color: #6b7280; line-height: 1.4;">
                    📍 位置：输入框右上角（类似Cursor）<br>
                    🎨 样式：白色卡片 + 简洁设计<br>
                    🚀 功能：全部接受/拒绝 + 关闭按钮
                </p>
            </div>
        `);
    }, 500);
}

// 完整展示新的预览系统功能
function showNewPreviewFeatures() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 16px; box-shadow: 0 8px 32px rgba(0,0,0,0.12); text-align: center;">
            <h2 style="margin: 0 0 16px 0;">🎉 全新预览系统上线！</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px; margin-bottom: 16px;">
                <h3 style="margin: 0 0 12px 0;">🔄 解决的问题</h3>
                <p style="margin: 0; font-size: 14px; opacity: 0.9;">
                    ❌ 之前：预览接受/拒绝后立即消失<br>
                    ❌ 之前：无法查看操作历史记录<br>
                    ❌ 之前：没有批量操作功能<br>
                    ❌ 之前：界面设计过于简陋
                </p>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">✅ 全新特性</h3>
                <p style="margin: 0; font-size: 14px; opacity: 0.9;">
                    ✅ <strong>持久化显示</strong>：预览保留在界面，可查看历史<br>
                    ✅ <strong>批量操作</strong>：多个预览时显示批量接受/拒绝<br>
                    ✅ <strong>状态跟踪</strong>：待处理→应用中→已完成，全程可视化<br>
                    ✅ <strong>现代化UI</strong>：渐变、动画、徽章，媲美顶级AI工具
                </p>
            </div>
        </div>
    `);

    setTimeout(() => {
        addAssistantMessage(`
            <div style="margin: 16px 0; padding: 16px; background: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 12px;">
                <h3>🧪 接下来会展示：</h3>
                <ol style="margin: 0; padding-left: 20px;">
                    <li>📝 插入内容预览</li>
                    <li>🎨 样式修改预览</li>
                    <li>📋 批量操作按钮自动显示</li>
                    <li>🔄 状态实时更新演示</li>
                </ol>
                <p style="margin: 16px 0 0 0;"><strong>注意观察：</strong>现在预览不会消失，状态会实时更新！</p>
            </div>
        `);
        testModernPreviewSystem();
    }, 1000);
}

// 测试现代化预览系统
function testModernPreviewSystem() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 16px; box-shadow: 0 8px 32px rgba(0,0,0,0.12);">
            <h3 style="margin: 0 0 16px 0;">🚀 现代化预览系统测试</h3>
            <p style="margin: 0 0 16px 0;"><strong>修复问题：</strong></p>
            <ul style="margin: 0; padding-left: 20px;">
                <li>✅ <strong>DOM元素错误</strong>：修复 querySelector 错误</li>
                <li>✅ <strong>批量操作显示</strong>：确保多个预览时自动显示</li>
                <li>✅ <strong>空项问题</strong>：修复Agent模式第一项为空的问题</li>
                <li>✅ <strong>预览持久化</strong>：保留操作历史记录</li>
            </ul>
            <p style="margin: 16px 0 0 0; opacity: 0.9; font-size: 14px;">接下来将展示两个预览，观察批量操作按钮的出现</p>
        </div>
    `);

    // 创建多个测试预览演示批量操作
    setTimeout(() => {
        console.log('创建第一个测试预览');
        // 第一个预览
        const testData1 = {
            action_type: 'insert_content',
            preview_mode: true,
            success: true,
            target_heading: '测试标题1',
            original_content: '这是第一个现代化预览测试内容',
            preview_content: '这是第一个现代化预览测试内容',
            format_type: 'paragraph',
            message: '预览：第一个插入内容操作'
        };

        const preview1 = previewManager.createPreview(testData1, {
            action_type: 'insert_content',
            parameters: { content: testData1.original_content, target_heading: testData1.target_heading }
        });

        const previewHtml1 = generateModernInlinePreview(testData1, preview1.id);
        const messageElement1 = addAssistantMessage(previewHtml1);
        preview1.element = messageElement1.querySelector('.modern-inline-preview');

        isPreviewPending = true;

        // 第二个预览
        setTimeout(() => {
            console.log('创建第二个测试预览');
            const testData2 = {
                action_type: 'modify_style',
                preview_mode: true,
                success: true,
                text_to_find: '测试文本',
                preview_styles: ['🔤 字体大小: 16磅', '🎨 字体颜色: blue', '🔲 粗体: 是'],
                message: '预览：修改文本样式'
            };

            const preview2 = previewManager.createPreview(testData2, {
                action_type: 'modify_style',
                parameters: { text_to_find: testData2.text_to_find, font_size: 16, font_color: 'blue', font_bold: true }
            });

            const previewHtml2 = generateModernInlinePreview(testData2, preview2.id);
            const messageElement2 = addAssistantMessage(previewHtml2);
            preview2.element = messageElement2.querySelector('.modern-inline-preview');

            // 模拟生成结束，检查是否显示悬浮批量按钮
            setTimeout(() => {
                console.log('模拟生成结束，检查悬浮批量按钮');
                checkAndShowBatchActions();
            }, 500);
        }, 1000);
    }, 1500);
}

// 生成现代化的内联预览HTML
function generateInlinePreview(data) {
    const previewId = 'preview_' + Date.now();
    let previewClass = 'preview-insert';
    let icon = '📝';
    let title = '插入内容预览';

    // 根据操作类型设置样式和图标
    switch (data.action_type) {
        case 'insert_content':
            previewClass = 'preview-insert';
            icon = '📝';
            title = '插入内容预览';
            break;
        case 'modify_style':
            previewClass = 'preview-modify';
            icon = '🎨';
            title = '样式修改预览';
            break;
        case 'extract_content':
            previewClass = 'preview-extract';
            icon = '📋';
            title = '内容提取预览';
            break;
        default:
            previewClass = 'preview-insert';
            icon = '🔧';
            title = '操作预览';
    }

    let previewContent = '';
    if (data.action_type === 'insert_content') {
        previewContent = generateInlineInsertPreview(data);
    } else if (data.action_type === 'modify_style') {
        previewContent = generateInlineStylePreview(data);
    } else {
        previewContent = generateGenericInlinePreview(data);
    }

    const previewHtml = `
        <div class="inline-preview ${previewClass}" id="${previewId}" onclick="event.stopPropagation()">
            <div class="preview-header">
                <div class="preview-title">
                    <span class="icon">${icon}</span>
                    <span>${title}</span>
                    <span class="preview-type-badge">${data.action_type.replace('_', ' ')}</span>
                </div>
            </div>
            <div class="preview-content" onclick="event.stopPropagation()">
                ${previewContent}
            </div>
            <div class="preview-actions" onclick="event.stopPropagation()">
                <button class="preview-btn reject-btn" onclick="rejectInlinePreview('${previewId}'); event.stopPropagation();">
                    <span class="icon">❌</span>
                    <span>拒绝</span>
                </button>
                <button class="preview-btn accept-btn" onclick="acceptInlinePreview('${previewId}'); event.stopPropagation();">
                    <span class="icon">✅</span>
                    <span>应用更改</span>
                </button>
            </div>
        </div>
    `;

    return previewHtml;
}

// 生成插入内容的内联预览
function generateInlineInsertPreview(data) {
    const targetHeading = data.target_heading || '光标位置';
    const formatType = data.format_type || 'paragraph';
    const formatName = getFormatTypeName(formatType);
    const content = data.original_content || '';
    const previewContent = data.preview_content || '';

    return `
        <div class="preview-section">
            <div class="preview-label">
                <span>📍</span>
                插入位置
            </div>
            <div class="preview-value">${targetHeading}</div>
        </div>
        
        <div class="preview-section">
            <div class="preview-label">
                <span>📝</span>
                内容格式
            </div>
            <div class="preview-meta">
                <div class="preview-meta-item">
                    <span class="label">格式:</span>
                    <span>${formatName}</span>
                </div>
                <div class="preview-meta-item">
                    <span class="label">缩进级别:</span>
                    <span>${data.indent_level || 0}</span>
                </div>
                <div class="preview-meta-item">
                    <span class="label">添加间距:</span>
                    <span>${data.add_spacing ? '是' : '否'}</span>
                </div>
            </div>
        </div>
        
        <div class="preview-section">
            <div class="preview-label">
                <span>👀</span>
                预览效果
            </div>
            <div class="preview-value preview-html">${previewContent || content}</div>
        </div>
    `;
}

// 生成样式修改的内联预览
function generateInlineStylePreview(data) {
    const textToFind = data.text_to_find || '';
    const previewStyles = data.preview_styles || [];

    return `
        <div class="preview-section">
            <div class="preview-label">
                <span>🔍</span>
                目标文本
            </div>
            <div class="preview-value">${textToFind}</div>
        </div>
        
        <div class="preview-section">
            <div class="preview-label">
                <span>🎨</span>
                应用样式
            </div>
            <ul class="style-preview-list">
                ${previewStyles.map(style => `
                    <li class="style-preview-item">
                        <span class="emoji">${getStyleEmoji(style)}</span>
                        <span>${style}</span>
                    </li>
                `).join('')}
            </ul>
        </div>
    `;
}

// 生成通用的内联预览
function generateGenericInlinePreview(data) {
    return `
        <div class="preview-section">
            <div class="preview-label">
                <span>ℹ️</span>
                操作描述
            </div>
            <div class="preview-value">${data.message || '准备执行操作'}</div>
        </div>
    `;
}

// 辅助函数：获取格式类型中文名称
function getFormatTypeName(formatType) {
    const formatNames = {
        'paragraph': '段落',
        'list': '列表',
        'table': '表格',
        'emphasis': '强调',
        'code': '代码'
    };
    return formatNames[formatType] || formatType;
}

// 辅助函数：获取样式对应的emoji
function getStyleEmoji(style) {
    if (style.includes('字体大小')) return '🔤';
    if (style.includes('粗体')) return '🔲';
    if (style.includes('斜体')) return '📐';
    if (style.includes('颜色')) return '🎨';
    if (style.includes('背景')) return '🖍️';
    if (style.includes('间距')) return '📏';
    if (style.includes('行间距')) return '📊';
    return '✨';
}

// 测试预览流程修复
function testPreviewFlow() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 16px; background: #fef7cd; border: 1px solid #f59e0b; border-radius: 8px;">
            <h3>🔄 预览流程修复测试</h3>
            <p><strong>问题说明:</strong></p>
            <p>之前的问题：内联预览显示后，AI回复仍在继续，导致预览被覆盖或与内容混合。</p>
            <p><strong>修复方案:</strong></p>
            <ul>
                <li>🛡️ <strong>预览时暂停</strong>：显示预览时设置 <code>isPreviewPending = true</code></li>
                <li>⏸️ <strong>内容追加暂停</strong>：<code>appendOutlineContent</code> 检查预览状态</li>
                <li>▶️ <strong>用户操作恢复</strong>：接受/拒绝后清除暂停状态</li>
                <li>🛟 <strong>保护机制</strong>：生成结束时自动清除预览状态</li>
            </ul>
            <p><strong>现在预览应该:</strong></p>
            <ul>
                <li>📌 稳定显示，不被AI回复覆盖</li>
                <li>⏳ 等待用户确认后才继续流程</li>
                <li>🔄 正确处理接受/拒绝操作</li>
            </ul>
        </div>
    `);
}

// 测试修复后的内联预览功能
function testInlinePreviewFixed() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 16px; background: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 8px;">
            <h3>🔧 内联预览修复测试</h3>
            <p><strong>修复问题:</strong></p>
            <ul>
                <li>✅ <strong>防止意外关闭</strong>：添加了事件阻止冒泡</li>
                <li>✅ <strong>消除空白区域</strong>：优化了元素移除动画和布局</li>
                <li>✅ <strong>改进用户体验</strong>：更优雅的淡入淡出效果</li>
                <li>✅ <strong>修复流程冲突</strong>：预览期间暂停内容追加，防止覆盖</li>
            </ul>
            <p><strong>测试说明:</strong></p>
            <ul>
                <li>🖱️ 点击预览区域外部不会关闭预览</li>
                <li>🎨 拒绝操作有平滑的收缩动画</li>
                <li>🧹 移除后不留空白区域</li>
                <li>⏸️ 预览显示时暂停AI回复，避免内容混乱</li>
            </ul>
        </div>
    `);

    // 延迟显示测试预览
    setTimeout(() => {
        currentPreviewedAction = {
            action_type: 'insert_content',
            parameters: {
                target_heading: '测试标题',
                content: '这是修复后的预览测试内容。\n现在点击预览外部不会意外关闭。\n拒绝操作也不会留下空白区域。',
                format_type: 'paragraph',
                indent_level: 0,
                add_spacing: true
            }
        };

        const testData = {
            action_type: 'insert_content',
            preview_mode: true,
            success: true,
            target_heading: '测试标题',
            original_content: '这是修复后的预览测试内容。\n现在点击预览外部不会意外关闭。\n拒绝操作也不会留下空白区域。',
            preview_content: '这是修复后的预览测试内容。\n现在点击预览外部不会意外关闭。\n拒绝操作也不会留下空白区域。',
            format_type: 'paragraph',
            indent_level: 0,
            add_spacing: true,
            message: '预览：修复测试 - 请尝试点击预览外部，应该不会关闭'
        };

        const previewHtml = generateInlinePreview(testData);
        addAssistantMessage(`
            <div style="margin: 16px 0;">
                <p><strong>🧪 测试预览</strong> - 请尝试：</p>
                <ol>
                    <li>点击预览区域外部（不应该关闭）</li>
                    <li>点击"拒绝"按钮（应该有平滑动画）</li>
                    <li>观察是否留下空白区域</li>
                </ol>
            </div>
            ${previewHtml}
        `);
    }, 1000);
}

// 演示内联预览效果的示例函数
function showInlinePreviewDemo() {
    // 演示插入内容预览
    const demoInsertData = {
        action_type: 'insert_content',
        preview_mode: true,
        success: true,
        target_heading: '项目概述',
        original_content: '这是一个示例内容，展示内联预览效果。\n- 支持多种格式\n- 现代化设计\n- 不遮挡视野',
        preview_content: '这是一个示例内容，展示内联预览效果。\n• 支持多种格式\n• 现代化设计\n• 不遮挡视野',
        format_type: 'list',
        indent_level: 0,
        add_spacing: true,
        message: '预览：将在标题 "项目概述" 下方插入列表格式的内容'
    };

    // 演示样式修改预览
    const demoStyleData = {
        action_type: 'modify_style',
        preview_mode: true,
        success: true,
        text_to_find: '重要提示',
        preview_styles: [
            '🔤 字体大小: 16磅',
            '🔲 粗体: 是',
            '🎨 字体颜色: red',
            '🖍️ 背景颜色: yellow'
        ],
        style_parameters: {
            text_to_find: '重要提示',
            font_size: 16,
            font_bold: true,
            font_color: 'red',
            background_color: 'yellow'
        },
        message: '预览：将为文本 "重要提示" 应用以下样式'
    };

    // 设置模拟的当前预览操作
    currentPreviewedAction = {
        action_type: 'insert_content',
        parameters: {
            target_heading: '项目概述',
            content: '这是一个示例内容，展示内联预览效果。\n- 支持多种格式\n- 现代化设计\n- 不遮挡视野',
            format_type: 'list',
            indent_level: 0,
            add_spacing: true
        }
    };

    // 生成并显示内联预览
    const previewHtml = generateInlinePreview(demoInsertData);
    addAssistantMessage(`
        <div style="margin: 16px 0;">
            <p><strong>🎉 新的内联预览效果展示</strong></p>
            <p>类似 <strong>Cursor</strong>、<strong>Cline</strong> 等现代AI工具的预览体验：</p>
            <ul>
                <li>✅ <strong>内联显示</strong>：预览直接嵌入聊天流中</li>
                <li>✅ <strong>不遮挡视野</strong>：不再使用弹出框</li>
                <li>✅ <strong>现代化设计</strong>：卡片式布局，优雅动画</li>
                <li>✅ <strong>智能反馈</strong>：实时状态更新</li>
            </ul>
        </div>
        ${previewHtml}
    `);

    // 2秒后再显示样式预览演示
    setTimeout(() => {
        currentPreviewedAction = {
            action_type: 'modify_style',
            parameters: demoStyleData.style_parameters
        };

        const stylePreviewHtml = generateInlinePreview(demoStyleData);
        addAssistantMessage(`
            <div style="margin: 16px 0;">
                <p><strong>🎨 样式修改预览演示</strong></p>
            </div>
            ${stylePreviewHtml}
        `);
    }, 2000);
}

// 生成插入内容的预览（支持渲染Markdown、表格、公式、代码、流程图）
function generateInsertPreviewContent(data) {
    // 元信息部分（移除emoji）
    let metaInfo = `
        <div class="preview-meta-section">
            <div class="preview-meta-item">
                <span class="meta-label">操作类型:</span>
                <span class="meta-value">${getActionTypeName(data.action_type)}</span>
            </div>
            <div class="preview-meta-item">
                <span class="meta-label">目标标题:</span>
                <span class="meta-value">"${escapeHtml(data.target_heading)}"</span>
            </div>
            <div class="preview-meta-item">
                <span class="meta-label">格式类型:</span>
                <span class="meta-value">${getFormatTypeName(data.format_type)}</span>
            </div>
    `;

    if (data.indent_level > 0) {
        metaInfo += `
            <div class="preview-meta-item">
                <span class="meta-label">缩进级别:</span>
                <span class="meta-value">${data.indent_level}</span>
            </div>
        `;
    }

    metaInfo += `</div>`;

    // 渲染预览内容（支持Markdown、表格、公式、代码、流程图）
    let previewContent = data.preview_content || '';
    let renderedContent = '';

    if (previewContent.trim()) {
        // 预处理：将 "• " 或 "・" 开头的行转换为标准Markdown列表
        previewContent = previewContent.replace(/^[•・]\s+/gm, '- ');

        // 使用现有的 renderMarkdown 函数渲染内容
        const renderedHTML = renderMarkdown(previewContent);

        // 创建临时容器用于后处理
        const tempDiv = document.createElement('div');
        tempDiv.className = 'preview-rendered-content markdown-content';
        tempDiv.innerHTML = renderedHTML;

        // 应用代码高亮、表格工具栏等处理（但不执行）
        try {
            // 处理代码块高亮
            tempDiv.querySelectorAll('pre code').forEach((block) => {
                if (typeof hljs !== 'undefined') {
                    hljs.highlightElement(block);
                }
            });

            // 处理表格（添加提示而非操作按钮）
            tempDiv.querySelectorAll('table').forEach((table, index) => {
                if (!table.closest('.table-container')) {
                    const container = document.createElement('div');
                    container.className = 'table-container';

                    const toolbar = document.createElement('div');
                    toolbar.className = 'table-toolbar';
                    toolbar.innerHTML = `
                        <p>表格预览</p>
                        <div>
                            <span style="color: #666; font-size: 12px;">接受后可插入到Word</span>
                        </div>
                    `;

                    table.parentNode.insertBefore(container, table);
                    container.appendChild(toolbar);
                    container.appendChild(table);
                }
            });

            // 处理代码块（添加提示而非操作按钮）
            tempDiv.querySelectorAll('pre code').forEach((codeBlock, index) => {
                const pre = codeBlock.parentElement;
                if (pre && !pre.previousElementSibling?.classList?.contains('code-toolbar')) {
                    const language = Array.from(codeBlock.classList)
                        .find(cls => cls.startsWith('language-'))
                        ?.replace('language-', '') || 'text';

                    const toolbar = document.createElement('div');
                    toolbar.className = 'code-toolbar';
                    toolbar.innerHTML = `
                        <p>${language}</p>
                        <div>
                            <span style="color: #666; font-size: 12px;">接受后可插入到Word</span>
                        </div>
                    `;

                    pre.parentNode.insertBefore(toolbar, pre);
                }
            });

        } catch (e) {
            console.warn('预览内容后处理失败:', e);
        }

        renderedContent = `
            <div class="preview-content-separator" onclick="toggleDetailedPreview(event)">
                <span class="preview-toggle-icon collapsed"></span>
                <span>预览效果</span>
            </div>
            <div class="preview-rendered-wrapper" style="display: none;">
                ${tempDiv.outerHTML}
            </div>
        `;
    } else {
        renderedContent = `
            <div class="preview-empty">
                <span style="color: #999;">（无预览内容）</span>
            </div>
        `;
    }

    // 插入设置（从 parameters 读取当前值，默认：清除空格=true，插入位置=end）
    const params = data.parameters || {};
    const trimSpaces = params.trim_spaces !== undefined ? params.trim_spaces : true; // 默认勾选
    const insertPos = params.insert_position || 'end'; // 默认末尾

    const insertSettingsHtml = `
        <div class="preview-content-separator">
            <span>插入设置</span>
        </div>
        <div class="preview-rendered-wrapper">
            <div style="display:flex;flex-direction:column;gap:12px;">
                <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
                    <input type="checkbox" class="insert-setting" data-setting-checkbox="trim_spaces" ${trimSpaces ? 'checked' : ''}>
                    <div style="display:flex;flex-direction:column;">
                        <span style="font-weight:500;">清除插入内容多余空格</span>
                        <span style="font-size:12px;color:#6b7280;">自动清理Agent填入内容的多余空行和格式</span>
                    </div>
                </label>
                <div style="display:flex;flex-direction:column;gap:6px;">
                    <label style="font-weight:500;">默认插入位置：</label>
                    <select class="insert-setting" data-setting-select="insert_position" style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;cursor:pointer;">
                        <option value="end" ${insertPos === 'end' ? 'selected' : ''}>标题下内容末尾（推荐）</option>
                        <option value="front" ${insertPos === 'front' ? 'selected' : ''}>标题下最前面</option>
                    </select>
                    <span style="font-size:12px;color:#6b7280;">设置formatted_insert_content工具的默认插入位置</span>
                </div>
            </div>
        </div>
    `;

    return metaInfo + renderedContent + insertSettingsHtml;
}

// 生成样式修改的预览（返回HTML格式）
function generateStylePreviewContent(data) {
    // 判断是否为范围批量模式（如果 text_to_find 为空或仅为引号）
    const isRangeMode = !data.text_to_find || data.text_to_find.trim() === '' || data.text_to_find === '""';

    // 元信息部分
    let metaInfo = '';
    if (!isRangeMode) {
        // 查找-匹配模式：显示目标文本
        metaInfo = `
            <div class="preview-meta-section">
                <div class="preview-meta-item">
                    <span class="meta-label">操作类型:</span>
                    <span class="meta-value">${getActionTypeName(data.action_type)}</span>
                </div>
                <div class="preview-meta-item">
                    <span class="meta-label">目标文本:</span>
                    <span class="meta-value">"${escapeHtml(data.text_to_find || '')}"</span>
                </div>
            </div>
        `;
    } else {
        // 范围批量模式：显示作用范围描述
        let scopeDesc = data.message || '全文正文';
        // 从 message 中提取范围描述（如"预览：将在全文正文应用以下样式"）
        const match = scopeDesc.match(/将在(.+?)应用以下样式/);
        if (match) {
            scopeDesc = match[1];
        }
        metaInfo = `
            <div class="preview-meta-section">
                <div class="preview-meta-item">
                    <span class="meta-label">操作类型:</span>
                    <span class="meta-value">${getActionTypeName(data.action_type)}</span>
                </div>
                <div class="preview-meta-item">
                    <span class="meta-label">作用范围:</span>
                    <span class="meta-value">${escapeHtml(scopeDesc)}</span>
                </div>
            </div>
        `;
    }

    // 样式内容：直接展开显示（不折叠）
    let stylesContent = '';
    if (data.preview_styles && data.preview_styles.length > 0) {
        const stylesList = data.preview_styles.map(style =>
            `<div style="padding: 8px 12px; background: #f0f9ff; border-left: 3px solid #3b82f6; margin-bottom: 6px; border-radius: 4px; font-size: 14px; color: #1e40af;">
                ${escapeHtml(style)}
            </div>`
        ).join('');

        stylesContent = `
            <div class="preview-content-separator">
                <span>样式内容</span>
            </div>
            <div class="preview-rendered-wrapper">
                ${stylesList}
            </div>
        `;
    } else {
        stylesContent = `
            <div class="preview-content-separator">
                <span>样式内容</span>
            </div>
            <div class="preview-rendered-wrapper">
                <div class="preview-empty">
                    <span>(无样式修改)</span>
                </div>
            </div>
        `;
    }

    // 作用对象设置（默认：正文=开，其余=关）
    const sp = data.style_parameters || {};
    const checked = (v, d) => (v === undefined ? d : !!v) ? 'checked' : '';
    const settingsHtml = `
        <div class="preview-content-separator">
            <span>作用范围</span>
        </div>
        <div class="preview-rendered-wrapper">
            <div class="settings-grid" style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;">
                <label style="display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" class="ms-setting" data-setting-checkbox="include_paragraphs" ${checked(sp.include_paragraphs, true)}>
                    <span>正文</span>
                </label>
                <label style="display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" class="ms-setting" data-setting-checkbox="include_headings" ${checked(sp.include_headings, false)}>
                    <span>标题</span>
                </label>
                <label style="display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" class="ms-setting" data-setting-checkbox="include_tables" ${checked(sp.include_tables, false)}>
                    <span>表格</span>
                </label>
                <label style="display:flex;align-items:center;gap:6px;">
                    <input type="checkbox" class="ms-setting" data-setting-checkbox="include_list_items" ${checked(sp.include_list_items, false)}>
                    <span>项目列表项</span>
                </label>
            </div>
            <div style="margin-top:8px;color:#6b7280;font-size:12px;">
                提示：默认仅修改正文；如需作用于标题/表格/列表项，请勾选对应项。
            </div>
        </div>
    `;

    return metaInfo + stylesContent + settingsHtml;
}

// 生成文本插入的预览


// 获取操作类型中文名称
function getActionTypeName(actionType) {
    switch (actionType) {
        case 'insert_content':
            return '格式化插入内容';
        case 'modify_style':
            return '修改文字样式';

        default:
            return actionType;
    }
}

// 获取格式类型中文名称
function getFormatTypeName(formatType) {
    switch (formatType) {
        case 'paragraph':
            return '段落';
        case 'list':
            return '列表';
        case 'table':
            return '表格';
        case 'emphasis':
            return '强调';
        default:
            return formatType;
    }
}

// 发送消息到C#
function sendMessageToCSharp(messageData) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(messageData);
    } else {
        console.error('WebView2环境不可用');
    }
}

// 接受预览操作
function acceptPreviewedAction() {
    if (!currentPreviewedAction) {
        console.error('没有当前预览操作');
        return;
    }

    console.log('接受预览操作:', currentPreviewedAction);

    // 诊断信息：检查操作类型和参数
    console.log('🔍 诊断信息:');
    console.log('- 操作类型:', currentPreviewedAction.action_type);
    console.log('- 参数个数:', Object.keys(currentPreviewedAction.parameters || {}).length);

    if (currentPreviewedAction.action_type === 'modify_style') {
        console.log('⚠️ 样式修改操作:');
        console.log('- text_to_find长度:', (currentPreviewedAction.parameters.text_to_find || '').length);
        console.log('- text_to_find内容:', (currentPreviewedAction.parameters.text_to_find || '').substring(0, 100) + '...');
    } else if (currentPreviewedAction.action_type === 'insert_content') {
        console.log('📝 插入内容操作:');
        console.log('- target_heading:', currentPreviewedAction.parameters.target_heading);
        console.log('- content长度:', (currentPreviewedAction.parameters.content || '').length);
    }

    // 发送应用操作请求到C#
    const messageData = {
        type: 'applyPreviewedAction',
        action_type: currentPreviewedAction.action_type,
        parameters: currentPreviewedAction.parameters
    };

    sendMessageToCSharp(messageData);
}

// 拒绝预览操作
function rejectPreviewedAction() {
    console.log('拒绝预览操作');

    // 隐藏预览面板
    hidePreviewPanel();

    // 在聊天界面显示拒绝消息
    addAssistantMessage(`<div style="color: #f59e0b; background: #fffbeb; padding: 12px; border-radius: 8px; border-left: 4px solid #f59e0b;"><strong>⚠️ 已取消:</strong> 操作已取消</div>`);
}

// ==================== 现代化工具预览功能 ====================

// 生成现代化工具预览HTML（支持渲染Markdown、表格、公式、代码、流程图）
function generateModernToolPreview(data, previewId) {
    const toolIcon = getToolIcon(data.action_type);
    const toolName = getToolName(data.action_type);
    const statusClass = data.success ? 'success' : 'error';

    // 根据操作类型生成预览内容（使用增强的渲染函数）
    let previewContentHtml = '';
    if (data.action_type === 'insert_content') {
        previewContentHtml = generateInsertPreviewContent(data);
    } else if (data.action_type === 'modify_style') {
        previewContentHtml = generateStylePreviewContent(data);
    } else {
        // 其他类型仍使用纯文本
        const previewContent = data.preview_content || data.original_content || '';
        const previewInfo = data.message || '工具调用预览';
        previewContentHtml = `
            <div class="preview-meta-section">
                <div class="preview-meta-item">
                    <span class="meta-label">预览信息:</span>
                    <span class="meta-value">${escapeHtml(previewInfo)}</span>
                </div>
            </div>
            <div class="preview-rendered-wrapper">
                <pre style="white-space: pre-wrap; font-family: 'Consolas', monospace; font-size: 13px;">${escapeHtml(previewContent)}</pre>
            </div>
        `;
    }

    return `
        <div class="tool-preview-container ${statusClass}" data-preview-id="${previewId}">
            <div class="tool-status-indicator"></div>
            
            <div class="tool-preview-header">
                <div class="tool-preview-title">
                    <div class="tool-preview-icon tool-icon-${data.action_type}">${toolIcon}</div>
                    <span>${toolName}</span>
                </div>
                <div class="tool-preview-actions">
                    <button class="tool-action-btn accept" title="接受修改" onclick="acceptToolPreview('${previewId}')">
                        接受
                    </button>
                    <button class="tool-action-btn reject" title="拒绝修改" onclick="rejectToolPreview('${previewId}')">
                        拒绝
                    </button>
                </div>
            </div>
            
            <div class="tool-preview-content">
                ${previewContentHtml}
            </div>
        </div>
    `;
}

// 切换详细预览内容的展开/折叠状态
function toggleDetailedPreview(event) {
    event.stopPropagation(); // 防止事件冒泡

    const separator = event.currentTarget;
    const toggleIcon = separator.querySelector('.preview-toggle-icon');
    const wrapper = separator.nextElementSibling;

    if (!wrapper || !wrapper.classList.contains('preview-rendered-wrapper')) return;

    if (wrapper.style.display === 'none') {
        // 展开
        wrapper.style.display = 'block';
        if (toggleIcon) {
            toggleIcon.classList.remove('collapsed');
        }

        // 在展开时触发渲染，确保MathJax在可见状态下排版更稳定
        try {
            const previewElement = separator.closest('.tool-preview-container, .modern-inline-preview, .inline-preview');
            if (previewElement) {
                renderPreviewContent(previewElement);
            }
        } catch (e) {
            console.warn('展开预览时渲染失败:', e);
        }
    } else {
        // 折叠
        wrapper.style.display = 'none';
        if (toggleIcon) {
            toggleIcon.classList.add('collapsed');
        }
    }
}

// 注意：ReAct思考过程现在由模型生成，不再在前端硬编码

// 生成工具参数显示
function generateToolParameters(data) {
    let parameters = [];

    if (data.action_type === 'insert_content') {
        if (data.target_heading) parameters.push(`目标标题: ${data.target_heading}`);
        if (data.format_type) parameters.push(`格式类型: ${getToolName(data.format_type)}`);
        if (data.indent_level !== undefined) parameters.push(`缩进级别: ${data.indent_level}`);
        if (data.add_spacing !== undefined) parameters.push(`添加间距: ${data.add_spacing ? '是' : '否'}`);
    } else if (data.action_type === 'modify_style') {
        if (data.style_parameters) {
            Object.entries(data.style_parameters).forEach(([key, value]) => {
                if (value !== undefined && value !== null && value !== '') {
                    const displayName = getParameterDisplayName(key);
                    parameters.push(`${displayName}: ${value}`);
                }
            });
        }
    }

    if (parameters.length === 0) return '';

    return `
        <div class="tool-parameters">
            <div class="tool-parameters-title">调用参数</div>
            <div class="tool-parameters-list">
                ${parameters.join('<br>')}
            </div>
        </div>
    `;
}

// 已移除工具详情网格功能，简化界面

// 渲染预览内容（支持MathJax公式、Mermaid流程图等）
function renderPreviewContent(previewElement) {
    if (!previewElement) return;

    try {
        // 查找预览渲染区域
        const renderedContent = previewElement.querySelector('.preview-rendered-content');
        if (!renderedContent) {
            console.log('预览中无需渲染的内容');
            return;
        }

        console.log('开始渲染预览内容中的公式和流程图...');

        // 1. 处理MathJax公式
        if (typeof MathJax !== 'undefined') {
            MathJax.Hub.Queue(["Typeset", MathJax.Hub, renderedContent]);

            // 渲染完成后添加公式工具栏（提示信息）
            MathJax.Hub.Queue(function () {
                renderedContent.querySelectorAll('script[type^="math/tex"]').forEach((script) => {
                    if (!script.parentElement.classList.contains('equation-container') &&
                        !script.hasAttribute('data-preview-processed')) {

                        const formula = script.textContent.trim();
                        if (formula) {
                            const container = document.createElement('div');
                            container.className = 'equation-container';

                            const toolbar = document.createElement('div');
                            toolbar.className = 'math-toolbar';
                            toolbar.innerHTML = `
                                <p>数学公式</p>
                                <div>
                                    <span style="color: #666; font-size: 12px;">接受后可插入到Word</span>
                                </div>
                            `;

                            if (script.parentNode) {
                                script.parentNode.insertBefore(container, script);
                                container.appendChild(toolbar);
                                container.appendChild(script);
                                script.setAttribute('data-preview-processed', 'true');
                            }
                        }
                    }
                });
                console.log('预览中的公式渲染完成');
            });
        }

        // 2. 处理Mermaid流程图
        if (typeof mermaid !== 'undefined') {
            const mermaidElements = renderedContent.querySelectorAll('.mermaid');
            if (mermaidElements.length > 0) {
                console.log(`发现 ${mermaidElements.length} 个Mermaid流程图，开始渲染...`);

                mermaidElements.forEach((element, index) => {
                    // 为每个Mermaid元素添加唯一ID
                    if (!element.id) {
                        element.id = `preview-mermaid-${Date.now()}-${index}`;
                    }

                    // 添加提示工具栏
                    if (!element.closest('.mermaid-container')) {
                        const container = document.createElement('div');
                        container.className = 'mermaid-container';

                        const toolbar = document.createElement('div');
                        toolbar.className = 'mermaid-toolbar';
                        toolbar.innerHTML = `
                            <p>流程图</p>
                            <div>
                                <span style="color: #666; font-size: 12px;">接受后可插入到Word</span>
                            </div>
                        `;

                        element.parentNode.insertBefore(container, element);
                        container.appendChild(toolbar);
                        container.appendChild(element);
                    }
                });

                // 初始化Mermaid渲染
                try {
                    mermaid.init(undefined, renderedContent.querySelectorAll('.mermaid'));
                    console.log('预览中的Mermaid流程图渲染完成');
                } catch (error) {
                    console.error('Mermaid渲染失败:', error);
                }
            }
        }

        console.log('预览内容渲染完成');

    } catch (error) {
        console.error('渲染预览内容时出错:', error);
    }
}

// 绑定工具预览事件处理器
function bindToolPreviewEvents(element, previewId, actionData) {
    if (!element) return;

    // 存储actionData到元素上，供后续使用
    element.setAttribute('data-action-data', JSON.stringify(actionData));

    // 绑定"样式修改"设置复选框变更 -> 同步到 actionData.parameters
    try {
        const checkboxes = element.querySelectorAll('input.ms-setting[data-setting-checkbox], input.insert-setting[data-setting-checkbox]');
        if (checkboxes && checkboxes.length > 0) {
            checkboxes.forEach(cb => {
                cb.addEventListener('change', () => {
                    try {
                        const key = cb.getAttribute('data-setting-checkbox');
                        const str = element.getAttribute('data-action-data') || '{}';
                        const ad = JSON.parse(str);
                        ad.parameters = ad.parameters || {};
                        ad.parameters[key] = cb.checked;
                        element.setAttribute('data-action-data', JSON.stringify(ad));
                        console.log('✅ 已更新设置:', key, cb.checked);
                    } catch (e) {
                        console.warn('更新设置失败:', e);
                    }
                });
            });
        }
    } catch (e) {
        console.warn('绑定复选框事件失败:', e);
    }

    // 绑定"插入设置"下拉框变更 -> 同步到 actionData.parameters
    try {
        const selects = element.querySelectorAll('select.insert-setting[data-setting-select]');
        if (selects && selects.length > 0) {
            selects.forEach(sel => {
                sel.addEventListener('change', () => {
                    try {
                        const key = sel.getAttribute('data-setting-select');
                        const str = element.getAttribute('data-action-data') || '{}';
                        const ad = JSON.parse(str);
                        ad.parameters = ad.parameters || {};
                        ad.parameters[key] = sel.value;
                        element.setAttribute('data-action-data', JSON.stringify(ad));
                        console.log('✅ 已更新设置:', key, sel.value);
                    } catch (e) {
                        console.warn('更新设置失败:', e);
                    }
                });
            });
        }
    } catch (e) {
        console.warn('绑定下拉框事件失败:', e);
    }

    console.log(`已绑定工具预览事件: ${previewId}`);
}

// 已移除工具详情切换功能，简化界面

// 接受工具预览
function acceptToolPreview(previewId) {
    console.log(`接受工具预览: ${previewId}`);

    const container = document.querySelector(`[data-preview-id="${previewId}"]`);
    if (!container) {
        console.error('未找到预览容器');
        return;
    }

    const actionDataStr = container.getAttribute('data-action-data');
    if (!actionDataStr) {
        console.error('未找到操作数据');
        return;
    }

    try {
        const actionData = JSON.parse(actionDataStr);

        // 从DOM控件补齐/覆盖设置，保证与UI一致
        try {
            const trimEl = container.querySelector('input.insert-setting[data-setting-checkbox="trim_spaces"], input.ms-setting[data-setting-checkbox="trim_spaces"]');
            const posEl = container.querySelector('select.insert-setting[data-setting-select="insert_position"], select.ms-setting[data-setting-select="insert_position"]');
            actionData.parameters = actionData.parameters || {};
            if (trimEl) actionData.parameters.trim_spaces = !!trimEl.checked;
            if (posEl && posEl.value) actionData.parameters.insert_position = posEl.value;
        } catch (_) { }

        // 先打印完整的 actionData（用于调试）
        console.log('🔍 读取到的 actionData:', JSON.stringify(actionData, null, 2));

        // 记录用户决策：接受
        recordPreviewDecision(actionData.action_type, 'accepted', previewId);

        // 打印样式修改的作用范围参数（用于调试）
        if (actionData.action_type === 'modify_style') {
            console.log('📋 样式修改参数:', {
                include_paragraphs: actionData.parameters.include_paragraphs,
                include_headings: actionData.parameters.include_headings,
                include_tables: actionData.parameters.include_tables,
                include_list_items: actionData.parameters.include_list_items
            });
        }

        // 发送应用操作请求到C#
        const messageData = {
            type: 'applyPreviewedAction',
            action_type: actionData.action_type,
            parameters: actionData.parameters,
            preview_id: previewId
        };

        sendMessageToCSharp(messageData);

        // 更新预览管理器状态
        previewManager.updateStatus(previewId, 'applying', '正在应用操作...');

        // 立即检查并更新批量操作按钮
        setTimeout(() => {
            showFloatingBatchActions();
        }, 50);

        console.log(`工具预览接受请求已发送: ${previewId}`);

    } catch (error) {
        console.error('解析操作数据失败:', error);
    }
}

// 拒绝工具预览
function rejectToolPreview(previewId) {
    console.log(`拒绝工具预览: ${previewId}`);

    // 获取预览信息用于发送给后端
    const container = document.querySelector(`[data-preview-id="${previewId}"]`);
    let actionType = "unknown";
    if (container) {
        const actionDataStr = container.getAttribute('data-action-data');
        if (actionDataStr) {
            try {
                const actionData = JSON.parse(actionDataStr);
                actionType = actionData.action_type || "unknown";
                // 记录用户决策：拒绝
                recordPreviewDecision(actionType, 'rejected', previewId);
            } catch (error) {
                console.error('解析操作数据失败:', error);
            }
        }
    }

    // 发送拒绝操作通知到后端
    const messageData = {
        type: 'rejectPreviewedAction',
        action_type: actionType,
        preview_id: previewId
    };

    sendMessageToCSharp(messageData);

    // 直接更新预览状态为已拒绝，保留卡片并替换为小字状态
    if (previewManager) {
        previewManager.updateStatus(previewId, 'rejected', '已拒绝');
    }

    // 立即刷新批量操作弹窗（被拒绝不再计入待处理）
    setTimeout(() => {
        showFloatingBatchActions();
    }, 50);
}

// 切换预览内容的显示/隐藏
function togglePreviewContent(previewId) {
    const content = document.getElementById(`preview-content-${previewId}`);
    const toggleBtn = document.querySelector(`.preview-toggle-btn[data-preview-id="${previewId}"] .toggle-icon`);

    if (!content || !toggleBtn) return;

    if (content.classList.contains('collapsed')) {
        // 展开
        content.classList.remove('collapsed');
        content.classList.add('expanded');
        toggleBtn.textContent = '▲';
        toggleBtn.style.transform = 'rotate(180deg)';
    } else {
        // 折叠
        content.classList.remove('expanded');
        content.classList.add('collapsed');
        toggleBtn.textContent = '▼';
        toggleBtn.style.transform = 'rotate(0deg)';
    }
}

// 获取工具图标
function getToolIcon(actionType) {
    const icons = {
        'insert_content': '📝',
        'modify_style': '🎨',
        'get_heading_content': '📋',
        'get_document_statistics': '📊',
        'get_document_images': '🖼️',
        'get_document_formulas': '🧮',
        'get_selected_text': '📄'
    };
    return icons[actionType] || '🔧';
}

// 获取工具名称
function getToolName(actionType) {
    const names = {
        'insert_content': '插入内容',
        'modify_style': '样式修改',
        'get_heading_content': '获取标题内容',
        'get_document_statistics': '文档统计',
        'get_document_images': '获取图片',
        'get_document_formulas': '获取公式',
        'get_selected_text': '获取选中文本',
        'paragraph': '段落',
        'list': '列表',
        'table': '表格',
        'emphasis': '强调'
    };
    return names[actionType] || actionType;
}

// 获取参数显示名称
function getParameterDisplayName(paramKey) {
    const displayNames = {
        'text_to_find': '查找文本',
        'font_size': '字体大小',
        'font_bold': '粗体',
        'font_italic': '斜体',
        'font_color': '字体颜色',
        'background_color': '背景颜色',
        'paragraph_spacing_before': '段前间距',
        'paragraph_spacing_after': '段后间距',
        'line_spacing': '行间距'
    };
    return displayNames[paramKey] || paramKey;
}

// ==================== ReAct 内容解析功能 ====================

// 解析并转换ReAct标记为可视化块
function parseReActContent(content) {
    if (!content || typeof content !== 'string') return content;

    // 解析THINKING标记
    content = content.replace(
        /<THINKING>([\s\S]*?)<\/THINKING>/gi,
        (match, thinking) => {
            return `<div class="react-thinking">
                <div class="react-thinking-content">${thinking.trim()}</div>
            </div>`;
        }
    );

    // 解析OBSERVATION标记
    content = content.replace(
        /<OBSERVATION>([\s\S]*?)<\/OBSERVATION>/gi,
        (match, observation) => {
            return `<div class="react-observation">
                <div class="react-observation-content">${observation.trim()}</div>
            </div>`;
        }
    );

    // 解析ACTION标记
    content = content.replace(
        /<ACTION>([\s\S]*?)<\/ACTION>/gi,
        (match, action) => {
            return `<div class="react-action">
                <div class="react-action-content">${action.trim()}</div>
            </div>`;
        }
    );

    return content;
}

// 检查内容是否包含ReAct标记
function hasReActContent(content) {
    if (!content || typeof content !== 'string') return false;

    return content.includes('<THINKING>') ||
        content.includes('<OBSERVATION>') ||
        content.includes('<ACTION>');
}

// 接受内联预览的操作
function acceptInlinePreview(previewId) {
    // 清除预览待处理标志，允许继续正常流程
    isPreviewPending = false;
    console.log('已清除预览待处理标志，恢复正常流程');

    if (!currentPreviewedAction) {
        console.log('没有预览操作可以接受');
        return;
    }

    console.log('接受内联预览操作:', currentPreviewedAction);

    // 禁用预览按钮，显示加载状态
    const previewElement = document.getElementById(previewId);
    if (previewElement) {
        const buttons = previewElement.querySelectorAll('.preview-btn');
        buttons.forEach(btn => {
            btn.disabled = true;
            if (btn.classList.contains('accept-btn')) {
                btn.innerHTML = '<span class="icon">⏳</span><span>正在应用...</span>';
            }
        });
    }

    // 发送应用操作请求到C#
    const messageData = {
        type: 'applyPreviewedAction',
        action_type: currentPreviewedAction.action_type,
        parameters: currentPreviewedAction.parameters
    };

    sendMessageToCSharp(messageData);
}

// 拒绝内联预览的操作
function rejectInlinePreview(previewId) {
    console.log('拒绝内联预览操作');

    // 清除预览待处理标志，允许继续正常流程
    isPreviewPending = false;
    console.log('已清除预览待处理标志，恢复正常流程');

    const previewElement = document.getElementById(previewId);
    if (previewElement) {
        // 添加拒绝动画
        previewElement.style.transition = 'all 0.3s ease';
        previewElement.style.opacity = '0.5';
        previewElement.style.transform = 'scale(0.98)';

        // 显示拒绝状态
        const header = previewElement.querySelector('.preview-title');
        if (header) {
            header.innerHTML = '<span class="icon">❌</span><span>操作已拒绝</span>';
        }

        const actions = previewElement.querySelector('.preview-actions');
        if (actions) {
            actions.style.display = 'none';
        }

        // 2秒后优雅地移除
        setTimeout(() => {
            previewElement.style.transition = 'all 0.5s ease';
            previewElement.style.opacity = '0';
            previewElement.style.transform = 'scale(0.95)';
            previewElement.style.maxHeight = '0';
            previewElement.style.margin = '0';
            previewElement.style.padding = '0';
            previewElement.style.overflow = 'hidden';

            setTimeout(() => {
                // 确保父容器布局正确
                const parentElement = previewElement.parentElement;
                previewElement.remove();

                // 清理父容器中可能残留的空白
                if (parentElement) {
                    // 如果父容器是消息容器且变成空的，清理它
                    if (parentElement.classList.contains('assistant-message') &&
                        parentElement.textContent.trim() === '') {
                        parentElement.style.display = 'none';
                        setTimeout(() => {
                            parentElement.remove();
                        }, 100);
                    }
                }
            }, 500);
        }, 1500);
    }

    // 清除当前预览操作
    currentPreviewedAction = null;
}

// 隐藏预览面板（保留向后兼容性，但现在不需要了）
function hidePreviewPanel() {
    // 现在使用内联预览，此函数主要用于向后兼容
    if (previewActionPanel) {
        previewActionPanel.style.display = 'none';
    }
    console.log('旧版预览面板已隐藏');
}

// 处理模型列表数据
function handleModelList(models) {
    console.log('收到模型列表:', models);

    if (!modelSelect) return;

    // 清空现有选项
    modelSelect.innerHTML = '';

    // 保存模型列表
    availableModels = models;

    if (models && models.length > 0) {
        // 添加模型选项
        models.forEach(model => {
            const option = document.createElement('option');
            option.value = model.id;
            option.textContent = `${model.name} (${model.modelName})`;

            // 构建详细的title提示信息
            let titleParts = [`模板: ${model.template}`, `地址: ${model.baseUrl}`];
            if (model.enableTools !== undefined) {
                titleParts.push(`工具调用: ${model.enableTools === 1 ? '✓ 已启用' : '✗ 未启用'}`);
            }
            option.title = titleParts.join(', ');

            modelSelect.appendChild(option);
        });

        // 尝试恢复之前保存的模型选择
        const savedModelId = parseInt(localStorage.getItem('selectedModelId')) || 0;
        const savedModel = models.find(m => m.id === savedModelId);

        if (savedModel) {
            selectedModelId = savedModelId;
            modelSelect.value = selectedModelId;
            console.log('恢复保存的模型选择:', savedModel.name, 'ID:', selectedModelId);
        } else if (models.length > 0) {
            // 如果没有保存的选择或找不到保存的模型，选择第一个
            selectedModelId = models[0].id;
            modelSelect.value = selectedModelId;
            localStorage.setItem('selectedModelId', selectedModelId.toString());
            console.log('默认选择模型:', models[0].name, 'ID:', selectedModelId);
        }

        // 检查当前选中的模型是否支持工具调用（如果在Agent模式下）
        checkCurrentModelSupportsTools();
    } else {
        // 没有模型时显示提示
        const option = document.createElement('option');
        option.value = '0';
        option.textContent = '暂无可用模型';
        option.disabled = true;
        modelSelect.appendChild(option);

        console.log('没有可用的对话模型');
    }
}

// 获取当前选中的模型信息
function getSelectedModelInfo() {
    if (selectedModelId > 0 && availableModels.length > 0) {
        return availableModels.find(model => model.id === selectedModelId);
    }
    return null;
}

// 检查当前选中的模型是否支持工具调用
function checkCurrentModelSupportsTools() {
    const currentModel = getSelectedModelInfo();
    if (!currentModel) {
        console.log('未找到当前选中的模型');
        return true; // 未找到时不阻止
    }

    const supportsTools = currentModel.enableTools === 1;
    console.log(`当前模型 "${currentModel.name}" 工具调用支持:`, supportsTools ? '已启用' : '未启用');

    // 如果当前是Agent模式且模型不支持工具
    if (currentChatMode === 'chat-agent' && !supportsTools) {
        // 自动切换回智能问答模式
        switchToChatMode(currentModel.name);
        return false;
    }

    return supportsTools;
}

// 自动切换到智能问答模式并显示提示
function switchToChatMode(modelName) {
    console.log('模型不支持工具调用，自动切换到智能问答模式');

    // 切换聊天模式选择器
    if (chatModeSelect && chatModeSelect.value === 'chat-agent') {
        chatModeSelect.value = 'chat';
        currentChatMode = 'chat';

        // 更新工具设置按钮可见性
        updateToolsSettingsVisibility();

        console.log('已自动切换到智能问答模式');
    }

    // 显示警告提示
    showModelToolsWarning(modelName, true);
}

// 显示模型不支持工具调用的警告
function showModelToolsWarning(modelName, autoSwitched = false) {
    const warningDiv = document.createElement('div');
    warningDiv.className = 'model-tools-warning';

    // 根据是否自动切换显示不同的消息
    const messageText = autoSwitched
        ? `<p>当前选择的模型 "${modelName}" 未启用工具调用功能，无法使用智能体（Agent）模式。</p>
           <p>已自动切换为"智能问答"模式。如需使用Agent模式，请切换到支持工具调用的模型。</p>`
        : `<p>当前选择的模型 "${modelName}" 未启用工具调用功能，无法使用智能体（Agent）模式。</p>
           <p>请切换到其他启用了工具调用的模型，或切换回"智能问答"模式。</p>`;

    warningDiv.innerHTML = `
        <div class="warning-content">
            <div class="warning-icon">⚠️</div>
            <div class="warning-text">
                <strong>模型不支持工具调用</strong>
                ${messageText}
            </div>
            <button class="warning-close" onclick="this.parentElement.parentElement.remove()">×</button>
        </div>
    `;

    // 添加样式
    if (!document.getElementById('model-tools-warning-style')) {
        const style = document.createElement('style');
        style.id = 'model-tools-warning-style';
        style.textContent = `
            .model-tools-warning {
                position: fixed;
                top: 20px;
                right: 20px;
                max-width: 400px;
                z-index: 10000;
                animation: slideInRight 0.3s ease;
            }
            
            .model-tools-warning .warning-content {
                background: linear-gradient(135deg, #fff3cd 0%, #ffe69c 100%);
                border: 2px solid #ffc107;
                border-radius: 12px;
                padding: 16px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                display: flex;
                gap: 12px;
                position: relative;
            }
            
            .model-tools-warning .warning-icon {
                font-size: 24px;
                flex-shrink: 0;
            }
            
            .model-tools-warning .warning-text {
                flex: 1;
            }
            
            .model-tools-warning .warning-text strong {
                display: block;
                color: #856404;
                font-size: 16px;
                margin-bottom: 8px;
            }
            
            .model-tools-warning .warning-text p {
                color: #856404;
                font-size: 13px;
                line-height: 1.5;
                margin: 4px 0;
            }
            
            .model-tools-warning .warning-close {
                position: absolute;
                top: 8px;
                right: 8px;
                background: none;
                border: none;
                font-size: 24px;
                color: #856404;
                cursor: pointer;
                padding: 0;
                width: 30px;
                height: 30px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: all 0.2s ease;
            }
            
            .model-tools-warning .warning-close:hover {
                background-color: rgba(133, 100, 4, 0.1);
            }
            
            @keyframes slideInRight {
                from {
                    transform: translateX(100%);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
        `;
        document.head.appendChild(style);
    }

    // 移除已存在的警告
    const existingWarning = document.querySelector('.model-tools-warning');
    if (existingWarning) {
        existingWarning.remove();
    }

    document.body.appendChild(warningDiv);

    // 5秒后自动移除
    setTimeout(() => {
        if (warningDiv.parentElement) {
            warningDiv.style.animation = 'slideOutRight 0.3s ease';
            setTimeout(() => {
                if (warningDiv.parentElement) {
                    warningDiv.remove();
                }
            }, 300);
        }
    }, 5000);
}

// 添加slideOutRight动画
if (!document.getElementById('model-tools-warning-animations')) {
    const style = document.createElement('style');
    style.id = 'model-tools-warning-animations';
    style.textContent = `
        @keyframes slideOutRight {
            from {
                transform: translateX(0);
                opacity: 1;
            }
            to {
                transform: translateX(100%);
                opacity: 0;
            }
        }
    `;
    document.head.appendChild(style);
}



// 智能体模式密码保护功能
let isAgentModeUnlocked = true;
const AGENT_PASSWORD = '123456';

// 初始化智能体模式锁定状态
function initializeAgentModeLock() {
    const chatModeSelect = document.getElementById('chat-mode');
    if (chatModeSelect) {
        // 确保智能体选项可用
        const agentOption = chatModeSelect.querySelector('option[value="chat-agent"]');
        if (agentOption) {
            agentOption.disabled = false;
            agentOption.textContent = '智能体（Chat+Agent模式）';
        }
    }
    // 隐藏旧的解锁按钮（如果存在于旧版本模板）
    const unlockBtn = document.getElementById('unlock-agent-btn');
    if (unlockBtn) {
        unlockBtn.style.display = 'none';
    }
    // 默认已解锁
    isAgentModeUnlocked = true;
}

// 解锁智能体模式
function unlockAgentMode() {
    if (isAgentModeUnlocked) {
        // 已解锁，直接切换到智能体模式
        const chatModeSelect = document.getElementById('chat-mode');
        chatModeSelect.value = 'chat-agent';
        onChatModeChange();
        return;
    }

    // 创建密码输入模态框
    const modal = document.createElement('div');
    modal.className = 'password-modal';
    modal.innerHTML = `
        <div class="password-modal-content">
            <h3>🔐 解锁智能体模式</h3>
            <div class="warning">
                ⚠️ <strong>测试阶段提醒</strong><br>
                智能体模式目前仍在测试阶段，功能尚不完善，可能存在以下问题：<br>
                • 响应速度较慢<br>
                • 工具调用可能不准确<br>
                • 偶尔出现错误或异常<br>
                建议您使用前先保存Word文档，并谨慎使用该功能。
            </div>
            <input type="password" id="agent-password" placeholder="请输入测试密码" maxlength="20">
            <div class="password-modal-buttons">
                <button class="btn-cancel" onclick="closePasswordModal()">取消</button>
                <button class="btn-confirm" onclick="verifyPassword()">确认</button>
            </div>
        </div>
    `;

    document.body.appendChild(modal);

    // 聚焦到密码输入框
    const passwordInput = document.getElementById('agent-password');
    passwordInput.focus();

    // 添加回车键支持
    passwordInput.addEventListener('keypress', function (e) {
        if (e.key === 'Enter') {
            verifyPassword();
        }
    });

    // 添加ESC键关闭支持
    modal.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            closePasswordModal();
        }
    });
}

// 验证密码
function verifyPassword() {
    const passwordInput = document.getElementById('agent-password');
    const password = passwordInput.value.trim();

    if (password === AGENT_PASSWORD) {
        // 密码正确，解锁智能体模式
        isAgentModeUnlocked = true;

        // 启用智能体选项
        const chatModeSelect = document.getElementById('chat-mode');
        const agentOption = chatModeSelect.querySelector('option[value="chat-agent"]');
        agentOption.disabled = false;
        agentOption.textContent = '智能体（Chat+Agent模式）✅';

        // 更新解锁按钮
        const unlockBtn = document.getElementById('unlock-agent-btn');
        unlockBtn.textContent = '✅ 已解锁';
        unlockBtn.style.background = 'linear-gradient(45deg, #10b981, #059669)';
        unlockBtn.title = '智能体模式已解锁';

        // 自动切换到智能体模式
        chatModeSelect.value = 'chat-agent';
        onChatModeChange();

        // 显示成功提示和详细说明
        showMessage('智能体模式已解锁！当前为测试版本，体验可能不佳，请谨慎使用。', 'success');

        // 延迟显示欢迎消息
        setTimeout(() => {
            addAssistantMessage('🎉 欢迎使用智能体模式！\n\n⚠️ **测试版本说明**：\n- 当前版本仍在开发测试中，功能尚不完善\n- 响应速度可能较慢，请耐心等待\n- 工具调用可能偶尔出现错误\n- 建议在使用前保存Word文档\n\n您可以尝试以下操作：\n- 获取文档信息和统计数据\n- 插入格式化内容到指定位置\n- 修改文字样式和格式\n- 提取文档中的图片、表格、公式信息\n\n有任何问题请切换回"智能问答模式"使用基础功能。');
        }, 1000);

        closePasswordModal();
    } else {
        // 密码错误
        passwordInput.style.borderColor = '#ef4444';
        passwordInput.style.boxShadow = '0 0 0 3px rgba(239, 68, 68, 0.1)';
        showMessage('密码错误，请重试。', 'error');

        // 清空输入并重新聚焦
        passwordInput.value = '';
        passwordInput.focus();

        // 1秒后恢复输入框样式
        setTimeout(() => {
            passwordInput.style.borderColor = '';
            passwordInput.style.boxShadow = '';
        }, 1000);
    }
}

// 关闭密码模态框
function closePasswordModal() {
    const modal = document.querySelector('.password-modal');
    if (modal) {
        modal.remove();
    }
}

// 显示消息提示
function showMessage(message, type = 'info') {
    const messageDiv = document.createElement('div');
    messageDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 6px;
        color: white;
        font-size: 14px;
        z-index: 10001;
        animation: slideInRight 0.3s ease;
        max-width: 300px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    `;

    switch (type) {
        case 'success':
            messageDiv.style.background = 'linear-gradient(45deg, #10b981, #059669)';
            break;
        case 'error':
            messageDiv.style.background = 'linear-gradient(45deg, #ef4444, #dc2626)';
            break;
        default:
            messageDiv.style.background = 'linear-gradient(45deg, #6b7280, #4b5563)';
    }

    messageDiv.textContent = message;
    document.body.appendChild(messageDiv);

    // 3秒后自动移除
    setTimeout(() => {
        messageDiv.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.parentNode.removeChild(messageDiv);
            }
        }, 300);
    }, 3000);
}

// 添加动画样式
if (!document.querySelector('#agent-unlock-animations')) {
    const style = document.createElement('style');
    style.id = 'agent-unlock-animations';
    style.textContent = `
        @keyframes slideInRight {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }
        @keyframes slideOutRight {
            from { transform: translateX(0); opacity: 1; }
            to { transform: translateX(100%); opacity: 0; }
        }
    `;
    document.head.appendChild(style);
}

// 全局函数，供HTML调用
window.hidePreviewPanel = hidePreviewPanel;
window.acceptPreviewedAction = acceptPreviewedAction;
window.rejectPreviewedAction = rejectPreviewedAction;
window.acceptInlinePreview = acceptInlinePreview;
window.rejectInlinePreview = rejectInlinePreview;
window.showInlinePreviewDemo = showInlinePreviewDemo;
window.testInlinePreviewFixed = testInlinePreviewFixed;
window.testPreviewFlow = testPreviewFlow;
// 现代化预览函数
window.acceptModernPreview = acceptModernPreview;
window.rejectModernPreview = rejectModernPreview;
window.acceptAllPreviews = acceptAllPreviews;
window.rejectAllPreviews = rejectAllPreviews;
window.testModernPreviewSystem = testModernPreviewSystem;
window.showNewPreviewFeatures = showNewPreviewFeatures;
window.testAllFixes = testAllFixes;
window.testFloatingButton = testFloatingButton;
window.testSinglePreview = testSinglePreview;

// 显示修复说明
function showFixInfo() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">🔧 悬浮按钮修复完成</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">✅ 问题解决：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>显示门槛降低</strong>：1个预览也会显示悬浮按钮</li>
                    <li><strong>按钮文案智能</strong>：单个预览显示"接受"/"拒绝"</li>
                    <li><strong>位置优化</strong>：输入框右上角，不遮挡内容</li>
                    <li><strong>即时响应</strong>：预览创建后立即显示按钮</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 8px; margin-top: 12px; font-size: 13px; opacity: 0.9;">
                                 <strong>🧪 测试命令：</strong><br>
                • <code>testTableInsertFix()</code> - 表格插入修复测试 🔥<br>
                • <code>testMarkdownHeaders()</code> - Markdown标题转换测试<br>
                • <code>testSimpleBatch()</code> - 简化批量测试<br>
                • <code>testAllFixes()</code> - 完整功能测试
            </div>
        </div>
    `);
}

window.showFixInfo = showFixInfo;

// 显示Markdown标题修复说明
function showMarkdownHeaderFix() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">✅ Markdown标题转换修复完成</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">🔧 修复详情：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>问题</strong>：# 符号显示为纯文本，未转换为Word标题</li>
                    <li><strong>原因</strong>：ConvertMarkdownToHtml 函数缺少标题转换逻辑</li>
                    <li><strong>修复</strong>：添加了完整的 # → &lt;h1&gt; 转换规则</li>
                    <li><strong>支持</strong>：1-6级标题 (# ## ### #### ##### ######)</li>
                    <li><strong>效果</strong>：Word会自动应用对应的标题样式</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 8px; margin-top: 12px;">
                <h3 style="margin: 0 0 8px 0;">🔄 转换流程：</h3>
                <div style="font-size: 13px; font-family: monospace; line-height: 1.4;">
                    <strong>1.</strong> Markdown: <code># 标题</code><br>
                    <strong>2.</strong> HTML: <code>&lt;h1&gt;标题&lt;/h1&gt;</code><br>
                    <strong>3.</strong> Word: <span style="background: rgba(255,255,255,0.2); padding: 2px 4px; border-radius: 3px;">标题 1 样式</span>
                </div>
            </div>
        </div>
    `);
}

window.showMarkdownHeaderFix = showMarkdownHeaderFix;

// 显示表格插入修复说明
function showTableInsertFix() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">✅ 表格插入修复完成</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">🐛 修复的问题：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>列对齐错误</strong>：数据向右偏移一列，第一列变空</li>
                    <li><strong>粗体格式丢失</strong>：**文本** 语法未转换为粗体</li>
                    <li><strong>解析不准确</strong>：边界|符号处理不当</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px; margin-top: 12px;">
                <h3 style="margin: 0 0 12px 0;">🔧 修复方案：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>SplitTableRow</strong>：智能分割表格行，正确处理边界|</li>
                    <li><strong>ProcessCellMarkdown</strong>：清理Markdown语法，保留格式信息</li>
                    <li><strong>SetCellContent</strong>：检测粗体标记并应用Word格式</li>
                    <li><strong>调试日志</strong>：详细记录解析过程便于排查</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 8px; margin-top: 12px;">
                <h3 style="margin: 0 0 8px 0;">📊 修复效果：</h3>
                <div style="font-size: 13px; line-height: 1.4;">
                    <strong>✅ 列对齐正确</strong>：每列数据在正确位置<br>
                    <strong>✅ 粗体显示</strong>：**文本** → <strong>粗体文本</strong><br>
                    <strong>✅ 格式规范</strong>：表头居中，数据左对齐<br>
                    <strong>✅ 边框完整</strong>：所有单元格都有边框
                </div>
            </div>
        </div>
    `);
}

window.showTableInsertFix = showTableInsertFix;

// 测试批量插入修复效果（改进版）
function testBatchInsertFix() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">🔧 批量插入修复测试 v2.0</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">✅ 改进内容：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>格式优化</strong>：修正换行符，改善内容结构</li>
                    <li><strong>位置指定</strong>：明确指定插入位置</li>
                    <li><strong>内容分层</strong>：标题、概述、详细对比清晰分离</li>
                    <li><strong>顺序处理</strong>：500ms间隔，避免并发冲突</li>
                </ul>
            </div>
        </div>
    `);

    // 创建改进的测试预览
    const previews = [];

    // 预览1：插入主标题（使用标题样式）
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📝 预览1 - 插入主标题'
    }, {
        action_type: 'insert_content',
        parameters: {
            target_heading: '', // 在当前光标位置插入
            content: 'MySQL 和 Redis 深度对比分析',
            format_type: 'emphasis', // 使用强调格式作为标题
            add_spacing: true
        }
    }));

    // 预览2：插入概述段落
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📄 预览2 - 插入概述内容'
    }, {
        action_type: 'insert_content',
        parameters: {
            target_heading: '',
            content: 'MySQL和Redis是两种不同类型的数据库系统，它们在数据存储、性能特点、适用场景等方面存在显著差异。本文将从多个维度深入分析这两种数据库的特点和适用场景。',
            format_type: 'paragraph',
            add_spacing: true
        }
    }));

    // 预览3：插入对比列表（修正格式）
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📋 预览3 - 插入详细对比'
    }, {
        action_type: 'insert_content',
        parameters: {
            target_heading: '',
            content: `数据模型差异
MySQL: 关系型数据库，使用表格结构存储数据，支持SQL查询
Redis: 键值存储数据库，数据以key-value形式存储，支持多种数据结构

性能特点
MySQL: 强一致性，支持ACID事务，保证数据一致性
Redis: 最终一致性，在某些配置下可能丢失数据

使用场景
MySQL: 适合存储结构化数据，如用户信息、订单数据等
Redis: 适合缓存、会话存储、消息队列、实时排行榜等场景

数据持久化
MySQL: 数据持久化到磁盘，重启后数据不丢失
Redis: 主要存储在内存中，可配置持久化策略`,
            format_type: 'list',
            add_spacing: true
        }
    }));

    setTimeout(() => {
        checkAndShowBatchActions();
        addAssistantMessage(`
            <div style="padding: 16px; background: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 10px; margin: 16px 0;">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                    <span style="font-size: 16px;">🚀</span>
                    <strong style="color: #0369a1;">改进版批量插入测试就绪</strong>
                </div>
                <p style="margin: 0; font-size: 13px; color: #0c4a6e; line-height: 1.4;">
                    📊 <strong>测试内容</strong>：主标题 + 概述段落 + 详细对比列表<br>
                    🎯 <strong>格式改进</strong>：正确的换行、间距和层次结构<br>
                    ⚡ <strong>处理方式</strong>：顺序处理，每个间隔500ms<br>
                    ✨ <strong>预期效果</strong>：格式规范、层次清晰、内容完整<br><br>
                    <strong>点击悬浮按钮的"全部接受"测试改进效果！</strong>
                </p>
            </div>
        `);
    }, 800);
}

// 简化版批量插入测试（用于快速验证）
function testSimpleBatch() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 16px; background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; border-radius: 12px; text-align: center;">
            <h3 style="margin: 0 0 8px 0;">⚡ 简化批量测试</h3>
            <p style="margin: 0; font-size: 14px;">3个简单段落，验证顺序处理效果</p>
        </div>
    `);

    // 创建3个简单的段落预览
    for (let i = 1; i <= 3; i++) {
        previewManager.createPreview({
            action_type: 'insert_content',
            message: `段落 ${i} - 简单测试内容`
        }, {
            action_type: 'insert_content',
            parameters: {
                content: `第${i}段：这是第${i}个测试段落，用于验证批量插入的顺序处理效果。内容会按照预定顺序依次插入到Word文档中。`,
                format_type: 'paragraph',
                add_spacing: true
            }
        });
    }

    setTimeout(() => {
        checkAndShowBatchActions();
        addAssistantMessage(`
            <div style="padding: 12px; background: #f0fdf4; border: 1px solid #10b981; border-radius: 8px; margin: 16px 0; text-align: center;">
                <strong style="color: #059669;">✅ 简化测试就绪 - 3个段落等待批量处理</strong>
            </div>
        `);
    }, 500);
}

window.testBatchInsertFix = testBatchInsertFix;
window.testSimpleBatch = testSimpleBatch;

// 测试Markdown标题转换功能
function testMarkdownHeaders() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">📝 Markdown标题转换测试</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">🔧 修复内容：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>标题识别</strong>：正确识别 # ## ### 等Markdown标题</li>
                    <li><strong>HTML转换</strong>：转换为 &lt;h1&gt; &lt;h2&gt; &lt;h3&gt; 标签</li>
                    <li><strong>Word样式</strong>：Word会自动应用对应的标题样式</li>
                    <li><strong>层级支持</strong>：支持1-6级标题转换</li>
                </ul>
            </div>
        </div>
    `);

    // 创建多级标题测试预览
    const previews = [];

    // 预览1：一级标题
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📝 预览1 - 一级标题'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: '# MySQL 和 Redis 深度对比分析',
            format_type: 'paragraph',
            add_spacing: true
        }
    }));

    // 预览2：二级标题
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📝 预览2 - 二级标题'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: '## 数据模型对比',
            format_type: 'paragraph',
            add_spacing: true
        }
    }));

    // 预览3：三级标题 + 内容
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📝 预览3 - 三级标题和内容'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: `### MySQL特点

**MySQL** 是一个关系型数据库管理系统，具有以下特点：

- 基于表格的关系模型
- 支持复杂的SQL查询
- 严格的数据结构和约束
- 支持事务和ACID特性`,
            format_type: 'paragraph',
            add_spacing: true
        }
    }));

    // 预览4：四级标题 + Redis内容
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📝 预览4 - Redis特点'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: `#### Redis特点

**Redis** 是一个内存数据结构存储系统：

- 键值对存储模型
- 支持多种数据结构：字符串、列表、集合、有序集合、哈希等
- 高性能内存操作
- 支持数据持久化`,
            format_type: 'paragraph',
            add_spacing: true
        }
    }));

    setTimeout(() => {
        checkAndShowBatchActions();
        addAssistantMessage(`
            <div style="padding: 16px; background: #faf5ff; border: 1px solid #8b5cf6; border-radius: 10px; margin: 16px 0;">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                    <span style="font-size: 16px;">✨</span>
                    <strong style="color: #7c3aed;">Markdown标题转换测试就绪</strong>
                </div>
                <p style="margin: 0; font-size: 13px; color: #581c87; line-height: 1.4;">
                    📊 <strong>测试内容</strong>：# 一级标题、## 二级标题、### 三级标题、#### 四级标题<br>
                    🔄 <strong>转换过程</strong>：Markdown → HTML → Word标题样式<br>
                    ✨ <strong>预期效果</strong>：Word中显示为正确的标题格式，而非纯文本<br><br>
                    <strong>点击"全部接受"查看标题转换效果！</strong>
                </p>
            </div>
        `);
    }, 800);
}

window.testMarkdownHeaders = testMarkdownHeaders;

// 测试表格插入修复功能
function testTableInsertFix() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">📊 表格插入修复测试</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">🔧 修复内容：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>列对齐修复</strong>：解决数据向右偏移一列的问题</li>
                    <li><strong>粗体格式</strong>：正确处理 **文本** 粗体语法</li>
                    <li><strong>智能分割</strong>：改进表格行解析，处理边界|符号</li>
                    <li><strong>调试信息</strong>：增加详细日志便于排查问题</li>
                </ul>
            </div>
        </div>
    `);

    // 创建表格测试预览
    const previews = [];

    // 预览1：标准Markdown表格
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📊 预览1 - 标准Markdown表格'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: `| **对比维度** | **MySQL** | **Redis** |
|---------|--------|-------|
| **数据库类型** | 关系型数据库(RDBMS) | 非关系型数据库(NoSQL) |
| **数据模型** | 表格结构，支持复杂关系 | 键值对存储，支持多种数据结构 |
| **存储介质** | 主要存储在磁盘 | 主要存储在内存 |
| **查询语言** | SQL | 专用命令 |
| **事务支持** | 完整 ACID 事务支持 | 有限的事务支持 |`,
            format_type: 'table',
            add_spacing: true
        }
    }));

    // 预览2：简化表格测试
    previews.push(previewManager.createPreview({
        action_type: 'insert_content',
        message: '📋 预览2 - 简化表格'
    }, {
        action_type: 'insert_content',
        parameters: {
            content: `| **特性** | **值** |
|------|-----|
| **性能** | 高性能 |
| **扩展性** | 水平扩展 |
| **可靠性** | 高可靠 |`,
            format_type: 'table',
            add_spacing: true
        }
    }));

    setTimeout(() => {
        checkAndShowBatchActions();
        addAssistantMessage(`
            <div style="padding: 16px; background: #fef2f2; border: 1px solid #dc2626; border-radius: 10px; margin: 16px 0;">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                    <span style="font-size: 16px;">🔧</span>
                    <strong style="color: #b91c1c;">表格插入修复测试就绪</strong>
                </div>
                <p style="margin: 0; font-size: 13px; color: #7f1d1d; line-height: 1.4;">
                    📊 <strong>测试内容</strong>：包含 **粗体** 格式的Markdown表格<br>
                    🔧 <strong>修复重点</strong>：列对齐 + 粗体格式处理<br>
                    ✨ <strong>预期效果</strong>：表格列正确对齐，**文本**显示为粗体<br><br>
                    <strong>点击"全部接受"查看表格修复效果！</strong>
                </p>
            </div>
        `);
    }, 800);
}

window.testTableInsertFix = testTableInsertFix;

// ==================== 上下文选择器相关函数 ====================

// 初始化上下文相关元素
function initializeContextElements() {
    contextBar = document.getElementById('context-bar');
    contextItems = document.getElementById('context-items');

    // 创建上下文选择器
    if (!contextSelector) {
        contextSelector = document.createElement('div');
        contextSelector.id = 'context-selector';
        contextSelector.className = 'context-selector';
        contextSelector.style.display = 'none';

        contextSelector.innerHTML = `
            <div class="context-selector-header">
                <span class="context-selector-title">📄 选择文档上下文</span>
                <button class="context-selector-close" onclick="hideContextSelector()">✕</button>
            </div>
            <div id="context-selector-content" class="context-selector-content">
                <div class="context-selector-loading">正在加载文档...</div>
            </div>
        `;

        document.body.appendChild(contextSelector);
        contextSelectorContent = document.getElementById('context-selector-content');
    }
}

// 显示上下文选择器
function showContextSelector() {
    if (!contextSelector) {
        initializeContextElements();
    }

    // 获取文档列表
    fetchDocuments();

    // 显示选择器
    contextSelector.style.display = 'block';

    // 定位选择器
    positionContextSelector();

    console.log('上下文选择器已显示');
}

// 隐藏上下文选择器
function hideContextSelector() {
    if (contextSelector) {
        contextSelector.style.display = 'none';
    }
}

// 定位上下文选择器
function positionContextSelector() {
    if (!contextSelector || !messageInput) return;

    const inputRect = messageInput.getBoundingClientRect();
    const selectorHeight = 300;
    const margin = 10;

    // 计算位置
    let top = inputRect.top - selectorHeight - margin;
    let left = inputRect.left;

    // 如果上方空间不够，显示在下方
    if (top < margin) {
        top = inputRect.bottom + margin;
    }

    // 确保不超出右边界
    const maxLeft = window.innerWidth - 400 - margin;
    if (left > maxLeft) {
        left = maxLeft;
    }

    contextSelector.style.position = 'fixed';
    contextSelector.style.top = `${top}px`;
    contextSelector.style.left = `${left}px`;
    contextSelector.style.width = '400px';
    contextSelector.style.maxHeight = `${selectorHeight}px`;
    contextSelector.style.zIndex = '10000';
}

// 获取文档列表
function fetchDocuments() {
    if (!contextSelectorContent) return;

    contextSelectorContent.innerHTML = '<div class="context-selector-loading">正在加载文档...</div>';

    // 发送获取文档列表的请求
    sendMessageToCSharp({
        type: 'getDocuments'
    });
}

// 显示文档列表
function showDocumentsInSelector(documents) {
    if (!contextSelectorContent) return;

    availableDocuments = documents;

    if (!documents || documents.length === 0) {
        contextSelectorContent.innerHTML = '<div class="context-selector-empty">暂无上传的文档</div>';
        return;
    }

    let html = '<div class="context-selector-section">';
    html += '<h4 class="context-section-title">📄 选择文档</h4>';

    documents.forEach((doc, index) => {
        const fileIcon = getFileIcon(doc.fileType);
        html += `
            <div class="context-document-item" data-type="document" data-id="${doc.id}" onclick="selectDocument(${doc.id}, '${escapeHtml(doc.fileName)}')">
                <div class="context-item-icon">${fileIcon}</div>
                <div class="context-item-info">
                    <div class="context-item-name">${escapeHtml(doc.fileName)}</div>
                    <div class="context-item-meta">${doc.fileType.toUpperCase()} • ${formatFileSize(doc.fileSize)} • ${doc.headingCount}个标题</div>
                </div>
                <div class="context-item-actions">
                    <button class="context-expand-btn" onclick="event.stopPropagation(); toggleDocumentHeadings(${doc.id}, this)" title="查看标题">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <polyline points="6,9 12,15 18,9" stroke="currentColor" stroke-width="2"/>
                        </svg>
                    </button>
                </div>
            </div>
            <div id="headings-${doc.id}" class="document-headings" style="display: none;"></div>
        `;
    });

    html += '</div>';
    contextSelectorContent.innerHTML = html;
}

// 切换文档标题显示
function toggleDocumentHeadings(documentId, button) {
    const headingsContainer = document.getElementById(`headings-${documentId}`);
    if (!headingsContainer) return;

    if (headingsContainer.style.display === 'none') {
        // 显示标题
        fetchDocumentHeadingsForContext(documentId);
        headingsContainer.style.display = 'block';
        button.classList.add('expanded');
    } else {
        // 隐藏标题
        headingsContainer.style.display = 'none';
        button.classList.remove('expanded');
    }
}

// 获取文档标题（用于上下文选择器）
function fetchDocumentHeadingsForContext(documentId) {
    const headingsContainer = document.getElementById(`headings-${documentId}`);
    if (!headingsContainer) return;

    headingsContainer.innerHTML = '<div class="context-selector-loading">正在加载标题...</div>';

    // 发送获取文档标题的请求
    sendMessageToCSharp({
        type: 'getDocumentContent',
        documentId: documentId
    });
}

// 显示文档标题列表
function showDocumentHeadingsInSelector(documentId, documentName, headings) {
    const headingsContainer = document.getElementById(`headings-${documentId}`);
    if (!headingsContainer) return;

    if (!headings || headings.length === 0) {
        headingsContainer.innerHTML = '<div class="context-selector-empty">该文档没有标题</div>';
        return;
    }

    let html = '';
    headings.forEach((heading, index) => {
        const indent = (heading.level - 1) * 20;
        const levelClass = `level-${heading.level}`;
        html += `
            <div class="context-heading-item ${levelClass}" 
                 data-type="heading" 
                 data-document-id="${documentId}" 
                 data-heading-id="${heading.id}"
                 style="padding-left: ${indent + 20}px;"
                 onclick="selectHeadingAsContext(${documentId}, ${heading.id}, '${escapeHtml(documentName)}', '${escapeHtml(heading.text)}')">
                <div class="heading-level">H${heading.level}</div>
                <div class="heading-text">${escapeHtml(heading.text)}</div>
                <div class="heading-meta">${heading.contentLength}字</div>
            </div>
        `;
    });

    headingsContainer.innerHTML = html;
}

// 选择文档作为上下文
function selectDocument(documentId, fileName) {
    const context = {
        type: 'document',
        id: documentId,
        name: fileName,
        displayText: `📄 ${fileName}`
    };

    addContext(context);
    hideContextSelector();

    // 选中文档后，自动清理触发用的 '#' 以及其后可能紧跟的空格
    try {
        if (messageInput) {
            const text = messageInput.value || '';
            const cursorPos = messageInput.selectionStart || 0;
            // 从光标往前寻找离光标最近的 '#'（且不跨行）
            let i = cursorPos - 1;
            while (i >= 0 && text[i] !== '\n') {
                if (text[i] === '#') break;
                i--;
            }
            if (i >= 0 && text[i] === '#') {
                const before = text.slice(0, i);
                // 若 # 后有一个空格也一起删除
                const afterStart = (i + 1 < text.length && text[i + 1] === ' ') ? i + 2 : i + 1;
                const after = text.slice(afterStart);
                messageInput.value = before + after;
                const newPos = before.length;
                messageInput.setSelectionRange(newPos, newPos);
                updateCharacterCount();
                autoResizeInput();
                updateInputHighlights();
            }
        }
    } catch (e) { console.warn('移除#触发符失败:', e); }
}

// 选择标题作为上下文
function selectHeadingAsContext(documentId, headingId, documentName, headingText) {
    const context = {
        type: 'heading',
        documentId: documentId,
        headingId: headingId,
        name: headingText,
        documentName: documentName,
        displayText: `📋 ${headingText} (${documentName})`
    };

    addContext(context);
    hideContextSelector();
}

// 添加上下文
function addContext(context) {
    // 检查是否已存在
    const existingIndex = selectedContexts.findIndex(ctx => {
        if (ctx.type !== context.type) return false;
        if (context.type === 'document') {
            return ctx.id === context.id;
        } else if (context.type === 'heading') {
            return ctx.documentId === context.documentId && ctx.headingId === context.headingId;
        }
        return false;
    });

    if (existingIndex !== -1) {
        console.log('上下文已存在，跳过添加');
        return;
    }

    // 添加到列表
    selectedContexts.push(context);

    // 更新显示
    updateContextDisplay();

    console.log('已添加上下文:', context);
}

// 移除上下文
function removeContext(index) {
    if (index >= 0 && index < selectedContexts.length) {
        selectedContexts.splice(index, 1);
        updateContextDisplay();
    }
}

// 清空所有上下文
function clearAllContexts() {
    selectedContexts = [];
    updateContextDisplay();
}

// 更新上下文显示
function updateContextDisplay() {
    if (!contextBar || !contextItems) {
        initializeContextElements();
        return;
    }

    if (selectedContexts.length === 0) {
        contextBar.style.display = 'none';
        return;
    }

    contextBar.style.display = 'block';

    let html = '';
    selectedContexts.forEach((context, index) => {
        const itemClass = context.type === 'document' ? 'context-item document' : 'context-item';
        // 移除图标，只显示文件名或标题名
        const displayName = context.type === 'document' ? context.name : context.name;

        html += `
            <div class="${itemClass}">
                <span class="context-item-text" title="${escapeHtml(context.displayText)}">${escapeHtml(displayName)}</span>
                <button class="context-item-remove" onclick="removeContext(${index})" title="移除">
                    <svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <line x1="18" y1="6" x2="6" y2="18" stroke="currentColor" stroke-width="2"/>
                        <line x1="6" y1="6" x2="18" y2="18" stroke="currentColor" stroke-width="2"/>
                    </svg>
                </button>
            </div>
        `;
    });

    contextItems.innerHTML = html;
}

// 键盘导航上下文选择器
function navigateContextSelector(direction) {
    const items = contextSelectorContent.querySelectorAll('.context-document-item, .context-heading-item');
    if (items.length === 0) return;

    let currentIndex = Array.from(items).findIndex(item => item.classList.contains('selected'));

    if (direction === 'down') {
        currentIndex = currentIndex < items.length - 1 ? currentIndex + 1 : 0;
    } else {
        currentIndex = currentIndex > 0 ? currentIndex - 1 : items.length - 1;
    }

    // 移除之前的选中状态
    items.forEach(item => item.classList.remove('selected'));

    // 设置新的选中状态
    if (items[currentIndex]) {
        items[currentIndex].classList.add('selected');
        items[currentIndex].scrollIntoView({ block: 'nearest' });
    }
}

// 选择当前高亮的上下文项目
function selectCurrentContextItem() {
    const selectedItem = contextSelectorContent.querySelector('.context-document-item.selected, .context-heading-item.selected');
    if (selectedItem) {
        selectedItem.click();
    }
}

// 获取文件图标
function getFileIcon(fileType) {
    const icons = {
        'docx': '📄',
        'doc': '📄',
        'md': '📝',
        'pdf': '📕'
    };
    return icons[fileType.toLowerCase()] || '📄';
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

// 暴露上下文相关函数到全局
window.clearAllContexts = clearAllContexts;
window.removeContext = removeContext;
window.hideContextSelector = hideContextSelector;
window.selectDocument = selectDocument;
window.selectHeading = selectHeading;
window.toggleDocumentHeadings = toggleDocumentHeadings;

// 悬浮批量操作按钮函数
window.hideFloatingBatchActions = hideFloatingBatchActions;
window.handleFloatingAcceptAll = handleFloatingAcceptAll;
window.handleFloatingRejectAll = handleFloatingRejectAll;
window.handleModelList = handleModelList;
window.getSelectedModelInfo = getSelectedModelInfo;
window.unlockAgentMode = unlockAgentMode;
window.closePasswordModal = closePasswordModal;
window.verifyPassword = verifyPassword;
window.onChatModeChange = onChatModeChange;

// 测试Chat和Agent模式下的标题格式一致性
function testMarkdownHeaderConsistency() {
    addAssistantMessage(`
        <div style="margin: 16px 0; padding: 20px; background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; border-radius: 16px;">
            <h2 style="margin: 0 0 16px 0;">🔧 标题格式一致性修复测试</h2>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px;">
                <h3 style="margin: 0 0 12px 0;">📋 问题描述：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li><strong>问题</strong>：Chat模式下#符号和标题之间缺少空格</li>
                    <li><strong>原因</strong>：Chat模式提示词缺少Markdown格式要求</li>
                    <li><strong>影响</strong>：导致两种模式下的输出格式不一致</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 16px; border-radius: 12px; margin-top: 12px;">
                <h3 style="margin: 0 0 12px 0;">✅ 修复方案：</h3>
                <ul style="margin: 0; padding-left: 20px; font-size: 14px; line-height: 1.6;">
                    <li>在Chat模式提示词中添加："使用标准的markdown回复，#符号和标题文本之间必须有空格"</li>
                    <li>在Agent模式提示词中也添加相同的格式要求</li>
                    <li>更新PromptManager.cs中的备用提示词</li>
                    <li>确保数据库中的默认提示词包含格式要求</li>
                </ul>
            </div>
            <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 8px; margin-top: 12px;">
                <h3 style="margin: 0 0 8px 0;">🧪 测试示例：</h3>
                <div style="font-size: 13px; font-family: monospace; line-height: 1.4;">
                    <strong>修复前 (Chat模式)：</strong><br>
                    <code style="background: rgba(255,255,255,0.2); padding: 2px 4px; border-radius: 3px;">#标题1</code> ❌<br>
                    <code style="background: rgba(255,255,255,0.2); padding: 2px 4px; border-radius: 3px;">##标题2</code> ❌<br><br>
                    <strong>修复后 (两种模式一致)：</strong><br>
                    <code style="background: rgba(255,255,255,0.2); padding: 2px 4px; border-radius: 3px;"># 标题1</code> ✅<br>
                    <code style="background: rgba(255,255,255,0.2); padding: 2px 4px; border-radius: 3px;">## 标题2</code> ✅
                </div>
            </div>
        </div>
    `);
}

window.testMarkdownHeaderConsistency = testMarkdownHeaderConsistency;

// 测试工具进度显示
function testToolProgress() {
    console.log('🧪 开始测试工具进度显示');

    // 模拟开始生成
    isGenerating = true;

    // 步骤1: 创建初始消息
    setTimeout(() => {
        // 创建初始内容（纯文本，不需要markdown渲染）
        const initialText = '我先获取文档的基本统计信息，了解标题结构概况：';
        const msg1 = addAssistantMessage(initialText);
        toolProgressHostMessage = msg1;
        currentMessageId = msg1.id;

        const markdownContent = msg1.querySelector('.message-content .markdown-content');

        // 步骤2: 第一个工具调用
        setTimeout(() => {
            handleToolProgress({ message: '执行工具: get_document_statistics', timestamp: '20:26:33.069' });

            // 步骤3: 第一个工具完成
            setTimeout(() => {
                handleToolProgress({ message: '工具 get_document_statistics 执行完成，返回数据长度: 250 字符', timestamp: '20:26:33.290' });

                // 步骤4: AI继续说话（在卡片后面追加文本）
                setTimeout(() => {
                    if (markdownContent) {
                        const p = document.createElement('p');
                        p.textContent = '根据文档统计和标题结构分析，当前文档的标题结构如下：';
                        markdownContent.appendChild(p);
                    }

                    // 步骤5: 第二个工具调用
                    setTimeout(() => {
                        handleToolProgress({ message: '执行工具: get_document_headings', timestamp: '20:26:36.697' });

                        // 步骤6: 第二个工具完成
                        setTimeout(() => {
                            handleToolProgress({ message: '工具 get_document_headings 执行完成，返回数据长度: 1828 字符', timestamp: '20:26:37.763' });

                            // 步骤7: 结束并显示最终结果
                            setTimeout(() => {
                                if (markdownContent) {
                                    // 添加最终的markdown内容
                                    const finalHtml = `
                                        <h2>文档基本信息</h2>
                                        <ul>
                                            <li>页数: 5</li>
                                            <li>字数: 2500</li>
                                            <li>标题总数: 12</li>
                                        </ul>
                                    `;
                                    const div = document.createElement('div');
                                    div.innerHTML = finalHtml;
                                    markdownContent.appendChild(div);
                                }

                                isGenerating = false;
                                console.log('✅ 工具进度测试完成');
                            }, 500);
                        }, 600);
                    }, 400);
                }, 500);
            }, 300);
        }, 200);
    }, 100);
}

window.testToolProgress = testToolProgress;

// 当前正在执行的工具卡片
let currentToolCard = null;
// 工具卡片缓存（在整段内容重新渲染时用于复原）
let toolCardsCache = [];
// 记录各工具最近一次“完成”的时间，用于去重显示
let lastCompletedToolAtByName = {};
// 记录工具名到当前卡片ID的映射（便于按名称更新状态）
let toolNameToCardId = {};

// 处理工具调用进度 - Cursor风格的独立卡片
function handleToolProgress(data) {
    console.log('收到工具进度:', data.message);

    // 判断消息类型
    const message = data.message;
    const isToolCall = message.includes('轮工具调用') || message.includes('执行工具:') || message.includes('执行完成');
    const isDebugInfo = message.includes('查找标题') || message.includes('提取关键词') || message.includes('选择最佳匹配') ||
        message.includes('找到候选标题') || message.includes('标题搜索完成') || message.includes('提取内容完成') ||
        message.includes('连接到已打开的Word实例') || message.includes('开始高效查找标题') ||
        message.includes('开始提取标题下的内容') || message.includes('找到标题:');

    // 检查是否应该显示此类消息
    const shouldShow = (isToolCall && agentConfig.showToolCalls) || (isDebugInfo && agentConfig.showDebugInfo);

    if (!shouldShow) {
        console.log('根据配置跳过显示此进度消息:', message);
        return;
    }

    // 确保在生成状态下才显示进度
    if (!isGenerating) {
        return;
    }

    // 根据消息类型决定如何显示
    if (message.includes('执行工具:')) {
        // 去重：若与当前正在运行的卡片工具同名，则合并为同一次调用
        const toolNameMatch = message.match(/执行工具:\s*(\w+)/);
        const toolName = toolNameMatch ? toolNameMatch[1] : '';
        if (currentToolCard) {
            const running = toolCardsCache.find(c => c.id === currentToolCard.id);
            if (running && running.name === toolName) {
                console.log(`检测到同名工具(${toolName})的连续启动，已合并为一次调用`);
                return;
            }
        }
        // 若刚刚完成同名工具，且时间间隔很短，则认为是重复展示，直接跳过
        if (toolName && lastCompletedToolAtByName[toolName] && (Date.now() - lastCompletedToolAtByName[toolName] < 1500)) {
            console.log(`同名工具(${toolName})在短时间内重复启动，已抑制重复显示`);
            return;
        }
        // 创建新的工具调用卡片
        createToolCallCard(message, data.timestamp);
    } else if (message.includes('执行完成')) {
        // 更新当前工具卡片状态为完成
        updateToolCallCard('completed', message, data.timestamp);
    } else {
        // 其他进度信息（调试信息等）附加到当前卡片
        appendToToolCallCard(message, data.timestamp);
    }

    // 滚动到底部
    scrollToBottom();
}

// 创建新的工具调用卡片（Cursor风格）
function createToolCallCard(message, timestamp) {
    // 提取工具名称
    const toolNameMatch = message.match(/执行工具:\s*(\w+)/);
    const toolName = toolNameMatch ? toolNameMatch[1] : '未知工具';

    // 生成唯一卡片ID
    const cardId = `tool_card_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;

    // 构建运行中卡片HTML（占位后渲染时替换）
    const runningHtml = `
        <div class="tool-call-card tool-call-running" data-card-id="${cardId}">
            <div class="tool-call-header">
                <span class="tool-call-icon">▶️</span>
                <span class="tool-call-name">${escapeHtml(toolName)}</span>
                <span class="tool-call-status">执行中</span>
                <span class="tool-call-time">${timestamp}</span>
            </div>
            <div class="tool-call-details" style="display: none;"></div>
        </div>
    `;

    // 写入缓存
    toolCardsCache.push({
        id: cardId,
        name: toolName,
        html: runningHtml,
        details: []
    });

    // 在文本流中插入占位符（使用HTML注释，避免Markdown转换）
    const placeholder = `\n\n<!--TOOL_CARD_${cardId}-->\n\n`;
    currentContent += placeholder;
    appendOutlineContent('');

    // 记录当前正在处理的卡片ID
    currentToolCard = { id: cardId };

    // 建立名称到卡片的映射（后续可按名称更新完成状态）
    if (toolName) {
        toolNameToCardId[toolName] = cardId;
    }
}

// 更新工具调用卡片状态
function updateToolCallCard(status, message, timestamp) {
    if (!currentToolCard || !currentToolCard.id) return;

    const cached = toolCardsCache.find(c => c.id === currentToolCard.id);
    if (!cached) return;

    if (status === 'completed') {
        // 不再显示"返回数据: xxx 字符"，保持简洁
        cached.html = `
            <div class="tool-call-card tool-call-completed no-animate" data-card-id="${cached.id}">
                <div class="tool-call-header">
                    <span class="tool-call-icon">✅</span>
                    <span class="tool-call-name">${escapeHtml(cached.name)}</span>
                    <span class="tool-call-status">完成</span>
                    <span class="tool-call-time">${timestamp}</span>
            </div>
                <div class="tool-call-details" style="display: none;"></div>
            </div>
        `;

        // 重新渲染，让改变生效
        appendOutlineContent('');

        // 记录完成时间，用于后续的重复调用去重
        lastCompletedToolAtByName[cached.name] = Date.now();
    }

    // 完成后清空当前卡片引用
    currentToolCard = null;
}

// 按工具名将最近一次卡片标记为“完成”
function markToolCardCompletedByName(toolName, timestamp) {
    const cardId = toolNameToCardId[toolName];
    if (!cardId) return false;
    const cached = toolCardsCache.find(c => c.id === cardId);
    if (!cached) return false;

    // 更新缓存的HTML
    cached.html = `
        <div class="tool-call-card tool-call-completed no-animate" data-card-id="${cached.id}">
            <div class="tool-call-header">
                <span class="tool-call-icon">✅</span>
                <span class="tool-call-name">${escapeHtml(cached.name)}</span>
                <span class="tool-call-status">完成</span>
                <span class="tool-call-time">${timestamp}</span>
            </div>
            <div class="tool-call-details" style="display: none;"></div>
        </div>
    `;

    // 如果当前处于预览挂起状态，appendOutlineContent 会被跳过
    // 这时直接更新现有的DOM卡片，确保UI立即反映为“完成”
    try {
        const el = document.querySelector(`.tool-call-card[data-card-id="${cached.id}"]`);
        if (el) {
            el.outerHTML = cached.html;
        }
    } catch (e) {
        console.warn('直接更新工具卡片DOM失败，将在下一次渲染时更新:', e);
    }

    // 仅在未挂起时触发重渲染
    if (!isPreviewPending) {
        appendOutlineContent('');
    }
    lastCompletedToolAtByName[cached.name] = Date.now();
    // 若当前卡片就是它，清空引用
    if (currentToolCard && currentToolCard.id === cardId) {
        currentToolCard = null;
    }
    return true;
}

// 向工具调用卡片添加详细信息
function appendToToolCallCard(message, timestamp) {
    if (!currentToolCard || !currentToolCard.id) return;
    const cached = toolCardsCache.find(c => c.id === currentToolCard.id);
    if (!cached) return;

    cached.details = cached.details || [];
    cached.details.push({ message, timestamp });

    // 重新构建HTML，展示详细信息
    const detailsHtml = cached.details.map(d => `
        <div class="tool-call-detail-item">
            <span class="detail-icon">ℹ️</span>
            <span class="detail-message">${escapeHtml(d.message)}</span>
            <span class="detail-time">${d.timestamp}</span>
        </div>
    `).join('');

    cached.html = `
        <div class="tool-call-card tool-call-running no-animate" data-card-id="${cached.id}">
            <div class="tool-call-header">
                <span class="tool-call-icon">▶️</span>
                <span class="tool-call-name">${escapeHtml(cached.name)}</span>
                <span class="tool-call-status">执行中</span>
                <span class="tool-call-time">${timestamp}</span>
            </div>
            <div class="tool-call-details" style="display: ${cached.details.length ? 'block' : 'none'};">
                ${detailsHtml}
            </div>
        </div>
    `;

    appendOutlineContent('');
}

// 兼容性：保留旧的函数名，但标记为废弃
function getOrCreateToolProgressContainer() {
    console.warn('getOrCreateToolProgressContainer() 已废弃，工具进度现在使用独立卡片模式');
    // 返回一个占位元素，避免旧代码报错
    let placeholder = document.createElement('div');
    placeholder.style.display = 'none';
    return placeholder;
}

// 兼容性：保留旧的函数名，但标记为废弃
function addToolProgressItem(container, message, timestamp) {
    console.warn('addToolProgressItem() 已废弃，工具进度现在使用独立卡片模式');
    // 不做任何操作
}

// ==================== 图片上传和管理相关函数 ====================

// 初始化图片相关元素
function initializeImageElements() {
    imageBar = document.getElementById('image-bar');
    imageItems = document.getElementById('image-items');

    // 添加粘贴事件监听
    const messageInput = document.getElementById('message-input');
    if (messageInput) {
        messageInput.addEventListener('paste', handlePasteImage);
    }

    // 添加拖放事件监听
    const inputWrapper = document.querySelector('.input-wrapper');
    if (inputWrapper) {
        inputWrapper.addEventListener('dragover', handleDragOver);
        inputWrapper.addEventListener('dragleave', handleDragLeave);
        inputWrapper.addEventListener('drop', handleDrop);
    }

    console.log('图片元素已初始化');
}

// 处理图片上传
function handleImageUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    console.log(`选择了 ${files.length} 个文件`);

    for (let i = 0; i < files.length; i++) {
        const file = files[i];

        // 检查文件类型
        if (!file.type.startsWith('image/')) {
            console.warn(`文件 ${file.name} 不是图片，已跳过`);
            continue;
        }

        // 检查文件大小（限制10MB）
        const maxSize = 10 * 1024 * 1024;
        if (file.size > maxSize) {
            showCustomAlert(`图片 ${file.name} 过大（超过10MB），请选择较小的图片`);
            continue;
        }

        // 读取并添加图片
        readImageFile(file);
    }

    // 清空文件输入框，允许重复选择同一文件
    event.target.value = '';
}

// 处理粘贴图片
function handlePasteImage(event) {
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;

    for (let i = 0; i < items.length; i++) {
        const item = items[i];

        if (item.type.startsWith('image/')) {
            event.preventDefault(); // 阻止默认粘贴行为

            const file = item.getAsFile();
            if (file) {
                console.log('粘贴了图片:', file.name || '未命名');
                readImageFile(file);
            }
        }
    }
}

// 处理拖放 - 拖动悬停
function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();

    const inputWrapper = event.currentTarget;
    inputWrapper.classList.add('drag-over');
}

// 处理拖放 - 离开
function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();

    const inputWrapper = event.currentTarget;
    inputWrapper.classList.remove('drag-over');
}

// 处理拖放 - 放下
function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();

    const inputWrapper = event.currentTarget;
    inputWrapper.classList.remove('drag-over');

    const files = event.dataTransfer.files;
    if (!files || files.length === 0) return;

    console.log(`拖放了 ${files.length} 个文件`);

    for (let i = 0; i < files.length; i++) {
        const file = files[i];

        if (!file.type.startsWith('image/')) {
            console.warn(`文件 ${file.name} 不是图片，已跳过`);
            continue;
        }

        const maxSize = 10 * 1024 * 1024;
        if (file.size > maxSize) {
            showCustomAlert(`图片 ${file.name} 过大（超过10MB），请选择较小的图片`);
            continue;
        }

        readImageFile(file);
    }
}

// 读取图片文件并转换为base64
function readImageFile(file) {
    const reader = new FileReader();

    reader.onload = function (e) {
        const base64Data = e.target.result;

        // 创建图片对象
        const imageObj = {
            id: ++imageIdCounter,
            name: file.name || `图片${imageIdCounter}`,
            base64: base64Data,
            size: file.size,
            type: file.type
        };

        // 添加到已选图片列表
        selectedImages.push(imageObj);

        console.log(`图片已添加: ${imageObj.name}, 大小: ${(imageObj.size / 1024).toFixed(2)} KB`);

        // 更新显示
        updateImageDisplay();
    };

    reader.onerror = function (error) {
        console.error('读取图片失败:', error);
        showCustomAlert('读取图片失败，请重试');
    };

    reader.readAsDataURL(file);
}

// 移除指定图片
function removeImage(imageId) {
    const index = selectedImages.findIndex(img => img.id === imageId);
    if (index >= 0) {
        const removedImage = selectedImages.splice(index, 1)[0];
        console.log(`移除了图片: ${removedImage.name}`);
        updateImageDisplay();
    }
}

// 清空所有图片
function clearAllImages() {
    selectedImages = [];
    console.log('已清空所有图片');
    updateImageDisplay();
}

// 更新图片显示
function updateImageDisplay() {
    if (!imageBar || !imageItems) {
        initializeImageElements();
        return;
    }

    if (selectedImages.length === 0) {
        imageBar.style.display = 'none';
        return;
    }

    imageBar.style.display = 'block';

    let html = '';
    selectedImages.forEach((image) => {
        html += `
            <div class="image-item" data-image-id="${image.id}" title="${image.name}" onclick="previewImage(${image.id})">
                <img src="${image.base64}" alt="${image.name}">
                <button class="image-item-remove" onclick="event.stopPropagation(); removeImage(${image.id})" title="删除图片">
                    ✕
                </button>
            </div>
        `;
    });

    imageItems.innerHTML = html;

    console.log(`图片显示已更新，当前有 ${selectedImages.length} 张图片`);
}

// 预览图片
function previewImage(imageId) {
    const image = selectedImages.find(img => img.id === imageId);
    if (!image) return;

    const modal = document.getElementById('image-preview-modal');
    const modalImg = document.getElementById('image-preview-content');
    const caption = document.getElementById('image-preview-caption');

    if (modal && modalImg && caption) {
        modal.style.display = 'block';
        modalImg.src = image.base64;
        caption.textContent = image.name;
    }
}

// 关闭图片预览
function closeImagePreview() {
    const modal = document.getElementById('image-preview-modal');
    if (modal) {
        modal.style.display = 'none';
    }
}

// 键盘事件：ESC关闭预览
document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
        const modal = document.getElementById('image-preview-modal');
        if (modal && modal.style.display === 'block') {
            closeImagePreview();
        }
    }
});

// 将图片上传按钮和函数暴露到全局作用域
window.handleImageUpload = handleImageUpload;
window.removeImage = removeImage;
window.clearAllImages = clearAllImages;
window.previewImage = previewImage;
window.closeImagePreview = closeImagePreview;