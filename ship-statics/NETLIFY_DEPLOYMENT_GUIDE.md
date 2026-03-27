# ══════════════════════════════════════════════════════════
# Netlify 部署指南 - 船舶静力学课程设计系统
# ══════════════════════════════════════════════════════════

## ⚠️ 重要警告

**你的项目是 Flask 应用，Netlify 无法直接托管！**

Flask 是 Python 动态后端，Netlify 只支持静态网站。

---

## 🎯 推荐部署方案

### 方案 1：Railway（推荐）

Railway 支持 Flask 应用，部署简单：

#### 步骤：

1. **创建 `requirements.txt`**（已有）
   ```
   Flask==3.0.0
   Flask-CORS==4.0.0
   gunicorn==21.2.0
   numpy==1.24.3
   scipy==1.11.4
   pandas==2.1.3
   openpyxl==3.1.2
   matplotlib==3.8.2
   python-docx==1.1.0
   Pillow==10.1.0
   ```

2. **创建 `runtime.txt`**
   ```
   python-3.10.13
   ```

3. **创建 `Procfile`**
   ```
   web: gunicorn app:app
   ```

4. **部署到 Railway**
   - 访问 https://railway.app
   - 连接 GitHub 仓库
   - 选择项目
   - 自动部署

**优点**：
- ✅ 完整支持 Flask
- ✅ 自动 HTTPS
- ✅ 免费额度充足
- ✅ 支持文件上传

---

### 方案 2：Render

Render 也支持 Flask：

#### 步骤：

1. **创建 `requirements.txt`**（同上）

2. **在 Render 配置**：
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`

3. **部署**
   - 访问 https://render.com
   - 创建 Web Service
   - 连接 GitHub 仓库

**优点**：
- ✅ 免费套餐
- ✅ 自动 HTTPS
- ✅ 支持 Flask

---

### 方案 3：PythonAnywhere

专门托管 Python 应用：

#### 步骤：

1. 注册 https://www.pythonanywhere.com
2. 创建 Web App
3. 选择 Flask
4. 上传代码
5. 配置 WSGI

**优点**：
- ✅ 专门为 Python 设计
- ✅ 简单易用

---

## 📝 如果坚持使用 Netlify（仅前端）

### 方案：将后端和前端分离

#### 1. 后端部署到 Railway/Render
```bash
# 部署 Flask 后端到 Railway
railway login
railway init
railway up
```

#### 2. 前端改造为纯静态
修改 `index.html` 中的 API 地址：

```javascript
// 将所有 fetch 请求改为指向后端服务器
const API_BASE = 'https://your-backend.railway.app';

fetch(`${API_BASE}/upload_offsets`, {
    method: 'POST',
    body: formData
})
```

#### 3. 部署前端到 Netlify
```bash
# 1. 确保 templates 目录包含所有静态文件
# 2. 在 Netlify 设置：
#    - Publish directory: templates
#    - 无需构建命令
```

---

## 🔧 Netlify 配置文件说明

### 1. `netlify.toml`

```toml
[build]
  publish = "templates"
  command = "echo 'No build needed'"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200
```

**作用**：
- `publish`: 指定发布目录为 templates
- `redirects`: SPA 路由支持，所有路由重定向到 index.html

### 2. `_redirects`

```
/*    /index.html   200
```

**作用**：
- 简化的重定向规则
- 解决刷新页面 404 问题

---

## 🐛 常见问题排查

### Q1: 404 Page not found

**原因**：
- 发布目录错误
- 缺少重定向配置
- Flask 路由无法工作

**解决**：
1. 检查 Netlify 设置 → Build & Deploy → Publish directory
   - 应设置为 `templates`（不是 `dist`）
   
2. 添加 `netlify.toml` 或 `_redirects`

3. **重要**：Flask 无法在 Netlify 运行！

### Q2: 刷新页面 404

**原因**：SPA 路由未配置

**解决**：
- 添加 `_redirects` 文件到发布目录
- 或在 `netlify.toml` 中配置重定向

### Q3: 静态资源加载失败

**原因**：
- 路径错误
- 大小写不匹配

**解决**：
1. 使用相对路径：
   ```html
   <link rel="stylesheet" href="./style.css">
   <script src="./script.js"></script>
   ```

2. 检查文件名大小写：
   - Netlify 区分大小写
   - 确保 URL 和文件名一致

---

## 📊 部署检查清单

### Flask 应用部署到 Railway/Render

- [ ] requirements.txt 完整
- [ ] Procfile 存在
- [ ] runtime.txt 指定 Python 版本
- [ ] 代码已推送到 GitHub
- [ ] Railway/Render 已连接仓库
- [ ] 环境变量已配置

### 静态前端部署到 Netlify

- [ ] templates 目录包含所有文件
- [ ] netlify.toml 已创建
- [ ] _redirects 已创建
- [ ] API 地址已改为后端服务器
- [ ] 所有路径使用相对路径

---

## 🚀 推荐方案（最佳实践）

**全栈部署架构**：

```
前端
    ├── 静态 HTML/CSS/JS
    ├── templates/index.html
    ├── templates/_redirects
    └── templates/netlify.toml

后端
    ├── app.py
    ├── requirements.txt
    ├── Procfile
    └── core/ (计算模块)
```

**优势**：
- ✅ 前端访问速度快（Netlify CDN）
- ✅ 后端功能完整
- ✅ 支持文件上传和计算
- ✅ 免费（两个平台免费额度）

---

## 📞 部署支持

如有问题，请检查：

1. **Railway 部署文档**: https://docs.railway.app
2. **Render 部署文档**: https://render.com/docs
3. **Netlify 部署文档**: https://docs.netlify.com

---

**建议**：优先使用 Railway 或 Render 部署完整 Flask 应用，避免前后端分离的复杂性。
