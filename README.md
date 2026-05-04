# office_auto

办公自动化工具 — Web 版

Excel 收集表批量导入总表 + Word 模板按条目导出，基于 Flask 构建，支持 Docker 一键部署。

## 功能

- **Excel 批量导入**：上传多个收集表（每人一份），自动校验表头并合并到总表，主键重复时更新、新增时追加
- **Word 模板导出**：选择总表中的条目，自动替换 Word 模板中的 `{{字段名}}` 占位符，生成并下载文档
- **总表管理**：在线预览总表内容，支持下载和覆盖上传
- **表头校验**：导入时自动比对收集表与总表的表头一致性，字段缺失或不匹配会给出明确提示
- **搜索定位**：导出页面支持按编号搜索，快速定位条目
- **说明**：因为需求简单，未使用数据库，简单的用excel表代替

## 快速开始

### Docker Compose 部署（推荐）

创建 `docker-compose.yml`：

```yaml
version: "3.8"

services:
  office-auto:
    image: ghcr.io/tszy333/office_auto:latest
    container_name: office-auto
    restart: unless-stopped
    ports:
      - "5000:5000"
    volumes:
      - ./data:/data
    environment:
      - TZ=Asia/Shanghai
```

启动服务：

```bash
docker compose up -d
```

服务启动后访问 `http://localhost:5000`。

数据持久化在 `./data` 目录，包含配置文件和总表文件。

### 本地运行

```bash
pip install -r requirements.txt
python app.py
```

## 使用说明

1. **配置路径**：进入「配置」页面，设置总表 Excel 路径和 Word 模板路径
2. **导入数据**：进入「导入」页面，批量上传收集表 Excel 文件（支持多文件）
3. **预览总表**：进入「预览」页面查看已导入的数据，支持下载和覆盖
4. **导出文档**：进入「导出」页面，选择条目编号，自动生成填充好的 Word 文档

### 收集表格式要求

- 收集表为单列两区域格式：第一列为字段名（表头），第二列为对应值
- 表头必须与总表列名完全一致
- 第一行第一列为主键字段，用于判断新增或更新

### Word 模板格式

在 Word 文档中使用 `{{字段名}}` 作为占位符，导出时会自动替换为对应数据。段落和表格内的占位符均支持替换。

## 项目结构

```
├── app.py                  # Flask 主程序
├── requirements.txt        # Python 依赖
├── Dockerfile              # 容器镜像构建
├── docker-compose.yml      # Docker Compose 部署配置
├── templates/
│   ├── base.html           # 基础布局
│   ├── index.html          # 首页
│   ├── config.html         # 配置页
│   ├── import.html         # Excel 导入页
│   ├── export.html         # Word 导出页
│   └── preview.html        # 总表预览页
└── .github/workflows/
    └── docker.yml          # GitHub Actions 自动构建镜像
```

## 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `CONFIG_FILE` | 配置文件路径 | `/data/config.ini` |
| `DATA_DIR` | 数据目录 | `/data` |
| `PORT` | 服务端口 | `5000` |

## 技术栈

- **后端**：Flask + pandas + python-docx + openpyxl
- **前端**：原生 HTML/CSS
- **部署**：Docker + Gunicorn
