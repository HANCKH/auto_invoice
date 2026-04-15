# 发票解析网页化整体路径

## 1. 目标

将当前 `invoice_to_excel.py` 升级为对外网页服务。

建议先做 **PDF-only MVP**：第一版只允许用户批量上传 PDF，不支持 ZIP。这样可以直接复用现有目录批处理逻辑，减少压缩包安全处理、目录结构兼容、乱码文件名等额外复杂度。ZIP 支持可以作为第二阶段能力追加。

1. 用户上传多个发票 PDF。
2. 后台自动解析、去重、按公司汇总输出。
3. 生成结果目录后打包为 ZIP。
4. 用户在网页端下载结果压缩包。

## 2. 总体架构

- 前端：上传页面（支持多 PDF 文件）
- API 服务：接收上传、创建任务、查询状态、下载结果
- 任务执行器：异步执行解析任务（MVP 可先用 FastAPI 后台线程，避免请求超时）
- Redis：任务队列与状态缓存（第二阶段引入）
- 文件存储：任务临时目录 + 结果 ZIP

建议部署形态：

- MVP：`web`（FastAPI，本地任务线程）
- 生产增强：`web`（FastAPI）+ `worker`（Celery）+ `redis`
- `nginx`（可选，做 HTTPS 和反向代理）

## 3. 核心流程（端到端）

1. 用户在网页上传多个 PDF。
2. 后端 API 校验文件类型与大小，只接受 `.pdf`。
3. API 创建 `task_id` 并落盘原始文件到任务目录。
4. API 投递异步任务，立即返回 `task_id`。
5. 后台任务执行：
   - 调用 `invoice_to_excel.py` 执行解析
   - 生成 `整理结果`
   - 打包 `整理结果` 为 `result.zip`
   - 更新任务状态为 `SUCCESS` 或 `FAILED`
6. 前端轮询任务状态接口。
7. 状态成功后显示下载按钮，调用下载接口获取 `result.zip`。
8. 到期任务自动清理临时文件。

## 4. API 设计建议

- `POST /api/tasks/upload`
  - 入参：`multipart/form-data`（支持多个 `pdf`）
  - 出参：`task_id`

- `GET /api/tasks/{task_id}/status`
  - 出参：`PENDING | RUNNING | SUCCESS | FAILED` + message

- `GET /api/tasks/{task_id}/download`
  - 成功时返回 `application/zip`

## 5. 后台任务执行细节

每个任务使用独立目录，例如：

- `work/tasks/{task_id}/input`
- `work/tasks/{task_id}/output`
- `work/tasks/{task_id}/result.zip`

执行命令示例：

```bash
python invoice_to_excel.py \
  --input-dir work/tasks/<task_id>/input \
  --template 模板文件.xlsx \
  --output-dir work/tasks/<task_id>/output/整理结果 \
  --rename
```

## 6. 安全与稳定性

1. 限制上传大小、文件类型（MVP 仅 PDF）。
2. 校验 PDF 扩展名和 MIME 类型，拒绝空文件与异常大文件。
3. 任务目录隔离，禁止跨任务访问。
4. 下载接口需校验 `task_id` 权限（若有账号体系）。
5. 失败日志保留，便于排障。
6. 定时清理过期任务目录与结果包。
7. 第二阶段支持 ZIP 时，再增加安全解压：拒绝 `../` 路径、绝对路径、嵌套压缩包炸弹和非 PDF 文件。

## 7. 上线步骤建议

1. 先做单机 PDF-only MVP（FastAPI + 本地线程任务）。
2. 跑通批量 PDF 上传、状态查询、结果 ZIP 下载。
3. 再按实际使用量升级到 Celery + Redis 异步队列。
4. 需要 ZIP 上传时，作为第二阶段增加安全解压。
5. Docker 化并使用 `docker compose` 部署。
6. 增加 HTTPS、监控、告警。

## 8. 交付物清单（建议）

- `web/`：前端上传与状态页面
- `api/`：FastAPI 服务
- `worker/`：任务执行器
- `deployment/`：Docker Compose、Nginx 配置
- `deployment_flow/WEB_INVOICE_FLOW.md`：本流程文档
