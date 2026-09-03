# 知识库导出脚本

本项目用于登录 Heren MindHub 知识库管理系统，查找指定知识库及其文件进度，并将知识条目导出为 Excel 文件。

项目包含两个入口脚本：

- `knowledge_query.py`：推荐使用。自动登录、搜索知识库、获取进度 ID，并批量导出 Excel。
- `export_knowledge.py`：手动模式。已知知识库 ID、进度 ID 和登录 Token 时，直接调用接口导出。

## 运行要求

- Windows 10/11
- Python 3.10 或更高版本
- Microsoft Edge
- 可访问知识库管理系统的账号

Python 依赖：

- `playwright`
- `requests`
- `openpyxl`

## 目录说明

```text
export_knowldege/
├─ README.md
├─ .venv/                         # 本地 Python 虚拟环境，不提交 Git
└─ export_knowldege/
   ├─ config.example.json         # 配置示例
   ├─ config.json                 # 本地真实配置，不提交 Git
   ├─ knowledge_query.py          # 自动登录和批量导出
   ├─ export_knowledge.py         # 根据 ID 和 Token 直接导出
   ├─ runtime_logging.py          # 统一运行日志配置
   ├─ logs/                       # 自动脚本默认日志目录
   └─ output/                     # 导出结果，不提交 Git
```

## 第一次运行

在 PowerShell 中进入仓库根目录：

```powershell
cd C:\Users\ASUS\Desktop\export_knowldege
```

创建并启用虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

安装依赖：

```powershell
python -m pip install --upgrade pip
python -m pip install playwright requests openpyxl
```

脚本会调用本机安装的 Microsoft Edge，一般不需要另外下载 Playwright 浏览器。

进入脚本目录，并从模板创建本地配置：

```powershell
cd .\export_knowldege
Copy-Item .\config.example.json .\config.json
```

编辑 `config.json`：

```json
{
  "hospitalName": "医院名称",
  "username": "登录账号",
  "password": "登录密码",
  "outputRoot": "./output",
  "createBgeKnowledgeBases": false,
  "uploadBgeKnowledgeBases": false,
  "knowledgeNames": [
    "医生信息知识库",
    "科室信息知识库",
    "医院信息知识库",
    "医院地址知识库"
  ]
}
```

`config.json` 包含登录信息，已通过 `.gitignore` 排除，不要手动提交到 Git。

## 推荐启动方式：自动批量导出

### 一键启动一个医院（一个账号）

项目根目录新增了 Windows 一键入口：

```text
start_hospital_full.cmd
```

双击该文件，或在 PowerShell 中执行：

```powershell
cd C:\Users\ASUS\Desktop\export_knowldege
.\start_hospital.ps1
```

它会自动完成：检查 `.venv` 和 `config.json`、检查 Python 依赖、启动 `knowledge_query.py`、登录一个账号、处理配置中的全部知识库、导出 Excel、保存日志和快照。

默认是否创建和上传 `_bge`，由 `config.json` 中的两个开关控制：

```json
"createBgeKnowledgeBases": false,
"uploadBgeKnowledgeBases": false
```

确认线上目标库和本地文件后，可以临时开启完整写入流程：

```powershell
.\start_hospital.ps1 -CreateBge -UploadBge -PauseOnExit
```

三个环节也可以分别启动：

```powershell
# 第一步：只读取源知识库并导出 Excel
.\start_hospital.ps1 -ExportOnly

# 第二步：只根据已有快照和 Excel 创建 _bge 知识库
.\start_hospital.ps1 -CreateOnly

# 第三步：只把已有 Excel 上传到已经存在的 _bge 知识库，并轮询导入状态
.\start_hospital.ps1 -UploadOnly
```

`-CreateOnly` 和 `-UploadOnly` 都会跳过源知识库读取及 Excel 导出，复用：

```text
output/<医院名称>/<知识库名称>/_metadata/knowledge_snapshot.json
output/<医院名称>/<知识库名称>/*.xlsx
```

因此，单独执行 `-UploadOnly` 前，需要先成功完成过导出，并确认目标 `_bge` 知识库已经存在；该模式不会创建目标库，也不会重新导出或覆盖本地 Excel。

无界面运行：

```powershell
.\start_hospital.ps1 -Headless
```

该一键脚本只适用于“一个 `config.json` 对应一个医院账号”的任务；切换医院时替换或通过 `-ConfigPath` 指定另一份配置文件。

确保当前目录是脚本目录：

```powershell
cd C:\Users\ASUS\Desktop\export_knowldege\export_knowldege
python .\knowledge_query.py
```

默认情况下脚本会打开可见的 Edge 窗口，依次执行：

1. 使用配置中的账号和密码登录。
2. 搜索 `knowledgeNames` 中配置的知识库。
3. 进入知识库详情页并只读检查向量模型，不会提交编辑。
4. 仅处理向量模型为 `text-embedding-v3` 的知识库。
5. 查询知识库文件的进度 ID。
6. 检查每条文件记录的导入状态，仅保留状态为 `导入成功` 的文件。
7. 将符合条件的文件对应的知识条目导出为 Excel。
8. 如果显式开启 `_bge` 创建，则在导出后按计划创建并复核新知识库。
9. 在终端输出每个知识库的文件记录数、导入成功数和实际导出数。

自动导出的筛选条件如下：

- 知识库向量模型必须是 `text-embedding-v3`。
- 文件进度记录的状态必须是 `导入成功`。
- 接口必须返回有效的 `knowledgeProcProgressId`。
- 对应文件必须能拉取到知识条目；空文件不会生成 Excel。

不符合条件的知识库或文件不会下载，跳过原因会直接显示在终端中。

### 创建 `_bge` 知识库

默认不会创建新知识库，也不会点击创建表单的“确认”按钮。需要实际创建时，在确认账号有权限、目标编码不存在后，使用：

```powershell
python .\knowledge_query.py --create-bge
```

如果只执行创建阶段（不重新导出、不上传文件），使用：

```powershell
python .\knowledge_query.py --create-only
```

也可以在 `config.json` 中设置：

```json
"createBgeKnowledgeBases": true
```

命令行参数优先于配置文件。每个源知识库只有同时满足以下条件才会创建：

- 源编辑页向量模型为 `text-embedding-v3`；
- 至少存在一个文件状态为 `导入成功`；
- 至少成功导出一个文件。

创建入口使用知识库列表页右上角 UI。新库复制源编辑页字段，仅将编码改为 `<原编码>_bge`，向量模型改为 `bge-m3`；创建阶段不上传文件。若目标编码已存在则跳过，不覆盖或修改已有知识库。创建后脚本会重新搜索目标库并复核字段，结果写入源知识库的 `_metadata\knowledge_snapshot.json`。源编辑页只读取并点击“取消”关闭，不会点击“确定”。

### 上传 Excel 到 `_bge` 知识库

默认不会选择本地文件，也不会调用上传接口。确认目标库和文件清单后，使用：

```powershell
python .\knowledge_query.py --upload-bge
```

也可以在 `config.json` 中设置：

```json
"uploadBgeKnowledgeBases": true
```

上传开关与创建开关相互独立。上传阶段只使用已经存在并通过只读校验的 `_bge` 知识库，不会因为目标库不存在而自动创建。脚本会扫描：

```text
output/<医院名称>/<知识库名称>/*.xlsx
```

每个 Excel 单独上传，目标库中已存在同名文件会跳过；接口没有返回进度 ID 时，会通过上传前后文件列表对账，无法唯一关联时记录 `uploaded_untracked`，不会盲目重传。全部上传请求完成后，脚本按 2 秒、5 秒、15 秒递增间隔轮询文件状态，单个文件最多等待 5 分钟；最终状态会记录为 `import_success`、`import_failed`、`timeout`、`skipped_existing`、`upload_rejected` 或 `uploaded_untracked`。结果写入对应源知识库的 `_metadata\knowledge_snapshot.json` 的 `bgeUpload` 字段。

`--upload-bge` 会向第三方线上系统传输本地 Excel。正式运行前请先确认目标知识库和本地文件清单；默认配置保持关闭。

如果只执行上传阶段，使用：

```powershell
python .\knowledge_query.py --upload-only
```

`--upload-only` 会强制关闭 `_bge` 创建，只扫描已有快照和 Excel，上传到已存在的目标库，然后继续轮询并记录最终导入状态。

### 运行日志

两个脚本都会同时把运行信息输出到终端和日志文件。每次运行会生成一个带时间戳的 UTF-8 日志文件。

自动批量导出的默认日志目录是：C:\Users\ASUS\Desktop\export_knowldege\logs。

手动导出的默认日志目录是：C:\Users\ASUS\Desktop\export_knowldege\export_knowldege\output\logs。

也可以通过 --log-dir 指定日志目录，例如：

    python .\knowledge_query.py --log-dir "D:\logs\knowledge"

    python .\export_knowledge.py --knowledgeId "知识库ID" --knowledgeProcProgressId "进度ID" --token "登录Token" --log-dir "D:\logs\knowledge"

日志会记录完整链路的阶段进度，包括登录、知识库搜索、编辑配置采集、详情和文件进度读取、导出、`_bge` 创建与复核、逐文件上传、上传结果对账、状态轮询、快照写入、跳过原因、接口错误和最终汇总；账号、密码和 token 不会主动写入日志。

只导出一个知识库：

```powershell
python .\knowledge_query.py --knowledge-name "医生信息知识库"
```

一次指定多个知识库：

```powershell
python .\knowledge_query.py `
  --knowledge-name "医生信息知识库" `
  --knowledge-name "科室信息知识库"
```

无界面运行：

```powershell
python .\knowledge_query.py --headless
```

使用其他配置文件：

```powershell
python .\knowledge_query.py --config "D:\configs\hospital.json"
```

临时覆盖登录账号：

```powershell
python .\knowledge_query.py --username "your-username"
```

也可以使用环境变量提供账号和密码。使用这种方式时，请删除或留空 `config.json` 中的 `username` 和 `password`，因为配置文件中的值优先于环境变量：

```powershell
$env:HEREN_USERNAME = "your-username"
$env:HEREN_PASSWORD = "your-password"
python .\knowledge_query.py
```

## 手动启动方式：根据 ID 直接导出

当你已经从浏览器或接口中获得以下信息时，可以直接运行 `export_knowledge.py`：

- `knowledgeId`：知识库 ID
- `knowledgeProcProgressId`：知识库文件进度 ID
- `token`：登录请求头中的 `user-token`

运行命令：

```powershell
python .\export_knowledge.py `
  --knowledgeId "知识库ID" `
  --knowledgeProcProgressId "进度ID" `
  --token "登录Token"
```

Token 属于敏感登录凭据，不要写入脚本、配置模板或 Git 提交。直接作为命令参数使用时还应注意 PowerShell 历史记录；Token 过期后需要重新登录并获取。

注意：`export_knowledge.py` 当前通过代码中的 `OUT_DIR` 常量指定输出目录。如果项目移动到其他位置，需要同步修改该常量；自动批量导出脚本不受此限制。

## 导出结果

自动批量导出默认写入：

```text
output/<医院名称>/<知识库名称>/<原文件名>.xlsx
```

例如：

```text
output/demo/医生信息知识库/医生信息知识库.xlsx
```

如果不同文件生成了相同文件名，脚本会在文件名中增加进度 ID，避免互相覆盖。

## 查看命令帮助

```powershell
python .\knowledge_query.py --help
python .\export_knowledge.py --help
```

## 常见问题

### 无法启动 Microsoft Edge

确认系统已安装 Edge，并确认当前虚拟环境中已安装 Playwright：

```powershell
python -m pip show playwright
```

### 提示未找到配置文件

确认 `config.json` 与 `knowledge_query.py` 位于同一目录，或者通过 `--config` 指定配置文件的完整路径。

### 登录页出现验证码

当前脚本不会自动识别或绕过验证码。出现验证码时脚本会安全退出，需要先解决登录验证问题后重新运行。

### 返回 403 或提示 Token 过期

登录状态或 `user-token` 已失效。重新登录系统后再次运行自动脚本，或者为手动脚本提供最新 Token。

### 找不到知识库

检查 `knowledgeNames` 或 `--knowledge-name` 是否与管理系统中显示的知识库名称完全一致，并确认当前账号有访问权限。

### 知识库被跳过

自动脚本只导出向量模型为 `text-embedding-v3` 的知识库。其他模型会被跳过；如果无法读取编辑页配置，脚本会跳过当前知识库，不使用默认值代替。

### 文件被跳过

脚本只处理状态为 `导入成功` 的文件记录。导入中、导入失败或状态未知的文件会被跳过，终端会显示文件名、进度 ID 和当前状态。

### 运行结束但导出数量较少

运行结束时会按知识库输出汇总，包括文件记录总数、导入成功数和成功导出数。三者可能不同：未导入成功的记录会在下载前被过滤，导入成功但没有知识条目的空文件也不会生成 Excel。

## 使用 uv 运行（可选）

两个脚本都声明了内联依赖。电脑已安装 `uv` 时，可以让 `uv` 自动准备依赖：

```powershell
cd C:\Users\ASUS\Desktop\export_knowldege\export_knowldege
uv run --script .\knowledge_query.py
```

手动导出：

```powershell
uv run --script .\export_knowledge.py `
  --knowledgeId "知识库ID" `
  --knowledgeProcProgressId "进度ID" `
  --token "登录Token"
```
