## 1. 目标与范围

`bench_fill_full_exe.py` 批量运行各仓库的首次 silent 审查，并输出批次清单；`verify_batch_reports.py` 独立读取指定清单、核验报告并回填 Excel F 列。ask 和后续审查由人工处理，这两个脚本不会运行 ask 或改写 G/H 列。

## 2. 批处理流程

1. 读取配置、Excel 和 `PACKAGE_MAP`（压缩包名到 Excel 行号的映射）。
2. 解压样本仓库，从 Excel B 列读取 commit。
3. 使用仓库、commit 和用户自定义的检查内容构造审查请求，不需要生成或查找 `review.patch`。
4. 每个仓库执行一次 silent 审查，最多同时运行 20 个 iCode 会话。
5. 等所有 silent 会话结束后，生成脚本输出唯一的批次清单，记录 Excel、样本、session、运行目录及失败信息，不核验报告或修改 Excel。
6. 单独运行核验脚本，指定清单后逐个核验报告并判定结果；运行失败或没有生成有效报告的样本记录失败，不阻塞其他样本。
7. 核验脚本将结果回填 Excel F 列，写入 `bench_runs.jsonl`，保留 G/H 列原有内容。
8. ask 和记忆后的审查由人工另行处理。

检查内容通过环境变量 `CO_REVIEW_REVIEW_REQUEST` 自定义，默认为“检查该提交引入的安全风险”。Excel D 列为待确认主题，仅供审查参考，不作为已确认事实或自动回答。

并发会话共用当前 Co-Review 资源和已有知识库；脚本不备份、恢复或隔离知识库。应避免在批处理期间同时进行会写入共享资源的人工操作，以免影响实验基线。

## 3. iCode 与 Co-Review

每个会话启动 `icode.exe run <prompt> -a <临时profile> -C <仓库路径> --json`。临时 profile 从 `~/.chrys/agents/Code.yaml` 复制，并指向当前 Co-Review MCP server。Agent 执行 `co-review prepare --input <request.json>`，使用返回的任务和 `report_path` 完成审查并生成 JSON 报告。`icode run` 不支持运行时交互；silent 模式不请求用户补充信息。

脚本从 JSON 输出读取 `session_id`，通过已安装的 Co-Review iCode Hooks 查找对应 run 绑定，并将 session ID 写入日志，便于追踪原会话。`icode run` 始终绕过工具审批；批量脚本使用 `CO_REVIEW_RUN_TIMEOUT_SEC` 限制每次运行的总时长。

## 4. 报告核验与结果判定

生成脚本在所有会话结束后输出清单路径。核验脚本读取清单中的 `run_dir`，从其中的 `execution-state.json` 找到 `report_path`，校验报告包含 `repository`、`findings`、`remaining_questions`、`confirmation_history`，再根据首个 finding 的 `risk_type` 判定：

- `confirmed_vulnerability`：`是(risk_type=confirmed_vulnerability)`
- `not_vulnerable`：`是(risk_type=not_vulnerable)`
- `potential_risk`：`否(risk_type=potential_risk)`
- 无 findings：存在待确认问题时按 `potential_risk`，否则按 `not_vulnerable`。
- 报告缺失、格式异常或运行失败：记录错误，不当作有效结果。

结果只写 Excel F 列。回填前会检查 Excel 中的 commit 是否仍与批次一致；不一致或出现重复行号时停止回填。`bench_runs.jsonl` 包含样本、commit、F 结果、session ID、run 目录、报告路径和错误信息，可用于人工检查和复盘。

## 5. Windows 运行

创建环境并配置路径：

```powershell
powershell -ExecutionPolicy Bypass -File .\co-review_bench\setup_windows_env.ps1
.\.venv\Scripts\Activate.ps1
$env:CO_REVIEW_ICODE_EXE="C:\path\to\icode.exe"
$env:CO_REVIEW_ASSET_ROOT="C:\path\to\co-review\co-review"
$env:CO_REVIEW_RUNTIME_SRC="C:\path\to\co-review\runtime-src\scripts"
$env:CO_REVIEW_BENCH_XLSX="C:\path\to\bench_info.xlsx"
$env:CO_REVIEW_PKG_DIR="C:\path\to\pkgs"
$env:CO_REVIEW_WORK_BASE="C:\path\to\work"
$env:CO_REVIEW_RUN_TIMEOUT_SEC="600"
$env:CO_REVIEW_REVIEW_REQUEST="检查该提交的权限校验"
python .\co-review_bench\bench_fill_full_exe.py
```

运行前还需在脚本中设置 `PACKAGE_MAP`。生成结束后，使用输出的实际批次清单路径单独运行：

```powershell
python .\co-review_bench\verify_batch_reports.py "C:\path\to\work\batches\<batch-id>.json"
```

也可以通过一键入口指定工作目录和检查内容，只执行生成阶段，再按照输出的清单路径单独核验：

```powershell
powershell -ExecutionPolicy Bypass -File .\co-review_bench\run_all.ps1 -RootDir "D:\co-review-work" -ReviewRequest "检查该提交的权限校验"
```

脚本支持 Windows 默认路径及 Linux 路径回退，以上路径均可通过环境变量覆盖。`generate_review_patches.py` 保留为独立工具，不属于本批处理的执行流程。

## 6. 故障排查

1. 生成后查看 `batches` 目录中的对应清单；核验后查看 `bench_runs.jsonl` 中的 `error`、`run_dir` 和 `report_path`。
2. 检查 run 目录中的 `execution-state.json` 及报告 JSON。
3. 超时可调整 `CO_REVIEW_RUN_TIMEOUT_SEC`；缺少 session 绑定通常意味着 prepare 或工具链未完成。

## 7. 流程图

```mermaid
flowchart TD
  A[读取配置与 Excel] --> B[解压仓库并读取 commit]
  B --> C[自定义检查内容]
  C --> D[最多 20 个 silent 会话并发生成报告]
  D --> E[等待所有会话结束并保存批次清单]
  E --> F[单独运行核验脚本并判定结果]
  F --> G[回填 F 列并记录日志]
  G --> H[人工处理 ask 和后续审查]
