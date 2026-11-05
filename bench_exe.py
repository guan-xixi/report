"""完全自动化的 co-review 三模式批量实验脚本（iCode .exe 版）。

每个 Excel 样本独立执行 F/G/H 三种模式，全部通过 iCode ACP Host 和 Co-Review MCP 完成审查。

流程（每个样本）：
  F 基线：silent 模式，使用样本基线知识库。
  G ask：由 iCode ACP Host 提问；默认对 Excel D 列主题回答未知，只有显式配置经核验的答案时才补充事实。
  H 记忆后 silent：等待 G 的存储 worker 结束后复用其记忆，分别记录召回与进入检查计划的情况。
  报告由同一 iCode 主 Agent 按 Co-Review Skill 工作流生成，风险类型不预设。
  隔离可靠：每个样本结束后恢复知识库，防止跨样本污染。
  结果可追溯：日志记录每个模式的运行目录、报告路径、问题、风险类型、存储和记忆状态。

使用前必须修改：
  - PKG_DIR: 压缩包文件夹绝对路径
  - PACKAGE_MAP: 压缩包文件名 -> Excel 行号（从第2行开始）

运行：
  /data/cyjia/iCode/.venv/bin/python bench_fill_full_exe.py
"""
from __future__ import annotations

import asyncio
import contextlib
import json
import os
import shlex
import shutil
import sys
import uuid
import zipfile
from pathlib import Path

import openpyxl


IS_WINDOWS = os.name == "nt"
HOME = Path.home()
if IS_WINDOWS:
  DEFAULT_BENCH_XLSX = HOME / "iCode" / "bench_info.xlsx"
  DEFAULT_PKG_DIR = HOME / "co-review-pkgs"
  DEFAULT_WORK_BASE = Path(os.environ.get("TEMP", str(HOME / "AppData" / "Local" / "Temp"))) / "co-review-bench"
  DEFAULT_ICODE_EXE = HOME / "iCode" / "icode.exe"
  DEFAULT_ASSET_ROOT = HOME / "co-review" / "co-review"
  DEFAULT_RUNTIME_SRC = HOME / "co-review" / "runtime-src" / "scripts"
else:
  DEFAULT_BENCH_XLSX = Path("/data/cyjia/iCode/bench_info.xlsx")
  DEFAULT_PKG_DIR = Path("/tmp")
  DEFAULT_WORK_BASE = Path("/tmp/co-review-bench")
  DEFAULT_ICODE_EXE = Path("/data/cyjia/iCode/icode.exe")
  DEFAULT_ASSET_ROOT = Path("/data/cyjia/co-review/co-review")
  DEFAULT_RUNTIME_SRC = Path("/data/cyjia/co-review/runtime-src/scripts")

# ================= 用户必须修改的配置 =================
BENCH_XLSX = Path(os.environ.get("CO_REVIEW_BENCH_XLSX", str(DEFAULT_BENCH_XLSX)))
SHEET_NAME = "Sheet1"
PKG_DIR = Path(os.environ.get("CO_REVIEW_PKG_DIR", str(DEFAULT_PKG_DIR)))
WORK_BASE = Path(os.environ.get("CO_REVIEW_WORK_BASE", str(DEFAULT_WORK_BASE)))
ICODE = Path(os.environ.get("CO_REVIEW_ICODE_EXE", str(DEFAULT_ICODE_EXE)))
RUN_TIMEOUT_SEC = int(os.environ.get("CO_REVIEW_RUN_TIMEOUT_SEC", "600"))
ASK_TIMEOUT_SEC = 300

# 压缩包文件名 -> Excel 行号（从数据行第2行开始）。必须显式指定，防止错行。
PACKAGE_MAP = {
  "wazuh-cwe78-44bf114.zip": 2,
}

# Excel 列号（1-based）
COL_NAME = 1
COL_COMMIT = 2
COL_MISSING_FLAG = 3
COL_MISSING_INFO = 4
COL_EXPECTED = 5
COL_F = 6          # 初始 silent
COL_G = 7          # ask 模式（自动回答）
COL_H = 8          # 有记忆后的 silent

# 只填写已由样本证据核验的补充事实；未配置时 ask 自动答复为未知。
SUPPLEMENTAL_ANSWERS = {}
# =====================================================

os.environ.setdefault("CO_REVIEW_ASSET_ROOT", str(DEFAULT_ASSET_ROOT))

SRC = os.environ.get("CO_REVIEW_RUNTIME_SRC", str(DEFAULT_RUNTIME_SRC))
if SRC not in sys.path:
  sys.path.insert(0, SRC)

from runtime_paths import asset_root  # noqa: E402
from input_recording import save_json, read_json, get_binding  # noqa: E402

SKILL_ROOT = asset_root()
KNOWLEDGE_PATH = SKILL_ROOT / "knowledge.json"
CHRYS_CONFIG_DIR = Path.home() / ".chrys"
STORE_TIMEOUT_SEC = 300


def find_review_patch(repo_path):
  """在仓库根目录及上一级目录查找 review.patch。"""
  candidates = list(repo_path.rglob("review.patch"))
  if candidates:
    return candidates[0]
  parent = repo_path.parent
  candidates = list(parent.rglob("review.patch"))
  if candidates:
    return candidates[0]
  return None


class BenchACPClient:
  def __init__(self, answer: str):
    self.answer = answer
    self.questions = []
    self.messages = []
    self.tool_updates = []

  async def ext_method(self, method, params):
    if method != "chrys/request_input":
      return {}
    questions = params.get("questions", [])
    self.questions.extend(questions)
    return {
      "answers": [{"values": [self.answer], "note": ""} for _ in questions],
      "cancelled": False,
    }

  async def session_update(self, session_id, update, **kwargs):
    payload = update.model_dump(mode="json", by_alias=True, exclude_none=True)
    if payload.get("sessionUpdate", "").startswith("tool_call"):
      self.tool_updates.append({
        key: payload.get(key)
        for key in ("sessionUpdate", "toolCallId", "title", "status", "kind")
        if payload.get(key) is not None
      })
    if payload.get("sessionUpdate") == "agent_message_chunk":
      text = payload.get("content", {}).get("text")
      if isinstance(text, str):
        self.messages.append(text)


class BenchACPTimeout(TimeoutError):
  def __init__(self, timeout_sec, diagnostics):
    super().__init__(f"iCode ACP turn exceeded {timeout_sec} seconds")
    self.diagnostics = diagnostics


async def run_icode_session(repo_path, prompt, answer, timeout_sec, wait_for_store):
  """Run one complete Co-Review turn through the iCode ACP host."""
  from acp import PROTOCOL_VERSION
  from acp.client import ClientSideConnection
  from acp.schema import ClientCapabilities, Implementation, TextContentBlock

  profile_name, profile_path = create_bench_profile()
  env = dict(os.environ)
  env["CHRYS_CONFIG_DIR"] = str(CHRYS_CONFIG_DIR)
  env["CO_REVIEW_ASSET_ROOT"] = str(SKILL_ROOT)
  env["PYTHONUTF8"] = "1"
  env["PYTHONIOENCODING"] = "utf-8"
  process = None
  connection = None
  stderr_task = None
  prompt_task = None
  timeout_diagnostics = None
  try:
    process = await asyncio.create_subprocess_exec(
      str(ICODE), "acp", "-a", profile_name, "--approval", "bypass",
      "--ask-user-timeout", str(ASK_TIMEOUT_SEC), "-C", str(repo_path),
      stdin=asyncio.subprocess.PIPE,
      stdout=asyncio.subprocess.PIPE,
      stderr=asyncio.subprocess.PIPE,
      cwd=str(repo_path),
      env=env,
    )
    client = BenchACPClient(answer)
    connection = ClientSideConnection(client, process.stdin, process.stdout)
    stderr_task = asyncio.create_task(process.stderr.read())
    await connection.initialize(
      PROTOCOL_VERSION,
      client_capabilities=ClientCapabilities(),
      client_info=Implementation(name="co-review-bench", version="1.0"),
    )
    session = await connection.new_session(cwd=str(repo_path))
    prompt_task = asyncio.create_task(
      connection.prompt(
        [TextContentBlock(type="text", text=prompt)],
        session_id=session.session_id,
      )
    )
    done, _ = await asyncio.wait({prompt_task}, timeout=timeout_sec)
    if not done:
      binding = get_binding(session.session_id, SKILL_ROOT)
      timeout_diagnostics = {
        "session_id": session.session_id,
        "run_dir": binding.get("run_dir") if binding else None,
        "questions": client.questions,
        "tool_updates": client.tool_updates[-20:],
        "answer_tail": "".join(client.messages)[-1200:],
      }
      raise BenchACPTimeout(timeout_sec, timeout_diagnostics)
    response = prompt_task.result()
    if response.stop_reason != "end_turn":
      raise RuntimeError(f"iCode turn ended with stop_reason={response.stop_reason}")
    binding = get_binding(session.session_id, SKILL_ROOT)
    store_status = (
      await wait_for_memory_store(Path(binding["run_dir"]))
      if binding and wait_for_store else "not_applicable"
    )
    return {
      "session_id": session.session_id,
      "answer": "".join(client.messages),
      "questions": client.questions,
      "store_status": store_status,
    }
  finally:
    if process and process.returncode is None:
      process.terminate()
      try:
        await asyncio.wait_for(process.wait(), timeout=10)
      except asyncio.TimeoutError:
        process.kill()
        await process.wait()
    if prompt_task and not prompt_task.done():
      prompt_task.cancel()
      await asyncio.gather(prompt_task, return_exceptions=True)
    if connection:
      with contextlib.suppress(Exception):
        await connection.close()
    if stderr_task:
      stderr = await stderr_task
      if timeout_diagnostics is not None:
        timeout_diagnostics["host_stderr_tail"] = stderr.decode("utf-8", "replace")[-2000:]
    profile_path.unlink(missing_ok=True)


async def wait_for_memory_store(run_dir):
  request_path = run_dir / "store-request.json"
  state_path = run_dir / "store-worker-state.json"
  deadline = asyncio.get_running_loop().time() + STORE_TIMEOUT_SEC
  while not request_path.exists():
    if asyncio.get_running_loop().time() >= deadline:
      return "not_started"
    await asyncio.sleep(0.25)
  while not state_path.exists():
    if asyncio.get_running_loop().time() >= deadline:
      return "timeout"
    await asyncio.sleep(0.25)
  while True:
    state = read_json(state_path)
    if state.get("status") in {"completed", "failed"}:
      return state["status"]
    if asyncio.get_running_loop().time() >= deadline:
      return "timeout"
    await asyncio.sleep(0.5)


def build_request(repo_path, repo_name, mode, commit, review_patch_path):
  patch_str = str(review_patch_path) if review_patch_path else "(未找到 review.patch)"
  return {
    "user_request": (
      f"使用 Co-Review 检视仓库 {repo_name} 的 commit {commit}，"
      f"检查 review.patch：{patch_str}。"
    ),
    "repository": repo_name,
    "repository_path": str(repo_path.resolve()),
    "interaction_mode": mode,
  }


def build_agent_prompt(request_path, repo_path, mode, missing_info, answer):
  command = f"{shlex.quote(str(SKILL_ROOT / 'co-review'))} prepare --input {shlex.quote(str(request_path))}"
  answer_instruction = (
    "收到 ask_user 问题时，只回答以下实验信息；不要推断或扩展：" + answer
    if mode == "ask" else "本轮为 silent，不向用户提问或请求补充信息。"
  )
  return (
    f"请阅读并遵循 Co-Review Skill：{SKILL_ROOT / 'SKILL.md'}。\n"
    f"在仓库 {repo_path} 中通过 shell 工具执行准备命令：\n{command}\n"
    "必须使用该命令返回的真实任务和 report_path，继续使用 Co-Review MCP 工具完成所有检查；"
    "不要自行伪造 Planner 任务、绕过 Co-Review 工作流或另起报告生成流程。\n"
    f"模式：{mode}。{answer_instruction}\n"
    "只做静态审查，不修改仓库文件。按 Skill 的报告结构，将基于代码证据的最终 JSON"
    "保存到准备材料返回的 report_path，并在结束前确认文件可解析。\n"
    f"样本记录的待确认主题（不是已确认事实）：{missing_info or '无'}。"
  )


def create_bench_profile():
  import yaml

  profile_dir = CHRYS_CONFIG_DIR / "agents"
  profile_dir.mkdir(parents=True, exist_ok=True)
  original = yaml.safe_load((profile_dir / "Code.yaml").read_text(encoding="utf-8-sig"))
  profile_name = "CoReviewBench-" + uuid.uuid4().hex[:8]
  profile = dict(original)
  profile["id"] = uuid.uuid4().hex[:12]
  profile["name"] = profile_name
  tools = dict(profile.get("tools", {}))
  mcp_servers = [dict(server) for server in tools.get("mcp", [])]
  co_review = [server for server in mcp_servers if server.get("name") == "co-review"]
  if len(co_review) != 1:
    raise RuntimeError("Code profile must contain exactly one Co-Review MCP server")
  co_review[0]["command"] = str(SKILL_ROOT / "co-review")
  tools["mcp"] = mcp_servers
  profile["tools"] = tools
  path = profile_dir / (profile_name + ".yaml")
  with path.open("x", encoding="utf-8") as stream:
    yaml.safe_dump(profile, stream, allow_unicode=True, sort_keys=False)
  return profile_name, path


def backup_knowledge():
  if KNOWLEDGE_PATH.exists():
    return read_json(KNOWLEDGE_PATH)
  return {"memories": []}


def restore_knowledge(backup):
  save_json(KNOWLEDGE_PATH, backup)


def count_memories_for_repo(repo_name):
  data = read_json(KNOWLEDGE_PATH)
  return sum(1 for m in data.get("memories", []) if m.get("source_repository") == repo_name)


def repo_memory_contents(repo_name):
  data = read_json(KNOWLEDGE_PATH)
  return {
    memory.get("content", "")
    for memory in data.get("memories", [])
    if memory.get("source_repository") == repo_name and memory.get("confirmed") is True
  }


def read_h_memory_application(run_dir, new_memory_contents):
  state_path = run_dir / "execution-state.json"
  if not state_path.exists():
    return False
  plan_path = run_dir / "plan.json"
  if not plan_path.exists():
    return False
  plan = read_json(plan_path)
  applied = {
    item
    for task in plan.get("tasks", [])
    for item in task.get("known_information", {}).get("repository", [])
  }
  return bool(new_memory_contents & applied)


def read_memory_recall(run_dir):
  """读取 memory-recall.jsonl 最后一行，返回召回情况。"""
  mr = run_dir / "memory-recall.jsonl"
  if not mr.exists():
    return None
  lines = [ln for ln in mr.read_text(encoding="utf-8").splitlines() if ln.strip()]
  if not lines:
    return None
  return json.loads(lines[-1])


def read_report_risk_type(run_dir):
  state_path = run_dir / "execution-state.json"
  if not state_path.exists():
    return None, "无 execution-state.json", None
  state = read_json(state_path)
  report_path = state.get("report_path")
  if not report_path:
    return None, "report_path 为空", None
  rp = Path(report_path)
  if not rp.exists():
    return None, f"报告文件不存在：{rp}", None
  try:
    data = read_json(rp)
    required = {"repository", "findings", "remaining_questions", "confirmation_history"}
    if not isinstance(data, dict) or not required <= data.keys():
      return None, "报告不符合 Co-Review 顶层结构", str(rp)
    if data["repository"] == "" or not isinstance(data["findings"], list):
      return None, "报告 repository/findings 无效", str(rp)
    if not isinstance(data["remaining_questions"], list) or not isinstance(data["confirmation_history"], list):
      return None, "报告问题或确认历史字段无效", str(rp)
    findings = data["findings"]
    if findings:
      risk_type = findings[0].get("risk_type")
      if risk_type not in {"confirmed_vulnerability", "potential_risk", "not_vulnerable"}:
        return None, f"未知 risk_type：{risk_type}", str(rp)
      return risk_type, None, str(rp)
    return ("potential_risk" if data["remaining_questions"] else "not_vulnerable"), None, str(rp)
  except Exception as exc:
    return None, f"报告解析失败：{exc}", str(rp)


def classify_risk(risk_type):
  if risk_type == "confirmed_vulnerability":
    return "是(risk_type=confirmed_vulnerability)"
  if risk_type == "not_vulnerable":
    return "是(risk_type=not_vulnerable)"
  if risk_type in ("potential", "potential_risk"):
    return f"否(risk_type={risk_type})"
  return "报告缺失/异常"


def run_single_mode(repo_path, repo_name, mode, commit, review_patch_path,
          missing_info="", timeout_sec=RUN_TIMEOUT_SEC):
  result = {
    "run_dir": None,
    "report_path": None,
    "risk_type": None,
    "val": "运行失败",
    "recall": None,
    "memories_before": count_memories_for_repo(repo_name),
    "memories_after": 0,
    "questions": [],
    "error": None,
  }
  answer = SUPPLEMENTAL_ANSWERS.get(repo_name, "无法确认该信息；请将该条件保留为未知，不要据此推断。")
  request_path = WORK_BASE / "requests" / f"{repo_name}-{mode}-{uuid.uuid4().hex}.json"
  save_json(request_path, build_request(repo_path, repo_name, mode, commit, review_patch_path))

  try:
    session_result = asyncio.run(run_icode_session(
      repo_path,
      build_agent_prompt(request_path, repo_path, mode, missing_info, answer),
      answer,
      timeout_sec,
      wait_for_store=mode == "ask",
    ))
    result["questions"] = session_result["questions"]
    result["store_status"] = session_result["store_status"]
    binding = get_binding(session_result["session_id"], SKILL_ROOT)
    if not binding:
      raise RuntimeError("iCode Host 没有为 Co-Review session 建立 run 绑定")
    run_dir = Path(binding["run_dir"])
    result["run_dir"] = str(run_dir)
    state_path = run_dir / "execution-state.json"
    if not state_path.exists():
      raise RuntimeError("Co-Review workflow 未生成 execution-state.json")
    state = read_json(state_path)
    report_path = Path(state["report_path"])
    result["report_path"] = str(report_path) if report_path.exists() else None
    risk, report_error, _ = read_report_risk_type(run_dir)
    result["risk_type"] = risk
    result["val"] = classify_risk(risk) if risk else f"报告无效({report_error})"
    result["recall"] = read_memory_recall(run_dir)
    result["memories_after"] = count_memories_for_repo(repo_name)
    if result["val"].startswith("报告无效"):
      result["error"] = report_error
  except BenchACPTimeout as exc:
    diagnostics = exc.diagnostics
    result["run_dir"] = diagnostics.get("run_dir")
    result["questions"] = diagnostics.get("questions", [])
    result["timeout_diagnostics"] = diagnostics
    result["error"] = str(exc)
    result["val"] = f"运行失败({exc})"
  except Exception as exc:
    result["error"] = str(exc)
    result["val"] = f"运行失败({exc})"

  return result


def find_repo_root(dest):
  subs = [d for d in dest.iterdir() if d.is_dir()]
  if len(subs) == 1 and not (dest / ".git").exists():
    return subs[0]
  return dest


def extract_archive(zf, dest):
  dest.mkdir(parents=True, exist_ok=True)
  name = zf.name.lower()
  if name.endswith(".zip"):
    zipfile.ZipFile(zf).extractall(dest)
  elif name.endswith((".tar.gz", ".tgz")):
    shutil.unpack_archive(str(zf), str(dest), "gztar")
  elif name.endswith(".tar"):
    shutil.unpack_archive(str(zf), str(dest), "tar")
  else:
    shutil.unpack_archive(str(zf), str(dest))
  return find_repo_root(dest)


def strip_ext(filename):
  name = filename
  for ext in (".tar.gz", ".tgz", ".tar", ".zip"):
    if name.lower().endswith(ext):
      return name[: -len(ext)]
  return Path(filename).stem


def main():
  WORK_BASE.mkdir(parents=True, exist_ok=True)
  log_path = WORK_BASE / "bench_runs.jsonl"

  wb = openpyxl.load_workbook(BENCH_XLSX)
  if SHEET_NAME not in wb.sheetnames:
    raise SystemExit(f"工作表 {SHEET_NAME} 不存在，当前：{wb.sheetnames}")
  ws = wb[SHEET_NAME]

  if not PKG_DIR.exists():
    raise SystemExit(f"压缩包目录不存在：{PKG_DIR}")

  if not ICODE.is_file():
    raise SystemExit(f"iCode executable not found: {ICODE}")

  original_knowledge = backup_knowledge()
  count = 0
  try:
    for zip_name, row in PACKAGE_MAP.items():
      zf = PKG_DIR / zip_name
      if not zf.is_file():
        print(f"[跳过] 压缩包不存在：{zf}")
        continue

      commit = ws.cell(row=row, column=COL_COMMIT).value
      commit = str(commit).strip() if commit is not None else ""
      missing_info = ws.cell(row=row, column=COL_MISSING_INFO).value
      missing_info = str(missing_info).strip() if missing_info is not None else ""
      expected = ws.cell(row=row, column=COL_EXPECTED).value
      repo_name = strip_ext(zip_name)

      print(f"\n===== [{count+1}] {zip_name} -> Excel 行 {row} =====")
      print(f"  commit={commit}, 待确认主题={missing_info}, 预期={expected}")

      dest = WORK_BASE / "samples" / f"{repo_name}-{uuid.uuid4().hex[:8]}"
      try:
        repo_path = extract_archive(zf, dest)
      except Exception as exc:
        print(f"[失败] 解压 {zip_name}: {exc}")
        continue

      review_patch = find_review_patch(repo_path)
      if review_patch is None:
        print("[失败] 仓库中未找到 review.patch，跳过 F/G/H")
        continue

      sample_knowledge_backup = backup_knowledge()
      f_res = run_single_mode(repo_path, repo_name, "silent", commit, review_patch,
                  missing_info=missing_info)
      restore_knowledge(sample_knowledge_backup)
      g_res = run_single_mode(repo_path, repo_name, "ask", commit, review_patch,
                  missing_info=missing_info)
      new_memory_contents = repo_memory_contents(repo_name) - {
        memory.get("content", "") for memory in sample_knowledge_backup.get("memories", [])
        if memory.get("source_repository") == repo_name and memory.get("confirmed") is True
      }
      h_res = run_single_mode(repo_path, repo_name, "silent", commit, review_patch,
                  missing_info=missing_info)
      h_res["memory_applied"] = bool(
        h_res["run_dir"] and read_h_memory_application(Path(h_res["run_dir"]), new_memory_contents)
      )

      ws.cell(row=row, column=COL_F, value=f_res["val"])
      ws.cell(row=row, column=COL_G, value=g_res["val"])
      ws.cell(row=row, column=COL_H, value=h_res["val"])
      log_row = {
        "zip": zip_name,
        "excel_row": row,
        "repo": repo_name,
        "commit": commit,
        "missing_info_topic": missing_info,
        "verified_answer_configured": repo_name in SUPPLEMENTAL_ANSWERS,
        "review_patch": str(review_patch),
        "g_new_memory_contents": sorted(new_memory_contents),
        "f": f_res,
        "g": g_res,
        "h": h_res | {"recall_hit": bool(h_res["recall"] and h_res["recall"].get("memory_indices"))},
      }
      with log_path.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(log_row, ensure_ascii=False) + "\n")

      restore_knowledge(sample_knowledge_backup)
      count += 1
      print(f"  F={f_res['val']}  G={g_res['val']}  H={h_res['val']} "
          f"(G 存储={g_res.get('store_status')}, H 记忆应用={h_res['memory_applied']})")
  finally:
    restore_knowledge(original_knowledge)

  wb.save(BENCH_XLSX)
  print(f"\n完成，共处理 {count} 个仓库。F/G/H 三列已填入 {BENCH_XLSX}。")
  print(f"详细旁证：{log_path}")


if __name__ == "__main__":
  main()
  main()
