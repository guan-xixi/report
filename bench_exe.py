#!/usr/bin/env python3
"""并发运行 silent 基线，生成供独立核验脚本使用的批次清单。

使用前必须修改：
  - PKG_DIR: 压缩包文件夹绝对路径
  - PACKAGE_MAP: 压缩包文件名 -> Excel 行号（从第2行开始）

运行：
  /data/cyjia/iCode/.venv/bin/python bench_fill_full_exe.py
"""
from __future__ import annotations

import asyncio
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
REVIEW_REQUEST = os.environ.get("CO_REVIEW_REVIEW_REQUEST", "检查该提交引入的安全风险")
MAX_CONCURRENT_SESSIONS = 20

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
# =====================================================

os.environ.setdefault("CO_REVIEW_ASSET_ROOT", str(DEFAULT_ASSET_ROOT))

SRC = os.environ.get("CO_REVIEW_RUNTIME_SRC", str(DEFAULT_RUNTIME_SRC))
if SRC not in sys.path:
  sys.path.insert(0, SRC)

from runtime_paths import asset_root  # noqa: E402
from input_recording import save_json, get_binding  # noqa: E402

SKILL_ROOT = asset_root()
CHRYS_CONFIG_DIR = Path.home() / ".chrys"


class BenchRunTimeout(TimeoutError):
  def __init__(self, timeout_sec, diagnostics):
    super().__init__(f"iCode run exceeded {timeout_sec} seconds")
    self.diagnostics = diagnostics


async def run_icode_session(repo_path, prompt, timeout_sec):
  profile_name, profile_path = create_bench_profile()
  env = dict(os.environ)
  env["CHRYS_CONFIG_DIR"] = str(CHRYS_CONFIG_DIR)
  env["CO_REVIEW_ASSET_ROOT"] = str(SKILL_ROOT)
  env["PYTHONUTF8"] = "1"
  env["PYTHONIOENCODING"] = "utf-8"
  process = None
  communicate_task = None
  try:
    process = await asyncio.create_subprocess_exec(
      str(ICODE), "run", prompt, "-a", profile_name, "-C", str(repo_path), "--json",
      stdin=asyncio.subprocess.DEVNULL,
      stdout=asyncio.subprocess.PIPE,
      stderr=asyncio.subprocess.PIPE,
      cwd=str(repo_path),
      env=env,
    )
    communicate_task = asyncio.create_task(process.communicate())
    try:
      stdout, stderr = await asyncio.wait_for(asyncio.shield(communicate_task), timeout=timeout_sec)
    except asyncio.TimeoutError:
      if process.returncode is None:
        process.kill()
      stdout, stderr = await communicate_task
      raise BenchRunTimeout(timeout_sec, {
        "stdout_tail": stdout.decode("utf-8", "replace")[-2000:],
        "stderr_tail": stderr.decode("utf-8", "replace")[-2000:],
      })
    if process.returncode != 0:
      detail = stderr.decode("utf-8", "replace").strip()
      raise RuntimeError(f"iCode run failed (exit {process.returncode}): {detail[-2000:]}")
    try:
      response = json.loads(stdout.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
      raise RuntimeError(f"iCode run 返回无效 JSON: {exc}") from exc
    if not isinstance(response, dict) or not isinstance(response.get("session_id"), str) or not response["session_id"]:
      raise RuntimeError("iCode run 返回结果缺少 session_id")
    return {
      "session_id": response["session_id"],
      "answer": response.get("result", ""),
      "questions": [],
    }
  finally:
    if process and process.returncode is None:
      process.terminate()
      try:
        await asyncio.wait_for(process.wait(), timeout=10)
      except asyncio.TimeoutError:
        process.kill()
        await process.wait()
    if communicate_task and not communicate_task.done():
      await communicate_task
    profile_path.unlink(missing_ok=True)


def build_request(repo_path, repo_name, commit):
  return {
    "user_request": (
      f"使用 Co-Review 检视仓库 {repo_name} 的 commit {commit}，"
      f"{REVIEW_REQUEST}。"
    ),
    "repository": repo_name,
    "repository_path": str(repo_path.resolve()),
    "interaction_mode": "silent",
  }


def build_agent_prompt(request_path, repo_path, missing_info):
  command = f"{shlex.quote(str(SKILL_ROOT / 'co-review'))} prepare --input {shlex.quote(str(request_path))}"
  return (
    f"请阅读并遵循 Co-Review Skill：{SKILL_ROOT / 'SKILL.md'}。\n"
    f"在仓库 {repo_path} 中通过 shell 工具执行准备命令：\n{command}\n"
    "必须使用该命令返回的真实任务和 report_path，继续使用 Co-Review MCP 工具完成所有检查；"
    "不要自行伪造 Planner 任务、绕过 Co-Review 工作流或另起报告生成流程。\n"
    "模式：silent，不向用户提问或请求补充信息。\n"
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


async def run_single_mode(repo_path, repo_name, commit, missing_info="", timeout_sec=RUN_TIMEOUT_SEC):
  result = {
    "session_id": None,
    "run_dir": None,
    "report_path": None,
    "risk_type": None,
    "val": "运行失败",
    "recall": None,
    "questions": [],
    "error": None,
  }
  try:
    request_path = WORK_BASE / "requests" / f"{repo_name}-silent-{uuid.uuid4().hex}.json"
    save_json(request_path, build_request(repo_path, repo_name, commit))
    session_result = await run_icode_session(
      repo_path,
      build_agent_prompt(request_path, repo_path, missing_info),
      timeout_sec,
    )
    result["session_id"] = session_result["session_id"]
    result["questions"] = session_result["questions"]
    binding = get_binding(session_result["session_id"], SKILL_ROOT)
    if not binding:
      raise RuntimeError("iCode run 没有为 Co-Review session 建立 run 绑定")
    run_dir = Path(binding["run_dir"])
    result["run_dir"] = str(run_dir)
    result["val"] = "待核验"
  except BenchRunTimeout as exc:
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


async def generate_reports(samples):
  semaphore = asyncio.Semaphore(MAX_CONCURRENT_SESSIONS)

  async def generate(sample):
    async with semaphore:
      try:
        zf = PKG_DIR / sample["zip"]
        if not zf.is_file():
          raise FileNotFoundError(f"压缩包不存在：{zf}")
        dest = WORK_BASE / "samples" / f"{sample['repo']}-{uuid.uuid4().hex[:8]}"
        repo_path = extract_archive(zf, dest)
        print(f"[生成] {sample['zip']} -> Excel 行 {sample['excel_row']}", flush=True)
        return await run_single_mode(repo_path, sample["repo"], sample["commit"], sample["missing_info_topic"])
      except Exception as exc:
        return {"session_id": None, "run_dir": None, "report_path": None, "risk_type": None,
            "val": f"运行失败({exc})", "recall": None, "questions": [], "error": str(exc)}

  return await asyncio.gather(*(generate(sample) for sample in samples))


def main():
  WORK_BASE.mkdir(parents=True, exist_ok=True)

  wb = openpyxl.load_workbook(BENCH_XLSX)
  if SHEET_NAME not in wb.sheetnames:
    raise SystemExit(f"工作表 {SHEET_NAME} 不存在，当前：{wb.sheetnames}")
  ws = wb[SHEET_NAME]

  if not PKG_DIR.exists():
    raise SystemExit(f"压缩包目录不存在：{PKG_DIR}")

  if not ICODE.is_file():
    raise SystemExit(f"iCode executable not found: {ICODE}")

  samples = []
  for zip_name, row in PACKAGE_MAP.items():
    commit = ws.cell(row=row, column=COL_COMMIT).value
    missing_info = ws.cell(row=row, column=COL_MISSING_INFO).value
    samples.append({
      "zip": zip_name,
      "excel_row": row,
      "repo": strip_ext(zip_name),
      "commit": str(commit).strip() if commit is not None else "",
      "missing_info_topic": str(missing_info).strip() if missing_info is not None else "",
    })

  results = asyncio.run(generate_reports(samples))
  manifest_path = WORK_BASE.resolve() / "batches" / f"{uuid.uuid4().hex}.json"
  save_json(manifest_path, {
    "workbook": str(BENCH_XLSX.resolve()),
    "sheet": SHEET_NAME,
    "work_base": str(WORK_BASE.resolve()),
    "records": [sample | {"f": result} for sample, result in zip(samples, results)],
  })
  print(f"\n所有 silent 会话已结束，共处理 {len(samples)} 个仓库，Excel 未修改。")
  print(f"批次清单：{manifest_path}")
  print("请使用 verify_batch_reports.py 和该清单单独核验结果。")


if __name__ == "__main__":
  main()
