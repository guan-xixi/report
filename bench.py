#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iCode `run --json` + Co-Review F/G/H 批量实验。

重要边界：
- 用 iCode CLI 主 Agent 执行任务，不直接调用模型 API。
- 每轮执行使用 `icode run PROMPT --agent Code --workdir REPO --json`。
- `--json` stdout 解析为 {session_id, result, duration}；stderr 单独保存。
- Co-Review 的 prepare/context/report API 在不同版本可能不同，因此这里不猜测私有 Python API；
  由 Co-Review Skill 在 iCode Agent 中运行标准工作流，并检查该 run 目录中是否产生正式报告。
- 如果你的 Co-Review MCP 工具要求 Hook 注入 `_co_review_context`，必须确保当前 CLI 的 Skill/MCP Hook 能注入；
  如果仍报 "The iCode Hook must inject the current host context"，脚本会将该模式标记失败，不会伪造报告。
- G 的 Excel D 列是预置补充线索，不等于事实；必须让 Agent 基于代码证据验证后再决定是否形成记忆。
- 知识库在 F/G/H 间按基线隔离：F 用初始知识；G 从初始知识运行；H 使用 G 运行后真实产生的知识；
  样本结束后恢复基线，避免跨样本污染。
- 仅当存在有效报告时才向 Excel 写入 risk_type；否则写入错误状态。

环境变量：
  ICODE_EXE=/path/to/icode             # Linux CLI 或 Windows exe 的实际路径
  ICODE_AGENT=Code                     # icode agent list 中的 Name/Display Name/ID
  ICODE_MODEL=                         # 可选；仅 active 模型 Agent 支持
  BENCH_XLSX=/data/cyjia/iCode/bench_info.xlsx
  PKG_DIR=/tmp
  WORK_BASE=/tmp/co-review-bench
  COREVIEW_ROOT=/data/cyjia/co-review
  COREVIEW_SKILL=/data/cyjia/co-review/SKILL.md
  ICODE_TIMEOUT=600

用法：
  python bench_fill_full.py --check
  python bench_fill_full.py --rows 2 --timeout 600
  python bench_fill_full.py --rows 2,3 --reset-results
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tarfile
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any

try:
    import openpyxl
except ImportError:
    raise SystemExit("缺少 openpyxl，请在当前 Python 环境安装：pip install openpyxl")


# ============================ 用户配置 ============================
BENCH_XLSX = Path(os.environ.get("BENCH_XLSX", "/data/cyjia/iCode/bench_info.xlsx"))
SHEET_NAME = os.environ.get("BENCH_SHEET", "Sheet1")
PKG_DIR = Path(os.environ.get("PKG_DIR", "/tmp"))
WORK_BASE = Path(os.environ.get("WORK_BASE", "/tmp/co-review-bench"))

# 文件名 -> Excel 行号；按你的实验数据继续补充映射。
PACKAGE_MAP = {
    "wazuh-cwe78-44bf114.zip": 2,
}

ICODE_EXE = os.environ.get("ICODE_EXE", "/data/cyjia/iCode/icode")
ICODE_AGENT = os.environ.get("ICODE_AGENT", "Code")
ICODE_MODEL = os.environ.get("ICODE_MODEL", "").strip()
ICODE_TIMEOUT = int(os.environ.get("ICODE_TIMEOUT", "600"))

COREVIEW_ROOT = Path(os.environ.get("COREVIEW_ROOT", "/data/cyjia/co-review"))
COREVIEW_SKILL = Path(os.environ.get("COREVIEW_SKILL", str(COREVIEW_ROOT / "SKILL.md")))
# 如果 Co-Review 仓库里的 Skill 实际路径不同，请设置 COREVIEW_SKILL。
# Excel 列号
COL_NAME = 1
COL_COMMIT = 2
COL_MISSING_FLAG = 3
COL_MISSING_INFO = 4
COL_EXPECTED = 5
COL_F = 6
COL_G = 7
COL_H = 8

# 支持的最终 risk_type。根据你实际 Co-Review 报告 schema 可补充。
VALID_RISK_TYPES = {
    "confirmed_vulnerability",
    "potential",
    "potential_risk",
    "not_vulnerable",
}
# ================================================================


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha256_file(path: Path) -> str | None:
    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def stable_hash(value: Any) -> str:
    return hashlib.sha256(
        json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    ).hexdigest()


def load_json(path: Path, default: Any = None) -> Any:
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp.replace(path)


def append_jsonl(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(value, ensure_ascii=False) + "\n")


def locate_executable(executable: str) -> str | None:
    p = Path(executable).expanduser()
    if p.exists() and p.is_file():
        return str(p.resolve())
    return shutil.which(executable)


def run_command(cmd: list[str], cwd: Path | None, timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd, cwd=str(cwd) if cwd else None,
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout, check=False,
    )


def check_environment() -> list[str]:
    errors = []
    exe = locate_executable(ICODE_EXE)
    if not exe:
        errors.append(f"找不到 iCode CLI：{ICODE_EXE}")
    elif not os.access(exe, os.X_OK) and os.name != "nt":
        errors.append(f"iCode 文件不可执行：{exe}（可尝试 chmod +x）")
    if not BENCH_XLSX.is_file():
        errors.append(f"Excel 文件不存在：{BENCH_XLSX}")
    if not PKG_DIR.is_dir():
        errors.append(f"压缩包目录不存在：{PKG_DIR}")
    if not COREVIEW_ROOT.is_dir():
        errors.append(f"Co-Review 仓库不存在：{COREVIEW_ROOT}")
    if not COREVIEW_SKILL.is_file():
        errors.append(
            f"Co-Review Skill 文件不存在：{COREVIEW_SKILL}。"
            "请将 COREVIEW_SKILL 配置为实际 SKILL.md 路径。"
        )
    if not errors and exe:
        try:
            r = run_command([exe, "run", "--help"], cwd=Path.cwd(), timeout=30)
            if r.returncode != 0:
                errors.append(
                    "iCode `run --help` 未成功："
                    + (r.stderr or r.stdout or f"exit={r.returncode}")[:1200]
                )
            else:
                help_text = r.stdout + "\n" + r.stderr
                for flag in ("--agent", "--json", "--workdir"):
                    if flag not in help_text:
                        errors.append(f"`icode run --help` 输出中未发现 {flag}；请确认 CLI 版本。")
        except Exception as e:
            errors.append(f"执行 `icode run --help` 失败：{type(e).__name__}: {e}")
    return errors


def safe_extract_zip(archive: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with zipfile.ZipFile(archive) as zf:
        for member in zf.infolist():
            target = (destination / member.filename).resolve()
            if target != root and root not in target.parents:
                raise ValueError(f"压缩包路径穿越风险：{member.filename}")
        zf.extractall(destination)


def extract_archive(archive: Path, destination: Path) -> Path:
    if destination.exists():
        shutil.rmtree(destination)
    destination.mkdir(parents=True, exist_ok=True)
    low = archive.name.lower()
    if low.endswith(".zip"):
        safe_extract_zip(archive, destination)
    elif low.endswith((".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tar.xz")):
        # Python tarfile 的 data filter 在支持的版本上阻止路径穿越/危险链接。
        with tarfile.open(archive, "r:*") as tf:
            try:
                tf.extractall(destination, filter="data")
            except TypeError:
                for member in tf.getmembers():
                    target = (destination / member.name).resolve()
                    root = destination.resolve()
                    if target != root and root not in target.parents:
                        raise ValueError(f"压缩包路径穿越风险：{member.name}")
                tf.extractall(destination)
    else:
        shutil.unpack_archive(str(archive), str(destination))

    children = [p for p in destination.iterdir() if p.is_dir()]
    if len(children) == 1 and not (destination / ".git").exists():
        return children[0]
    return destination


def find_repo_root(path: Path) -> Path:
    """优先找到含 .git 的仓库；没有 .git 时用唯一一级目录。"""
    try:
        result = run_command(["git", "-C", str(path), "rev-parse", "--show-toplevel"], path, 10)
        if result.returncode == 0:
            return Path(result.stdout.strip()).resolve()
    except Exception:
        pass
    git_dirs = [p.parent for p in path.rglob(".git") if p.is_dir()]
    if len(git_dirs) == 1:
        return git_dirs[0].resolve()
    children = [p for p in path.iterdir() if p.is_dir()]
    if len(children) == 1:
        return children[0].resolve()
    return path.resolve()


def find_patch(repo: Path) -> Path | None:
    direct = repo / "review.patch"
    if direct.is_file():
        return direct
    candidates = sorted(p for p in repo.rglob("review.patch") if p.is_file())
    return candidates[0] if candidates else None


def find_knowledge_files(repo: Path) -> list[Path]:
    """收集 Co-Review 仓库/样本目录中常见的持久化知识文件用于基线快照。
    不会自动改写记忆格式；只快照 JSON 文件名中含 knowledge/memory 的文件。
    """
    roots = [repo / ".co-review", COREVIEW_ROOT]
    found: set[Path] = set()
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("*.json"):
            low = p.name.lower()
            if "knowledge" in low or "memory" in low:
                found.add(p.resolve())
    return sorted(found)


def snapshot_files(paths: list[Path]) -> dict[str, dict[str, Any]]:
    snap = {}
    for p in paths:
        if p.is_file():
            snap[str(p)] = {
                "exists": True,
                "bytes": p.read_bytes(),
                "sha256": sha256_file(p),
            }
        else:
            snap[str(p)] = {"exists": False, "bytes": None, "sha256": None}
    return snap


def restore_files(snapshot: dict[str, dict[str, Any]]) -> None:
    for raw_path, item in snapshot.items():
        p = Path(raw_path)
        if item["exists"]:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_bytes(item["bytes"])
        elif p.exists():
            p.unlink()


def create_task_file(run_dir: Path, prompt: str) -> Path:
    tasks = run_dir / "bench-prompts"
    tasks.mkdir(parents=True, exist_ok=True)
    task_path = tasks / f"task-{uuid.uuid4().hex[:10]}.md"
    task_path.write_text(prompt + "\n", encoding="utf-8")
    return task_path


def prepare_run_dir(repo: Path, repo_name: str, mode: str) -> Path:
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    run_dir = repo / ".co-review" / "runs" / f"{stamp}-{mode}-{uuid.uuid4().hex[:8]}"
    run_dir.mkdir(parents=True, exist_ok=False)
    return run_dir.resolve()


def make_prompt(
    repo: Path, run_dir: Path, skill_path: Path, mode: str,
    commit: str, patch: Path | None, missing_info: str, expected: str,
) -> str:
    patch_text = str(patch.resolve()) if patch else "未找到 review.patch，请明确报告输入缺失，不得虚构审查范围。"
    mode_text = "silent" if mode in ("F", "H") else "ask"
    prompt = f"""你是 iCode 的主 Agent，正在执行受控的 Co-Review 代码安全审查实验。请真实执行工具和审查流程，不要只给出操作建议。

【实验参数】
- 实验模式：{mode}（interaction_mode={mode_text}）
- 仓库：{repo}
- 目标 commit：{commit}
- 审查范围：{patch_text}
- Co-Review Skill：{skill_path}
- 本次实验目录：{run_dir}

【必须遵守】
1. 先完整读取并遵循 Co-Review Skill：{skill_path}。
2. 使用 Co-Review 定义的正式工作流/MCP 工具执行审查。若需要 load_review_task、recall_known_information 等工具，必须实际调用，不要仅阅读 review-materials。
3. 审查目标 commit 的变更。检查 patch 对应的实际文件及必要上下文，结合可验证代码证据判断风险。
4. 不得修改业务代码、提交代码或执行破坏性命令。
5. 最终输出需包含：风险类型、漏洞类别/CWE（如能确定）、受影响文件和行号、数据流/触发条件、可利用性分析、关键证据、反证/已有防护、置信度、仍缺少的信息。
6. 不得预设漏洞存在；证据不足应明确写“证据不足/需进一步核实”，不得强行判定安全或漏洞。
7. 请把本次正式审查报告写入 Co-Review 工作流要求的 report_path / run 目录，并更新对应 execution-state。只在最终回复里写结论而不保存正式报告，不算完成。
8. 如工具/Hook/MCP 报错（例如“The iCode Hook must inject the current host context”），明确报告失败原因，不得伪造工具调用、报告路径或成功状态。

【模式要求】
- F：初始 silent 基线。不得引用本次 G 阶段才产生的记忆。
- G：ask 全自动执行。iCode run 无法等待真人输入，所以把下面 Excel D 列内容当作预先提供的补充线索/回答，继续完成审查；它只是待验证线索，不能直接当成事实或直接写入 confirmed memory。只有经代码证据或可信依据验证后，才可按 Co-Review 正式流程创建/更新知识。
- H：silent。应使用 G 阶段真实形成且经过验证的记忆，先检查 recall 工具的实际返回，并在报告中说明是否命中相关记忆及其如何影响判断。

【Excel D 列：缺失信息/补充线索】
{missing_info or "（未填写）"}

【Excel E 列：预期核查目标】
{expected or "（未填写）"}

【正式报告要求】
完成后，最终回复给出简短 JSON，至少包含 report_path、risk_type、finding_count、key_evidence、memory_recall_status、tools_used、completion_status。JSON 必须描述真实执行结果，不得臆造。
"""
    return prompt


def build_icode_command(exe: str, prompt: str, repo: Path, task_file: Path) -> list[str]:
    # 用 --task 避免长 prompt 被 shell 引号/命令行长度破坏。
    cmd = [exe, "run", "--task", str(task_file), "--agent", ICODE_AGENT,
           "--workdir", str(repo), "--json"]
    if ICODE_MODEL:
        cmd.extend(["--model", ICODE_MODEL])
    return cmd


def parse_json_output(stdout: str) -> tuple[dict[str, Any] | None, str | None]:
    text = stdout.strip()
    if not text:
        return None, "stdout 为空"
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            return obj, None
        return None, "JSON stdout 顶层不是对象"
    except json.JSONDecodeError as e:
        return None, f"stdout 不是单个 JSON 对象：{e}"


def run_icode(exe: str, repo: Path, run_dir: Path, prompt: str, timeout: int) -> dict[str, Any]:
    task_file = create_task_file(run_dir, prompt)
    cmd = build_icode_command(exe, prompt, repo, task_file)
    log_dir = WORK_BASE / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    prefix = log_dir / f"{run_dir.name}"
    started = time.time()
    result: dict[str, Any] = {
        "command": cmd[:-1] + ["--json"],
        "task_file": str(task_file),
        "stdout_log": str(prefix.with_suffix(".stdout.json")),
        "stderr_log": str(prefix.with_suffix(".stderr.jsonl")),
    }
    try:
        proc = subprocess.run(
            cmd, cwd=str(repo), capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout, check=False,
        )
        result.update({
            "returncode": proc.returncode,
            "elapsed_seconds": round(time.time() - started, 2),
        })
        Path(result["stdout_log"]).write_text(proc.stdout, encoding="utf-8")
        Path(result["stderr_log"]).write_text(proc.stderr, encoding="utf-8")
        parsed, parse_error = parse_json_output(proc.stdout)
        result["json"] = parsed
        result["parse_error"] = parse_error
        result["stderr_excerpt"] = proc.stderr[-5000:]
        if proc.returncode != 0:
            result["error"] = f"icode run 退出码 {proc.returncode}: {proc.stderr[-1500:]}"
        elif parsed is None:
            result["error"] = parse_error
        elif parsed.get("error"):
            result["error"] = f"{parsed.get('code', 'icode_error')}: {parsed['error']}"
        elif not isinstance(parsed.get("result"), str):
            result["error"] = "JSON 输出缺少字符串 result"
        else:
            result["final_reply"] = parsed["result"]
            result["session_id"] = parsed.get("session_id")
            result["duration"] = parsed.get("duration")
    except subprocess.TimeoutExpired as e:
        result["returncode"] = 124
        result["elapsed_seconds"] = round(time.time() - started, 2)
        result["error"] = f"iCode run 超时（{timeout}s）"
        Path(result["stdout_log"]).write_text(
            (e.stdout.decode(errors="replace") if isinstance(e.stdout, bytes) else e.stdout or ""),
            encoding="utf-8",
        )
        Path(result["stderr_log"]).write_text(
            (e.stderr.decode(errors="replace") if isinstance(e.stderr, bytes) else e.stderr or ""),
            encoding="utf-8",
        )
    except Exception as e:
        result["error"] = f"{type(e).__name__}: {e}"
    return result


def extract_json_object(text: str) -> dict[str, Any] | None:
    """从 Agent 最终回复中提取 JSON 对象；只解析，不把文本本身当报告。"""
    try:
        obj = json.loads(text.strip())
        return obj if isinstance(obj, dict) else None
    except Exception:
        pass
    # 允许模型在 JSON 前后加 Markdown，但只接受完整可解析对象。
    decoder = json.JSONDecoder()
    for i, ch in enumerate(text):
        if ch != "{":
            continue
        try:
            obj, _ = decoder.raw_decode(text[i:])
            if isinstance(obj, dict):
                return obj
        except Exception:
            continue
    return None


def locate_formal_report(repo: Path, run_dir: Path, final_reply: str) -> Path | None:
    # 先使用 Agent JSON 中明确声明的报告路径。
    obj = extract_json_object(final_reply)
    if obj:
        raw = obj.get("report_path")
        if isinstance(raw, str) and raw.strip():
            p = Path(raw)
            if not p.is_absolute():
                p = (repo / p).resolve()
            if p.is_file():
                return p
    # 再仅在本次 run_dir 内查找标准报告文件，避免捡到旧 run 的报告。
    preferred_names = ("report.json", "review-report.json", "final-report.json", "report.md", "review-report.md")
    for name in preferred_names:
        p = run_dir / name
        if p.is_file() and p.stat().st_size:
            return p
    # execution-state 里可能声明 report_path。
    state = load_json(run_dir / "execution-state.json", {})
    if isinstance(state, dict):
        raw = state.get("report_path")
        if isinstance(raw, str) and raw.strip():
            p = Path(raw)
            if not p.is_absolute():
                p = (run_dir / p).resolve()
            if p.is_file() and p.stat().st_size:
                return p
    return None


def validate_formal_report(path: Path | None, final_reply: str) -> tuple[bool, str | None, str | None]:
    """优先验证报告文件；Markdown 报告不自动推断 risk_type。"""
    if path is None:
        return False, None, "未找到本次 run 的正式报告文件"
    if not path.is_file() or path.stat().st_size == 0:
        return False, None, "报告文件不存在或为空"
    if path.suffix.lower() == ".json":
        obj = load_json(path)
        if not isinstance(obj, dict):
            return False, None, "报告 JSON 无效"
        risk = obj.get("risk_type")
        findings = obj.get("findings")
        if isinstance(findings, list) and findings and isinstance(findings[0], dict):
            risk = findings[0].get("risk_type", risk)
        if risk not in VALID_RISK_TYPES:
            # 某些 Co-Review 版本使用不同 schema；要求 Agent JSON 有明确风险标签。
            reply_obj = extract_json_object(final_reply)
            if reply_obj and reply_obj.get("risk_type") in VALID_RISK_TYPES:
                risk = reply_obj["risk_type"]
            else:
                return False, None, f"报告中 risk_type 缺失或不支持：{risk!r}"
        return True, risk, None
    # Markdown 文件必须有 Agent 结构化回复提供的 risk_type，避免猜测标签。
    reply_obj = extract_json_object(final_reply)
    risk = reply_obj.get("risk_type") if reply_obj else None
    if risk not in VALID_RISK_TYPES:
        return False, None, "Markdown 报告缺少可验证的结构化 risk_type"
    return True, risk, None


def execute_mode(
    exe: str, repo: Path, repo_name: str, mode: str, commit: str,
    patch: Path | None, skill_path: Path, missing_info: str,
    expected: str, timeout: int,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "mode": mode, "status": "FAILED", "started_at": utc_now(),
        "risk_type": None, "report_path": None, "error": None,
    }
    try:
        run_dir = prepare_run_dir(repo, repo_name, mode)
        result["run_dir"] = str(run_dir)
        result["review_patch_sha256"] = sha256_file(patch) if patch else None
        prompt = make_prompt(repo, run_dir, skill_path, mode, commit, patch, missing_info, expected)
        cli = run_icode(exe, repo, run_dir, prompt, timeout)
        result["cli"] = cli
        result["session_id"] = cli.get("session_id")
        if cli.get("error"):
            result["status"] = "CLI_FAILED"
            result["error"] = cli["error"]
            result["final_reply"] = cli.get("final_reply")
            return result

        final_reply = cli.get("final_reply", "")
        result["final_reply"] = final_reply
        reply_obj = extract_json_object(final_reply)
        report_path = locate_formal_report(repo, run_dir, final_reply)
        result["report_path"] = str(report_path) if report_path else None
        ok, risk, reason = validate_formal_report(report_path, final_reply)
        result["agent_report_json"] = reply_obj
        if not ok:
            result["status"] = "INVALID_REPORT"
            result["error"] = reason
        else:
            result["status"] = "SUCCESS"
            result["risk_type"] = risk
            result["report_sha256"] = sha256_file(report_path)
            result["memory_recall_status"] = (
                reply_obj.get("memory_recall_status") if reply_obj else None
            )
            result["tools_used"] = reply_obj.get("tools_used") if reply_obj else None
            result["finding_count"] = reply_obj.get("finding_count") if reply_obj else None
    except Exception as e:
        result["status"] = "EXECUTION_FAILED"
        result["error"] = f"{type(e).__name__}: {e}"
    finally:
        result["finished_at"] = utc_now()
    return result


def excel_result(result: dict[str, Any]) -> str:
    if result.get("status") == "SUCCESS":
        return str(result.get("risk_type"))
    error = result.get("error") or "未生成有效正式报告"
    return f"{result.get('status', 'FAILED')}: {error}"[:32000]


def main() -> int:
    parser = argparse.ArgumentParser(description="iCode CLI JSON + Co-Review F/G/H 批量实验")
    parser.add_argument("--check", action="store_true", help="检查 iCode CLI、Co-Review、Excel 和路径")
    parser.add_argument("--rows", default="", help="只运行指定 Excel 行，如 2 或 2,3")
    parser.add_argument("--timeout", type=int, default=ICODE_TIMEOUT, help="每轮 iCode run 超时秒数")
    parser.add_argument("--reset-results", action="store_true", help="清空所选样本 F/G/H 旧结果")
    parser.add_argument("--keep-extracted", action="store_true", help="保留已有解压目录")
    args = parser.parse_args()

    exe = locate_executable(ICODE_EXE)
    errors = check_environment()
    if errors:
        print("环境检查失败：", file=sys.stderr)
        for error in errors:
            print(f" - {error}", file=sys.stderr)
        return 2
    assert exe is not None
    if args.check:
        print("基础检查通过。")
        print(f"iCode CLI: {exe}")
        print(f"Agent: {ICODE_AGENT}")
        print(f"Co-Review Skill: {COREVIEW_SKILL}")
        print("注意：基础检查不代表 Co-Review MCP Hook、Session 上下文及正式报告已通过端到端验证。")
        return 0

    if args.timeout <= 0:
        print("--timeout 必须大于 0", file=sys.stderr)
        return 2
    WORK_BASE.mkdir(parents=True, exist_ok=True)
    log_path = WORK_BASE / "bench_runs.jsonl"

    wb = openpyxl.load_workbook(BENCH_XLSX)
    if SHEET_NAME not in wb.sheetnames:
        raise SystemExit(f"工作表 {SHEET_NAME!r} 不存在；当前工作表：{wb.sheetnames}")
    ws = wb[SHEET_NAME]
    selected_rows = {int(x.strip()) for x in args.rows.split(",") if x.strip()} if args.rows else None
    original_row_numbers = {row for row in PACKAGE_MAP.values()}
    if selected_rows is not None:
        unknown = selected_rows - original_row_numbers
        if unknown:
            print(f"警告：这些行不在 PACKAGE_MAP 中，无法匹配压缩包：{sorted(unknown)}")

    total = 0
    for archive_name, row in PACKAGE_MAP.items():
        if selected_rows is not None and row not in selected_rows:
            continue
        archive = PKG_DIR / archive_name
        if not archive.is_file():
            print(f"[SKIP] 压缩包不存在：{archive}", file=sys.stderr)
            continue
        commit = str(ws.cell(row, COL_COMMIT).value or "").strip()
        missing_info = str(ws.cell(row, COL_MISSING_INFO).value or "").strip()
        expected = str(ws.cell(row, COL_EXPECTED).value or "").strip()
        repo_label = str(ws.cell(row, COL_NAME).value or archive_name)
        if not commit:
            print(f"[SKIP] Excel 第 {row} 行未填写 commit", file=sys.stderr)
            continue

        repo_name = Path(archive_name).name
        for suffix in (".tar.gz", ".tgz", ".tar", ".zip"):
            if repo_name.lower().endswith(suffix):
                repo_name = repo_name[:-len(suffix)]
                break
        dest = WORK_BASE / repo_name
        if dest.exists() and args.keep_extracted:
            repo = find_repo_root(dest)
        else:
            repo = extract_archive(archive, dest)
            repo = find_repo_root(repo)

        patch = find_patch(repo)
        if patch is None:
            print(f"[WARN] {repo_label}: 没找到 review.patch；Agent 会被要求报告输入缺失。")
        print(f"\n===== {repo_label} / Excel 行 {row} =====")
        print(f"repo={repo}\ncommit={commit}\npatch={patch or '(missing)'}")

        # 每个样本的知识文件集合和基线。只快照识别出的 knowledge/memory JSON。
        knowledge_files = find_knowledge_files(repo)
        baseline = snapshot_files(knowledge_files)
        if args.reset_results:
            for col in (COL_F, COL_G, COL_H):
                ws.cell(row, col, None)

        try:
            # F: 初始 silent；从基线开始。
            restore_files(baseline)
            f = execute_mode(exe, repo, repo_name, "F", commit, patch,
                             COREVIEW_SKILL, missing_info, expected, args.timeout)
            ws.cell(row, COL_F, excel_result(f))
            print(f"F: {f['status']} risk_type={f.get('risk_type')} error={f.get('error')}")

            # G: ask；仍从初始基线开始。
            restore_files(baseline)
            g = execute_mode(exe, repo, repo_name, "G", commit, patch,
                             COREVIEW_SKILL, missing_info, expected, args.timeout)
            ws.cell(row, COL_G, excel_result(g))
            # 记忆文件清单不变时，可捕捉已存在 JSON 的更新。新建文件需再次扫描。
            after_g_paths = sorted(set(knowledge_files + find_knowledge_files(repo)))
            after_g = snapshot_files(after_g_paths)
            print(f"G: {g['status']} risk_type={g.get('risk_type')} error={g.get('error')}")

            # H: 从 G 结束后的真实文件状态开始。
            restore_files(after_g)
            h = execute_mode(exe, repo, repo_name, "H", commit, patch,
                             COREVIEW_SKILL, missing_info, expected, args.timeout)
            ws.cell(row, COL_H, excel_result(h))
            print(f"H: {h['status']} risk_type={h.get('risk_type')} "
                  f"memory_recall={h.get('memory_recall_status')} error={h.get('error')}")

            record = {
                "timestamp": utc_now(),
                "archive": archive_name,
                "excel_row": row,
                "repo_label": repo_label,
                "repo_path": str(repo),
                "commit": commit,
                "missing_info": missing_info,
                "expected": expected,
                "baseline_memory_hashes": {
                    p: v.get("sha256") for p, v in baseline.items()
                },
                "after_g_memory_hashes": {
                    p: v.get("sha256") for p, v in after_g.items()
                },
                "f": f, "g": g, "h": h,
                "warning": (
                    "SUCCESS 只表示存在可验证的报告文件且风险标签可解析；"
                    "仍需抽查 report 内容与实际 Co-Review 工具调用。"
                ),
            }
            append_jsonl(log_path, record)
            wb.save(BENCH_XLSX)
            total += 1
        finally:
            # 样本结束恢复基线，避免不同样本之间互相污染。
            restore_files(baseline)
            wb.save(BENCH_XLSX)

    print(f"\n完成样本数：{total}")
    print(f"Excel：{BENCH_XLSX}")
    print(f"实验日志：{log_path}")
    print(f"iCode CLI 日志：{WORK_BASE / 'logs'}")
    print("如果结果为 CLI_FAILED / INVALID_REPORT，请查看 stderr 日志和 Agent 最终回复。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
