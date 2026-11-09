#!/usr/bin/env python3
"""核验生成批次中的报告并回填 Excel F 列。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import openpyxl


def read_json(path):
  return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def read_memory_recall(run_dir):
  path = run_dir / "memory-recall.jsonl"
  if not path.exists():
    return None
  lines = [line for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
  return json.loads(lines[-1]) if lines else None


def read_report_risk_type(run_dir):
  state_path = run_dir / "execution-state.json"
  if not state_path.exists():
    return None, "无 execution-state.json", None
  state = read_json(state_path)
  report_path = state.get("report_path")
  if not report_path:
    return None, "report_path 为空", None
  path = Path(report_path)
  if not path.exists():
    return None, f"报告文件不存在：{path}", None
  try:
    data = read_json(path)
    required = {"repository", "findings", "remaining_questions", "confirmation_history"}
    if not isinstance(data, dict) or not required <= data.keys():
      return None, "报告不符合 Co-Review 顶层结构", str(path)
    if data["repository"] == "" or not isinstance(data["findings"], list):
      return None, "报告 repository/findings 无效", str(path)
    if not isinstance(data["remaining_questions"], list) or not isinstance(data["confirmation_history"], list):
      return None, "报告问题或确认历史字段无效", str(path)
    if data["findings"]:
      risk_type = data["findings"][0].get("risk_type")
      if risk_type not in {"confirmed_vulnerability", "potential_risk", "not_vulnerable"}:
        return None, f"未知 risk_type：{risk_type}", str(path)
      return risk_type, None, str(path)
    return ("potential_risk" if data["remaining_questions"] else "not_vulnerable"), None, str(path)
  except Exception as exc:
    return None, f"报告解析失败：{exc}", str(path)


def classify_risk(risk_type):
  if risk_type == "confirmed_vulnerability":
    return "是(risk_type=confirmed_vulnerability)"
  if risk_type == "not_vulnerable":
    return "是(risk_type=not_vulnerable)"
  if risk_type in ("potential", "potential_risk"):
    return f"否(risk_type={risk_type})"
  return "报告缺失/异常"


def validate_result(result):
  if result["error"]:
    return
  try:
    run_dir = Path(result["run_dir"])
    risk, report_error, report_path = read_report_risk_type(run_dir)
    result["report_path"] = report_path
    result["risk_type"] = risk
    result["val"] = classify_risk(risk) if risk else f"报告无效({report_error})"
    result["error"] = report_error
    result["recall"] = read_memory_recall(run_dir)
  except Exception as exc:
    result["error"] = str(exc)
    result["val"] = f"报告无效({exc})"


def verify_batch(manifest_path):
  manifest = read_json(manifest_path)
  workbook_path = Path(manifest["workbook"])
  workbook = openpyxl.load_workbook(workbook_path)
  sheet = workbook[manifest["sheet"]]
  records = manifest["records"]
  rows = [record["excel_row"] for record in records]
  if len(rows) != len(set(rows)):
    raise ValueError("批次清单包含重复的 Excel 行号")
  for record in records:
    commit = sheet.cell(row=record["excel_row"], column=2).value
    if (str(commit).strip() if commit is not None else "") != record["commit"]:
      raise ValueError(f"Excel 第 {record['excel_row']} 行 commit 已变化，停止回填")

  for record in records:
    validate_result(record["f"])
    sheet.cell(row=record["excel_row"], column=6, value=record["f"]["val"])
    print(f"  {record['zip']}: F={record['f']['val']}")
  workbook.save(workbook_path)
  log_path = Path(manifest["work_base"]) / "bench_runs.jsonl"
  with log_path.open("a", encoding="utf-8") as stream:
    for record in records:
      stream.write(json.dumps(record, ensure_ascii=False) + "\n")
  print(f"完成，共核验 {len(records)} 个样本。F 列已填入 {workbook_path}，G/H 列未修改。")
  print(f"详细旁证：{log_path}")


def main():
  parser = argparse.ArgumentParser(description=__doc__)
  parser.add_argument("manifest", type=Path, help="报告生成阶段输出的批次清单")
  args = parser.parse_args()
  verify_batch(args.manifest)


if __name__ == "__main__":
  main()
