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

        if args.reset_results:
            for col in (COL_F, COL_G, COL_H):
                ws.cell(row, col, None)

        # 注意：这里不快照/恢复 Co-Review 知识库。记忆的创建、更新与持久化
        # 完全由 Co-Review 自身工作流处理，F→G→H 按顺序在同一环境中执行。
        f = execute_mode(
            exe, repo, repo_name, "F", commit, patch, COREVIEW_SKILL,
            missing_info, expected, args.timeout, repo_label
        )
        ws.cell(row, COL_F, excel_result(f))
        print(f"F: {f['status']} risk_type={f.get('risk_type')} error={f.get('error')}")

        g = execute_mode(
            exe, repo, repo_name, "G", commit, patch, COREVIEW_SKILL,
            missing_info, expected, args.timeout, repo_label
        )
        ws.cell(row, COL_G, excel_result(g))
        print(f"G: {g['status']} risk_type={g.get('risk_type')} error={g.get('error')}")

        h = execute_mode(
            exe, repo, repo_name, "H", commit, patch, COREVIEW_SKILL,
            missing_info, expected, args.timeout, repo_label
        )
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
            "f": f, "g": g, "h": h,
            "memory_management": (
                "由 Co-Review 自身工作流管理；本脚本不快照、修改或恢复知识库。"
            ),
            "warning": (
                "SUCCESS 只表示找到可校验的报告文件且风险标签可解析；"
                "仍需抽查报告内容及 Co-Review 实际工具调用。"
            ),
        }
        append_jsonl(log_path, record)
        wb.save(BENCH_XLSX)
        total += 1

    print(f"\n完成样本数：{total}")
    print(f"Excel：{BENCH_XLSX}")
    print(f"实验日志：{log_path}")
    print(f"iCode CLI 日志：{WORK_BASE / 'logs'}")
    print("知识库写入/更新由 Co-Review 自身管理；本脚本不对其记忆文件做快照或恢复。")
    print("如果结果为 CLI_FAILED / INVALID_REPORT，请查看 stderr 日志和 Agent 最终回复。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
