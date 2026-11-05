"""批量为压缩包中的仓库生成 review.patch。

适配 co-review bench 脚本：
- 读取 bench 脚本中的 PACKAGE_MAP（压缩包名 -> Excel 行号）
- 从 Excel 指定行读取 commit
- 解压仓库，执行 git show 生成 review.patch
- 将 review.patch 写回压缩包内仓库根目录

默认读取：
- bench 脚本: /data/cyjia/co-review_bench/bench_fill_full_exe.py
- Excel: CO_REVIEW_BENCH_XLSX 或 /data/cyjia/iCode/bench_info.xlsx
- 包目录: CO_REVIEW_PKG_DIR 或 /tmp
"""

from __future__ import annotations

import argparse
import ast
import shutil
import subprocess
import tarfile
import tempfile
import zipfile
from pathlib import Path

import openpyxl

DEFAULT_BENCH_SCRIPT = Path("/data/cyjia/co-review_bench/bench_fill_full_exe.py")
DEFAULT_BENCH_XLSX = Path("/data/cyjia/iCode/bench_info.xlsx")
DEFAULT_PKG_DIR = Path("/tmp")


def _literal_assignments(py_file: Path) -> dict[str, object]:
    tree = ast.parse(py_file.read_text(encoding="utf-8"), filename=str(py_file))
    out: dict[str, object] = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
            try:
                out[name] = ast.literal_eval(node.value)
            except Exception:
                continue
    return out


def load_bench_map(bench_script: Path) -> tuple[dict[str, int], str, int]:
    values = _literal_assignments(bench_script)
    pkg_map = values.get("PACKAGE_MAP")
    if not isinstance(pkg_map, dict) or not pkg_map:
        raise RuntimeError(f"无法从 {bench_script} 读取 PACKAGE_MAP")

    normalized: dict[str, int] = {}
    for key, value in pkg_map.items():
        if not isinstance(key, str) or not isinstance(value, int):
            raise RuntimeError("PACKAGE_MAP 必须是 {str: int} 结构")
        normalized[key] = value

    sheet_name = values.get("SHEET_NAME", "Sheet1")
    if not isinstance(sheet_name, str):
        sheet_name = "Sheet1"

    col_commit = values.get("COL_COMMIT", 2)
    if not isinstance(col_commit, int):
        col_commit = 2

    return normalized, sheet_name, col_commit


def extract_archive(archive_path: Path, dest: Path) -> None:
    name = archive_path.name.lower()
    if name.endswith(".zip"):
        with zipfile.ZipFile(archive_path, "r") as zf:
            zf.extractall(dest)
        return
    if name.endswith((".tar.gz", ".tgz", ".tar")):
        with tarfile.open(archive_path, "r:*") as tf:
            tf.extractall(dest)
        return
    raise RuntimeError(f"不支持的压缩格式: {archive_path}")


def repack_archive(src_dir: Path, archive_path: Path) -> None:
    name = archive_path.name.lower()
    if name.endswith(".zip"):
        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in sorted(src_dir.rglob("*")):
                if p.is_file():
                    zf.write(p, p.relative_to(src_dir))
        return

    if name.endswith(".tar.gz") or name.endswith(".tgz"):
        mode = "w:gz"
    elif name.endswith(".tar"):
        mode = "w"
    else:
        raise RuntimeError(f"不支持的压缩格式: {archive_path}")

    with tarfile.open(archive_path, mode) as tf:
        for p in sorted(src_dir.rglob("*")):
            tf.add(p, arcname=p.relative_to(src_dir), recursive=False)


def find_repo_root(extract_root: Path) -> Path:
    # 与 bench 脚本保持一致：若仅有一个子目录且根目录无 .git，则下钻一层。
    subs = [d for d in extract_root.iterdir() if d.is_dir()]
    if len(subs) == 1 and not (extract_root / ".git").exists():
        return subs[0]
    return extract_root


def generate_patch(repo_path: Path, commit: str, out_file: Path) -> None:
    git_dir = repo_path / ".git"
    if not git_dir.exists():
        raise RuntimeError(f"仓库无 .git，无法生成 patch: {repo_path}")

    cmd = ["git", "-C", str(repo_path), "show", commit]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"git show 失败(commit={commit}): {err}")
    out_file.write_text(proc.stdout, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="批量生成 review.patch 并写回压缩包")
    parser.add_argument("--bench-script", type=Path, default=DEFAULT_BENCH_SCRIPT)
    parser.add_argument("--bench-xlsx", type=Path, default=Path(__import__("os").environ.get("CO_REVIEW_BENCH_XLSX", str(DEFAULT_BENCH_XLSX))))
    parser.add_argument("--pkg-dir", type=Path, default=Path(__import__("os").environ.get("CO_REVIEW_PKG_DIR", str(DEFAULT_PKG_DIR))))
    parser.add_argument("--backup", action="store_true", help="覆盖前生成 .bak 备份")
    args = parser.parse_args()

    pkg_map, sheet_name, col_commit = load_bench_map(args.bench_script)

    wb = openpyxl.load_workbook(args.bench_xlsx)
    if sheet_name not in wb.sheetnames:
        raise RuntimeError(f"Excel 不存在工作表 {sheet_name}，当前: {wb.sheetnames}")
    ws = wb[sheet_name]

    ok = 0
    fail = 0

    for archive_name, row in pkg_map.items():
        archive_path = args.pkg_dir / archive_name
        print(f"\n==> {archive_name} (row={row})")

        if not archive_path.is_file():
            print(f"[失败] 压缩包不存在: {archive_path}")
            fail += 1
            continue

        commit_cell = ws.cell(row=row, column=col_commit).value
        commit = str(commit_cell).strip() if commit_cell is not None else ""
        if not commit:
            print("[失败] Excel commit 为空")
            fail += 1
            continue

        with tempfile.TemporaryDirectory(prefix="review-patch-gen-") as tmp:
            tmp_dir = Path(tmp)
            extract_root = tmp_dir / "extract"
            extract_root.mkdir(parents=True, exist_ok=True)

            try:
                extract_archive(archive_path, extract_root)
                repo_root = find_repo_root(extract_root)
                patch_file = repo_root / "review.patch"
                generate_patch(repo_root, commit, patch_file)

                if args.backup:
                    backup_path = archive_path.with_suffix(archive_path.suffix + ".bak")
                    shutil.copy2(archive_path, backup_path)
                    print(f"[备份] {backup_path}")

                repack_archive(extract_root, archive_path)
                print(f"[成功] 已写入 review.patch -> {archive_path}")
                ok += 1
            except Exception as exc:
                print(f"[失败] {exc}")
                fail += 1

    print(f"\n完成: 成功 {ok}，失败 {fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
