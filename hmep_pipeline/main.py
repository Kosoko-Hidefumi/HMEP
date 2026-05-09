"""
HMEP パイプライン統合エントリ（F6）
① extract → ② rename → ③ upload を個別実行または `all` で一括実行する。

各ステップは従来どおり単体スクリプトでも実行可能。フェーズ境界と終了コードは
`logs/pipeline_YYYY-MM-DD.log` に記録する（要件定義 5）。
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import yaml

PIPELINE_ROOT = Path(__file__).resolve().parent

for _stage in ("01_extract", "02_rename", "03_upload"):
    _p = str(PIPELINE_ROOT / _stage)
    if _p not in sys.path:
        sys.path.insert(0, _p)

import extract_lectures  # noqa: E402
import rename_videos  # noqa: E402
import upload_youtube  # noqa: E402


def load_config_dict(config_path: Path) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def pipeline_log_dir(config_path: Path) -> Path:
    cfg = load_config_dict(config_path)
    paths = cfg.get("paths") or {}
    rel = paths.get("logs_dir") or "logs"
    p = Path(rel)
    if p.is_absolute():
        return p.resolve()
    return (config_path.parent / p).resolve()


def setup_pipeline_logging(log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"pipeline_{datetime.now().strftime('%Y-%m-%d')}.log"
    log = logging.getLogger("hmep.pipeline")
    log.setLevel(logging.INFO)
    if log.handlers:
        return log
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    log.propagate = False
    return log


def main(argv: Optional[list[str]] = None) -> int:
    argv = sys.argv[1:] if argv is None else argv

    root_parser = argparse.ArgumentParser(
        description="HMEP pipeline: extract (F2) / rename (F4) / upload (F5) / all",
    )
    root_parser.add_argument(
        "--config",
        type=Path,
        default=PIPELINE_ROOT / "config.yaml",
        help="config.yaml path (default: hmep_pipeline/config.yaml)",
    )
    sub = root_parser.add_subparsers(dest="command", required=True)

    p_ex = sub.add_parser("extract", help="Outlook -> lectures.xlsx")
    p_ex.add_argument("--dry-run", action="store_true")
    p_ex.add_argument("--fixture", type=Path, default=None, metavar="PATH")

    p_rn = sub.add_parser("rename", help="copy + rename videos from lectures.xlsx")
    p_rn.add_argument("--dry-run", action="store_true")
    p_rn.add_argument("--force", action="store_true")
    p_rn.add_argument("--limit", type=int, default=None, metavar="N")
    p_rn.add_argument("--time-window-hours", type=float, default=None, metavar="H")
    p_rn.add_argument("--similarity-threshold", type=float, default=None, metavar="T")
    p_rn.add_argument("--similarity-fallback", type=float, default=None, metavar="T")

    p_up = sub.add_parser("upload", help="YouTube upload from videos_renamed/")
    p_up.add_argument("--dry-run", action="store_true")
    p_up.add_argument("--force", action="store_true")
    p_up.add_argument("--limit", type=int, default=None, metavar="N")
    p_up.add_argument(
        "--auth-only",
        action="store_true",
        help="OAuth only -> token.json (same as upload_youtube.py --auth-only)",
    )

    p_all = sub.add_parser("all", help="run extract, then rename, then upload")
    p_all.add_argument("--dry-run", action="store_true")
    p_all.add_argument("--force", action="store_true", help="pass to rename and upload")
    p_all.add_argument("--rename-limit", type=int, default=None, metavar="N")
    p_all.add_argument("--upload-limit", type=int, default=None, metavar="N")

    args = root_parser.parse_args(argv)
    config_path = args.config.resolve()
    if not config_path.is_file():
        print(f"config not found: {config_path}", file=sys.stderr)
        return 2

    plog = setup_pipeline_logging(pipeline_log_dir(config_path))

    if args.command == "extract":
        ex_args: list[str] = ["--config", str(config_path)]
        if args.dry_run:
            ex_args.append("--dry-run")
        if args.fixture is not None:
            ex_args.extend(["--fixture", str(args.fixture.resolve())])
        plog.info("command=extract argv=%s", ex_args)
        rc = extract_lectures.main(ex_args)
        plog.info("command=extract finished rc=%s", rc)
        return rc

    if args.command == "rename":
        rn_args: list[str] = ["--config", str(config_path)]
        if args.dry_run:
            rn_args.append("--dry-run")
        if args.force:
            rn_args.append("--force")
        if args.limit is not None:
            rn_args.extend(["--limit", str(args.limit)])
        if args.time_window_hours is not None:
            rn_args.extend(["--time-window-hours", str(args.time_window_hours)])
        if args.similarity_threshold is not None:
            rn_args.extend(["--similarity-threshold", str(args.similarity_threshold)])
        if args.similarity_fallback is not None:
            rn_args.extend(["--similarity-fallback", str(args.similarity_fallback)])
        plog.info("command=rename argv=%s", rn_args)
        rc = rename_videos.main(rn_args)
        plog.info("command=rename finished rc=%s", rc)
        return rc

    if args.command == "upload":
        up_args: list[str] = ["--config", str(config_path)]
        if args.auth_only:
            up_args.append("--auth-only")
        if args.dry_run:
            up_args.append("--dry-run")
        if args.force:
            up_args.append("--force")
        if args.limit is not None:
            up_args.extend(["--limit", str(args.limit)])
        plog.info("command=upload argv=%s", up_args)
        rc = upload_youtube.main(up_args)
        plog.info("command=upload finished rc=%s", rc)
        return rc

    if args.command == "all":
        plog.info("command=all dry_run=%s force=%s", args.dry_run, args.force)

        ex_all = ["--config", str(config_path)]
        if args.dry_run:
            ex_all.append("--dry-run")
        plog.info("phase=extract start argv=%s", ex_all)
        rc = extract_lectures.main(ex_all)
        plog.info("phase=extract end rc=%s", rc)
        if rc != 0:
            plog.warning("pipeline stopped after extract")
            return rc

        rn_all = ["--config", str(config_path)]
        if args.dry_run:
            rn_all.append("--dry-run")
        if args.force:
            rn_all.append("--force")
        if args.rename_limit is not None:
            rn_all.extend(["--limit", str(args.rename_limit)])
        plog.info("phase=rename start argv=%s", rn_all)
        rc = rename_videos.main(rn_all)
        plog.info("phase=rename end rc=%s", rc)
        if rc != 0:
            plog.warning("pipeline stopped after rename")
            return rc

        up_all = ["--config", str(config_path)]
        if args.dry_run:
            up_all.append("--dry-run")
        if args.force:
            up_all.append("--force")
        if args.upload_limit is not None:
            up_all.extend(["--limit", str(args.upload_limit)])
        plog.info("phase=upload start argv=%s", up_all)
        rc = upload_youtube.main(up_all)
        plog.info("phase=upload end rc=%s", rc)
        if rc != 0:
            plog.warning("pipeline finished with upload errors (see upload log)")
        return rc

    print(f"internal error: unknown command {args.command!r}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
