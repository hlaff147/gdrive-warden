#!/usr/bin/env python3
"""
Idempotent local Write-Ahead Log (WAL) inspector & rollback generator for gdrive-warden.
Reads a JSONL WAL log (reports/wal_log.jsonl) and prints or executes reverse `gdrive mutate mv`
commands in LIFO order to restore items to their exact original parent folder IDs.
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Inspect or rollback gdrive-warden WAL moves.")
    parser.add_argument("--wal", default="reports/wal_log.jsonl", help="Path to WAL JSONL file")
    parser.add_argument("--wave", default="ALL", help="Filter by wave name (or ALL)")
    parser.add_argument("--execute", action="store_true", help="Execute `gdrive mutate mv` (default is dry-run)")
    args = parser.parse_args()

    wal_path = Path(args.wal)
    if not wal_path.exists():
        print(f"WAL file not found: {wal_path}", file=sys.stderr)
        sys.exit(1)

    entries = []
    with wal_path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                entries.append(json.loads(line))

    # Build latest status by log_id
    by_id = {}
    order = []
    for e in entries:
        lid = e.get("log_id")
        if lid not in by_id:
            order.append(lid)
        by_id[lid] = e

    completed = [
        by_id[lid] for lid in reversed(order)
        if by_id[lid].get("status") == "CONCLUIDO"
        and (args.wave == "ALL" or by_id[lid].get("wave") == args.wave)
    ]

    print(f"Found {len(completed)} CONCLUIDO operations matching wave={args.wave} (mode={'EXECUTE' if args.execute else 'DRY_RUN'}).")
    reverted = 0
    errors = 0
    for op in completed:
        fid = op["file_id"]
        name = op["name"]
        orig_parent = op["old_parent_id"]
        dest_parent = op["new_parent_id"]
        if not args.execute:
            print(f"[DRY_RUN] gdrive mutate mv {fid} {orig_parent}  # {name} (from {dest_parent})")
            continue

        cp = subprocess.run(["gdrive", "mutate", "mv", fid, orig_parent, "--json"], capture_output=True, text=True)
        if cp.returncode == 0:
            reverted += 1
            print(f"[REVERTIDO] {fid} ({name}) -> {orig_parent}")
        else:
            errors += 1
            print(f"[FALHA_ROLLBACK] {fid} ({name}): {cp.stderr.strip()}", file=sys.stderr)
            if errors >= 5:
                print("[CIRCUIT BREAKER] Aborting after 5 rollback errors.", file=sys.stderr)
                sys.exit(2)

    if args.execute:
        print(f"Rollback finished: {reverted} reverted, {errors} errors.")


if __name__ == "__main__":
    main()
