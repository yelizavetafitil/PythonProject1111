#!/usr/bin/env python3
"""Prepare a fresh portal clone from an existing production checkout.

Secrets are moved to the ignored .env.production file and are never printed.
The SQLite backup API is used so the source portal may remain online while its
database is copied.
"""

from __future__ import annotations

import argparse
import ast
import os
from pathlib import Path
import secrets
import shutil
import sqlite3
import subprocess
import time


PROJECT_DIR = Path(__file__).resolve().parent.parent


def read_env(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip()
    return values


def legacy_code_settings(path: Path) -> dict[str, str]:
    """Read only known legacy defaults; never execute the old application."""
    if not path.is_file():
        return {}
    tree = ast.parse(path.read_text(encoding="utf-8-sig"))
    found: dict[str, str] = {}
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        names = [target.id for target in node.targets if isinstance(target, ast.Name)]
        if "LDAP_CONFIG" in names and isinstance(node.value, ast.Dict):
            try:
                ldap = ast.literal_eval(node.value)
            except (ValueError, TypeError):
                ldap = {}
            mapping = {
                "uri": "LDAP_URI",
                "base": "LDAP_BASE_DN",
                "bind_dn": "LDAP_BIND_DN",
                "bind_password": "LDAP_BIND_PASSWORD",
                "user_attr": "LDAP_USER_ATTRIBUTE",
            }
            for old_key, new_key in mapping.items():
                value = ldap.get(old_key)
                if isinstance(value, str) and value:
                    found[new_key] = value
        if "SMTP_PASSWORD" in names and isinstance(node.value, ast.Call):
            args = node.value.args
            if len(args) >= 2 and isinstance(args[1], ast.Constant) and isinstance(args[1].value, str):
                found["SMTP_PASSWORD"] = args[1].value
    return found


def write_env(path: Path, values: dict[str, str]) -> None:
    for key, value in values.items():
        if "\n" in value or "\r" in value:
            raise ValueError(f"Variable {key} contains a newline")
    if path.exists():
        stamp = time.strftime("%Y%m%d-%H%M%S")
        shutil.copy2(path, path.with_name(f"{path.name}.before-bootstrap-{stamp}"))
    content = "\n".join(f"{key}={value}" for key, value in sorted(values.items())) + "\n"
    path.write_text(content, encoding="utf-8")
    path.chmod(0o600)


def backup_sqlite(source: Path, target: Path) -> None:
    if not source.is_file() or source.resolve() == target.resolve():
        return
    backup_dir = PROJECT_DIR / "backups"
    backup_dir.mkdir(mode=0o700, exist_ok=True)
    if target.exists():
        stamp = time.strftime("%Y%m%d-%H%M%S")
        shutil.copy2(target, backup_dir / f"database_from_clone_{stamp}.db")
    temporary = target.with_suffix(".migrating.db")
    temporary.unlink(missing_ok=True)
    source_connection = sqlite3.connect(f"file:{source}?mode=ro", uri=True)
    target_connection = sqlite3.connect(temporary)
    try:
        source_connection.backup(target_connection)
        if target_connection.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise RuntimeError("The migrated SQLite database failed integrity_check")
    finally:
        target_connection.close()
        source_connection.close()
    temporary.replace(target)
    target.chmod(0o600)


def copy_runtime_files(legacy_dir: Path) -> None:
    backup_sqlite(legacy_dir / "database.db", PROJECT_DIR / "database.db")
    cache_source = legacy_dir / "tabel_portal_cache.json"
    if cache_source.is_file():
        shutil.copy2(cache_source, PROJECT_DIR / cache_source.name)
    uploads_source = legacy_dir / "static" / "uploads" / "news"
    uploads_target = PROJECT_DIR / "static" / "uploads" / "news"
    if uploads_source.is_dir() and uploads_source.resolve() != uploads_target.resolve():
        shutil.copytree(uploads_source, uploads_target, dirs_exist_ok=True)


def copy_container_uploads(legacy_dir: Path) -> None:
    compose_file = legacy_dir / "docker-compose.yml"
    if not compose_file.is_file() or not shutil.which("docker"):
        return
    command = [
        "docker", "compose", "-f", str(compose_file),
        "--project-directory", str(legacy_dir), "ps", "-q", "web",
    ]
    result = subprocess.run(command, text=True, capture_output=True, check=False)
    container_id = result.stdout.strip()
    if not container_id:
        return
    uploads_target = PROJECT_DIR / "static" / "uploads" / "news"
    uploads_target.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        ["docker", "cp", f"{container_id}:/app/static/uploads/news/.", str(uploads_target)],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Prepare a fresh portal clone using an existing deployment"
    )
    parser.add_argument("legacy_dir", type=Path, help="path to the currently deployed portal")
    parser.add_argument(
        "--ai-dir",
        type=Path,
        default=PROJECT_DIR.parent / "ai_belnipi",
        help="path to the AI clone; default: sibling ai_belnipi",
    )
    args = parser.parse_args()
    legacy_dir = args.legacy_dir.expanduser().resolve()
    if not (legacy_dir / "app.py").is_file():
        parser.error(f"portal app.py was not found in {legacy_dir}")
    if legacy_dir == PROJECT_DIR.resolve():
        parser.error("legacy_dir must point to the old deployment, not this clone")

    target_env_path = PROJECT_DIR / ".env.production"
    values: dict[str, str] = {}
    values.update(legacy_code_settings(legacy_dir / "app.py"))
    values.update(read_env(legacy_dir / ".env"))
    values.update(read_env(legacy_dir / ".env.production"))
    values.update(read_env(target_env_path))

    ai_env_path = args.ai_dir.expanduser().resolve() / ".env.production"
    ai_values = read_env(ai_env_path)
    shared_secret = values.get("AI_SSO_SHARED_SECRET") or ai_values.get("AI_SSO_SHARED_SECRET")
    if not shared_secret or shared_secret.startswith("CHANGE_ME") or len(shared_secret) < 32:
        shared_secret = secrets.token_hex(32)

    portal_secret = values.get("PORTAL_SECRET_KEY", "")
    if not portal_secret or portal_secret.startswith("CHANGE_ME") or len(portal_secret) < 32:
        portal_secret = secrets.token_hex(32)

    values.update({
        "PORTAL_SECRET_KEY": portal_secret,
        "SESSION_COOKIE_SECURE": "1",
        "AI_ASSISTANT_PUBLIC_URL": "https://ai.energoprom.by",
        "AI_ASSISTANT_CALLBACK_URL": "https://ai.energoprom.by/sso-login",
        "AI_SSO_SHARED_SECRET": shared_secret,
        "AI_SSO_MAX_AGE_SEC": "300",
    })
    required = ("LDAP_BIND_DN", "LDAP_BIND_PASSWORD")
    missing = [key for key in required if not values.get(key) or values[key].startswith("CHANGE_ME")]
    if missing:
        raise SystemExit(
            "Cannot migrate required Active Directory settings: " + ", ".join(missing)
        )

    write_env(target_env_path, values)
    copy_runtime_files(legacy_dir)
    copy_container_uploads(legacy_dir)

    prepare_script = args.ai_dir.expanduser().resolve() / "deploy" / "prepare-sso-production.sh"
    if prepare_script.is_file():
        subprocess.run([str(prepare_script), str(PROJECT_DIR)], check=True)
        ai_status = f"AI configuration synchronized in {args.ai_dir.expanduser().resolve()}"
    else:
        ai_status = "AI clone was not found; securely synchronize AI_SSO_SHARED_SECRET before launch"

    print(f"Portal clone prepared: {PROJECT_DIR}")
    print("LDAP/SMTP values and SSO secrets were migrated without printing them")
    print("database.db, tabel cache and news uploads were copied when present")
    print(ai_status)
    print("Next command: bash deploy/update-docker-production.sh")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
