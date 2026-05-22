# -*- coding: utf-8 -*-
"""
Кодировки имён файлов/папок табеля на Linux (utf-8, cp1251, cp866).

Пути в настройках остаются как на Windows (\\\\srv-doc\\ТАБЕЛЬ\\...).
На Linux они автоматически читаются через тот же mount (/mnt/tabel) — менять
TABEL_BASE_DIR в конфиге не нужно.
"""
from __future__ import annotations

import os
import sys
from typing import Iterator, Tuple

TABEL_FILENAME_ENCODINGS = ('utf-8', 'cp1251', 'cp866')

# Префиксы шары в настройках (как на Windows).
TABEL_UNC_ROOT_MARKERS = (
    r'\\srv-doc\ТАБЕЛЬ',
    r'//srv-doc/ТАБЕЛЬ',
)


def configure_tabel_locale() -> None:
    if sys.platform == 'win32':
        return
    os.environ.setdefault('PYTHONUTF8', '1')
    for var in ('LANG', 'LC_ALL', 'LC_CTYPE'):
        os.environ.setdefault(var, 'C.UTF-8')


def tabel_linux_mount_dir() -> str:
    return os.environ.get('TABEL_LINUX_DEFAULT_BASE', '/mnt/tabel').rstrip('/')


def tabel_default_base_dir() -> str:
    return r'\\srv-doc\ТАБЕЛЬ'


def tabel_default_leaders_file() -> str:
    return r'\\srv-doc\ТАБЕЛЬ\ОЦ\Список руководителей.xlsx'


def is_windows_unc_path(path: str | None) -> bool:
    p = (path or '').strip()
    return p.startswith('\\\\') or (len(p) > 1 and p.startswith('//') and not p.startswith('///'))


def _unc_path_tail(path: str) -> str | None:
    """Часть пути после \\\\srv-doc\\ТАБЕЛЬ (для подстановки mount на Linux)."""
    if not path:
        return None
    normalized = path.strip().replace('/', '\\')
    for marker in TABEL_UNC_ROOT_MARKERS:
        m = marker.replace('/', '\\')
        low = normalized.lower()
        pos = low.find(m.lower())
        if pos >= 0:
            tail = normalized[pos + len(m):].lstrip('\\')
            return tail.replace('\\', '/')
    if normalized.startswith('\\\\'):
        parts = [p for p in normalized.split('\\') if p]
        if len(parts) >= 2 and parts[0].lower() == 'srv-doc':
            tail_parts = parts[2:] if len(parts) > 2 and parts[1] else parts[1:]
            return '/'.join(tail_parts)
    return None


def resolve_tabel_path(path: str) -> str:
    """
    Windows: путь как в настройках.
    Linux: \\\\srv-doc\\ТАБЕЛЬ\\... → /mnt/tabel/... (без смены конфига).
    """
    if not path:
        return path
    if sys.platform == 'win32':
        return path
    tail = _unc_path_tail(path)
    if tail is not None:
        mount = tabel_linux_mount_dir()
        return os.path.join(mount, tail) if tail else mount
    return os.path.normpath(path.replace('\\', '/'))


# Совместимость с app.py
normalize_tabel_path = resolve_tabel_path


def tabel_access_hint(base_dir: str, leaders_file: str) -> str | None:
    if sys.platform == 'win32':
        return None
    resolved_base = resolve_tabel_path(base_dir)
    if os.path.isdir(resolved_base):
        return None
    mount = tabel_linux_mount_dir()
    return (
        f'Шара недоступна ({resolved_base}). На Linux один раз смонтируйте '
        f'//srv-doc/ТАБЕЛЬ в {mount} с iocharset=utf8. '
        f'Пути \\\\srv-doc\\... в настройках менять не нужно — только кодировка и mount.'
    )


def linux_unc_misconfiguration(base_dir: str, leaders_file: str) -> str | None:
    return tabel_access_hint(base_dir, leaders_file)


def decode_tabel_name(name) -> str:
    if isinstance(name, bytes):
        for encoding in TABEL_FILENAME_ENCODINGS:
            try:
                return name.decode(encoding)
            except UnicodeDecodeError:
                continue
        return name.decode('utf-8', errors='replace')
    text = str(name)
    if sys.platform == 'win32':
        return text
    if _cyrillic_ratio(text) >= 0.15:
        return text
    for encoding in ('cp1251', 'cp866', 'utf-8'):
        try:
            candidate = text.encode('latin-1').decode(encoding)
            if _cyrillic_ratio(candidate) > _cyrillic_ratio(text):
                return candidate
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return text


def _cyrillic_ratio(text: str) -> float:
    if not text:
        return 0.0
    cyr = sum(1 for ch in text if '\u0400' <= ch <= '\u04FF')
    return cyr / max(len(text), 1)


def tabel_path_exists(path: str) -> bool:
    path = resolve_tabel_path(path)
    if not path:
        return False
    try:
        return os.path.exists(path)
    except OSError:
        return False


def join_tabel_path(*parts: str) -> str:
    if not parts:
        return ''
    if sys.platform == 'win32':
        return os.path.join(*parts)
    first = resolve_tabel_path(parts[0])
    rest = [str(p).replace('\\', '/') for p in parts[1:]]
    return os.path.join(first, *rest) if rest else first


def listdir_tabel(directory: str) -> list[str]:
    directory = resolve_tabel_path(directory)
    return [decode_tabel_name(n) for n in os.listdir(directory)]


def walk_tabel_excel(base_dir: str) -> Iterator[Tuple[str, str, str]]:
    base_dir = resolve_tabel_path(base_dir)
    if not base_dir or not tabel_path_exists(base_dir):
        return
    for root, _, files in os.walk(base_dir):
        rel = os.path.relpath(root, base_dir)
        dept = rel.split(os.sep)[0] if rel != '.' else 'Общий'
        for raw_name in files:
            filename = decode_tabel_name(raw_name)
            if not filename.lower().endswith(('.xls', '.xlsx')) or filename.startswith('~$'):
                continue
            yield root, dept, filename


def tabel_path_status(base_dir: str, leaders_file: str) -> dict:
    resolved_base = resolve_tabel_path(base_dir)
    resolved_leaders = resolve_tabel_path(leaders_file)
    base_exists = tabel_path_exists(base_dir)
    leaders_exists = tabel_path_exists(leaders_file)
    hint = tabel_access_hint(base_dir, leaders_file)
    if not hint and not leaders_exists:
        hint = f'Не найден файл руководителей: {resolved_leaders}'
    return {
        'platform': sys.platform,
        'base_dir': base_dir,
        'leaders_file': leaders_file,
        'resolved_base_dir': resolved_base,
        'resolved_leaders_file': resolved_leaders,
        'base_dir_exists': base_exists,
        'leaders_file_exists': leaders_exists,
        'setup_hint': hint,
        'filename_encodings': list(TABEL_FILENAME_ENCODINGS),
    }


configure_tabel_locale()
