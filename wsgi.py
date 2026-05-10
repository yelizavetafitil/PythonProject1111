"""WSGI entrypoint for Gunicorn / production servers."""
import os

from app import (
    KNOWLEDGE_BASE_INSTRUCTIONS_DIR,
    app,
    init_db,
    _tabel_load_cache,
    _tabel_rebuild_index_from_cache,
)

os.makedirs(KNOWLEDGE_BASE_INSTRUCTIONS_DIR, exist_ok=True)
init_db()
_tabel_load_cache()
_tabel_rebuild_index_from_cache()
