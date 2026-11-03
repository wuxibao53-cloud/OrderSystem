# -*- coding: utf-8 -*-
"""Small helpers for durable JSON persistence."""
import json
import os
import tempfile


def save_json_atomic(path, data, indent=1):
    """Write JSON beside its destination and replace it only after serialization succeeds."""
    path = os.path.abspath(path)
    directory = os.path.dirname(path)
    fd, temporary = tempfile.mkstemp(
        prefix=".%s." % os.path.basename(path),
        suffix=".tmp",
        dir=directory,
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=indent)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except BaseException:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise
