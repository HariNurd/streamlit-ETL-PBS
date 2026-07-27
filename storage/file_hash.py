import hashlib
from pathlib import Path


def compute_file_hash(file_path, chunk_size=1024 * 1024):
    file_path = Path(file_path)
    digest = hashlib.sha256()

    with file_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)

    return digest.hexdigest()
