from pathlib import Path


PROJECT_DIRS = [
    Path("extractors"),
    Path("transformers"),
    Path("validators"),
    Path("loaders"),
    Path("services"),
    Path("storage"),
    Path("dashboards"),
    Path("data/raw"),
    Path("data/staging"),
    Path("data/dwh"),
    Path("data/mart"),
    Path("data/metadata"),
    Path("logs"),
    Path("uploads"),
    Path("outputs"),
    Path("cache"),
    Path("cache/staging"),
    Path("cache/staging/slik"),
    Path("cache/staging/bank_pdf"),
    Path("cache/staging/bank_txt"),
    Path("temp"),
]


def ensure_project_dirs(extra_dirs=None):
    for folder in [*PROJECT_DIRS, *(extra_dirs or [])]:
        Path(folder).mkdir(parents=True, exist_ok=True)
