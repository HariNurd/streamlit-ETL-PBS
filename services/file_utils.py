import re
import uuid
import zipfile
import logging
import os
import time
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ALLOWED_CLEANUP_DIRS = {
    "temp": PROJECT_ROOT / "temp",
    "uploads": PROJECT_ROOT / "uploads",
    "cache": PROJECT_ROOT / "cache",
}
PROTECTED_DIR_NAMES = {
    ".git",
    ".streamlit",
    "data",
    "outputs",
    "logs",
    "env",
    "env-win",
    "venv",
    ".venv",
    "extractors",
    "transformers",
    "validators",
    "loaders",
    "services",
    "storage",
    "dashboards",
    "converters",
}


def safe_filename(name):
    path_name = Path(str(name)).name
    cleaned = re.sub(r"[\x00-\x1f<>:\"/\\|?*]+", "_", path_name).strip(" ._")
    return cleaned or "file"


def create_session_dir(prefix, base_dir="temp"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = Path(base_dir) / f"{safe_filename(prefix)}_{timestamp}_{uuid.uuid4().hex[:8]}"
    session_dir.mkdir(parents=True, exist_ok=True)
    return session_dir


def save_uploaded_files(uploaded_files, target_dir):
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)

    if uploaded_files is None:
        return []
    if not isinstance(uploaded_files, (list, tuple)):
        uploaded_files = [uploaded_files]

    saved_files = []
    used_names = set()
    for uploaded_file in uploaded_files:
        original_name = safe_filename(uploaded_file.name)
        path = Path(original_name)
        stem = path.stem or "file"
        suffix = path.suffix
        candidate_name = original_name
        counter = 1

        while candidate_name.lower() in used_names or (target_dir / candidate_name).exists():
            candidate_name = f"{stem}_{counter}{suffix}"
            counter += 1

        used_names.add(candidate_name.lower())
        output_path = target_dir / candidate_name
        output_path.write_bytes(uploaded_file.getbuffer())
        saved_files.append(output_path)

    return saved_files


def zip_files(file_paths, output_zip):
    output_zip = Path(output_zip)
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    used_names = set()

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in file_paths:
            file_path = Path(file_path)
            if not file_path.exists():
                continue

            arcname = file_path.name
            if arcname.lower() in used_names:
                arcname = f"{file_path.parent.name}_{file_path.name}"
            used_names.add(arcname.lower())
            archive.write(file_path, arcname=arcname)

    return output_zip


def _new_cleanup_summary():
    return {
        "deleted_files": 0,
        "deleted_dirs": 0,
        "skipped_files": 0,
        "errors": [],
        "targets": {},
    }


def _target_summary(target_folder):
    return {
        "target_folder": target_folder,
        "deleted_files": 0,
        "deleted_dirs": 0,
        "skipped_files": 0,
        "errors": [],
    }


def _merge_cleanup_summary(target, source):
    for key in ("deleted_files", "deleted_dirs", "skipped_files"):
        target[key] += int(source.get(key, 0) or 0)
    target["errors"].extend(source.get("errors", []))

    for folder, folder_summary in source.get("targets", {}).items():
        if folder not in target["targets"]:
            target["targets"][folder] = _target_summary(folder)
        _merge_cleanup_summary(target["targets"][folder], folder_summary)

    return target


def combine_cleanup_summaries(*summaries):
    combined = _new_cleanup_summary()
    for summary in summaries:
        if summary:
            _merge_cleanup_summary(combined, summary)
    return combined


def _add_cleanup_error(summary, path, message):
    summary["errors"].append({"path": str(path), "error": str(message)})
    return summary


def _absolute_path(path):
    path = Path(path)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.absolute()


def _norm_path(path):
    return os.path.normcase(os.path.normpath(str(path)))


def _is_inside_path(candidate, parent):
    candidate_norm = _norm_path(candidate)
    parent_norm = _norm_path(parent)
    try:
        return os.path.commonpath([candidate_norm, parent_norm]) == parent_norm
    except ValueError:
        return False


def _cleanup_target_for_path(path):
    absolute = _absolute_path(path)
    for target_name, target_dir in ALLOWED_CLEANUP_DIRS.items():
        if _is_inside_path(absolute, target_dir):
            return target_name
    return None


def _is_allowed_cleanup_path(path):
    absolute = _absolute_path(path)
    if _norm_path(absolute) == _norm_path(PROJECT_ROOT):
        return False
    if absolute.name in PROTECTED_DIR_NAMES:
        return False
    return _cleanup_target_for_path(absolute) is not None


def _should_keep_file(path, keep_gitkeep=True, older_than_hours=None):
    if keep_gitkeep and Path(path).name == ".gitkeep":
        return True

    if older_than_hours is None:
        return False

    try:
        threshold_seconds = float(older_than_hours) * 3600
        age_seconds = time.time() - Path(path).stat(follow_symlinks=False).st_mtime
        return age_seconds < threshold_seconds
    except (OSError, TypeError, ValueError):
        return False


def _is_top_level_cleanup_dir(path):
    absolute = _absolute_path(path)
    return any(_norm_path(absolute) == _norm_path(folder) for folder in ALLOWED_CLEANUP_DIRS.values())


def safe_delete_file(path, keep_gitkeep=True, older_than_hours=None):
    summary = _new_cleanup_summary()
    path = _absolute_path(path)
    target_folder = _cleanup_target_for_path(path) or "unknown"
    summary["targets"][target_folder] = _target_summary(target_folder)

    if not _is_allowed_cleanup_path(path):
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        _add_cleanup_error(summary, path, "Skipped: path is outside allowed cleanup folders.")
        _add_cleanup_error(summary["targets"][target_folder], path, "Skipped: path is outside allowed cleanup folders.")
        return summary

    try:
        if not path.exists() and not path.is_symlink():
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            return summary
        if path.is_symlink():
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            _add_cleanup_error(summary, path, "Skipped: symlinks are not deleted or followed.")
            _add_cleanup_error(summary["targets"][target_folder], path, "Skipped: symlinks are not deleted or followed.")
            return summary
        if path.is_dir():
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            _add_cleanup_error(summary, path, "Skipped: expected a file, got a directory.")
            _add_cleanup_error(summary["targets"][target_folder], path, "Skipped: expected a file, got a directory.")
            return summary
        if _should_keep_file(path, keep_gitkeep=keep_gitkeep, older_than_hours=older_than_hours):
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            return summary

        path.unlink()
        summary["deleted_files"] += 1
        summary["targets"][target_folder]["deleted_files"] += 1
    except OSError as exc:
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        _add_cleanup_error(summary, path, exc)
        _add_cleanup_error(summary["targets"][target_folder], path, exc)

    return summary


def _delete_empty_dir(path, keep_gitkeep=True, older_than_hours=None):
    summary = _new_cleanup_summary()
    path = _absolute_path(path)
    target_folder = _cleanup_target_for_path(path) or "unknown"
    summary["targets"][target_folder] = _target_summary(target_folder)

    if not _is_allowed_cleanup_path(path) or _is_top_level_cleanup_dir(path):
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        return summary

    try:
        if not path.exists() and not path.is_symlink():
            return summary
        if path.is_symlink():
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            _add_cleanup_error(summary, path, "Skipped: symlink directory is not deleted or followed.")
            _add_cleanup_error(summary["targets"][target_folder], path, "Skipped: symlink directory is not deleted or followed.")
            return summary
        if not path.is_dir():
            return summary
        if _should_keep_file(path, keep_gitkeep=False, older_than_hours=older_than_hours):
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            return summary

        children = list(path.iterdir())
        if keep_gitkeep and any(child.name == ".gitkeep" for child in children):
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            return summary

        remaining = [child.name for child in children]
        if remaining:
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            return summary

        path.rmdir()
        summary["deleted_dirs"] += 1
        summary["targets"][target_folder]["deleted_dirs"] += 1
    except OSError as exc:
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        _add_cleanup_error(summary, path, exc)
        _add_cleanup_error(summary["targets"][target_folder], path, exc)

    return summary


def safe_delete_dir_contents(dir_path, keep_gitkeep=True, older_than_hours=None):
    summary = _new_cleanup_summary()
    dir_path = _absolute_path(dir_path)
    target_folder = _cleanup_target_for_path(dir_path) or "unknown"
    summary["targets"][target_folder] = _target_summary(target_folder)

    if not _is_allowed_cleanup_path(dir_path):
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        _add_cleanup_error(summary, dir_path, "Skipped: directory is outside allowed cleanup folders.")
        _add_cleanup_error(summary["targets"][target_folder], dir_path, "Skipped: directory is outside allowed cleanup folders.")
        return summary

    try:
        if not dir_path.exists() and not dir_path.is_symlink():
            return summary
        if dir_path.is_symlink():
            summary["skipped_files"] += 1
            summary["targets"][target_folder]["skipped_files"] += 1
            _add_cleanup_error(summary, dir_path, "Skipped: symlink directory is not deleted or followed.")
            _add_cleanup_error(summary["targets"][target_folder], dir_path, "Skipped: symlink directory is not deleted or followed.")
            return summary
        if not dir_path.is_dir():
            return _merge_cleanup_summary(
                summary,
                safe_delete_file(dir_path, keep_gitkeep=keep_gitkeep, older_than_hours=older_than_hours),
            )

        children = sorted(dir_path.iterdir(), key=lambda item: (not item.is_dir(), str(item).lower()))
        for child in children:
            if child.is_symlink():
                child_summary = _new_cleanup_summary()
                child_summary["skipped_files"] += 1
                child_summary["targets"][target_folder] = _target_summary(target_folder)
                child_summary["targets"][target_folder]["skipped_files"] += 1
                _add_cleanup_error(child_summary, child, "Skipped: symlinks are not deleted or followed.")
                _add_cleanup_error(child_summary["targets"][target_folder], child, "Skipped: symlinks are not deleted or followed.")
                _merge_cleanup_summary(summary, child_summary)
            elif child.is_dir():
                _merge_cleanup_summary(
                    summary,
                    safe_delete_dir_contents(
                        child,
                        keep_gitkeep=keep_gitkeep,
                        older_than_hours=older_than_hours,
                    ),
                )
                _merge_cleanup_summary(
                    summary,
                    _delete_empty_dir(
                        child,
                        keep_gitkeep=keep_gitkeep,
                        older_than_hours=older_than_hours,
                    ),
                )
            else:
                _merge_cleanup_summary(
                    summary,
                    safe_delete_file(
                        child,
                        keep_gitkeep=keep_gitkeep,
                        older_than_hours=older_than_hours,
                    ),
                )
    except OSError as exc:
        summary["skipped_files"] += 1
        summary["targets"][target_folder]["skipped_files"] += 1
        _add_cleanup_error(summary, dir_path, exc)
        _add_cleanup_error(summary["targets"][target_folder], dir_path, exc)

    return summary


def cleanup_paths(file_paths=None, dir_paths=None, keep_gitkeep=True, older_than_hours=None, remove_empty_dirs=True):
    summary = _new_cleanup_summary()

    for file_path in file_paths or []:
        _merge_cleanup_summary(
            summary,
            safe_delete_file(
                file_path,
                keep_gitkeep=keep_gitkeep,
                older_than_hours=older_than_hours,
            ),
        )

    for dir_path in dir_paths or []:
        _merge_cleanup_summary(
            summary,
            safe_delete_dir_contents(
                dir_path,
                keep_gitkeep=keep_gitkeep,
                older_than_hours=older_than_hours,
            ),
        )
        if remove_empty_dirs:
            _merge_cleanup_summary(
                summary,
                _delete_empty_dir(
                    dir_path,
                    keep_gitkeep=keep_gitkeep,
                    older_than_hours=older_than_hours,
                ),
            )

    return summary


def cleanup_working_dirs(
    clean_temp=True,
    clean_uploads=True,
    clean_cache=False,
    older_than_hours=None,
):
    """
    Safely remove temporary working files from temp/, uploads/, and optionally cache/.
    """
    summary = _new_cleanup_summary()
    targets = {
        "temp": clean_temp,
        "uploads": clean_uploads,
        "cache": clean_cache,
    }

    for target_name, enabled in targets.items():
        if not enabled:
            continue
        target_dir = ALLOWED_CLEANUP_DIRS[target_name]
        target_dir.mkdir(parents=True, exist_ok=True)
        _merge_cleanup_summary(
            summary,
            safe_delete_dir_contents(
                target_dir,
                keep_gitkeep=True,
                older_than_hours=older_than_hours,
            ),
        )

    return summary


def _cleanup_logger():
    log_dir = PROJECT_ROOT / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("etl.cleanup")
    logger.setLevel(logging.INFO)

    log_path = log_dir / "app.log"
    if not any(isinstance(handler, logging.FileHandler) and Path(handler.baseFilename) == log_path for handler in logger.handlers):
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        logger.addHandler(handler)

    return logger


def log_cleanup_summary(summary, cleanup_type="manual", job_id=None):
    logger = _cleanup_logger()
    logger.info(
        "cleanup_type=%s job_id=%s deleted_files=%s deleted_dirs=%s skipped_files=%s error_count=%s",
        cleanup_type,
        job_id,
        summary.get("deleted_files", 0),
        summary.get("deleted_dirs", 0),
        summary.get("skipped_files", 0),
        len(summary.get("errors", [])),
    )
    for error in summary.get("errors", []):
        logger.warning(
            "cleanup_type=%s job_id=%s path=%s error=%s",
            cleanup_type,
            job_id,
            error.get("path"),
            error.get("error"),
        )


def cleanup_old_temp_files():
    return cleanup_working_dirs(clean_temp=True, clean_uploads=False, clean_cache=False)
