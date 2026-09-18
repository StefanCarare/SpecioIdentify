"""Portable catalog paths. Stored image paths are relative to catalog.db's folder."""
from pathlib import Path, PureWindowsPath, PurePosixPath


def store_catalog_path(path: Path, output_dir: Path) -> str:
    return Path(path).resolve().relative_to(Path(output_dir).resolve()).as_posix()


def store_original_path(path: Path, output_dir: Path) -> str:
    """Keep project-local provenance relative; external originals remain absolute."""
    path = Path(path).resolve()
    try:
        return path.relative_to(Path(output_dir).resolve().parent).as_posix()
    except ValueError:
        return str(path)


def resolve_catalog_path(value: str, output_dir: Path) -> Path:
    """Resolve locally, including legacy absolute output/category/species/file paths.

    Never fall back to an old installation, even if it still exists. Unknown
    absolute layouts are rejected rather than guessed or modified.
    """
    root = Path(output_dir).resolve()
    value = str(value)
    if not value.strip():
        raise ValueError("Empty catalog image path")
    windows = PureWindowsPath(value)
    posix = PurePosixPath(value)
    path = Path(value.replace('\\', '/'))
    if path.is_absolute() or windows.is_absolute() or posix.is_absolute():
        try:
            relative = path.relative_to(root)
        except ValueError:
            parts = windows.parts if windows.is_absolute() else posix.parts
            # The standard legacy layout has exactly three components below output.
            if len(parts) < 4 or parts[-4].casefold() != root.name.casefold():
                raise ValueError(f"Unsupported legacy catalog path: {value}")
            relative = Path(*parts[-3:])
    else:
        if windows.drive or windows.root:
            raise ValueError(f"Invalid catalog path: {value}")
        relative = path
    resolved = (root / relative).resolve()
    if not resolved.is_relative_to(root) or resolved == root:
        raise ValueError(f"Catalog path is outside its root: {value}")
    return resolved


def resolved_catalog_item(row, output_dir: Path) -> dict:
    """Preserve the API's absolute path contract for GUI and other callers."""
    item = dict(row)
    item['catalog_path'] = str(resolve_catalog_path(item['catalog_path'], output_dir))
    return item
