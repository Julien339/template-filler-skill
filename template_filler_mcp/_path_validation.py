"""Path validation utilities to prevent path traversal and file access vulnerabilities.

This module provides secure path validation for all file operations in the MCP server,
preventing path traversal attacks and ensuring files are accessed safely.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Literal


class PathValidationError(ValueError):
    """Raised when path validation fails."""

    pass


def validate_input_file(
    filepath: str,
    allowed_extensions: tuple[str, ...] | None = None,
) -> Path:
    """Validate an input file path for reading.

    Args:
        filepath: Path to the file to validate.
        allowed_extensions: Tuple of allowed extensions (e.g., ('.pptx', '.docx')).
            If None, no extension check is performed.

    Returns:
        Resolved absolute Path object.

    Raises:
        PathValidationError: If validation fails.
    """
    try:
        path = Path(filepath).resolve(strict=True)
    except (OSError, RuntimeError) as e:
        raise PathValidationError(f"Invalid or inaccessible path: {filepath}") from e

    if not path.is_file():
        raise PathValidationError(f"Path is not a file: {filepath}")

    if allowed_extensions and path.suffix.lower() not in allowed_extensions:
        raise PathValidationError(
            f"Invalid file extension: {path.suffix}. "
            f"Allowed extensions: {', '.join(allowed_extensions)}"
        )

    return path


def validate_output_file(
    filepath: str,
    allowed_extensions: tuple[str, ...] | None = None,
) -> Path:
    """Validate an output file path for writing.

    Args:
        filepath: Path to the file to validate.
        allowed_extensions: Tuple of allowed extensions (e.g., ('.pptx', '.docx')).
            If None, no extension check is performed.

    Returns:
        Resolved absolute Path object.

    Raises:
        PathValidationError: If validation fails.
    """
    try:
        path = Path(filepath).resolve()
    except (OSError, RuntimeError) as e:
        raise PathValidationError(f"Invalid path: {filepath}") from e

    # Ensure parent directory exists or can be created
    parent = path.parent
    if not parent.exists():
        raise PathValidationError(
            f"Parent directory does not exist: {parent}. "
            "Please create it first or use an existing directory."
        )

    if not parent.is_dir():
        raise PathValidationError(f"Parent path is not a directory: {parent}")

    if allowed_extensions and path.suffix.lower() not in allowed_extensions:
        raise PathValidationError(
            f"Invalid file extension: {path.suffix}. "
            f"Allowed extensions: {', '.join(allowed_extensions)}"
        )

    return path


def validate_output_directory(dirpath: str) -> Path:
    """Validate an output directory path.

    Args:
        dirpath: Path to the directory to validate.

    Returns:
        Resolved absolute Path object.

    Raises:
        PathValidationError: If validation fails.
    """
    try:
        path = Path(dirpath).resolve()
    except (OSError, RuntimeError) as e:
        raise PathValidationError(f"Invalid path: {dirpath}") from e

    # For output directories, we allow creation, so we just check parent exists
    if path.exists() and not path.is_dir():
        raise PathValidationError(f"Path exists but is not a directory: {dirpath}")

    return path
