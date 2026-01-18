"""Pytest configuration and fixtures for sg_results tests."""

import sys
from pathlib import Path

import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sg_results import SGResults


FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def minimal_file_path() -> Path:
    """Path to minimal test fixture file."""
    return FIXTURES_DIR / "minimal_sg_output.txt"


@pytest.fixture
def example_file_path() -> Path:
    """Path to full example fixture file."""
    return FIXTURES_DIR / "example_sg_output.txt"


@pytest.fixture
def minimal_results(minimal_file_path: Path) -> SGResults:
    """SGResults instance loaded from minimal fixture."""
    return SGResults(str(minimal_file_path))


@pytest.fixture
def example_results(example_file_path: Path) -> SGResults:
    """SGResults instance loaded from full example fixture."""
    return SGResults(str(example_file_path))
