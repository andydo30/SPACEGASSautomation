# SPACEGASSautomation justfile

# Default recipe - show available commands
default:
    @just --list

# Run all tests
test:
    uv run pytest tests/

# Run tests with verbose output
test-verbose:
    uv run pytest tests/ -v

# Run tests with coverage report
test-cov:
    uv run pytest tests/ --cov=src --cov-report=term-missing

# Run tests with HTML coverage report
test-cov-html:
    uv run pytest tests/ --cov=src --cov-report=html
    @echo "Coverage report generated in htmlcov/index.html"

# Run a specific test file
test-file FILE:
    uv run pytest {{FILE}} -v

# Run tests matching a pattern
test-match PATTERN:
    uv run pytest tests/ -v -k "{{PATTERN}}"

# Install dev dependencies
install:
    uv sync --group dev

# Format code with black
format:
    uv run black src/ tests/

# Check formatting without changes
format-check:
    uv run black src/ tests/ --check
