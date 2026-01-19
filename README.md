# SPACEGASS Automation

Help run and analyse SPACEGASS structural analyses.

## Overview

<!-- TODO: Add project background and purpose -->

## Prerequisites

- **Python 3.13+**
- **[uv](https://docs.astral.sh/uv/)** - Fast Python package manager
  ```bash
  # Install uv (Windows PowerShell)
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

  # Install uv (macOS/Linux)
  curl -LsSf https://astral.sh/uv/install.sh | sh
  ```
- **[just](https://github.com/casey/just)** - Command runner (optional, for dev commands)
  ```bash
  # Install just (Windows - requires cargo)
  cargo install just

  # Install just (macOS)
  brew install just

  # Install just (Linux)
  cargo install just
  # or via package manager (e.g., apt install just)
  ```

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd SPACEGASSautomation

# Install dependencies
uv sync

# Install with dev dependencies
uv sync --group dev
```

## Quick Start

```python
from sg_results import SGResults

# Load a SPACEGASS output file
results = SGResults('path/to/output.txt')

# Access parsed data as DataFrames
nodes = results.nodes
members = results.members
sections = results.sections
materials = results.materials

# Get units from the file
units = results.units  # {'LENGTH': 'm', 'FORCE': 'kN', ...}

# Query member forces and moments
forces = results.query_forces_moments(load_case_id=1, member_id=[1, 2, 3])

# Query members with their section properties
member_sections = results.query_member_sections(member_id=[1, 2, 3])
```

## API Reference

### SGResults

Main class for parsing SPACEGASS text output files.

#### Constructor

```python
SGResults(filepath: str)
```

#### Properties

| Property | Returns | Description |
|----------|---------|-------------|
| `units` | `Dict[str, str]` | Units dictionary from file |
| `nodes` | `DataFrame` | Node coordinates |
| `members` | `DataFrame` | Member definitions |
| `sections` | `DataFrame` | Section properties |
| `materials` | `DataFrame` | Material properties |
| `restraints` | `DataFrame` | Node restraints |
| `titles` | `DataFrame` | Load case titles |
| `combinations` | `DataFrame` | Load combinations |
| `displacements` | `DataFrame` | Node displacements |
| `reactions` | `DataFrame` | Support reactions |
| `member_forces_moments` | `DataFrame` | Member end forces |
| `member_int_forces_moments` | `DataFrame` | Intermediate forces |
| `member_stresses` | `DataFrame` | Member stresses |
| `steel_members` | `DataFrame` | Steel member data |

#### Methods

##### `query_forces_moments(load_case_id=None, member_id=None)`

Filter member forces and moments by load case and/or member ID.

```python
# All forces
forces = results.query_forces_moments()

# Filter by load case
forces = results.query_forces_moments(load_case_id=1)

# Filter by member
forces = results.query_forces_moments(member_id=[1, 2, 3])
```

##### `query_member_sections(member_id=None)`

Query member information joined with section properties.

```python
# All members with sections
data = results.query_member_sections()

# Specific members
data = results.query_member_sections(member_id=[1, 2, 3])
```

Returns a DataFrame with all member columns joined with all section columns. Warns if any requested member IDs don't exist.

##### `summary()`

Returns a string summary of all loaded sections and their sizes.

## Configuration

<!-- TODO: Add any configuration options -->

## Development

This project uses [just](https://github.com/casey/just) as a command runner.

### Available Commands

```bash
just                      # List all available commands
just install              # Install dev dependencies
just test                 # Run all tests
just test-verbose         # Run tests with verbose output
just test-cov             # Run tests with coverage report
just test-cov-html        # Run tests with HTML coverage report
just test-file FILE       # Run a specific test file
just test-match PATTERN   # Run tests matching a pattern
just format               # Format code with black
just format-check         # Check formatting without changes
```

### Project Structure

```
SPACEGASSautomation/
├── src/
│   ├── sg_results.py    # Main parser class
│   └── script.py        # Legacy automation script
├── tests/
│   ├── conftest.py      # Test fixtures
│   ├── test_sg_results.py
│   └── fixtures/        # Test data files
├── pyproject.toml
└── README.md
```

## Contributing

<!-- TODO: Add contribution guidelines -->

## License

<!-- TODO: Add license information -->
