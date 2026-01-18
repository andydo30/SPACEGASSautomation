"""
SPACEGASS Results Parser

A class for parsing and interacting with SPACEGASS structural analysis output files.
"""

import csv
import re
from io import StringIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd


class SGResults:
    """
    Parser for SPACEGASS text output files.

    Usage:
        results = SGResults('path/to/output.txt')
        nodes_df = results.nodes
        members_df = results.members
        units = results.units  # {'LENGTH': 'm', 'FORCE': 'kN', ...}
    """

    # Mapping of section names in file to attribute names
    SECTION_MAP = {
        'FILTERS': 'filters',
        'NODES': 'nodes',
        'MEMBERS': 'members',
        'RESTRAINTS': 'restraints',
        'SECTIONS': 'sections',
        'MATERIALS': 'materials',
        'NODELOADS': 'node_loads',
        'MEMBFORCES': 'member_forces',
        'SELFWEIGHT': 'self_weight',
        'HARMONIC LOADS': 'harmonic_loads',
        'COMBINATIONS': 'combinations',
        'TITLES': 'titles',
        'LOAD CASE GROUPS': 'load_case_groups',
        'LOAD CATEGORIES': 'load_categories',
        'DISPLACEMENTS': 'displacements',
        'MEMBER FORCES AND MOMENTS': 'member_forces_moments',
        'REACTIONS': 'reactions',
        'MEMBER INTERMEDIATE DISPLACEMENTS': 'member_int_displacements',
        'MEMBER INTERMEDIATE FORCES AND MOMENTS': 'member_int_forces_moments',
        'MEMBER STRESSES': 'member_stresses',
    }

    # Sections that have multi-line records
    MULTILINE_SECTIONS = {
        'RESTRAINTS': 5,  # 1 main + 4 continuation lines
        'SECTIONS': 4,    # 1 main + 3 continuation lines
    }

    # Pre-compiled regex for restraint code validation (performance optimization)
    _RESTRAINT_CODE_PATTERN = re.compile(r'^[VRF]{6}$')

    # -------------------------------------------------------------------------
    # Column name mappings for each section
    # Update these lists with meaningful column names for your needs.
    # If a section is not listed here, or has fewer names than columns,
    # remaining columns will use placeholder names (col_0, col_1, etc.)
    # -------------------------------------------------------------------------
    SECTION_COLUMNS = {
        'NODES': ['node_id', 'x', 'y', 'z'],
        'MEMBERS': [
            'member_id', 'col_1', 'col_2', 'col_3', 'col_4',
            'node_i', 'node_j', 'section_id', 'material_id',
            'releases_i', 'releases_j',
            'col_11', 'col_12', 'col_13', 'col_14', 'col_15',
            'col_16', 'col_17', 'col_18', 'col_19',
        ],
        'RESTRAINTS': [
            'node_id', 'restraint_code', 'col_2', 'col_3', 'col_4',
            'col_5', 'col_6', 'col_7', 'col_8', 'col_9', 'col_10',
            'col_11', 'col_12', 'col_13', 'col_14', 'col_15', 'col_16',
            'col_17', 'col_18', 'col_19', 'col_20',
            'cont1_col_0', 'cont1_col_1', 'cont1_col_2', 'cont1_col_3',
            'cont2_col_0', 'cont2_col_1', 'cont2_col_2', 'cont2_col_3',
            'cont3_col_0', 'cont3_col_1', 'cont3_col_2', 'cont3_col_3',
            'cont4_col_0', 'cont4_col_1', 'cont4_col_2', 'cont4_col_3',
        ],
        'SECTIONS': [
            'section_id', 'name', 'col_2', 'short_name',
            'area', 'Ixx', 'Iyy', 'J', 'col_8', 'col_9', 'col_10',
            'col_11', 'col_12', 'col_13', 'col_14', 'col_15',
        ],
        'MATERIALS': [
            'material_id', 'name', 'col_2', 'E', 'poisson', 'density',
            'thermal_coeff', 'col_7',
        ],
        'NODELOADS': [
            'load_id', 'node_id', 'fx', 'fy', 'fz', 'mx', 'my', 'mz', 'col_8',
        ],
        'COMBINATIONS': ['combo_id', 'load_case_id', 'factor'],
        'TITLES': ['load_case_id', 'title'],
        'DISPLACEMENTS': [
            'load_case_id', 'node_id', 'dx', 'dy', 'dz', 'rx', 'ry', 'rz',
        ],
        'REACTIONS': [
            'load_case_id', 'node_id', 'fx', 'fy', 'fz', 'mx', 'my', 'mz',
        ],
        'MEMBER FORCES AND MOMENTS': [
            'load_case_id', 'member_id', 'col_0', 'fx', 'fy', 'fz', 'mx', 'my', 'mz',
            'col_1'
        ],
        # Add more section column mappings as needed...
    }

    def __init__(self, filepath: str):
        """
        Initialize and parse a SPACEGASS output file.

        Args:
            filepath: Path to the SPACEGASS text output file.
        """
        self.filepath = Path(filepath)
        self._units: Dict[str, str] = {}
        self._dataframes: Dict[str, pd.DataFrame] = {}

        # Parse the file
        self._parse_file()

    @property
    def units(self) -> Dict[str, str]:
        """Return the units dictionary."""
        return self._units

    # -------------------------------------------------------------------------
    # Each returns a DataFrame (or empty DataFrame if missing)
    # -------------------------------------------------------------------------

    @property
    def filters(self) -> pd.DataFrame:
        return self._dataframes.get('filters', pd.DataFrame())

    @property
    def nodes(self) -> pd.DataFrame:
        return self._dataframes.get('nodes', pd.DataFrame())

    @property
    def members(self) -> pd.DataFrame:
        return self._dataframes.get('members', pd.DataFrame())

    @property
    def restraints(self) -> pd.DataFrame:
        return self._dataframes.get('restraints', pd.DataFrame())

    @property
    def sections(self) -> pd.DataFrame:
        return self._dataframes.get('sections', pd.DataFrame())

    @property
    def materials(self) -> pd.DataFrame:
        return self._dataframes.get('materials', pd.DataFrame())

    @property
    def node_loads(self) -> pd.DataFrame:
        return self._dataframes.get('node_loads', pd.DataFrame())

    @property
    def member_forces(self) -> pd.DataFrame:
        return self._dataframes.get('member_forces', pd.DataFrame())

    @property
    def self_weight(self) -> pd.DataFrame:
        return self._dataframes.get('self_weight', pd.DataFrame())

    @property
    def harmonic_loads(self) -> pd.DataFrame:
        return self._dataframes.get('harmonic_loads', pd.DataFrame())

    @property
    def combinations(self) -> pd.DataFrame:
        return self._dataframes.get('combinations', pd.DataFrame())

    @property
    def titles(self) -> pd.DataFrame:
        return self._dataframes.get('titles', pd.DataFrame())

    @property
    def load_case_groups(self) -> pd.DataFrame:
        return self._dataframes.get('load_case_groups', pd.DataFrame())

    @property
    def load_categories(self) -> pd.DataFrame:
        return self._dataframes.get('load_categories', pd.DataFrame())

    @property
    def displacements(self) -> pd.DataFrame:
        return self._dataframes.get('displacements', pd.DataFrame())

    @property
    def member_forces_moments(self) -> pd.DataFrame:
        return self._dataframes.get('member_forces_moments', pd.DataFrame())

    @property
    def reactions(self) -> pd.DataFrame:
        return self._dataframes.get('reactions', pd.DataFrame())

    @property
    def member_int_displacements(self) -> pd.DataFrame:
        return self._dataframes.get('member_int_displacements', pd.DataFrame())

    @property
    def member_int_forces_moments(self) -> pd.DataFrame:
        return self._dataframes.get('member_int_forces_moments', pd.DataFrame())

    @property
    def member_stresses(self) -> pd.DataFrame:
        return self._dataframes.get('member_stresses', pd.DataFrame())

    # -------------------------------------------------------------------------
    # Query methods
    # -------------------------------------------------------------------------

    def query_forces_moments(
        self,
        load_case_id: Optional[Union[int, List[int]]] = None,
        member_id: Optional[Union[int, List[int]]] = None,
    ) -> pd.DataFrame:
        """
        Filter member forces and moments by load case and/or member ID.

        Args:
            load_case_id: Single load case ID or list of IDs to filter by.
            member_id: Single member ID or list of IDs to filter by.

        Returns:
            DataFrame filtered by the specified criteria. Returns all data
            if no filters are provided.
        """
        df = self.member_forces_moments

        if df.empty:
            return df

        # Build filter mask
        mask = pd.Series(True, index=df.index)

        if load_case_id is not None:
            if isinstance(load_case_id, int):
                mask &= df['load_case_id'] == load_case_id
            else:
                mask &= df['load_case_id'].isin(load_case_id)

        if member_id is not None:
            if isinstance(member_id, int):
                mask &= df['member_id'] == member_id
            else:
                mask &= df['member_id'].isin(member_id)

        return df[mask]

    # -------------------------------------------------------------------------
    # Parsing methods
    # -------------------------------------------------------------------------

    def _parse_file(self) -> None:
        """Parse the entire SPACEGASS output file."""
        # Use streaming approach to avoid loading entire file into memory
        raw_sections: Dict[str, List[str]] = {}
        current_section: Optional[str] = None
        current_lines: List[str] = []
        units_parsed = False

        # Use set for O(1) header lookup
        section_headers_set = set(self.SECTION_MAP.keys())

        try:
            with open(self.filepath, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f):
                    stripped = line.strip()

                    # Parse units from early in the file (usually line 14)
                    if not units_parsed and line_num < 20:
                        if stripped.startswith('UNITS '):
                            self._parse_units_line(stripped)
                            units_parsed = True
                            continue

                    # Skip empty lines and comments
                    if not stripped or stripped.startswith('#'):
                        continue

                    # Check if this line is a section header
                    if stripped in section_headers_set:
                        # Save previous section
                        if current_section is not None:
                            raw_sections[current_section] = current_lines
                        # Start new section
                        current_section = stripped
                        current_lines = []
                        continue

                    # Check for END marker
                    if stripped == 'END':
                        if current_section is not None:
                            raw_sections[current_section] = current_lines
                        break

                    # Add to current section
                    if current_section is not None:
                        current_lines.append(line)

        except FileNotFoundError:
            raise FileNotFoundError(f"SPACEGASS output file not found: {self.filepath}")
        except PermissionError:
            raise PermissionError(f"Permission denied reading file: {self.filepath}")
        except UnicodeDecodeError as e:
            raise ValueError(f"File contains invalid UTF-8 encoding: {self.filepath}") from e

        # Parse each section into a DataFrame
        for section_name, section_lines in raw_sections.items():
            attr_name = self.SECTION_MAP.get(section_name)
            if attr_name and section_lines:
                df = self._parse_section(section_name, section_lines)
                self._dataframes[attr_name] = df

    def _parse_units_line(self, line: str) -> None:
        """Parse a UNITS line into the units dictionary."""
        # Format: "UNITS LENGTH:m, SECTION:mm, STRENGTH:MPa, ..."
        units_str = line[6:]  # Remove "UNITS "
        pairs = units_str.split(',')
        for pair in pairs:
            pair = pair.strip()
            if ':' in pair:
                key, value = pair.split(':', 1)
                self._units[key.strip()] = value.strip()

    def _get_column_names(self, section_name: str, num_cols: int) -> List[str]:
        """
        Get column names for a section, using defined names where available.

        Args:
            section_name: Name of the section (e.g., 'NODES', 'MEMBERS')
            num_cols: Total number of columns needed

        Returns:
            List of column names, using defined names first then placeholders
        """
        defined_names = self.SECTION_COLUMNS.get(section_name, [])
        columns = []

        for i in range(num_cols):
            if i < len(defined_names):
                columns.append(defined_names[i])
            else:
                columns.append(f'col_{i}')

        return columns

    def _parse_section(self, section_name: str, lines: List[str]) -> pd.DataFrame:
        """
        Parse a section's lines into a DataFrame.

        Args:
            section_name: Name of the section (e.g., 'NODES', 'RESTRAINTS')
            lines: List of content lines for this section

        Returns:
            DataFrame with parsed data
        """
        if section_name in self.MULTILINE_SECTIONS:
            return self._parse_multiline_section(section_name, lines)
        else:
            return self._parse_simple_section(section_name, lines)

    def _parse_simple_section(self, section_name: str, lines: List[str]) -> pd.DataFrame:
        """Parse a section where each line is a single record."""
        if not lines:
            return pd.DataFrame()

        rows = []
        max_cols = 0

        for line in lines:
            parsed = self._parse_csv_line(line)
            if parsed:
                rows.append(parsed)
                max_cols = max(max_cols, len(parsed))

        if not rows:
            return pd.DataFrame()

        # Pad rows to have consistent column count
        for row in rows:
            row.extend([None] * (max_cols - len(row)))

        # Create DataFrame with column names (defined or placeholder)
        columns = self._get_column_names(section_name, max_cols)
        df = pd.DataFrame(rows, columns=columns)

        # Try to convert numeric columns
        return self._convert_numeric_columns(df)

    def _parse_multiline_section(self, section_name: str, lines: List[str]) -> pd.DataFrame:
        """Parse a section where each record spans multiple lines."""
        if not lines:
            return pd.DataFrame()

        lines_per_record = self.MULTILINE_SECTIONS[section_name]

        rows = []
        max_cols = 0
        i = 0

        while i < len(lines):
            # Collect lines for this record
            record_lines = []
            for j in range(lines_per_record):
                if i + j < len(lines):
                    record_lines.append(lines[i + j])

            if not record_lines:
                break

            # Parse main line
            main_parsed = self._parse_csv_line(record_lines[0])
            if not main_parsed:
                i += 1
                continue

            # For RESTRAINTS: verify this is a main record by checking for restraint code
            if section_name == 'RESTRAINTS':
                if len(main_parsed) < 2 or not self._is_restraint_code(main_parsed[1]):
                    # This is not a main record, skip this line
                    i += 1
                    continue

            # For SECTIONS: verify this is a main record by checking for quoted name
            if section_name == 'SECTIONS':
                if len(main_parsed) < 2 or not isinstance(main_parsed[1], str):
                    # This is not a main record, skip this line
                    i += 1
                    continue

            # Parse continuation lines and flatten into the row
            row = list(main_parsed)
            for cont_line in record_lines[1:]:
                cont_parsed = self._parse_csv_line(cont_line)
                if cont_parsed:
                    row.extend(cont_parsed)

            rows.append(row)
            max_cols = max(max_cols, len(row))
            i += lines_per_record

        if not rows:
            return pd.DataFrame()

        # Pad rows to have consistent column count
        for row in rows:
            row.extend([None] * (max_cols - len(row)))

        # Create DataFrame with column names (defined or placeholder)
        columns = self._get_column_names(section_name, max_cols)
        df = pd.DataFrame(rows, columns=columns)

        # Try to convert numeric columns
        return self._convert_numeric_columns(df)

    def _is_restraint_code(self, value) -> bool:
        """Check if a value is a valid restraint code (e.g., VRVRRR, VFVRRR)."""
        if not isinstance(value, str):
            return False
        # Restraint codes are 6 characters containing V, R, or F
        return bool(self._RESTRAINT_CODE_PATTERN.match(value.strip()))

    def _parse_csv_line(self, line: str) -> Optional[List]:
        """
        Parse a comma-separated line, handling quoted fields.

        Returns:
            List of parsed values, or None if line is empty/invalid
        """
        line = line.strip()
        if not line:
            return None

        try:
            # Fast path: no quotes - use simple split (much faster for large files)
            if '"' not in line:
                cleaned = [val.strip() or None for val in line.split(',')]
                return cleaned if cleaned else None

            # Slow path: has quotes - use csv module to handle properly
            reader = csv.reader(StringIO(line))
            row = next(reader)
            cleaned = [val.strip() or None for val in row]
            return cleaned if cleaned else None
        except (csv.Error, StopIteration):
            return None

    def _convert_numeric_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Try to convert columns to numeric types where possible.

        This improves memory usage and enables numeric operations.
        """
        if df.empty:
            return df

        # Build converted columns dict to avoid repeated DataFrame copies
        converted_cols = {}
        for col in df.columns:
            try:
                numeric_col = pd.to_numeric(df[col], errors='coerce')
                # Only convert if most values are numeric (>50%)
                non_null_count = numeric_col.notna().sum()
                original_non_null = df[col].notna().sum()
                if original_non_null > 0 and non_null_count / original_non_null > 0.5:
                    converted_cols[col] = numeric_col
                else:
                    converted_cols[col] = df[col]
            except (TypeError, ValueError):
                converted_cols[col] = df[col]

        return pd.DataFrame(converted_cols)

    def __repr__(self) -> str:
        """Return a string representation of the SGResults object."""
        sections_loaded = [name for name, df in self._dataframes.items() if not df.empty]
        return f"SGResults(filepath='{self.filepath}', sections={sections_loaded})"

    def summary(self) -> str:
        """Return a summary of all loaded sections and their sizes."""
        lines = [f"SGResults: {self.filepath}", ""]
        lines.append(f"Units: {self.units}")
        lines.append("")
        lines.append("Sections:")

        for section_name, attr_name in sorted(self.SECTION_MAP.items(), key=lambda x: x[1]):
            df = self._dataframes.get(attr_name, pd.DataFrame())
            if not df.empty:
                lines.append(f"  {attr_name}: {len(df)} rows x {len(df.columns)} cols")
            else:
                lines.append(f"  {attr_name}: (empty)")

        return "\n".join(lines)
