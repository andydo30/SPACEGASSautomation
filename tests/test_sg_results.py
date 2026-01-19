"""Tests for sg_results module."""

from pathlib import Path

import pandas as pd
import pytest

from sg_results import SGResults


class TestFileLoading:
    """Tests for file loading and error handling."""

    def test_load_minimal_file(self, minimal_file_path: Path):
        """Test loading a minimal SPACEGASS output file."""
        results = SGResults(str(minimal_file_path))
        assert results.filepath == minimal_file_path

    def test_load_example_file(self, example_file_path: Path):
        """Test loading the full example file."""
        results = SGResults(str(example_file_path))
        assert results.filepath == example_file_path

    def test_file_not_found(self, tmp_path: Path):
        """Test that FileNotFoundError is raised for missing files."""
        fake_path = tmp_path / "nonexistent.txt"
        with pytest.raises(FileNotFoundError):
            SGResults(str(fake_path))

    def test_repr(self, minimal_results: SGResults):
        """Test string representation."""
        repr_str = repr(minimal_results)
        assert "SGResults" in repr_str
        assert "filepath" in repr_str

    def test_summary(self, minimal_results: SGResults):
        """Test summary method returns string with section info."""
        summary = minimal_results.summary()
        assert "SGResults" in summary
        assert "Units" in summary
        assert "Sections" in summary


class TestUnitsParsing:
    """Tests for UNITS line parsing."""

    def test_units_parsed(self, minimal_results: SGResults):
        """Test that units are parsed from the file."""
        units = minimal_results.units
        assert isinstance(units, dict)
        assert len(units) > 0

    def test_units_values(self, minimal_results: SGResults):
        """Test specific unit values are correct."""
        units = minimal_results.units
        assert units.get("LENGTH") == "m"
        assert units.get("FORCE") == "kN"
        assert units.get("MOMENT") == "kNm"
        assert units.get("SECTION") == "mm"


class TestNodesParsing:
    """Tests for NODES section parsing."""

    def test_nodes_dataframe(self, minimal_results: SGResults):
        """Test that nodes property returns a DataFrame."""
        nodes = minimal_results.nodes
        assert isinstance(nodes, pd.DataFrame)
        assert not nodes.empty

    def test_nodes_columns(self, minimal_results: SGResults):
        """Test that nodes DataFrame has expected columns."""
        nodes = minimal_results.nodes
        assert "node_id" in nodes.columns
        assert "x" in nodes.columns
        assert "y" in nodes.columns
        assert "z" in nodes.columns

    def test_nodes_count(self, minimal_results: SGResults):
        """Test correct number of nodes parsed."""
        nodes = minimal_results.nodes
        assert len(nodes) == 4

    def test_nodes_values(self, minimal_results: SGResults):
        """Test specific node values."""
        nodes = minimal_results.nodes
        node_1 = nodes[nodes["node_id"] == 1].iloc[0]
        assert node_1["x"] == 0.0
        assert node_1["y"] == 0.0
        assert node_1["z"] == 0.0


class TestMembersParsing:
    """Tests for MEMBERS section parsing."""

    def test_members_dataframe(self, minimal_results: SGResults):
        """Test that members property returns a DataFrame."""
        members = minimal_results.members
        assert isinstance(members, pd.DataFrame)
        assert not members.empty

    def test_members_columns(self, minimal_results: SGResults):
        """Test that members DataFrame has expected columns."""
        members = minimal_results.members
        assert "member_id" in members.columns
        assert "node_i" in members.columns
        assert "node_j" in members.columns

    def test_members_count(self, minimal_results: SGResults):
        """Test correct number of members parsed."""
        members = minimal_results.members
        assert len(members) == 3


class TestMaterialsParsing:
    """Tests for MATERIALS section parsing."""

    def test_materials_dataframe(self, minimal_results: SGResults):
        """Test that materials property returns a DataFrame."""
        materials = minimal_results.materials
        assert isinstance(materials, pd.DataFrame)
        assert not materials.empty

    def test_materials_columns(self, minimal_results: SGResults):
        """Test that materials DataFrame has expected columns."""
        materials = minimal_results.materials
        assert "material_id" in materials.columns
        assert "name" in materials.columns
        assert "E" in materials.columns

    def test_materials_values(self, minimal_results: SGResults):
        """Test specific material values."""
        materials = minimal_results.materials
        mat_1 = materials[materials["material_id"] == 1].iloc[0]
        assert mat_1["name"] == "Concrete"
        assert mat_1["E"] == 32000


class TestSectionsParsing:
    """Tests for SECTIONS section parsing (multiline)."""

    def test_sections_dataframe(self, minimal_results: SGResults):
        """Test that sections property returns a DataFrame."""
        sections = minimal_results.sections
        assert isinstance(sections, pd.DataFrame)
        assert not sections.empty

    def test_sections_columns(self, minimal_results: SGResults):
        """Test that sections DataFrame has expected columns."""
        sections = minimal_results.sections
        assert "section_id" in sections.columns
        assert "name" in sections.columns

    def test_sections_values(self, minimal_results: SGResults):
        """Test specific section values."""
        sections = minimal_results.sections
        sec_1 = sections[sections["section_id"] == 1].iloc[0]
        assert sec_1["name"] == "Test Section"


class TestRestraintsParsing:
    """Tests for RESTRAINTS section parsing (multiline)."""

    def test_restraints_dataframe(self, minimal_results: SGResults):
        """Test that restraints property returns a DataFrame."""
        restraints = minimal_results.restraints
        assert isinstance(restraints, pd.DataFrame)
        assert not restraints.empty

    def test_restraints_columns(self, minimal_results: SGResults):
        """Test that restraints DataFrame has expected columns."""
        restraints = minimal_results.restraints
        assert "node_id" in restraints.columns
        assert "restraint_code" in restraints.columns

    def test_restraints_values(self, minimal_results: SGResults):
        """Test specific restraint values."""
        restraints = minimal_results.restraints
        rest_1 = restraints[restraints["node_id"] == 1].iloc[0]
        assert rest_1["restraint_code"] == "VVVRRR"


class TestTitlesParsing:
    """Tests for TITLES section parsing."""

    def test_titles_dataframe(self, minimal_results: SGResults):
        """Test that titles property returns a DataFrame."""
        titles = minimal_results.titles
        assert isinstance(titles, pd.DataFrame)
        assert not titles.empty

    def test_titles_count(self, minimal_results: SGResults):
        """Test correct number of titles parsed."""
        titles = minimal_results.titles
        assert len(titles) == 3

    def test_titles_values(self, minimal_results: SGResults):
        """Test specific title values."""
        titles = minimal_results.titles
        assert "Dead Load" in titles["title"].values
        assert "Live Load" in titles["title"].values


class TestDisplacementsParsing:
    """Tests for DISPLACEMENTS section parsing."""

    def test_displacements_dataframe(self, minimal_results: SGResults):
        """Test that displacements property returns a DataFrame."""
        displacements = minimal_results.displacements
        assert isinstance(displacements, pd.DataFrame)
        assert not displacements.empty

    def test_displacements_columns(self, minimal_results: SGResults):
        """Test that displacements DataFrame has expected columns."""
        displacements = minimal_results.displacements
        assert "load_case_id" in displacements.columns
        assert "node_id" in displacements.columns
        assert "dx" in displacements.columns
        assert "dy" in displacements.columns
        assert "dz" in displacements.columns


class TestReactionsParsing:
    """Tests for REACTIONS section parsing."""

    def test_reactions_dataframe(self, minimal_results: SGResults):
        """Test that reactions property returns a DataFrame."""
        reactions = minimal_results.reactions
        assert isinstance(reactions, pd.DataFrame)
        assert not reactions.empty

    def test_reactions_columns(self, minimal_results: SGResults):
        """Test that reactions DataFrame has expected columns."""
        reactions = minimal_results.reactions
        assert "load_case_id" in reactions.columns
        assert "node_id" in reactions.columns
        assert "fx" in reactions.columns
        assert "fy" in reactions.columns
        assert "fz" in reactions.columns


class TestMemberForcesMomentsParsing:
    """Tests for MEMBER FORCES AND MOMENTS section parsing."""

    def test_member_forces_moments_dataframe(self, minimal_results: SGResults):
        """Test that member_forces_moments property returns a DataFrame."""
        forces = minimal_results.member_forces_moments
        assert isinstance(forces, pd.DataFrame)
        assert not forces.empty

    def test_member_forces_moments_columns(self, minimal_results: SGResults):
        """Test that member_forces_moments DataFrame has expected columns."""
        forces = minimal_results.member_forces_moments
        assert "load_case_id" in forces.columns
        assert "member_id" in forces.columns


class TestQueryForcesMoments:
    """Tests for query_forces_moments method."""

    def test_query_all(self, minimal_results: SGResults):
        """Test query with no filters returns all data."""
        result = minimal_results.query_forces_moments()
        assert len(result) == len(minimal_results.member_forces_moments)

    def test_query_by_load_case(self, minimal_results: SGResults):
        """Test filtering by single load case."""
        result = minimal_results.query_forces_moments(load_case_id=1)
        assert all(result["load_case_id"] == 1)

    def test_query_by_member(self, minimal_results: SGResults):
        """Test filtering by single member."""
        result = minimal_results.query_forces_moments(member_id=1)
        assert all(result["member_id"] == 1)

    def test_query_by_multiple_load_cases(self, minimal_results: SGResults):
        """Test filtering by multiple load cases."""
        result = minimal_results.query_forces_moments(load_case_id=[1, 2])
        assert all(result["load_case_id"].isin([1, 2]))


class TestQueryMemberSections:
    """Tests for query_member_sections method."""

    def test_query_all_members(self, minimal_results: SGResults):
        """Test query with no filters returns all members with sections."""
        result = minimal_results.query_member_sections()
        # Should have same number of rows as members
        assert len(result) == len(minimal_results.members)
        # Should have columns from both tables
        assert 'member_id' in result.columns
        assert 'section_id' in result.columns
        assert 'name' in result.columns  # from sections
        assert 'area' in result.columns  # from sections

    def test_query_single_member(self, minimal_results: SGResults):
        """Test filtering by single member ID."""
        result = minimal_results.query_member_sections(member_id=1)
        assert len(result) == 1
        assert result.iloc[0]['member_id'] == 1

    def test_query_multiple_members(self, minimal_results: SGResults):
        """Test filtering by multiple member IDs."""
        result = minimal_results.query_member_sections(member_id=[1, 2])
        assert len(result) == 2
        assert set(result['member_id']) == {1, 2}

    def test_query_invalid_member_warns(self, minimal_results: SGResults):
        """Test that invalid member IDs produce warnings."""
        with pytest.warns(UserWarning, match="Member IDs not found"):
            result = minimal_results.query_member_sections(member_id=[1, 999])
        # Should still return results for valid ID
        assert len(result) == 1
        assert result.iloc[0]['member_id'] == 1

    def test_query_all_invalid_returns_empty(self, minimal_results: SGResults):
        """Test that all invalid IDs returns empty DataFrame."""
        with pytest.warns(UserWarning):
            result = minimal_results.query_member_sections(member_id=[999, 1000])
        assert len(result) == 0
        assert isinstance(result, pd.DataFrame)

    def test_section_properties_included(self, minimal_results: SGResults):
        """Test that section properties are properly joined."""
        result = minimal_results.query_member_sections(member_id=1)
        # Check specific section properties exist
        assert 'area' in result.columns
        assert 'Ixx' in result.columns
        assert 'Iyy' in result.columns
        # Verify values are present (not NaN)
        assert pd.notna(result.iloc[0]['area'])

    def test_query_with_tuple_input(self, minimal_results: SGResults):
        """Test that tuple input works like list input."""
        result = minimal_results.query_member_sections(member_id=(1, 2))
        assert len(result) == 2
        assert set(result['member_id']) == {1, 2}

    def test_string_input_raises_error(self, minimal_results: SGResults):
        """Test that string input raises TypeError."""
        with pytest.raises(TypeError, match="must be an int or list of ints"):
            minimal_results.query_member_sections(member_id="1")

    def test_empty_members_returns_empty(self, minimal_results: SGResults):
        """Test that empty members table returns empty DataFrame."""
        # Clear members DataFrame
        minimal_results._dataframes['members'] = pd.DataFrame()
        result = minimal_results.query_member_sections()
        assert isinstance(result, pd.DataFrame)
        assert result.empty

    def test_empty_sections_returns_members_only(self, minimal_results: SGResults):
        """Test that empty sections returns members without section properties."""
        # Clear sections DataFrame
        minimal_results._dataframes['sections'] = pd.DataFrame()
        result = minimal_results.query_member_sections(member_id=1)
        assert len(result) == 1
        assert 'member_id' in result.columns
        # Section columns should not be present
        assert 'area' not in result.columns


class TestEmptySections:
    """Tests for empty/missing section handling."""

    def test_empty_section_returns_empty_dataframe(self, minimal_results: SGResults):
        """Test that missing sections return empty DataFrames."""
        # harmonic_loads is not in the minimal fixture
        harmonic = minimal_results.harmonic_loads
        assert isinstance(harmonic, pd.DataFrame)
        assert harmonic.empty


class TestExampleFileIntegration:
    """Integration tests using the full example file."""

    def test_example_loads_successfully(self, example_results: SGResults):
        """Test that the full example file loads without errors."""
        assert example_results is not None

    def test_example_has_nodes(self, example_results: SGResults):
        """Test that example file has nodes parsed."""
        assert not example_results.nodes.empty
        assert len(example_results.nodes) > 10

    def test_example_has_members(self, example_results: SGResults):
        """Test that example file has members parsed."""
        assert not example_results.members.empty
        assert len(example_results.members) > 10

    def test_example_has_sections(self, example_results: SGResults):
        """Test that example file has sections parsed."""
        assert not example_results.sections.empty

    def test_example_has_materials(self, example_results: SGResults):
        """Test that example file has materials parsed."""
        assert not example_results.materials.empty

    def test_example_units_complete(self, example_results: SGResults):
        """Test that example file has all expected unit types."""
        units = example_results.units
        expected_keys = ["LENGTH", "FORCE", "MOMENT", "SECTION"]
        for key in expected_keys:
            assert key in units
