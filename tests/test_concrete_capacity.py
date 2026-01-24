"""Tests for concrete_capacity module."""

from pathlib import Path

import pandas as pd
import pytest

from concrete_capacity import (
    AppliedLoads,
    ConcreteCapacityAnalyser,
    ConcreteProperties,
    ReinforcementLayer,
    SectionGeometry,
    SpreadsheetCells,
    UtilisationResult,
    VALID_BAR_SIZES,
    MAX_REINFORCEMENT_LAYERS,
)


class TestSectionGeometry:
    """Tests for SectionGeometry dataclass."""

    def test_valid_geometry(self):
        """Test creating valid geometry."""
        geom = SectionGeometry(depth=500, width=300)
        assert geom.depth == 500
        assert geom.width == 300

    def test_zero_depth_raises(self):
        """Test that zero depth raises ValueError."""
        with pytest.raises(ValueError, match="depth must be positive"):
            SectionGeometry(depth=0, width=300)

    def test_negative_depth_raises(self):
        """Test that negative depth raises ValueError."""
        with pytest.raises(ValueError, match="depth must be positive"):
            SectionGeometry(depth=-100, width=300)

    def test_zero_width_raises(self):
        """Test that zero width raises ValueError."""
        with pytest.raises(ValueError, match="width must be positive"):
            SectionGeometry(depth=500, width=0)

    def test_negative_width_raises(self):
        """Test that negative width raises ValueError."""
        with pytest.raises(ValueError, match="width must be positive"):
            SectionGeometry(depth=500, width=-100)

    def test_frozen(self):
        """Test that geometry is immutable."""
        geom = SectionGeometry(depth=500, width=300)
        with pytest.raises(AttributeError):
            geom.depth = 600


class TestConcreteProperties:
    """Tests for ConcreteProperties dataclass."""

    def test_valid_strength(self):
        """Test creating valid concrete properties."""
        concrete = ConcreteProperties(strength=40)
        assert concrete.strength == 40

    def test_typical_strengths(self):
        """Test common concrete strengths."""
        for strength in [25, 32, 40, 50, 65, 80]:
            concrete = ConcreteProperties(strength=strength)
            assert concrete.strength == strength

    def test_zero_strength_raises(self):
        """Test that zero strength raises ValueError."""
        with pytest.raises(ValueError, match="strength must be positive"):
            ConcreteProperties(strength=0)

    def test_negative_strength_raises(self):
        """Test that negative strength raises ValueError."""
        with pytest.raises(ValueError, match="strength must be positive"):
            ConcreteProperties(strength=-40)

    def test_excessive_strength_raises(self):
        """Test that excessively high strength raises ValueError."""
        with pytest.raises(ValueError, match="seems too high"):
            ConcreteProperties(strength=150)

    def test_boundary_strength(self):
        """Test boundary value of 100 MPa."""
        concrete = ConcreteProperties(strength=100)
        assert concrete.strength == 100


class TestReinforcementLayer:
    """Tests for ReinforcementLayer dataclass."""

    def test_valid_layer(self):
        """Test creating valid reinforcement layer."""
        layer = ReinforcementLayer(bar_size=20, spacings=(150, 200))
        assert layer.bar_size == 20
        assert layer.spacings == (150, 200)

    def test_all_valid_bar_sizes(self):
        """Test all valid Australian bar sizes."""
        for size in VALID_BAR_SIZES:
            layer = ReinforcementLayer(bar_size=size, spacings=(150,))
            assert layer.bar_size == size

    def test_invalid_bar_size_raises(self):
        """Test that invalid bar size raises ValueError."""
        with pytest.raises(ValueError, match="bar_size must be one of"):
            ReinforcementLayer(bar_size=15, spacings=(150,))

    def test_empty_spacings(self):
        """Test layer with no spacings."""
        layer = ReinforcementLayer(bar_size=20, spacings=())
        assert layer.spacings == ()

    def test_single_spacing(self):
        """Test layer with single spacing."""
        layer = ReinforcementLayer(bar_size=20, spacings=(150,))
        assert layer.spacings == (150,)

    def test_max_spacings(self):
        """Test layer with maximum allowed spacings."""
        spacings = tuple([150] * MAX_REINFORCEMENT_LAYERS)
        layer = ReinforcementLayer(bar_size=20, spacings=spacings)
        assert len(layer.spacings) == MAX_REINFORCEMENT_LAYERS

    def test_too_many_spacings_raises(self):
        """Test that too many spacings raises ValueError."""
        spacings = tuple([150] * (MAX_REINFORCEMENT_LAYERS + 1))
        with pytest.raises(ValueError, match="maximum .* layers allowed"):
            ReinforcementLayer(bar_size=20, spacings=spacings)

    def test_negative_spacing_raises(self):
        """Test that negative spacing raises ValueError."""
        with pytest.raises(ValueError, match="must be non-negative"):
            ReinforcementLayer(bar_size=20, spacings=(150, -100))

    def test_zero_spacing_allowed(self):
        """Test that zero spacing is allowed (no bars in that layer)."""
        layer = ReinforcementLayer(bar_size=20, spacings=(150, 0, 200))
        assert layer.spacings == (150, 0, 200)

    def test_from_spacings_factory(self):
        """Test from_spacings class method."""
        layer = ReinforcementLayer.from_spacings(20, 150, 200, 250)
        assert layer.bar_size == 20
        assert layer.spacings == (150, 200, 250)

    def test_frozen(self):
        """Test that layer is immutable."""
        layer = ReinforcementLayer(bar_size=20, spacings=(150,))
        with pytest.raises(AttributeError):
            layer.bar_size = 24


class TestAppliedLoads:
    """Tests for AppliedLoads dataclass."""

    def test_default_values(self):
        """Test that all loads default to zero."""
        loads = AppliedLoads()
        assert loads.fx == 0.0
        assert loads.fy == 0.0
        assert loads.fz == 0.0
        assert loads.mx == 0.0
        assert loads.my == 0.0
        assert loads.mz == 0.0

    def test_partial_loads(self):
        """Test creating loads with only some values."""
        loads = AppliedLoads(mz=250, fx=50)
        assert loads.mz == 250
        assert loads.fx == 50
        assert loads.fy == 0.0

    def test_all_loads(self):
        """Test creating loads with all values."""
        loads = AppliedLoads(fx=10, fy=20, fz=30, mx=40, my=50, mz=60)
        assert loads.fx == 10
        assert loads.fy == 20
        assert loads.fz == 30
        assert loads.mx == 40
        assert loads.my == 50
        assert loads.mz == 60

    def test_negative_loads(self):
        """Test that negative loads are valid."""
        loads = AppliedLoads(mz=-250, fx=-50)
        assert loads.mz == -250
        assert loads.fx == -50

    def test_from_series_lowercase(self):
        """Test creating loads from pandas Series with lowercase columns."""
        series = pd.Series({
            "fx": 10, "fy": 20, "fz": 30,
            "mx": 40, "my": 50, "mz": 60
        })
        loads = AppliedLoads.from_series(series)
        assert loads.fx == 10
        assert loads.mz == 60

    def test_from_series_uppercase(self):
        """Test creating loads from pandas Series with uppercase columns."""
        series = pd.Series({
            "Fx": 10, "Fy": 20, "Fz": 30,
            "Mx": 40, "My": 50, "Mz": 60
        })
        loads = AppliedLoads.from_series(series)
        assert loads.fx == 10
        assert loads.mz == 60

    def test_from_series_missing_columns(self):
        """Test that missing columns default to zero."""
        series = pd.Series({"mz": 100})
        loads = AppliedLoads.from_series(series)
        assert loads.mz == 100
        assert loads.fx == 0.0


class TestUtilisationResult:
    """Tests for UtilisationResult dataclass."""

    def test_create_result(self):
        """Test creating a utilisation result."""
        result = UtilisationResult(
            ultimate_utilisation=0.85,
            ultimate_strength=500,
            serviceability_stress=200,
            serviceability_utilisation=0.75,
        )
        assert result.ultimate_utilisation == 0.85
        assert result.ultimate_strength == 500

    def test_is_adequate_passing(self):
        """Test is_adequate returns True when both checks pass."""
        result = UtilisationResult(
            ultimate_utilisation=0.85,
            ultimate_strength=500,
            serviceability_stress=200,
            serviceability_utilisation=0.75,
        )
        assert result.is_adequate is True

    def test_is_adequate_ultimate_failing(self):
        """Test is_adequate returns False when ultimate fails."""
        result = UtilisationResult(
            ultimate_utilisation=1.05,
            ultimate_strength=500,
            serviceability_stress=200,
            serviceability_utilisation=0.75,
        )
        assert result.is_adequate is False

    def test_is_adequate_serviceability_failing(self):
        """Test is_adequate returns False when serviceability fails."""
        result = UtilisationResult(
            ultimate_utilisation=0.85,
            ultimate_strength=500,
            serviceability_stress=200,
            serviceability_utilisation=1.10,
        )
        assert result.is_adequate is False

    def test_is_adequate_boundary(self):
        """Test is_adequate at exactly 1.0."""
        result = UtilisationResult(
            ultimate_utilisation=1.0,
            ultimate_strength=500,
            serviceability_stress=200,
            serviceability_utilisation=1.0,
        )
        assert result.is_adequate is True


class TestSpreadsheetCells:
    """Tests for SpreadsheetCells configuration."""

    def test_default_cells(self):
        """Test default cell references."""
        cells = SpreadsheetCells()
        assert cells.depth == "D8"
        assert cells.width == "D9"
        assert cells.top_bar_size == "D15"
        assert cells.bottom_bar_size == "D16"

    def test_custom_cells(self):
        """Test custom cell references."""
        cells = SpreadsheetCells(depth="E10", width="E11")
        assert cells.depth == "E10"
        assert cells.width == "E11"
        # Others should still be default
        assert cells.top_bar_size == "D15"


class TestConcreteCapacityAnalyser:
    """Tests for ConcreteCapacityAnalyser class."""

    def test_init_with_missing_spreadsheet(self, tmp_path: Path):
        """Test that missing spreadsheet raises FileNotFoundError."""
        fake_path = tmp_path / "nonexistent.xlsm"
        with pytest.raises(FileNotFoundError, match="Spreadsheet not found"):
            ConcreteCapacityAnalyser(spreadsheet_path=fake_path)

    def test_repr(self, tmp_path: Path):
        """Test string representation."""
        # Create a dummy spreadsheet file
        spreadsheet = tmp_path / "test.xlsm"
        spreadsheet.touch()

        analyser = ConcreteCapacityAnalyser(spreadsheet_path=spreadsheet)
        repr_str = repr(analyser)
        assert "ConcreteCapacityAnalyser" in repr_str
        assert "test.xlsm" in repr_str

    def test_custom_cells(self, tmp_path: Path):
        """Test analyser with custom cell mapping."""
        spreadsheet = tmp_path / "test.xlsm"
        spreadsheet.touch()

        custom_cells = SpreadsheetCells(depth="A1", width="A2")
        analyser = ConcreteCapacityAnalyser(
            spreadsheet_path=spreadsheet,
            cells=custom_cells
        )
        assert analyser.cells.depth == "A1"


# Integration tests (require actual spreadsheet)
# These tests are skipped by default unless the spreadsheet exists

SPREADSHEET_PATH = Path(__file__).parent.parent / "src" / "RC Beam Design to AS3600 - 2018.xlsm"


@pytest.mark.skipif(
    not SPREADSHEET_PATH.exists(),
    reason="Spreadsheet not found for integration tests"
)
class TestConcreteCapacityAnalyserIntegration:
    """Integration tests requiring the actual spreadsheet."""

    @pytest.fixture
    def analyser(self) -> ConcreteCapacityAnalyser:
        """Create analyser with real spreadsheet."""
        return ConcreteCapacityAnalyser(spreadsheet_path=SPREADSHEET_PATH)

    @pytest.fixture
    def typical_inputs(self):
        """Typical input values for testing."""
        return {
            "geometry": SectionGeometry(depth=500, width=300),
            "concrete": ConcreteProperties(strength=40),
            "top_reo": ReinforcementLayer(bar_size=16, spacings=(200, 200)),
            "bottom_reo": ReinforcementLayer(bar_size=20, spacings=(150, 150)),
            "loads": AppliedLoads(mz=250, fx=50),
        }

    def test_calculate_returns_result(self, analyser, typical_inputs):
        """Test that calculate returns a UtilisationResult."""
        result = analyser.calculate(
            geometry=typical_inputs["geometry"],
            concrete=typical_inputs["concrete"],
            top_reinforcement=typical_inputs["top_reo"],
            bottom_reinforcement=typical_inputs["bottom_reo"],
            loads=typical_inputs["loads"],
        )
        assert isinstance(result, UtilisationResult)
        assert result.ultimate_utilisation > 0
        assert result.ultimate_strength > 0

    def test_calculate_with_negative_moment(self, analyser, typical_inputs):
        """Test calculation with negative moment (bars flipped)."""
        loads = AppliedLoads(mz=-250, fx=50)
        result = analyser.calculate(
            geometry=typical_inputs["geometry"],
            concrete=typical_inputs["concrete"],
            top_reinforcement=typical_inputs["top_reo"],
            bottom_reinforcement=typical_inputs["bottom_reo"],
            loads=loads,
        )
        assert isinstance(result, UtilisationResult)

    def test_calculate_batch(self, analyser, typical_inputs):
        """Test batch calculation."""
        loads_list = [
            AppliedLoads(mz=100),
            AppliedLoads(mz=200),
            AppliedLoads(mz=300),
        ]
        results = analyser.calculate_batch(
            geometry=typical_inputs["geometry"],
            concrete=typical_inputs["concrete"],
            top_reinforcement=typical_inputs["top_reo"],
            bottom_reinforcement=typical_inputs["bottom_reo"],
            loads_list=loads_list,
        )
        assert len(results) == 3
        assert all(isinstance(r, UtilisationResult) for r in results)
        # Higher moment should give higher utilisation
        assert results[0].ultimate_utilisation < results[2].ultimate_utilisation

    def test_calculate_batch_empty(self, analyser, typical_inputs):
        """Test batch calculation with empty list."""
        results = analyser.calculate_batch(
            geometry=typical_inputs["geometry"],
            concrete=typical_inputs["concrete"],
            top_reinforcement=typical_inputs["top_reo"],
            bottom_reinforcement=typical_inputs["bottom_reo"],
            loads_list=[],
        )
        assert results == []
