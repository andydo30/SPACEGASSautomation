"""
Concrete Section Capacity Analyser

A class for analysing reinforced concrete beam section capacity using
an Excel spreadsheet (RC Beam Design to AS3600 - 2018.xlsm).
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import xlwings as xw


# -----------------------------------------------------------------------------
# Input Dataclasses
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class SectionGeometry:
    """
    Concrete beam section geometry.

    Attributes:
        depth: Section depth in mm (must be positive)
        width: Section width in mm (must be positive)

    Example:
        >>> geometry = SectionGeometry(depth=500, width=300)
        >>> geometry.depth
        500
    """

    depth: float
    width: float

    def __post_init__(self) -> None:
        if self.depth <= 0:
            raise ValueError(f"depth must be positive, got {self.depth}")
        if self.width <= 0:
            raise ValueError(f"width must be positive, got {self.width}")


@dataclass(frozen=True)
class ConcreteProperties:
    """
    Concrete material properties.

    Attributes:
        strength: Characteristic compressive strength f'c in MPa (must be positive)

    Example:
        >>> concrete = ConcreteProperties(strength=40)
        >>> concrete.strength
        40
    """

    strength: float

    def __post_init__(self) -> None:
        if self.strength <= 0:
            raise ValueError(f"strength must be positive, got {self.strength}")
        if self.strength > 100:
            raise ValueError(
                f"strength {self.strength} MPa seems too high, expected <= 100 MPa"
            )


# Valid Australian standard bar sizes (diameter in mm)
VALID_BAR_SIZES = frozenset({10, 12, 16, 20, 24, 28, 32, 36, 40})

# Maximum number of reinforcement layers supported by spreadsheet
MAX_REINFORCEMENT_LAYERS = 5


@dataclass(frozen=True)
class ReinforcementLayer:
    """
    Reinforcement layer configuration.

    Attributes:
        bar_size: Bar diameter in mm (must be valid Australian standard size)
        spacings: List of bar spacings in mm for each layer (1-5 layers).
                  Use 0 or omit to indicate no bars in that layer.

    Examples:
        Create a single layer of N20 bars at 150mm spacing:

        >>> layer = ReinforcementLayer(bar_size=20, spacings=(150,))

        Create two layers with different spacings:

        >>> layer = ReinforcementLayer(bar_size=16, spacings=(200, 250))

        Use the convenience method for multiple layers:

        >>> layer = ReinforcementLayer.from_spacings(20, 150, 200)
    """

    bar_size: int
    spacings: tuple[float, ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        if self.bar_size not in VALID_BAR_SIZES:
            raise ValueError(
                f"bar_size must be one of {sorted(VALID_BAR_SIZES)}, got {self.bar_size}"
            )

        if len(self.spacings) > MAX_REINFORCEMENT_LAYERS:
            raise ValueError(
                f"maximum {MAX_REINFORCEMENT_LAYERS} spacing layers allowed, "
                f"got {len(self.spacings)}"
            )

        for i, spacing in enumerate(self.spacings):
            if spacing < 0:
                raise ValueError(
                    f"spacing[{i}] must be non-negative, got {spacing}"
                )

    @classmethod
    def from_spacings(cls, bar_size: int, *spacings: float) -> "ReinforcementLayer":
        """
        Create a ReinforcementLayer from variable spacing arguments.

        Example:
            layer = ReinforcementLayer.from_spacings(20, 150, 200)
        """
        return cls(bar_size=bar_size, spacings=tuple(spacings))


@dataclass(frozen=True)
class AppliedLoads:
    """
    Applied loads on the section.

    All forces in kN, moments in kNm.
    Sign convention follows SPACEGASS output.

    Attributes:
        fx: Axial force (positive = tension)
        fy: Shear force in y-direction
        fz: Shear force in z-direction
        mx: Torsional moment
        my: Bending moment about y-axis
        mz: Bending moment about z-axis (primary bending).
            Note: When mz == 0, the section is analysed with standard orientation
            (compression at top, tension at bottom).

    Examples:
        Create loads manually:

        >>> loads = AppliedLoads(mz=250, fx=50)

        Create from a pandas Series (e.g., SPACEGASS output row):

        >>> import pandas as pd
        >>> row = pd.Series({'fx': 10, 'fy': 5, 'fz': 20, 'mx': 0, 'my': 0, 'mz': 150})
        >>> loads = AppliedLoads.from_series(row)
    """

    fx: float = 0.0
    fy: float = 0.0
    fz: float = 0.0
    mx: float = 0.0
    my: float = 0.0
    mz: float = 0.0

    @classmethod
    def from_series(cls, series) -> "AppliedLoads":
        """
        Create AppliedLoads from a pandas Series (e.g., from SPACEGASS output).

        Expected columns: fx, fy, fz, mx, my, mz (case-insensitive)

        Example:
            >>> from sg_results import SGResults
            >>> results = SGResults('output.txt')
            >>> forces = results.query_forces_moments(load_case_id=1, member_id=5)
            >>> loads = AppliedLoads.from_series(forces.iloc[0])
        """
        # Build case-insensitive index lookup
        index_map = {str(idx).lower(): idx for idx in series.index}

        def get_value(name: str) -> float:
            key = name.lower()
            if key in index_map:
                return float(series[index_map[key]])
            return 0.0

        return cls(
            fx=get_value("fx"),
            fy=get_value("fy"),
            fz=get_value("fz"),
            mx=get_value("mx"),
            my=get_value("my"),
            mz=get_value("mz"),
        )


# -----------------------------------------------------------------------------
# Output Dataclass
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class UtilisationResult:
    """
    Results from capacity analysis.

    Attributes:
        ultimate_utilisation: Ratio of applied moment to ultimate capacity (<=1.0 is OK)
        ultimate_strength: Ultimate moment capacity in kNm
        serviceability_stress: Bottom bar stress under serviceability loads in MPa
        serviceability_utilisation: Serviceability utilisation ratio

    Example:
        >>> result = analyser.calculate(geometry, concrete, top_reo, bottom_reo, loads)
        >>> print(f"Ultimate utilisation: {result.ultimate_utilisation:.1%}")
        >>> print(f"Capacity: {result.ultimate_strength:.1f} kNm")
    """

    ultimate_utilisation: float
    ultimate_strength: float
    serviceability_stress: float
    serviceability_utilisation: float

    @property
    def is_adequate(self) -> bool:
        """
        Check if section passes both ultimate and serviceability checks.

        Example:
            >>> if result.is_adequate:
            ...     print("Section OK")
            ... else:
            ...     print("Section inadequate - increase reinforcement")
        """
        return self.ultimate_utilisation <= 1.0 and self.serviceability_utilisation <= 1.0


# -----------------------------------------------------------------------------
# Excel Cell Mapping
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class SpreadsheetCells:
    """
    Mapping of spreadsheet cell references.

    This allows the cell references to be customised if the spreadsheet
    layout changes.
    """

    # Section geometry
    depth: str = "D8"
    width: str = "D9"

    # Concrete properties
    concrete_strength: str = "D7"

    # Bar sizes
    top_bar_size: str = "D15"
    bottom_bar_size: str = "D16"

    # Top bar spacings (layers 1-5)
    top_spacings: tuple[str, ...] = ("K27", "K28", "K29", "K30", "K31")

    # Bottom bar spacings (layers 1-5)
    bottom_spacings: tuple[str, ...] = ("K34", "K35", "K36", "K37", "K38")

    # Applied loads
    ultimate_moment: str = "D33"
    axial_force: str = "D35"
    shear_force: str = "D36"
    torsion: str = "D37"
    serviceability_moment: str = "D38"

    # Results
    result_ultimate_utilisation: str = "L8"
    result_ultimate_strength: str = "J8"
    result_serviceability_stress: str = "J19"
    result_serviceability_utilisation: str = "L19"

    # Macro name
    solver_macro: str = "Solvefordn"


# Default cell mapping
DEFAULT_CELLS = SpreadsheetCells()


# -----------------------------------------------------------------------------
# Main Analyser Class
# -----------------------------------------------------------------------------


class ConcreteCapacityAnalyser:
    """
    Analyses reinforced concrete beam section capacity.

    Uses an Excel spreadsheet to perform the calculations according to AS3600-2018.

    Usage:
        analyser = ConcreteCapacityAnalyser()

        geometry = SectionGeometry(depth=500, width=300)
        concrete = ConcreteProperties(strength=40)
        top_reo = ReinforcementLayer(bar_size=16, spacings=(200, 200))
        bottom_reo = ReinforcementLayer(bar_size=20, spacings=(150, 150))
        loads = AppliedLoads(mz=250, fx=50)

        result = analyser.calculate(geometry, concrete, top_reo, bottom_reo, loads)
        print(f"Utilisation: {result.ultimate_utilisation:.2%}")

    Note on moment sign convention:
        - Positive mz: tension at bottom (standard beam configuration)
        - Negative mz: tension at top (hogging moment)
        - Zero mz: treated as positive moment (standard configuration)

    Attributes:
        spreadsheet_path: Path to the Excel spreadsheet file
        cells: Cell mapping configuration
    """

    # Default spreadsheet location (relative to this module)
    DEFAULT_SPREADSHEET = "RC Beam Design to AS3600 - 2018.xlsm"

    def __init__(
        self,
        spreadsheet_path: Optional[Path | str] = None,
        cells: Optional[SpreadsheetCells] = None,
    ) -> None:
        """
        Initialize the analyser.

        Args:
            spreadsheet_path: Path to the Excel spreadsheet. If None, uses
                              the default spreadsheet in the src directory.
            cells: Custom cell mapping. If None, uses default mapping.
        """
        if spreadsheet_path is None:
            # Default to spreadsheet in same directory as this module
            self.spreadsheet_path = Path(__file__).parent / self.DEFAULT_SPREADSHEET
        else:
            self.spreadsheet_path = Path(spreadsheet_path)

        if not self.spreadsheet_path.exists():
            raise FileNotFoundError(
                f"Spreadsheet not found: {self.spreadsheet_path}"
            )

        self.cells = cells or DEFAULT_CELLS

    def calculate(
        self,
        geometry: SectionGeometry,
        concrete: ConcreteProperties,
        top_reinforcement: ReinforcementLayer,
        bottom_reinforcement: ReinforcementLayer,
        loads: AppliedLoads,
    ) -> UtilisationResult:
        """
        Calculate section capacity utilisation.

        Args:
            geometry: Section dimensions
            concrete: Concrete material properties
            top_reinforcement: Top (compression) reinforcement configuration
            bottom_reinforcement: Bottom (tension) reinforcement configuration
            loads: Applied forces and moments

        Returns:
            UtilisationResult with capacity check results

        Raises:
            RuntimeError: If Excel calculation fails or returns invalid results

        Example:
            >>> analyser = ConcreteCapacityAnalyser()
            >>> geometry = SectionGeometry(depth=500, width=300)
            >>> concrete = ConcreteProperties(strength=40)
            >>> top_reo = ReinforcementLayer(bar_size=16, spacings=(200,))
            >>> bottom_reo = ReinforcementLayer(bar_size=20, spacings=(150, 150))
            >>> loads = AppliedLoads(mz=250, fx=50)
            >>> result = analyser.calculate(geometry, concrete, top_reo, bottom_reo, loads)
            >>> print(f"Utilisation: {result.ultimate_utilisation:.1%}")
        """
        workbook = None
        try:
            workbook = xw.Book(str(self.spreadsheet_path))
            sheet = workbook.sheets[0]

            # Set section geometry and concrete properties
            sheet[self.cells.depth].value = geometry.depth
            sheet[self.cells.width].value = geometry.width
            sheet[self.cells.concrete_strength].value = concrete.strength

            # Perform the calculation
            return self._perform_single_calculation(
                sheet, workbook, top_reinforcement, bottom_reinforcement, loads
            )
        finally:
            if workbook is not None:
                workbook.close()

    def calculate_batch(
        self,
        geometry: SectionGeometry,
        concrete: ConcreteProperties,
        top_reinforcement: ReinforcementLayer,
        bottom_reinforcement: ReinforcementLayer,
        loads_list: list[AppliedLoads],
    ) -> list[UtilisationResult]:
        """
        Calculate capacity for multiple load cases efficiently.

        Opens the workbook once and processes all load cases.

        Args:
            geometry: Section dimensions (same for all cases)
            concrete: Concrete properties (same for all cases)
            top_reinforcement: Top reinforcement (same for all cases)
            bottom_reinforcement: Bottom reinforcement (same for all cases)
            loads_list: List of load cases to analyse

        Returns:
            List of UtilisationResult, one per load case

        Raises:
            RuntimeError: If Excel calculation fails or returns invalid results

        Example:
            >>> # Load forces from SPACEGASS for multiple load cases
            >>> forces_df = results.query_forces_moments(member_id=5)
            >>> loads_list = [AppliedLoads.from_series(row) for _, row in forces_df.iterrows()]
            >>> results = analyser.calculate_batch(geometry, concrete, top_reo, bottom_reo, loads_list)
            >>> for i, result in enumerate(results):
            ...     print(f"Case {i+1}: {result.ultimate_utilisation:.1%}")
        """
        if not loads_list:
            return []

        results = []
        workbook = None

        try:
            workbook = xw.Book(str(self.spreadsheet_path))
            sheet = workbook.sheets[0]

            # Set geometry and concrete once (they don't change)
            sheet[self.cells.depth].value = geometry.depth
            sheet[self.cells.width].value = geometry.width
            sheet[self.cells.concrete_strength].value = concrete.strength

            # Process each load case
            for loads in loads_list:
                result = self._perform_single_calculation(
                    sheet, workbook, top_reinforcement, bottom_reinforcement, loads
                )
                results.append(result)

        finally:
            if workbook is not None:
                workbook.close()

        return results

    def _perform_single_calculation(
        self,
        sheet: xw.Sheet,
        workbook: xw.Book,
        top_reinforcement: ReinforcementLayer,
        bottom_reinforcement: ReinforcementLayer,
        loads: AppliedLoads,
    ) -> UtilisationResult:
        """
        Perform a single capacity calculation on an already-open workbook.

        This is the core calculation logic, extracted to avoid duplication
        between calculate() and calculate_batch().

        Args:
            sheet: The active worksheet
            workbook: The open workbook (needed for macro execution)
            top_reinforcement: Top reinforcement configuration
            bottom_reinforcement: Bottom reinforcement configuration
            loads: Applied forces and moments

        Returns:
            UtilisationResult with capacity check results
        """
        # Determine if bars need to be flipped based on moment direction.
        # Negative moment means tension is at top.
        # Zero moment is treated as positive (standard configuration).
        flip_bars = loads.mz < 0

        if flip_bars:
            tension_reo = top_reinforcement
            compression_reo = bottom_reinforcement
        else:
            tension_reo = bottom_reinforcement
            compression_reo = top_reinforcement

        # Set bar sizes (tension bar goes to D16, compression to D15)
        sheet[self.cells.top_bar_size].value = compression_reo.bar_size
        sheet[self.cells.bottom_bar_size].value = tension_reo.bar_size

        # Set bar spacings
        self._set_spacings(sheet, self.cells.top_spacings, compression_reo.spacings)
        self._set_spacings(sheet, self.cells.bottom_spacings, tension_reo.spacings)

        # Set applied loads (use absolute values)
        sheet[self.cells.ultimate_moment].value = abs(loads.mz)
        sheet[self.cells.serviceability_moment].value = abs(loads.mz)
        sheet[self.cells.axial_force].value = abs(loads.fx)
        sheet[self.cells.shear_force].value = abs(loads.fz)
        sheet[self.cells.torsion].value = abs(loads.mx)

        # Run the solver macro
        workbook.macro(self.cells.solver_macro)()

        # Read and validate results
        return self._read_results(sheet)

    def _set_spacings(
        self,
        sheet: xw.Sheet,
        cells: tuple[str, ...],
        spacings: tuple[float, ...],
    ) -> None:
        """Set spacing values in spreadsheet cells, padding with zeros."""
        for i, cell in enumerate(cells):
            if i < len(spacings):
                sheet[cell].value = spacings[i]
            else:
                sheet[cell].value = 0

    def _read_results(self, sheet: xw.Sheet) -> UtilisationResult:
        """
        Read and validate results from the spreadsheet.

        Args:
            sheet: The worksheet containing results

        Returns:
            UtilisationResult with validated values

        Raises:
            RuntimeError: If any result cell contains None or invalid data
        """
        ultimate_util = sheet[self.cells.result_ultimate_utilisation].value
        ultimate_str = sheet[self.cells.result_ultimate_strength].value
        service_stress = sheet[self.cells.result_serviceability_stress].value
        service_util = sheet[self.cells.result_serviceability_utilisation].value

        # Validate that all results are present
        results = {
            "ultimate_utilisation": ultimate_util,
            "ultimate_strength": ultimate_str,
            "serviceability_stress": service_stress,
            "serviceability_utilisation": service_util,
        }

        missing = [name for name, value in results.items() if value is None]
        if missing:
            raise RuntimeError(
                f"Excel calculation failed - result cells contain None: {missing}"
            )

        return UtilisationResult(
            ultimate_utilisation=float(ultimate_util),
            ultimate_strength=float(ultimate_str),
            serviceability_stress=float(service_stress),
            serviceability_utilisation=float(service_util),
        )

    def __repr__(self) -> str:
        return f"ConcreteCapacityAnalyser(spreadsheet_path='{self.spreadsheet_path}')"
