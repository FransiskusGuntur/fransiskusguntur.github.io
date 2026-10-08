# -*- coding: utf-8 -*-
"""SIwave 2024 R2 IronPython starter for PCB AC/PWM conductor loss.

RUN INSIDE STANDALONE SIWAVE: Tools > Scripting > Run Script.
This is not an AEDT oDesktop script and does not require PyAEDT or NumPy.
API signatures were checked against the official 2024 R2 help. The script
has not been run against an installed SIwave solver.

QUICK START
1. Open an imported, correctly classified and connected PCB in SIwave.
   Remove component models that are represented outside the PCB model.
2. Configure OUTPUT_DIR, SOURCES, GROUND_SOURCE and GROUND_TERMINAL below.
   Supply real source spectra; this script does not calculate or normalize FFTs.
3. Run MODE='setup'. A new project copy is saved and sources are placed.
4. In that copy, inspect source placement/polarity. Open Compute DC IR, select
   'Use sources defined in project', review DC temperature and solver options,
   and Save Settings. This selector is deliberately not automated: its DC
   scripting control was not verified. Record the generated project/run name.
5. Run MODE='solve' with the prepared copy open. If automatic report export
   says the simulation is not yet available, wait for completion in SIwave,
   set MODE='export', and run again using the same SIMULATION_NAME.

SOURCE FILE FORMAT (space-separated, no column header)
SOURCE_RESISTANCE <source resistance in ohms>
<frequency in Hz> <real current component> <imaginary current component>
All source files must have identical frequency grids. Include 0 Hz for DC.
Use spectra with a verified Ansys-compatible amplitude convention and shared
time reference. Do not assume raw FFT coefficients are RMS or peak phasors.

TERMINAL FORMAT
(0, part_name, reference_designator, pin_name)       # pin
(1, part_name, reference_designator, pin_group_name) # pin group
(2, x_string, y_string, metal_layer_name)            # coordinate
The coordinate form follows Ansys's IPY example. Verify layout units in SIwave.
Positive-to-negative is the current direction through the source.

SCOPE
Automates source placement, validation, adaptive DC mesh enablement,
power-data generation, solver launch and report export. The frequency-dependent
source workflow uses SYZ resistance inside the DC IR analysis; this script does
not substitute constant-R RMS loss or automate Ansys Circuit Push Excitations.
Spatial loss-map accuracy retains the limitations of the solver workflow.
"""

from __future__ import print_function
import math
import os
from datetime import datetime

# ---------------------------- Configuration ----------------------------
MODE = "setup"                         # "setup", "solve", or "export"
OUTPUT_DIR = r"C:\SIwave_AC_Loss"
SIMULATION_NAME = ""                    # populated in setup for this session;
                                        # copy printed name here across sessions
GROUND_SOURCE = ""                      # one configured source name
GROUND_TERMINAL = 1                     # 1 = negative, 2 = positive

# Replace this empty list with entries matching your board, for example:
# SOURCES = [{
#     "name": "I_INPUT",
#     "positive": (0, "YOUR_PART_NAME", "YOUR_REFDES", "YOUR_PIN"),
#     "negative": (0, "YOUR_PART_NAME", "YOUR_REFDES", "YOUR_RETURN_PIN"),
#     "spectrum": r"C:\YourData\input_current.txt",
# }]
# Those names are placeholders, not a suggested connection for your PCB.
SOURCES = []


def require_success_bool(result, operation):
    # Several SIwave commands use 1 for success, unlike solve/export commands.
    if not result:
        raise RuntimeError(operation + " failed (BOOL return).");


def require_success_code(result, operation):
    if result != 0:
        raise RuntimeError("%s failed (return code %s)." % (operation, result))


def finite(value):
    return not (math.isnan(value) or math.isinf(value))


def read_spectrum(path):
    """Check syntax/grid only; do not change or renormalize current data."""
    if not os.path.isfile(path):
        raise ValueError("Spectrum not found: " + path)
    frequencies = []
    has_ac = False
    resistance_seen = False
    with open(path, "r") as handle:
        for line_number, line in enumerate(handle, 1):
            fields = line.strip().split()
            if not fields:
                continue
            if fields[0] == "SOURCE_RESISTANCE":
                if resistance_seen or frequencies or len(fields) != 2:
                    raise ValueError("Bad SOURCE_RESISTANCE header: " + path)
                resistance = float(fields[1])
                if not finite(resistance) or resistance <= 0:
                    raise ValueError("Source resistance must be finite and positive.")
                resistance_seen = True
                continue
            if not resistance_seen or len(fields) != 3:
                raise ValueError("Expected Freq Re Im at %s:%s" % (path, line_number))
            frequency, real, imag = [float(item) for item in fields]
            if not all(finite(item) for item in (frequency, real, imag)):
                raise ValueError("Non-finite spectrum value: " + path)
            if frequency < 0 or (frequencies and frequency <= frequencies[-1]):
                raise ValueError("Frequencies must be nonnegative and strictly increasing.")
            if frequency == 0 and imag != 0:
                raise ValueError("DC current must have zero imaginary component.")
            frequencies.append(frequency)
            if frequency > 0 and (real != 0 or imag != 0):
                has_ac = True
    if not frequencies or frequencies[0] != 0:
        raise ValueError("Include the DC point, even if its current is zero: " + path)
    return frequencies, has_ac


def validate_configuration():
    if MODE not in ("setup", "solve", "export"):
        raise ValueError("MODE must be setup, solve or export.")
    if not SOURCES:
        raise ValueError("Configure SOURCES with your board terminals and current files.")
    names = [source["name"] for source in SOURCES]
    if any(not name for name in names) or len(set(names)) != len(names):
        raise ValueError("Source names must be nonempty and unique.")
    if GROUND_SOURCE not in names or GROUND_TERMINAL not in (1, 2):
        raise ValueError("Choose a configured GROUND_SOURCE and terminal 1 or 2.")
    common_grid = None
    any_ac = False
    for source in SOURCES:
        for key in ("positive", "negative"):
            terminal = source[key]
            if len(terminal) != 4 or terminal[0] not in (0, 1, 2):
                raise ValueError("Terminal must contain connection type and three strings.")
            if any(not str(value).strip() for value in terminal[1:]):
                raise ValueError("Terminal details cannot be empty.")
        grid, has_ac = read_spectrum(source["spectrum"])
        if common_grid is not None and grid != common_grid:
            raise ValueError("All sources must have identical frequency points.")
        common_grid = grid
        any_ac = any_ac or has_ac
    if not any_ac:
        raise ValueError("No nonzero AC current was supplied; this would be a DC-only case.")
    if not os.path.isdir(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)


def get_document():
    document = globals().get("oDoc")
    if document is None:
        application = globals().get("oApp")
        if application is None:
            raise RuntimeError("Run this file inside standalone SIwave, not regular Python.")
        document = application.GetActiveProject()
    if document is None:
        raise RuntimeError("Open the PCB project before running this script.")
    return document


def export_report(document, simulation_name):
    path = os.path.join(OUTPUT_DIR, simulation_name + ".html")
    code = document.ScrExportDcSimReport(simulation_name, "white", path)
    if code == 1:
        print("No completed result is available under that name yet.")
        print("After completion, run MODE='export' with SIMULATION_NAME='" + simulation_name + "'.")
        return
    require_success_code(code, "ScrExportDcSimReport")
    print("Report export requested: " + path)
    print("Verify completion and the output file in SIwave before using the result.")


def main():
    global SIMULATION_NAME
    validate_configuration()
    document = get_document()
    if MODE == "export":
        if not SIMULATION_NAME:
            raise ValueError("Set SIMULATION_NAME to the completed simulation name.")
        export_report(document, SIMULATION_NAME)
        return

    existing = set(str(name) for name in document.ScrGetComponentList("current sources"))
    desired = set(source["name"] for source in SOURCES)

    if MODE == "setup":
        if existing.intersection(desired):
            raise ValueError("Configured sources already exist. Inspect them and use solve mode.")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        SIMULATION_NAME = "AC_Loss_" + stamp
        # The official API expects a project filename WITHOUT its extension.
        project_basename = os.path.join(OUTPUT_DIR, SIMULATION_NAME)
        require_success_bool(document.ScrSaveProjectAs(project_basename), "Save project copy")
        for source in SOURCES:
            positive = source["positive"]
            negative = source["negative"]
            result = document.ScrPlaceFreqDependentSrc(
                source["name"], 4,
                positive[0], str(positive[1]), str(positive[2]), str(positive[3]),
                negative[0], str(negative[1]), str(negative[2]), str(negative[3]),
                os.path.abspath(source["spectrum"]))
            require_success_bool(result, "Place source " + source["name"])
    else:
        if not desired.issubset(existing):
            raise ValueError("Prepared project is missing one or more configured sources.")
        if not SIMULATION_NAME:
            raise ValueError("Set SIMULATION_NAME to the name printed by setup mode.")

    require_success_bool(document.ScrSetSimulationName("dc", SIMULATION_NAME), "Set name")
    require_success_bool(document.ScrSetIdealGroundNodeInDcSimulation(
        GROUND_SOURCE, GROUND_TERMINAL), "Set ideal ground")
    document.ScrSetRefineDcSimulation(1)
    document.ScrSetPlotAfterDcSimulation(True)
    document.ScrExportDcPowerDataToIcepak(True)
    document.Save()

    if MODE == "setup":
        print("Prepared project copy: " + project_basename + ".siw")
        print("SIMULATION_NAME = '" + SIMULATION_NAME + "'")
        print("Inspect terminal mapping and source polarity, and validate spectra normalization.")
        print("In Compute DC IR, select Use sources defined in project, review temperature,")
        print("and Save Settings. Then run MODE='solve' on this prepared copy.")
        return

    # Inspect all validation diagnostics in SIwave. This does not prove the
    # selected terminals, operating currents or spectral amplitudes are correct.
    validation = document.ScrRunValidationCheck()
    errors, warnings = int(validation[0]), int(validation[1])
    print("Validation: %s errors, %s warnings." % (errors, warnings))
    if errors:
        raise RuntimeError("Resolve validation errors before solving.")
    require_success_code(document.ScrRunDcSimulation(1), "ScrRunDcSimulation")
    document.Save()
    export_report(document, SIMULATION_NAME)


if __name__ == "__main__":
    main()

# Official 2024 R2 sources (append each filename to BASE_URL):
# BASE_URL = https://ansyshelp.ansys.com/public/Views/Secured/Electronics/v242/en/Subsystems/SIwave/Content/
# SIwaveScripting.htm; RunningaScript.htm; ScrPlaceFreqDependentSrc.htm;
# ScrSaveProjectAs.htm; ScrGetComponentList.htm; ScrSetSimulationName.htm;
# ScrSetIdealGroundNodeInDcSimulation.htm; ScrSetRefineDcSimulation.htm;
# ScrSetPlotAfterDcSimulation.htm; ScrExportDcPowerDataToIcepak.htm;
# ScrRunValidationCheck.htm; ScrRunDcSimulation.htm; ScrExportDcSimReport.htm;
# GetActiveProject.htm; Save.htm; electrothermalflowforpowerelectronics.htm.
