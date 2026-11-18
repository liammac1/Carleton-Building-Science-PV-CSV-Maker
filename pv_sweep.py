#!/usr/bin/env python3

from __future__ import annotations

import argparse
import concurrent.futures as cf
import csv
import datetime
import math
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field, asdict
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required:  pip install openpyxl")


# =============================================================================
# 1. ASSUMPTIONS - edit here (they are also editable later in the workbook)
# =============================================================================
# Each entry: key -> (value, unit, description, citation IDs [see CITATIONS])
ASSUMPTIONS = {
    "elec_rate": (0.3538, "$/kWh",
                  "Yellowknife all-in residential electricity rate, before the Feb-2026 increase", "[1]"),
    "customer_charge": (0.00, "$/month",
                        "Fixed monthly customer charge. 0 because the all-in rate already folds in "
                        "fixed charges; the separate charge is $18/month if you prefer to split it out.",
                        "[1], [2]"),
    "export_credit": (0.3538, "$/kWh",
                      "Value of each kWh of surplus electricity exported to the grid "
                      "(net metering at the retail rate)", "[1]"),
    "annual_credit_cap": ("Yes", "Yes/No",
                          "Yearly electricity cost cannot go below $0 (unused export credit is forfeited). "
                          "No = exports can earn net revenue.", "Modelling choice"),
    "propane_price": (1.08, "$/L",
                      "Residential heating propane price (Whitehorse, used as proxy for Yellowknife)", "[3]"),
    "propane_mj_per_l": (25.31, "MJ/L",
                         "Propane energy content (higher heating value)", "[4]"),
    "grid_ef": (0.18, "kg CO2e/kWh",
                "NWT grid electricity emission intensity", "[5]"),
    "propane_ef": (1.544, "kg CO2e/L",
                   "Propane combustion, residential: 1,515 g CO2 + 0.027 g CH4 + 0.108 g N2O per litre, "
                   "converted with AR5 GWPs (CH4 28, N2O 265)", "[6]"),
    "credit_exports_ghg": ("Yes", "Yes/No",
                           "Exported PV electricity counts as displacing grid emissions "
                           "(Yes = GHG on net grid electricity; No = GHG on purchased electricity only)",
                           "Modelling choice"),
}

# Numbered sources shown on the "Notes & Citations" sheet.
CITATIONS = [
    ("[1]", "Electricity rate 35.38 c/kWh (Yellowknife all-in, before 2026 increase)",
     "Cabin Radio, \"Wait, the bad news got less bad: your YK power rate hike might be 17%\". NTPC figure: "
     "cost of power for a Yellowknife resident 35.38 c/kWh before the proposed increase.",
     "https://cabinradio.ca/210157/news/yellowknife/wait-the-bad-news-got-less-bad-your-yk-power-rate-hike-might-be-17/"),
    ("[2]", "Customer charge $18.00/month",
     "Canada Energy Regulator, Canadian Residential Electricity Bill - Northwest Territories "
     "(NWT utilities charge an $18.00 monthly customer charge).",
     "https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/electricity/report/canadian-residential-electricity-bill/northwest-territories.html"),
    ("[3]", "Propane $1.08/L",
     "Yukon Bureau of Statistics, Yukon Fuel Price Survey, December 2024: Whitehorse propane "
     "(home heating) 108.0 c/L, including taxes.",
     "https://yukon.ca/sites/default/files/2024-12/fin-yukon-fuel-price-survey-december-2024.pdf"),
    ("[4]", "Propane 25.31 MJ/L",
     "Statistics Canada, Report on Energy Supply and Demand in Canada - Energy conversion factors "
     "(propane 25.31 TJ per megalitre).",
     "https://www150.statcan.gc.ca/n1/pub/57-003-x/2022001/dq-qd/dq-qd-4-eng.htm"),
    ("[5]", "Grid emission factor 0.18 kg CO2e/kWh",
     "Canada Energy Regulator, Provincial and Territorial Energy Profiles - Northwest Territories: "
     "NWT electricity emissions intensity 180 g CO2e/kWh in 2020 (from ECCC National Inventory Report).",
     "https://www.neb-one.gc.ca/nrg/ntgrtd/mrkt/nrgsstmprfls/nt-eng.html"),
    ("[6]", "Propane emission factors (g/L)",
     "ECCC, Emission Factors and Reference Values (Federal GHG Offset System), Table 3 - "
     "natural gas liquids: propane, residential 1515 g CO2, 0.027 g CH4, 0.108 g N2O per litre. "
     "GWPs from IPCC AR5 as used in Canada's NIR.",
     "https://www.canada.ca/en/environment-climate-change/services/climate-change/pricing-pollution-how-it-will-work/output-based-pricing-system/federal-greenhouse-gas-offset-system/emission-factors-reference-values.html"),
    ("[7]", "PV, inverter and load-center modelling",
     "EnergyPlus Input Output Reference and Engineering Reference (v26.1): Generator:Photovoltaic, "
     "PhotovoltaicPerformance:Simple, PhotovoltaicPerformance:EquivalentOne-Diode, "
     "ElectricLoadCenter:Inverter:Simple, ElectricLoadCenter:Distribution.",
     "https://energyplus.net/documentation"),
    ("[8]", "One-diode module parameters (mono_400, poly_280)",
     "Representative datasheet values for generic 400 W mono-Si (108 half-cell) and 280 W poly-Si "
     "(60-cell) modules; not a specific product. Replace with your chosen module's datasheet.",
     "-"),
    ("[9]", "Weather",
     "Canadian Weather for Energy Calculations (CWEC) file for Yellowknife, as passed to the script "
     "(file name recorded under 'Run information'). Available from climate.onebuilding.org.",
     "https://climate.onebuilding.org"),
]


# =============================================================================
# 2. PV MODULE LIBRARY (for the EquivalentOne-Diode model)
# =============================================================================
# Datasheet-style values. area_m2 = full module area.
MODULES = {
    "mono_400": dict(
        label="Generic mono-Si 400 W (108 half-cell)", cell_type="CrystallineSilicon",
        cells_series=54, area_m2=1.95, tau_alpha=0.95, bandgap=1.12, r_shunt=1_000_000,
        isc=13.70, voc=37.0, t_ref=25, g_ref=1000, imp=12.90, vmp=31.0,
        alpha_isc=0.0066, beta_voc=-0.100,             # A/K, V/K
        noct_tamb=20, noct_tcell=45, noct_g=800, u_loss=30, heat_cap=50000),
    "poly_280": dict(
        label="Generic poly-Si 280 W (60-cell)", cell_type="CrystallineSilicon",
        cells_series=60, area_m2=1.63, tau_alpha=0.95, bandgap=1.12, r_shunt=1_000_000,
        isc=9.40, voc=38.5, t_ref=25, g_ref=1000, imp=8.92, vmp=31.4,
        alpha_isc=0.0047, beta_voc=-0.119,
        noct_tamb=20, noct_tcell=46, noct_g=800, u_loss=30, heat_cap=50000),
}


# =============================================================================
# 3. SCENARIOS
# =============================================================================
@dataclass
class Scenario:
    name: str
    mount: str = "roof"            # roof | rack
    model: str = "simple"          # simple | onediode
    coverage: float = 0.85         # roof mount only
    efficiency: float = 0.18       # simple model only
    module: str = "mono_400"       # onediode only
    area_m2: float = 20.0          # rack + simple
    n_modules: int = 10            # rack + onediode
    tilt_deg: float = 45.0         # rack only (roof mount uses the roof's own tilt)
    azimuth_deg: float = 180.0     # rack only
    heat_transfer: str = "Decoupled"
    inverter_eff: float = 0.96
    # filled in when the IDF is generated
    pv_area_m2: float = 0.0
    rated_kw: float = 0.0
    layout_note: str = ""


def default_scenarios() -> list[Scenario]:
    s: list[Scenario] = []
    # A. Flat on the roof, simple model: area (coverage) x module efficiency
    for cov in (0.25, 0.50, 0.85):
        for eff in (0.15, 0.18, 0.21):
            s.append(Scenario(f"Roof {int(cov*100)}% cover, eff {int(eff*100)}%",
                              coverage=cov, efficiency=eff))
    # B. Flat on the roof, single-diode model with real module parameters
    for mod in ("mono_400", "poly_280"):
        for cov in (0.50, 0.85):
            s.append(Scenario(f"Roof {int(cov*100)}% cover, 1-diode {mod}",
                              model="onediode", module=mod, coverage=cov))
    # C. Single-diode model with the dynamic (thermal-mass) cell temperature model
    s.append(Scenario("Roof 85% cover, 1-diode mono_400, Uleberg cell temp",
                      model="onediode", module="mono_400", coverage=0.85,
                      heat_transfer="DecoupledUllebergDynamic"))
    # D. Tilted rack, south-facing: tilt sweep at fixed area
    for tilt in (15, 30, 45, 60, 75):
        s.append(Scenario(f"Rack 20 m2, tilt {tilt}", mount="rack",
                          area_m2=20, tilt_deg=tilt))
    # E. Tilted rack: area sweep at 45 deg
    for area in (10, 30):
        s.append(Scenario(f"Rack {area} m2, tilt 45", mount="rack",
                          area_m2=area, tilt_deg=45))
    # F. Tilted rack, single-diode modules
    s.append(Scenario("Rack 10 x mono_400, tilt 45, 1-diode", mount="rack",
                      model="onediode", module="mono_400", n_modules=10, tilt_deg=45))
    s.append(Scenario("Rack 15 x mono_400, tilt 60, 1-diode", mount="rack",
                      model="onediode", module="mono_400", n_modules=15, tilt_deg=60))
    return s


def load_scenarios_csv(path: Path) -> list[Scenario]:
    out = []
    types = {f.name: f.type for f in Scenario.__dataclass_fields__.values()}
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            kw = {}
            for k, v in row.items():
                k = (k or "").strip()
                v = (v or "").strip()
                if not k or v == "" or k not in types:
                    continue
                t = types[k]
                kw[k] = float(v) if t == "float" else int(float(v)) if t == "int" else v
            if "name" not in kw:
                raise ValueError(f"Scenario row without a name: {row}")
            out.append(Scenario(**kw))
    return out


# =============================================================================
# 4. MINIMAL IDF HANDLING (keeps the file's formatting and comments)
# =============================================================================
class IdfObject:
    def __init__(self, text: str):
        self.text = text
        code = "\n".join(line.split("!", 1)[0] for line in text.splitlines())
        code = code.strip().rstrip(";")
        self.fields = [f.strip() for f in code.split(",")]
        self.cls = self.fields[0].lower()

    @staticmethod
    def build(cls_name: str, fields: list, comments: list[str] | None = None) -> "IdfObject":
        lines = [f"{cls_name},"]
        for i, val in enumerate(fields):
            end = ";" if i == len(fields) - 1 else ","
            entry = f"    {val}{end}"
            if comments and i < len(comments) and comments[i]:
                entry = entry.ljust(29) + "!- " + comments[i]
            lines.append(entry)
        return IdfObject("\n".join(lines) + "\n")


def parse_idf(text: str) -> list:
    """Return a list whose items are either raw str (comments/blank lines) or IdfObject."""
    items, buf, between = [], [], []
    for line in text.splitlines(keepends=True):
        code = line.split("!", 1)[0]
        if not buf and not code.strip():
            between.append(line)
            continue
        if between:
            items.append("".join(between))
            between = []
        buf.append(line)
        if ";" in code:
            items.append(IdfObject("".join(buf)))
            buf = []
    if between:
        items.append("".join(between))
    if buf:
        raise ValueError("IDF ends inside an object (missing ';').")
    return items


def write_idf(items: list) -> str:
    return "".join(i if isinstance(i, str) else i.text for i in items)


PV_CLASSES = {
    "generator:photovoltaic", "generator:pvwatts",
    "photovoltaicperformance:simple", "photovoltaicperformance:equivalentone-diode",
    "photovoltaicperformance:sandia",
    "electricloadcenter:generators", "electricloadcenter:distribution",
    "electricloadcenter:inverter:simple", "electricloadcenter:inverter:pvwatts",
    "electricloadcenter:inverter:lookuptable", "electricloadcenter:inverter:functionofpower",
}
RACK_PREFIX = "PVRack_"

METERS = [
    "Electricity:Facility", "ElectricityPurchased:Facility", "ElectricitySurplusSold:Facility",
    "ElectricityNet:Facility", "ElectricityProduced:Facility", "Photovoltaic:ElectricityProduced",
    "Propane:Facility", "Heating:EnergyTransfer", "Cooling:EnergyTransfer",
]


def prepare_baseline(idf_text: str) -> tuple[list, dict]:
    """Strip PV, de-duplicate SimulationControl, add outputs. Returns (items, info)."""
    items = parse_idf(idf_text)
    info = {"removed": [], "warnings": [], "roof_surfaces": {}}
    kept, seen_simctrl = [], False
    for it in items:
        if isinstance(it, IdfObject):
            if it.cls in PV_CLASSES or it.cls in ("output:sqlite",):
                info["removed"].append(f"{it.fields[0]} {it.fields[1] if len(it.fields) > 1 else ''}")
                continue
            if it.cls == "shading:building:detailed" and it.fields[1].startswith(RACK_PREFIX):
                continue
            if it.cls == "output:meter" and it.fields[1] in METERS:
                continue
            if it.cls == "simulationcontrol":
                if seen_simctrl:
                    info["warnings"].append("Extra SimulationControl object removed (only one is allowed).")
                    continue
                seen_simctrl = True
                f = it.fields[1:]
                f += [""] * (7 - len(f))
                f[3] = "No"   # Run Simulation for Sizing Periods: not needed for annual results
                f[4] = "Yes"  # Run Simulation for Weather File Run Periods
                it = IdfObject.build("SimulationControl", f[:7], [
                    "Do Zone Sizing Calculation", "Do System Sizing Calculation",
                    "Do Plant Sizing Calculation", "Run Simulation for Sizing Periods",
                    "Run Simulation for Weather File Run Periods",
                    "Do HVAC Sizing Simulation for Sizing Periods",
                    "Maximum Number of HVAC Sizing Simulation Passes"])
            if it.cls == "building" and len(it.fields) > 2 and it.fields[2] not in ("", "0", "-0"):
                info["warnings"].append("Building north axis is not 0: rack azimuth is relative to building axes.")
            if (it.cls == "buildingsurface:detailed" and len(it.fields) > 11
                    and it.fields[2].lower() == "roof" and it.fields[6].lower() == "outdoors"):
                xyz = [float(v) for v in it.fields[12:] if v != ""]
                verts = [tuple(xyz[i:i + 3]) for i in range(0, len(xyz) - 2, 3)]
                info["roof_surfaces"][it.fields[1]] = verts
        kept.append(it)
    if not info["roof_surfaces"]:
        raise ValueError("No outdoor roof surfaces found in the IDF.")
    kept.append("\n\n! ---- Outputs added by pv_sweep.py ----\n")
    kept.append(IdfObject.build("Output:SQLite", ["SimpleAndTabular"]))
    for m in METERS:
        kept.append(IdfObject.build("Output:Meter", [m, "Timestep"]))
    return kept, info


# ---- geometry helpers --------------------------------------------------------
def polygon_area(verts) -> float:
    n = len(verts)
    cx = cy = cz = 0.0
    for i in range(n):
        x1, y1, z1 = verts[i]
        x2, y2, z2 = verts[(i + 1) % n]
        cx += y1 * z2 - z1 * y2
        cy += z1 * x2 - x1 * z2
        cz += x1 * y2 - y1 * x2
    return 0.5 * math.sqrt(cx * cx + cy * cy + cz * cz)


def rack_rows(roof: dict, area: float, tilt: float, az: float,
              setback=0.2, standoff=0.3, max_slope_len=2.0):
    """Lay out tilted rows of panels over the roof footprint. Returns list of 4-vertex rows."""
    pts = [v for verts in roof.values() for v in verts]
    xmin, xmax = min(p[0] for p in pts), max(p[0] for p in pts)
    ymin, ymax = min(p[1] for p in pts), max(p[1] for p in pts)
    z0 = max(p[2] for p in pts) + standoff
    cx, cy = (xmin + xmax) / 2, (ymin + ymax) / 2
    a = math.radians(az)
    f = (math.sin(a), math.cos(a))      # horizontal facing direction
    w = (-f[1], f[0])                   # along-row direction (to the viewer's right)
    corners = [(xmin, ymin), (xmin, ymax), (xmax, ymin), (xmax, ymax)]
    ext_f = max(abs((x - cx) * f[0] + (y - cy) * f[1]) for x, y in corners) * 2
    ext_w = max(abs((x - cx) * w[0] + (y - cy) * w[1]) for x, y in corners) * 2
    width, depth = ext_w - 2 * setback, ext_f - 2 * setback
    b = math.radians(tilt)
    rows = max(1, math.ceil(area / (width * max_slope_len)))
    slope = area / (rows * width)
    pitch = depth / rows
    if slope * math.cos(b) > pitch + 1e-9:
        raise ValueError(f"{area:.1f} m2 at {tilt} deg tilt does not fit on the roof "
                         f"({rows} rows x {width:.2f} m wide, {depth:.2f} m deep). "
                         "Reduce the area or raise the tilt.")

    def P(s, t, z):
        return (round(cx + f[0] * s + w[0] * t, 4), round(cy + f[1] * s + w[1] * t, 4), round(z, 4))

    out = []
    for i in range(rows):
        s_front = depth / 2 - i * pitch
        s_back = s_front - slope * math.cos(b)
        z_top = z0 + slope * math.sin(b)
        out.append([P(s_back, -width / 2, z_top), P(s_front, -width / 2, z0),
                    P(s_front, width / 2, z0), P(s_back, width / 2, z_top)])
    # sanity: outward normal must point up and toward the facing direction
    v = out[0]
    a1 = [v[1][k] - v[0][k] for k in range(3)]
    a2 = [v[2][k] - v[1][k] for k in range(3)]
    n = (a1[1] * a2[2] - a1[2] * a2[1], a1[2] * a2[0] - a1[0] * a2[2], a1[0] * a2[1] - a1[1] * a2[0])
    assert n[2] > -1e-9 and (tilt < 1 or n[0] * f[0] + n[1] * f[1] > 0), "rack normal check failed"
    note = f"{rows} row(s) x {width:.2f} m wide x {slope:.2f} m slope, pitch {pitch:.2f} m"
    return out, note


# ---- PV object generation ---------------------------------------------------
VALID_HT = {
    ("simple", "roof"): {"Decoupled"},
    ("simple", "rack"): {"Decoupled"},
    ("onediode", "roof"): {"Decoupled", "DecoupledUllebergDynamic"},
    ("onediode", "rack"): {"Decoupled", "DecoupledUllebergDynamic"},
}


def onediode_perf(name: str, m: dict) -> IdfObject:
    vals = [name, m["cell_type"], m["cells_series"], m["area_m2"], m["tau_alpha"], m["bandgap"],
            m["r_shunt"], m["isc"], m["voc"], m["t_ref"], m["g_ref"], m["imp"], m["vmp"],
            m["alpha_isc"], m["beta_voc"], m["noct_tamb"], m["noct_tcell"], m["noct_g"],
            m["u_loss"], m["heat_cap"]]
    names = ["Name", "Cell type", "Number of Cells in Series", "Active Area {m2}",
             "Transmittance Absorptance Product", "Semiconductor Bandgap {eV}", "Shunt Resistance {ohms}",
             "Short Circuit Current {A}", "Open Circuit Voltage {V}", "Reference Temperature {C}",
             "Reference Insolation {W/m2}", "Module Current at Maximum Power {A}",
             "Module Voltage at Maximum Power {V}", "Temperature Coefficient of Short Circuit Current {A/K}",
             "Temperature Coefficient of Open Circuit Voltage {V/K}",
             "NOCT Test Ambient Temperature {C}", "NOCT Test Cell Temperature {C}",
             "NOCT Test Insolation {W/m2}", "Module Heat Loss Coefficient {W/m2-K}",
             "Total Heat Capacity {J/m2-K}"]
    return IdfObject.build("PhotovoltaicPerformance:EquivalentOne-Diode", vals, names)


def scenario_objects(sc: Scenario, roof: dict) -> list:
    if sc.mount not in ("roof", "rack"):
        raise ValueError(f"mount must be roof or rack, got {sc.mount!r}")
    if sc.model not in ("simple", "onediode"):
        raise ValueError(f"model must be simple or onediode, got {sc.model!r}")
    if sc.heat_transfer not in VALID_HT[(sc.model, sc.mount)]:
        raise ValueError(f"heat_transfer {sc.heat_transfer!r} not valid for {sc.model}/{sc.mount}; "
                         f"use one of {sorted(VALID_HT[(sc.model, sc.mount)])}")
    objs: list = [f"\n! ---- PV system for scenario: {sc.name} ----\n"]
    hosts: list[tuple[str, float]] = []           # (surface name, gross area)
    if sc.mount == "rack":
        if sc.model == "onediode":
            mod = MODULES[sc.module]
            target_area = sc.n_modules * mod["area_m2"]
        else:
            target_area = sc.area_m2
        rows, note = rack_rows(roof, target_area, sc.tilt_deg, sc.azimuth_deg)
        sc.layout_note = note
        per_row_modules = None
        if sc.model == "onediode":
            per_row_modules = [sc.n_modules // len(rows) + (1 if i < sc.n_modules % len(rows) else 0)
                               for i in range(len(rows))]
        for i, verts in enumerate(rows):
            nm = f"{RACK_PREFIX}{i + 1}"
            flat = [c for v in verts for c in v]
            objs.append(IdfObject.build("Shading:Building:Detailed", [nm, "", 4] + flat))
            hosts.append((nm, polygon_area(verts)))
    else:
        for nm, verts in roof.items():
            hosts.append((nm, polygon_area(verts)))
        sc.layout_note = f"{len(hosts)} roof surfaces, {sc.coverage:.0%} covered"

    gens, total_area, total_w = [], 0.0, 0.0
    if sc.model == "simple":
        frac = 1.0 if sc.mount == "rack" else sc.coverage
        objs.append(IdfObject.build("PhotovoltaicPerformance:Simple",
                                    ["PVPerf", frac, "Fixed", sc.efficiency],
                                    ["Name", "Fraction of Surface Area with Active Solar Cells",
                                     "Conversion Efficiency Input Mode", "Value for Cell Efficiency if Fixed"]))
        for i, (srf, area) in enumerate(hosts):
            a_act = area * frac
            gens.append((f"PV_{i + 1}", srf, 1, a_act * sc.efficiency * 1000.0))
            total_area += a_act
    else:
        mod = MODULES[sc.module]
        objs.append(onediode_perf("PVPerf", mod))
        p_mod = mod["imp"] * mod["vmp"]
        for i, (srf, area) in enumerate(hosts):
            if sc.mount == "rack":
                n = per_row_modules[i]
            else:
                n = int(math.floor(sc.coverage * area / mod["area_m2"] + 1e-9))
            if n < 1:
                continue
            gens.append((f"PV_{i + 1}", srf, n, n * p_mod))
            total_area += n * mod["area_m2"]
        if not gens:
            raise ValueError("No module fits on any surface for this scenario.")
        sc.layout_note += f"; {sum(g[2] for g in gens)} modules"

    perf_cls = ("PhotovoltaicPerformance:Simple" if sc.model == "simple"
                else "PhotovoltaicPerformance:EquivalentOne-Diode")
    for gname, srf, n, watts in gens:
        objs.append(IdfObject.build("Generator:Photovoltaic",
                                    [gname, srf, perf_cls, "PVPerf", sc.heat_transfer, 1, n],
                                    ["Name", "Surface Name", "Photovoltaic Performance Object Type",
                                     "Module Performance Name", "Heat Transfer Integration Mode",
                                     "Number of Series Strings in Parallel", "Number of Modules in Series"]))
        total_w += watts
    glist = ["PV_Generators"]
    for gname, _, _, watts in gens:
        glist += [gname, "Generator:Photovoltaic", round(watts, 1), "", ""]
    objs.append(IdfObject.build("ElectricLoadCenter:Generators", glist[:-2]))
    objs.append(IdfObject.build("ElectricLoadCenter:Inverter:Simple",
                                ["PV_Inverter", "", "", 0, sc.inverter_eff],
                                ["Name", "Availability Schedule Name", "Zone Name",
                                 "Radiative Fraction", "Inverter Efficiency"]))
    objs.append(IdfObject.build("ElectricLoadCenter:Distribution",
                                ["PV_Distribution", "PV_Generators", "Baseload", "", "", "",
                                 "DirectCurrentWithInverter", "PV_Inverter"],
                                ["Name", "Generator List Name", "Generator Operation Scheme Type",
                                 "Demand Limit", "Track Schedule", "Track Meter",
                                 "Electrical Buss Type", "Inverter Name"]))
    sc.pv_area_m2 = total_area
    sc.rated_kw = total_w / 1000.0
    return objs


# =============================================================================
# 5. RUNNING ENERGYPLUS
# =============================================================================
def find_energyplus(user_path: str | None) -> str:
    cands = [user_path] if user_path else []
    cands += [shutil.which("energyplus"), shutil.which("energyplus.exe")]
    for root in ("/usr/local", "/Applications", "C:\\", "/opt", str(Path.home())):
        for pat in ("EnergyPlus*/energyplus", "EnergyPlus*/energyplus.exe", "EnergyPlusV*/energyplus.exe"):
            cands += sorted(str(p) for p in Path(root).glob(pat))
    for c in cands:
        if c and Path(c).is_file():
            return c
    sys.exit("Could not find EnergyPlus. Pass --energyplus /path/to/energyplus")


def run_one(exe: str, idf_text: str, epw: Path, rundir: Path) -> dict:
    rundir.mkdir(parents=True, exist_ok=True)
    (rundir / "in.idf").write_text(idf_text, encoding="utf-8")
    proc = subprocess.run([exe, "-w", str(epw), "-d", str(rundir), str(rundir / "in.idf")],
                          capture_output=True, text=True)
    err = (rundir / "eplusout.err").read_text(errors="replace") if (rundir / "eplusout.err").exists() else ""
    ok = proc.returncode == 0 and "EnergyPlus Completed Successfully" in err
    severe = [ln.strip() for ln in err.splitlines() if "** Severe  **" in ln or "**  Fatal  **" in ln]
    return {"ok": ok, "warnings": err.count("** Warning **"), "severe": severe,
            "sql": rundir / "eplusout.sql"}


# =============================================================================
# 6. READING RESULTS
# =============================================================================
J_PER_KWH = 3.6e6


def read_results(sql: Path) -> dict:
    con = sqlite3.connect(str(sql))
    q = """
        SELECT rdd.Name, t.Month, t.Day, t.Hour, t.Minute, t.Interval, rd.Value
        FROM ReportData rd
        JOIN ReportDataDictionary rdd ON rd.ReportDataDictionaryIndex = rdd.ReportDataDictionaryIndex
        JOIN Time t ON rd.TimeIndex = t.TimeIndex
        JOIN EnvironmentPeriods ep ON t.EnvironmentPeriodIndex = ep.EnvironmentPeriodIndex
        WHERE ep.EnvironmentType = 3 AND rdd.IsMeter = 1
          AND rdd.ReportingFrequency LIKE '%Timestep%'
    """
    annual = {m: 0.0 for m in METERS}
    monthly = {m: [0.0] * 12 for m in METERS}
    peak = {m: (0.0, "") for m in METERS}
    for name, mo, dy, hr, mi, interval, val in con.execute(q):
        if name not in annual:
            continue
        annual[name] += val
        monthly[name][mo - 1] += val
        watts = val / (interval * 60.0) if interval else 0.0
        if watts > peak[name][0]:
            h, m = (hr - 1, mi) if mi < 60 else (hr, 0)   # E+ stamps end of interval
            peak[name] = (watts, f"{mo:02d}/{dy:02d} {h:02d}:{m:02d}")
    con.close()
    return {"kwh": {k: v / J_PER_KWH for k, v in annual.items()},
            "monthly_kwh": {k: [x / J_PER_KWH for x in v] for k, v in monthly.items()},
            "peak_kw": {k: (v[0] / 1000.0, v[1]) for k, v in peak.items()}}


# =============================================================================
# 7. WORKBOOK
# =============================================================================
FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
BOLD = Font(name=FONT, bold=True)
NORMAL = Font(name=FONT)
GREEN = Font(name=FONT, color="008000")
HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
GROUP_FILL = PatternFill("solid", fgColor="B8CCE4")
BASE_FILL = PatternFill("solid", fgColor="F2F2F2")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_INT = '#,##0;(#,##0);"-"'
FMT_1 = '#,##0.0;(#,##0.0);"-"'
FMT_2 = '#,##0.00;(#,##0.00);"-"'
FMT_USD = '$#,##0;($#,##0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'


def _style_header(cell):
    cell.font, cell.fill, cell.border = BOLD, HEAD_FILL, BOX
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def build_workbook(rows: list[dict], monthly: list[dict], log: list[dict], out: Path, meta: dict):
    import re
    wb = Workbook()

    # ---------------- Assumptions ----------------
    wa = wb.active
    wa.title = "Assumptions"
    wa["A1"] = "Assumptions - edit the blue values; every cost, GHG and baseline comparison updates."
    wa["A1"].font = Font(name=FONT, bold=True, size=12)
    wa["A2"] = "Sources are listed on the 'Notes & Citations' sheet."
    wa["A2"].font = Font(name=FONT, italic=True)
    for c, h in enumerate(["Key", "Value", "Unit", "Description", "Source"], 1):
        _style_header(wa.cell(row=3, column=c, value=h))
    ref = {}
    for r, (key, (val, unit, desc, src)) in enumerate(ASSUMPTIONS.items(), start=4):
        wa.cell(row=r, column=1, value=key)
        v = wa.cell(row=r, column=2, value=val)
        v.font, v.fill = BLUE, INPUT_FILL
        if isinstance(val, float):
            v.number_format = "0.0000" if val < 1 else "0.00"
        wa.cell(row=r, column=3, value=unit)
        wa.cell(row=r, column=4, value=desc)
        wa.cell(row=r, column=5, value=src)
        for c in range(1, 6):
            cell = wa.cell(row=r, column=c)
            cell.border = BOX
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c != 2:
                cell.font = NORMAL
        ref[key] = f"Assumptions!$B${r}"
    for col, wdt in zip("ABCDE", (20, 11, 13, 80, 16)):
        wa.column_dimensions[col].width = wdt

    # ---------------- Results ----------------
    ws = wb.create_sheet("Results", 0)
    A = ref
    # (id, group, header, kind, value-key or formula, number format, header note)
    # Formulas reference other columns as [id] (this row) or [id@b] (baseline row).
    cols = [
        ("name", "Scenario", "Scenario", "v", "name", None, None),
        ("mount", "Scenario", "Mounting", "v", "mount", None, None),
        ("model", "Scenario", "PV model", "v", "model_label", None,
         "simple = PhotovoltaicPerformance:Simple; 1-diode = PhotovoltaicPerformance:EquivalentOne-Diode"),
        ("ht", "Scenario", "Heat transfer mode", "v", "heat_transfer", None, None),
        ("tilt", "Scenario", "Tilt (deg)", "v", "tilt", FMT_INT, "Roof mount follows the roof (flat = 0)."),
        ("layout", "Scenario", "Layout", "v", "layout_note", None, None),
        ("area", "Scenario", "PV module area (m2)", "v", "pv_area", FMT_1, None),
        ("kw", "Scenario", "Rated DC capacity (kW)", "v", "rated_kw", FMT_2,
         "STC rating: area x efficiency x 1 kW/m2, or modules x Vmp x Imp."),
        ("pv", "Annual energy use", "PV AC generation (kWh/yr)", "v", "pv_ac", FMT_INT,
         "ElectricityProduced:Facility (after inverter losses)."),
        ("yield", "Annual energy use", "Specific yield (kWh/kWp)", "f", "=IFERROR([pv]/[kw],0)", FMT_INT, None),
        ("elec", "Annual energy use", "Electricity use, gross (kWh/yr)", "v", "elec", FMT_INT,
         "Electricity:Facility - all electricity the house uses, whatever its source."),
        ("self", "Annual energy use", "PV used on site (kWh/yr)", "f", "=[pv]-[sold]", FMT_INT, None),
        ("purch", "Annual energy use", "Grid electricity purchased (kWh/yr)", "v", "purch", FMT_INT,
         "ElectricityPurchased:Facility"),
        ("netgrid", "Annual energy use", "Net grid electricity (kWh/yr)", "f", "=[purch]-[sold]", FMT_INT,
         "Purchased minus exported. Negative = net exporter over the year."),
        ("prop", "Annual energy use", "Propane (kWh/yr)", "v", "propane", FMT_INT, "Propane:Facility (heating)"),
        ("gross", "Annual energy use", "Total site energy, gross (kWh/yr)", "f", "=[elec]+[prop]", FMT_INT,
         "All energy consumed on site, before subtracting PV."),
        ("net", "Annual energy use", "Net site energy (kWh/yr)", "f", "=[netgrid]+[prop]", FMT_INT,
         "Gross site energy minus PV generation (EnergyPlus 'Net Site Energy')."),
        ("dnet", "Change vs baseline", "Change in net site energy (kWh/yr)", "f", "=[net]-[net@b]", FMT_INT,
         "Negative = less energy than baseline."),
        ("dnetp", "Change vs baseline", "Change in net site energy (%)", "f", "=IFERROR([dnet]/[net@b],0)",
         FMT_PCT, None),
        ("sold", "Electricity surplus for exportation", "Surplus electricity exported (kWh/yr)", "v", "sold",
         FMT_INT, "ElectricitySurplusSold:Facility - PV output not used in the house at that moment."),
        ("soldp", "Electricity surplus for exportation", "Share of PV generation exported (%)", "f",
         "=IFERROR([sold]/[pv],0)", FMT_PCT, None),
        ("soldv", "Electricity surplus for exportation", "Export credit value ($/yr)", "f",
         f"=[sold]*{A['export_credit']}", FMT_USD, "Before the annual credit cap is applied."),
        ("pkexp", "Electricity surplus for exportation", "Peak export to grid (kW)", "v", "pk_export", FMT_2,
         None),
        ("litres", "Annual energy cost", "Propane volume (L/yr)", "f", f"=[prop]*3.6/{A['propane_mj_per_l']}",
         FMT_INT, None),
        ("ecost", "Annual energy cost", "Electricity cost ($/yr)", "f",
         f"=IF({A['annual_credit_cap']}=\"Yes\",MAX(0,[purch]*{A['elec_rate']}-[soldv]),"
         f"[purch]*{A['elec_rate']}-[soldv])+12*{A['customer_charge']}", FMT_USD,
         "Purchased x rate - export credit + 12 x customer charge."),
        ("pcost", "Annual energy cost", "Propane / other fuels cost ($/yr)", "f",
         f"=[litres]*{A['propane_price']}", FMT_USD, None),
        ("tcost", "Annual energy cost", "Total energy cost ($/yr)", "f", "=[ecost]+[pcost]", FMT_USD, None),
        ("save", "Annual energy cost", "Cost savings vs baseline ($/yr)", "f", "=[tcost@b]-[tcost]", FMT_USD,
         None),
        ("eghg", "Annual GHG", "Electricity GHG (kg CO2e/yr)", "f",
         f"=IF({A['credit_exports_ghg']}=\"Yes\",[netgrid],[purch])*{A['grid_ef']}", FMT_INT, None),
        ("pghg", "Annual GHG", "Propane GHG (kg CO2e/yr)", "f", f"=[litres]*{A['propane_ef']}", FMT_INT, None),
        ("ghg", "Annual GHG", "Total GHG (t CO2e/yr)", "f", "=([eghg]+[pghg])/1000", FMT_2, None),
        ("dghg", "Annual GHG", "GHG reduction vs baseline (t CO2e/yr)", "f", "=[ghg@b]-[ghg]", FMT_2, None),
        ("dghgp", "Annual GHG", "GHG reduction vs baseline (%)", "f", "=IFERROR([dghg]/[ghg@b],0)", FMT_PCT,
         None),
        ("heat", "Heating & cooling", "Heating delivered (kWh/yr)", "v", "heat", FMT_INT,
         "Heating:EnergyTransfer"),
        ("dheat", "Heating & cooling", "Change in heating vs baseline (kWh/yr)", "f", "=[heat]-[heat@b]",
         FMT_INT, None),
        ("cool", "Heating & cooling", "Cooling delivered (kWh/yr)", "v", "cool", FMT_INT,
         "Cooling:EnergyTransfer (the model has no cooling equipment, so 0)."),
        ("dcool", "Heating & cooling", "Change in cooling vs baseline (kWh/yr)", "f", "=[cool]-[cool@b]",
         FMT_INT, None),
        ("pkh", "Peaks", "Peak heating (kW)", "v", "pk_heat", FMT_2, None),
        ("pkht", "Peaks", "Peak heating time (mm/dd hh:mm)", "v", "pk_heat_t", None, None),
        ("pkc", "Peaks", "Peak cooling (kW)", "v", "pk_cool", FMT_2, None),
        ("pke", "Peaks", "Peak building electricity demand (kW)", "v", "pk_elec", FMT_2, "Electricity:Facility"),
        ("pket", "Peaks", "Peak electricity time (mm/dd hh:mm)", "v", "pk_elec_t", None, None),
        ("pkg", "Peaks", "Peak grid electricity draw (kW)", "v", "pk_grid", FMT_2,
         "ElectricityPurchased:Facility"),
        ("dpkg", "Peaks", "Change in peak grid draw vs baseline (kW)", "f", "=[pkg]-[pkg@b]", FMT_2, None),
    ]
    letter = {c[0]: get_column_letter(i) for i, c in enumerate(cols, 1)}
    base_row = 3

    def expand(tpl: str, r: int) -> str:
        def sub(m):
            col, b = m.group(1), m.group(2)
            return f"{letter[col]}${base_row}" if b else f"{letter[col]}{r}"
        return re.sub(r"\[(\w+)(@b)?\]", sub, tpl)

    prev, start = None, 1
    for i, c in enumerate(cols + [(None,) * 7], 1):
        grp = c[1]
        if grp != prev:
            if prev is not None:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=i - 1)
                cell = ws.cell(row=1, column=start, value=prev)
                cell.font, cell.fill = BOLD, GROUP_FILL
                cell.alignment = Alignment(horizontal="center")
            prev, start = grp, i
        if grp is None:
            break
        cell = ws.cell(row=2, column=i, value=c[2])
        _style_header(cell)
        if c[6]:
            cell.comment = Comment(c[6], "pv_sweep")
    for n, rowdata in enumerate(rows):
        r = base_row + n
        for i, (cid, grp, h, kind, key, fmt, com) in enumerate(cols, 1):
            val = rowdata.get(key) if kind == "v" else expand(key, r)
            cell = ws.cell(row=r, column=i, value=val)
            cell.font = BOLD if (n == 0 and i == 1) else NORMAL
            cell.border = BOX
            if fmt:
                cell.number_format = fmt
            if n == 0:
                cell.fill = BASE_FILL
    ws.row_dimensions[2].height = 62
    widths = {"name": 44, "mount": 9, "model": 22, "ht": 22, "tilt": 7, "layout": 40}
    for c in cols:
        ws.column_dimensions[letter[c[0]]].width = widths.get(c[0], 14)
    ws.freeze_panes = "B3"
    ws.cell(row=base_row + len(rows) + 1, column=1,
            value="Grey row = baseline (no PV). Headers with a red corner have notes. Costs and GHG use the "
                  "Assumptions sheet; sources are on 'Notes & Citations'.").font = Font(name=FONT, italic=True)

    # ---------------- Monthly ----------------
    wm = wb.create_sheet("Monthly")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    r = 1
    for title, key in [("PV AC generation (kWh)", "pv"),
                       ("Surplus electricity exported (kWh)", "sold"),
                       ("Grid electricity purchased (kWh)", "purch"),
                       ("Heating delivered (kWh)", "heat")]:
        wm.cell(row=r, column=1, value=title).font = Font(name=FONT, bold=True, size=11)
        r += 1
        for c, h in enumerate(["Scenario"] + months + ["Year"], 1):
            _style_header(wm.cell(row=r, column=c, value=h))
        r += 1
        for m in monthly:
            wm.cell(row=r, column=1, value=m["name"]).font = NORMAL
            for c, v in enumerate(m[key], 2):
                cell = wm.cell(row=r, column=c, value=round(v, 1))
                cell.number_format, cell.font = FMT_INT, NORMAL
            cell = wm.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
            cell.number_format, cell.font = FMT_INT, NORMAL
            r += 1
        r += 1
    wm.column_dimensions["A"].width = 44
    for c in range(2, 15):
        wm.column_dimensions[get_column_letter(c)].width = 9
    wm.freeze_panes = "B1"

    # ---------------- Notes & Citations ----------------
    wn = wb.create_sheet("Notes & Citations")
    wn["A1"] = "Notes & Citations"
    wn["A1"].font = Font(name=FONT, bold=True, size=12)
    r = 3
    wn.cell(row=r, column=1, value="Citations").font = Font(name=FONT, bold=True, size=11)
    r += 1
    for c, h in enumerate(["ID", "Figure used", "Source", "Link"], 1):
        _style_header(wn.cell(row=r, column=c, value=h))
    for cid, fig, src, url in CITATIONS:
        r += 1
        for c, v in enumerate([cid, fig, src, url], 1):
            cell = wn.cell(row=r, column=c, value=v)
            cell.font, cell.border = NORMAL, BOX
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if url.startswith("http"):
            wn.cell(row=r, column=4).hyperlink = url
            wn.cell(row=r, column=4).font = Font(name=FONT, color="0563C1", underline="single")
    r += 2
    wn.cell(row=r, column=1, value="Method notes").font = Font(name=FONT, bold=True, size=11)
    notes = [
        "Baseline = the input model with every PV, inverter and electric load center object removed.",
        "Each scenario adds one PV system: generators, a PV performance object, a simple inverter and a "
        "DC load center with inverter [7].",
        "Roof mount: PV lies flat on the existing roof surfaces; 'coverage' is the fraction of each roof "
        "surface covered by modules. 'Decoupled' mode: the PV has no thermal effect on the roof.",
        "Rack mount: tilted rows of panels are modelled as shading surfaces 0.3 m above the roof, "
        "0.2 m in from the edges. They shade the roof (changing heating) and each other (row-to-row shading).",
        "Simple model: output = irradiance x area x fixed module efficiency x inverter efficiency; "
        "no temperature losses. 1-diode model: computes cell temperature and voltage/current from module "
        "datasheet values [8], so it captures temperature and low-light effects.",
        "Energy figures are site energy from EnergyPlus timestep meters over the weather-file run period. "
        "kWh = J / 3,600,000.",
        "Surplus electricity for exportation = PV output that exceeds the house's demand in the same "
        "timestep (ElectricitySurplusSold:Facility).",
        "Peaks are the highest single-timestep (10-minute) average rate in the year.",
        "Heating/cooling delivered = Heating:EnergyTransfer / Cooling:EnergyTransfer meters (energy the "
        "HVAC delivered to the zones). The model has heating only, so cooling is 0.",
        "Electricity cost = purchased x rate - exported x export credit (+ customer charge), floored at $0 "
        "per year when the annual credit cap is 'Yes'.",
        "GHG = electricity x grid factor + propane litres x propane factor. With 'credit exports' = Yes the "
        "grid factor applies to net grid electricity, so exports offset purchases.",
        "Limitations: EnergyPlus does not model snow cover on PV, so winter output at 62 N is likely "
        "overstated. The roof surfaces are non-convex, so shading from racks onto them is approximate.",
    ]
    for line in notes:
        r += 1
        cell = wn.cell(row=r, column=1, value="-")
        cell = wn.cell(row=r, column=2, value=line)
        wn.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        cell.font = NORMAL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        wn.row_dimensions[r].height = 30
    r += 2
    wn.cell(row=r, column=1, value="Run information").font = Font(name=FONT, bold=True, size=11)
    for k, v in [("Input IDF", meta["idf"]), ("Weather file", meta["epw"]), ("EnergyPlus", meta["exe"]),
                 ("Run date", meta["date"])]:
        r += 1
        wn.cell(row=r, column=1, value=k).font = BOLD
        wn.cell(row=r, column=2, value=v).font = NORMAL
    for col, wdt in zip("ABCD", (16, 38, 70, 60)):
        wn.column_dimensions[col].width = wdt

    # ---------------- Run log ----------------
    wl = wb.create_sheet("Run log")
    for c, h in enumerate(["Scenario", "Status", "Warnings", "Severe / fatal messages", "Run folder"], 1):
        _style_header(wl.cell(row=1, column=c, value=h))
    for r, e in enumerate(log, 2):
        for c, v in enumerate([e["name"], e["status"], e["warnings"], e["severe"], e["dir"]], 1):
            wl.cell(row=r, column=c, value=v).font = NORMAL
    for col, wdt in zip("ABCDE", (44, 10, 10, 90, 50)):
        wl.column_dimensions[col].width = wdt

    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    write_cached_values(wb, out)


# =============================================================================
# 7b. CACHED VALUES FOR FORMULAS
# =============================================================================
# openpyxl saves formulas without their results, so previews, phones, Google Drive
# viewers and Excel's Protected View show those cells blank. This computes every
# formula in Python and writes the result next to it. Excel still recalculates on
# open, so edits on the Assumptions sheet keep working.
import re as _re
import zipfile as _zip
from xml.sax.saxutils import escape as _esc

_TOKEN = _re.compile(r"""\s*(?:
    (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?)|
    (?P<str>"[^"]*")|
    (?P<ref>(?:(?:'[^']+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|
    (?P<func>[A-Z][A-Z0-9.]*)(?=\()|
    (?P<op><=|>=|<>|[-+*/(),=<>&]))""", _re.X)


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


class _FormulaEvaluator:
    def __init__(self, wb):
        self.wb, self.memo, self.busy = wb, {}, set()

    def value(self, sheet: str, coord: str):
        key = (sheet, coord)
        if key in self.memo:
            return self.memo[key]
        raw = self.wb[sheet][coord].value
        if isinstance(raw, str) and raw.startswith("="):
            if key in self.busy:
                raise ValueError(f"circular reference at {sheet}!{coord}")
            self.busy.add(key)
            try:
                val = self.evaluate(raw[1:], sheet)
            finally:
                self.busy.discard(key)
        else:
            val = raw
        self.memo[key] = val
        return val

    # --- tiny recursive-descent parser: returns closures so IF/IFERROR are lazy ---
    def evaluate(self, formula: str, sheet: str):
        toks, pos = [], 0
        while pos < len(formula):
            m = _TOKEN.match(formula, pos)
            if not m or m.end() == pos:
                if formula[pos:].strip() == "":
                    break
                raise ValueError(f"cannot parse formula: {formula}")
            kind = m.lastgroup
            toks.append((kind, m.group(kind)))
            pos = m.end()
        self.toks, self.i, self.sheet = toks, 0, sheet
        node = self._cmp()
        if self.i != len(toks):
            raise ValueError(f"cannot parse formula: {formula}")
        return node()

    def _peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def _take(self, expected=None):
        tok = self._peek()
        if expected and tok[1] != expected:
            raise ValueError(f"expected {expected!r}")
        self.i += 1
        return tok

    def _cmp(self):
        left = self._add()
        ops = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b, "<": lambda a, b: a < b,
               ">": lambda a, b: a > b, "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}
        if self._peek()[1] in ops:
            fn = ops[self._take()[1]]
            right = self._add()
            return lambda l=left, r=right: fn(l(), r())
        return left

    def _add(self):
        node = self._mul()
        while self._peek()[1] in ("+", "-"):
            op = self._take()[1]
            rhs = self._mul()
            node = (lambda a=node, b=rhs: _num(a()) + _num(b())) if op == "+" else \
                   (lambda a=node, b=rhs: _num(a()) - _num(b()))
        return node

    def _mul(self):
        node = self._unary()
        while self._peek()[1] in ("*", "/"):
            op = self._take()[1]
            rhs = self._unary()
            node = (lambda a=node, b=rhs: _num(a()) * _num(b())) if op == "*" else \
                   (lambda a=node, b=rhs: _num(a()) / _num(b()))
        return node

    def _unary(self):
        if self._peek()[1] == "-":
            self._take()
            inner = self._unary()
            return lambda: -_num(inner())
        if self._peek()[1] == "+":
            self._take()
        return self._primary()

    def _primary(self):
        kind, text = self._take()
        if kind == "num":
            v = float(text)
            return lambda: v
        if kind == "str":
            s = text[1:-1]
            return lambda: s
        if kind == "op" and text == "(":
            node = self._cmp()
            self._take(")")
            return node
        if kind == "ref":
            sheet, ref = (text.rsplit("!", 1) if "!" in text else (self.sheet, text))
            sheet = sheet.strip("'")
            ref = ref.replace("$", "")
            if ":" in ref:
                return lambda s=sheet, r=ref: self._range(s, r)
            return lambda s=sheet, r=ref: self.value(s, r)
        if kind == "func":
            self._take("(")
            args = []
            if self._peek()[1] != ")":
                args.append(self._cmp())
                while self._peek()[1] == ",":
                    self._take()
                    args.append(self._cmp())
            self._take(")")
            return self._func(text, args)
        raise ValueError(f"unexpected token {text!r}")

    def _range(self, sheet, ref):
        a, b = ref.split(":")
        ca, ra = _re.match(r"([A-Z]+)(\d+)", a).groups()
        cb, rb = _re.match(r"([A-Z]+)(\d+)", b).groups()
        out = []
        for r in range(int(ra), int(rb) + 1):
            for c in range(_col_num(ca), _col_num(cb) + 1):
                out.append(self.value(sheet, f"{get_column_letter(c)}{r}"))
        return out

    @staticmethod
    def _func(name, args):
        def flat(vals):
            for v in vals:
                if isinstance(v, list):
                    yield from (x for x in v if isinstance(x, (int, float)))
                else:
                    yield _num(v)
        if name == "IF":
            return lambda: (args[1]() if args[0]() else (args[2]() if len(args) > 2 else False))
        if name == "IFERROR":
            def f():
                try:
                    return args[0]()
                except (ZeroDivisionError, ValueError, TypeError):
                    return args[1]()
            return f
        if name == "MAX":
            return lambda: max(flat(a() for a in args))
        if name == "MIN":
            return lambda: min(flat(a() for a in args))
        if name == "SUM":
            return lambda: sum(flat(a() for a in args))
        raise ValueError(f"function {name} not supported by the value cache")


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    raise TypeError(f"not a number: {v!r}")


def write_cached_values(wb, path: Path) -> int:
    """Compute every formula in wb and store the results inside the saved file at path."""
    ev = _FormulaEvaluator(wb)
    per_sheet = []
    for ws in wb.worksheets:
        vals = {}
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    v = ev.value(ws.title, cell.coordinate)
                    if isinstance(v, float) and (v != v or v in (float("inf"), float("-inf"))):
                        continue
                    vals[cell.coordinate] = v
        per_sheet.append(vals)
    tmp = path.with_suffix(".tmp.xlsx")
    count = 0
    with _zip.ZipFile(path) as zin, _zip.ZipFile(tmp, "w", _zip.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = _re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if m and int(m.group(1)) <= len(per_sheet):
                vals = per_sheet[int(m.group(1)) - 1]
                text = data.decode("utf-8")

                def fill(mo):
                    nonlocal count
                    attrs, coord, formula = mo.group(1), mo.group(2), mo.group(3)
                    if coord not in vals:
                        return mo.group(0)
                    v = vals[coord]
                    count += 1
                    if isinstance(v, bool):
                        return f'<c{attrs} t="b"><f>{formula}</f><v>{int(v)}</v></c>'
                    if isinstance(v, (int, float)):
                        return f"<c{attrs}><f>{formula}</f><v>{repr(float(v))}</v></c>"
                    return f'<c{attrs} t="str"><f>{formula}</f><v>{_esc(str(v))}</v></c>'

                text = _re.sub(r'<c((?: (?!t=)[a-z]+="[^"]*")*? r="([A-Z]+\d+)"(?: (?!t=)[a-z]+="[^"]*")*)>'
                               r"<f>(.*?)</f><v ?/?>(?:</v>)?</c>", fill, text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)
    return count


# =============================================================================
# 8. MAIN
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--idf", required=True, type=Path)
    ap.add_argument("--epw", type=Path)
    ap.add_argument("--out", type=Path, default=Path("pv_results.xlsx"))
    ap.add_argument("--energyplus")
    ap.add_argument("--jobs", type=int, default=max(1, (os.cpu_count() or 2) - 1))
    ap.add_argument("--scenarios", type=Path, help="CSV of scenarios (replaces the built-in list)")
    ap.add_argument("--workdir", type=Path, help="Where run folders go (default: a temp folder)")
    ap.add_argument("--keep-runs", action="store_true")
    ap.add_argument("--list", action="store_true", help="Print scenarios and exit")
    args = ap.parse_args()

    scenarios = load_scenarios_csv(args.scenarios) if args.scenarios else default_scenarios()
    base_items, info = prepare_baseline(args.idf.read_text(encoding="utf-8", errors="replace"))
    for w in info["warnings"]:
        print("NOTE:", w)
    print(f"Roof surfaces found: {', '.join(info['roof_surfaces'])} "
          f"({sum(polygon_area(v) for v in info['roof_surfaces'].values()):.1f} m2)")
    if info["removed"]:
        print(f"Removed {len(info['removed'])} PV / load-center / output objects for the baseline.")

    jobs = [("Baseline (no PV)", None, write_idf(base_items))]
    for sc in scenarios:
        try:
            objs = scenario_objects(sc, info["roof_surfaces"])
        except (ValueError, KeyError) as e:
            print(f"SKIP {sc.name}: {e}")
            continue
        jobs.append((sc.name, sc, write_idf(base_items + objs)))

    if args.list:
        for name, sc, _ in jobs:
            print(f"- {name}" + (f"  [{sc.rated_kw:.2f} kW, {sc.pv_area_m2:.1f} m2, {sc.layout_note}]" if sc else ""))
        return
    if not args.epw:
        sys.exit("--epw is required to run simulations")

    exe = find_energyplus(args.energyplus)
    workdir = args.workdir or Path(tempfile.mkdtemp(prefix="pv_sweep_"))
    workdir.mkdir(parents=True, exist_ok=True)
    print(f"EnergyPlus: {exe}\nRunning {len(jobs)} simulations with {args.jobs} parallel jobs in {workdir}")

    results = {}
    with cf.ThreadPoolExecutor(max_workers=args.jobs) as pool:
        futs = {pool.submit(run_one, exe, text, args.epw.resolve(), workdir / f"run_{i:03d}"): (i, name)
                for i, (name, _, text) in enumerate(jobs)}
        for fut in cf.as_completed(futs):
            i, name = futs[fut]
            res = fut.result()
            results[i] = res
            print(f"  [{'ok ' if res['ok'] else 'FAIL'}] {name}  ({res['warnings']} warnings)")

    if not results[0]["ok"]:
        sys.exit(f"Baseline failed - see {workdir / 'run_000' / 'eplusout.err'}")

    rows, monthly, log = [], [], []
    for i, (name, sc, _) in enumerate(jobs):
        res = results[i]
        rundir = workdir / f"run_{i:03d}"
        log.append({"name": name, "status": "ok" if res["ok"] else "FAILED", "warnings": res["warnings"],
                    "severe": " | ".join(res["severe"])[:2000], "dir": str(rundir)})
        if not res["ok"]:
            continue
        d = read_results(res["sql"])
        k, pk, mk = d["kwh"], d["peak_kw"], d["monthly_kwh"]
        row = {
            "name": name,
            "mount": sc.mount if sc else "-",
            "model_label": ("-" if not sc else "simple, eff %.0f%%" % (sc.efficiency * 100) if sc.model == "simple"
                            else f"1-diode, {MODULES[sc.module]['label']}"),
            "heat_transfer": sc.heat_transfer if sc else "-",
            "tilt": (sc.tilt_deg if sc.mount == "rack" else 0) if sc else None,
            "layout_note": sc.layout_note if sc else "-",
            "pv_area": round(sc.pv_area_m2, 2) if sc else 0,
            "rated_kw": round(sc.rated_kw, 3) if sc else 0,
            "pv_ac": round(k["ElectricityProduced:Facility"], 1),
            "elec": round(k["Electricity:Facility"], 1),
            "purch": round(k["ElectricityPurchased:Facility"], 1),
            "sold": round(k["ElectricitySurplusSold:Facility"], 1),
            "propane": round(k["Propane:Facility"], 1),
            "heat": round(k["Heating:EnergyTransfer"], 1),
            "cool": round(k["Cooling:EnergyTransfer"], 1),
            "pk_heat": round(pk["Heating:EnergyTransfer"][0], 3),
            "pk_heat_t": pk["Heating:EnergyTransfer"][1] or "-",
            "pk_cool": round(pk["Cooling:EnergyTransfer"][0], 3),
            "pk_elec": round(pk["Electricity:Facility"][0], 3),
            "pk_elec_t": pk["Electricity:Facility"][1] or "-",
            "pk_grid": round(pk["ElectricityPurchased:Facility"][0], 3),
            "pk_export": round(pk["ElectricitySurplusSold:Facility"][0], 3),
        }
        rows.append(row)
        monthly.append({"name": name, "pv": mk["ElectricityProduced:Facility"],
                        "purch": mk["ElectricityPurchased:Facility"],
                        "sold": mk["ElectricitySurplusSold:Facility"], "heat": mk["Heating:EnergyTransfer"]})

    build_workbook(rows, monthly, log, args.out,
                   {"idf": str(args.idf), "epw": str(args.epw), "exe": exe,
                    "date": datetime.date.today().isoformat()})
    print(f"\nWrote {args.out}  ({len(rows)} successful runs of {len(jobs)})")

    if not args.keep_runs:
        for i, res in results.items():
            if res["ok"]:
                shutil.rmtree(workdir / f"run_{i:03d}", ignore_errors=True)
        if not any(workdir.iterdir()):
            workdir.rmdir()
        else:
            print(f"Failed run folders kept in {workdir}")


if __name__ == "__main__":
    main()
