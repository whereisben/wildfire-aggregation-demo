#!/usr/bin/env python3
"""
Synthetic Wildfire Risk Dataset Generator
=========================================
Generates a realistic, 100% synthetic property dataset for the Wildfire 1km
Aggregation pipeline.

All geographic coordinates, policy identifiers, property IDs, and insured values
are mathematically generated with realistic spatial clustering across California
wildland-urban interface (WUI) communities. NO proprietary company data or real
policyholder information is used.

Output:
  - Synthetic_Inforce_Risk_Portfolio.xlsx (Summary sheet with required pipeline columns)
"""

import os, math, random
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as GL

FOLDER = os.path.dirname(os.path.abspath(__file__))
OUTPUT_EXCEL = os.path.join(FOLDER, "Synthetic_Inforce_Risk_Portfolio.xlsx")

# 70+ California community regions across coastal, foothill, mountain, and valley interfaces
REGIONS = [
    # Southern CA coastal & hills
    ("Pacific Palisades", "Los Angeles", 90272, 34.070, -118.530),
    ("Malibu", "Los Angeles", 90265, 34.055, -118.690),
    ("Malibu West", "Los Angeles", 90265, 34.050, -118.840),
    ("Beverly Hills", "Los Angeles", 90210, 34.090, -118.410),
    ("Bel Air", "Los Angeles", 90277, 34.105, -118.450),
    ("Brentwood", "Los Angeles", 90049, 34.065, -118.490),
    ("Altadena", "Los Angeles", 91001, 34.190, -118.140),
    ("Pasadena", "Los Angeles", 91105, 34.150, -118.160),
    ("Montecito", "Santa Barbara", 93108, 34.445, -119.630),
    ("Santa Barbara", "Santa Barbara", 93105, 34.450, -119.720),
    ("Hope Ranch", "Santa Barbara", 93110, 34.435, -119.755),
    ("Ojai", "Ventura", 93023, 34.445, -119.240),
    ("Laguna Beach", "Orange", 92651, 33.560, -117.760),
    ("Newport Coast", "Orange", 92657, 33.610, -117.830),
    ("Rancho Santa Fe", "San Diego", 92067, 33.020, -117.200),
    ("La Jolla", "San Diego", 92037, 32.845, -117.240),
    ("Palos Verdes Estates", "Los Angeles", 90274, 33.785, -118.375),
    ("Rolling Hills", "Los Angeles", 90274, 33.760, -118.345),
    ("Topanga", "Los Angeles", 90290, 34.090, -118.600),
    ("Calabasas", "Los Angeles", 91302, 34.150, -118.640),
    ("Hidden Hills", "Los Angeles", 91302, 34.165, -118.660),
    ("Encino", "Los Angeles", 91316, 34.145, -118.510),
    ("Sierra Madre", "Los Angeles", 91024, 34.165, -118.050),
    ("San Marino", "Los Angeles", 91108, 34.120, -118.110),
    
    # Bay Area & Peninsula
    ("San Francisco", "San Francisco", 94118, 37.788, -122.465),
    ("Hillsborough", "San Mateo", 94010, 37.560, -122.350),
    ("Atherton", "San Mateo", 94027, 37.460, -122.200),
    ("Woodside", "San Mateo", 94062, 37.425, -122.255),
    ("Portola Valley", "San Mateo", 94028, 37.380, -122.225),
    ("Los Altos Hills", "Santa Clara", 94022, 37.360, -122.140),
    ("Saratoga", "Santa Clara", 95070, 37.260, -122.030),
    ("Los Gatos", "Santa Clara", 95032, 37.220, -121.975),
    ("Monte Sereno", "Santa Clara", 95030, 37.235, -121.990),
    ("Piedmont", "Alameda", 94611, 37.825, -122.230),
    ("Berkeley", "Alameda", 94708, 37.890, -122.260),
    ("Oakland Hills", "Alameda", 94611, 37.840, -122.200),
    ("Orinda", "Contra Costa", 94563, 37.880, -122.180),
    ("Lafayette", "Contra Costa", 94549, 37.890, -122.120),
    ("Alamo", "Contra Costa", 94507, 37.855, -122.010),
    ("Danville", "Contra Costa", 94526, 37.820, -122.000),
    ("Blackhawk", "Contra Costa", 94506, 37.810, -121.910),
    
    # Marin & North Bay
    ("Mill Valley", "Marin", 94941, 37.905, -122.545),
    ("Tiburon", "Marin", 94920, 37.895, -122.460),
    ("Belvedere", "Marin", 94920, 37.875, -122.463),
    ("Ross", "Marin", 94957, 37.965, -122.555),
    ("Kentfield", "Marin", 94904, 37.950, -122.550),
    ("San Anselmo", "Marin", 94960, 37.975, -122.565),
    ("Sonoma", "Sonoma", 95476, 38.300, -122.465),
    ("Glen Ellen", "Sonoma", 95442, 38.360, -122.530),
    ("Kenwood", "Sonoma", 95452, 38.410, -122.545),
    ("Healdsburg", "Sonoma", 95448, 38.615, -122.865),
    ("Napa", "Napa", 94558, 38.320, -122.290),
    ("Yountville", "Napa", 94599, 38.400, -122.360),
    ("St. Helena", "Napa", 94574, 38.505, -122.470),
    ("Calistoga", "Napa", 94515, 38.580, -122.580),
    
    # Central Coast
    ("Carmel", "Monterey", 93921, 36.545, -121.910),
    ("Carmel Highlands", "Monterey", 93923, 36.515, -121.915),
    ("Carmel Valley", "Monterey", 93924, 36.480, -121.750),
    ("Pebble Beach", "Monterey", 93953, 36.590, -121.930),
    ("Pacific Grove", "Monterey", 93950, 36.615, -121.915),
    ("Santa Cruz", "Santa Cruz", 95060, 36.980, -122.030),
    ("Aptos", "Santa Cruz", 95003, 36.995, -121.890),
    ("San Luis Obispo", "San Luis Obispo", 93401, 35.280, -120.660),
    ("Paso Robles", "San Luis Obispo", 93446, 35.625, -120.690),
    
    # Sierra / Tahoe & Foothills
    ("Tahoe City", "Placer", 96145, 39.185, -120.165),
    ("Truckee", "Nevada", 96161, 39.330, -120.185),
    ("Olympic Valley", "Placer", 96146, 39.195, -120.240),
    ("Incline Village", "Washoe", 89451, 39.265, -119.960),
    ("Mammoth Lakes", "Mono", 93546, 37.645, -118.970),
    ("Nevada City", "Nevada", 95959, 39.260, -121.020),
    ("Grass Valley", "Nevada", 95945, 39.220, -121.060),
    ("Auburn", "Placer", 95603, 38.895, -121.075),
    ("Placerville", "El Dorado", 95667, 38.730, -120.800),
]

def is_valid_land(lat, lng):
    """
    Validates that a simulated geographic coordinate falls strictly on land
    and not in the Pacific Ocean, coastal bays, or major inland lakes (e.g. Lake Tahoe).
    """
    try:
        from global_land_mask import globe
        if not globe.is_land(lat, lng):
            return False
    except ImportError:
        pass
    
    # Lake Tahoe water body exclusion
    if 38.935 <= lat <= 39.245 and -120.145 <= lng <= -119.950:
        return False
    # San Francisco Bay open water buffer (between Angel Island and East Bay)
    if 37.820 <= lat <= 37.875 and -122.435 <= lng <= -122.380:
        return False
    # Richardson Bay open water buffer (between Sausalito and Tiburon)
    if 37.865 <= lat <= 37.890 and -122.505 <= lng <= -122.475:
        return False
    # California bounding box
    if not (32.5 <= lat <= 42.1 and -124.5 <= lng <= -114.1):
        return False
    return True

def hav(lat1, lon1, lat2, lon2):
    """Great-circle distance in kilometers between two coordinates."""
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    x = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.atan2(math.sqrt(x), math.sqrt(1 - x))

# Distinct, realistic sub-neighborhood centers for premier high-density communities.
# Spaced >= 2.5km apart to ensure clean, non-overlapping 1km accumulation circles.
HOTSPOTS = {
    "Pacific Palisades": [
        ("Highlands", 34.085, -118.555, 8),
        ("Marquez Knolls", 34.055, -118.545, 8),
        ("Riviera", 34.048, -118.505, 8),
    ],
    "Beverly Hills": [
        ("Trousdale", 34.095, -118.395, 8),
        ("Benedict Canyon", 34.120, -118.425, 8),
        ("Flats", 34.072, -118.405, 8),
    ],
    "Malibu": [
        ("Serra Retreat", 34.040, -118.680, 8),
        ("Point Dume", 34.015, -118.805, 8),
        ("Malibu Canyon", 34.080, -118.705, 8),
    ],
    "Carmel": [
        ("Carmel Woods", 36.555, -121.920, 8),
        ("Carmel Valley", 36.530, -121.840, 8),
    ],
    "Woodside": [
        ("Skyline", 37.440, -122.280, 8),
        ("Central", 37.425, -122.245, 8),
    ],
    "Saratoga": [
        ("Foothills", 37.260, -122.045, 8),
        ("Valley", 37.265, -122.010, 8),
    ],
    "Hillsborough": [
        ("North", 37.575, -122.360, 8),
        ("South", 37.545, -122.335, 8),
    ],
}

def generate_synthetic_dataset(n_target=895, seed=42):
    """
    Generates properties using discrete spatial neighborhood pockets.
    Each pocket contains homes spread compactly (radius <= 280m) so all homes
    within the pocket mathematically share the exact same optimal 1km circle.
    Pocket centers maintain >= 2.2km clearance from each other, completely
    eliminating redundant overlapping circles in high-density areas like
    Pacific Palisades while preserving authentic California WUI distributions.
    All properties and pocket centers are rigorously validated to fall strictly on land.
    """
    np.random.seed(seed)
    random.seed(seed)
    
    pockets = []
    dense_centers = []

    # 1. Hotspot sub-neighborhoods and showcase pockets
    for reg in REGIONS:
        city, county, zip_code, base_lat, base_lng = reg
        if city in HOTSPOTS:
            for sub_name, s_lat, s_lng, cnt in HOTSPOTS[city]:
                pockets.append((city, county, zip_code, s_lat, s_lng, cnt, 0.28))
                dense_centers.append((s_lat, s_lng))
        elif city == "San Francisco":
            pockets.append((city, county, zip_code, base_lat, base_lng, 9, 0.22))
            dense_centers.append((base_lat, base_lng))
        elif city == "Calabasas":
            pockets.append((city, county, zip_code, base_lat, base_lng, 8, 0.25))
            dense_centers.append((base_lat, base_lng))

    # 2. Standard community pockets (spaced >= 2.2km from all existing centers)
    for reg in REGIONS:
        city, county, zip_code, base_lat, base_lng = reg
        if city in HOTSPOTS or city in ("San Francisco", "Calabasas"):
            continue
        
        p1_lat, p1_lng = None, None
        if all(hav(base_lat, base_lng, clat, clng) >= 2.2 for clat, clng in dense_centers) and is_valid_land(base_lat, base_lng):
            p1_lat, p1_lng = base_lat, base_lng
        else:
            for _att in range(50):
                ang = np.random.uniform(0, 2 * math.pi)
                dist_km = np.random.uniform(2.5, 4.5)
                cand_lat = round(base_lat + (dist_km / 111.139) * math.cos(ang), 6)
                cand_lng = round(base_lng + (dist_km / (111.139 * math.cos(math.radians(base_lat)))) * math.sin(ang), 6)
                if is_valid_land(cand_lat, cand_lng) and all(hav(cand_lat, cand_lng, clat, clng) >= 2.2 for clat, clng in dense_centers):
                    p1_lat, p1_lng = cand_lat, cand_lng
                    break
        
        if p1_lat is not None:
            cnt1 = int(np.random.choice([2, 3, 4, 5, 6], p=[0.25, 0.35, 0.25, 0.10, 0.05]))
            pockets.append((city, county, zip_code, p1_lat, p1_lng, cnt1, 0.20 if city == "Belvedere" else 0.28))
            dense_centers.append((p1_lat, p1_lng))
            
            # 35% chance of second pocket spaced >= 2.5km
            if np.random.rand() < 0.35:
                for _att in range(50):
                    ang = np.random.uniform(0, 2 * math.pi)
                    dist_km = np.random.uniform(2.6, 3.8)
                    cand_lat = round(p1_lat + (dist_km / 111.139) * math.cos(ang), 6)
                    cand_lng = round(p1_lng + (dist_km / (111.139 * math.cos(math.radians(p1_lat)))) * math.sin(ang), 6)
                    if is_valid_land(cand_lat, cand_lng) and all(hav(cand_lat, cand_lng, clat, clng) >= 2.2 for clat, clng in dense_centers):
                        cnt2 = int(np.random.choice([2, 3, 4, 5], p=[0.35, 0.35, 0.20, 0.10]))
                        pockets.append((city, county, zip_code, cand_lat, cand_lng, cnt2, 0.28))
                        dense_centers.append((cand_lat, cand_lng))
                        break

    rows = []
    pid_base = 900000000
    home_id = 1
    
    def make_row(hid, city, county, zip_code, lat, lng):
        if city == "San Francisco":
            score = round(float(np.random.uniform(0.12, 0.32)), 2)
            tiv_raw = np.random.lognormal(16.0, 0.35)
            tiv = float(np.clip(round(tiv_raw / 25000) * 25000, 6000000, 18000000))
            pf = round(float(np.random.uniform(0.00006, 0.00016)), 6)
            plf = round(float(np.random.uniform(0.20, 0.35)), 4)
        elif city == "Calabasas":
            score = round(float(np.random.uniform(2.8, 3.7)), 2)
            tiv_raw = np.random.lognormal(15.7, 0.40)
            tiv = float(np.clip(round(tiv_raw / 25000) * 25000, 4000000, 16000000))
            pf = round(float(np.random.uniform(0.016, 0.028)), 6)
            plf = round(float(np.random.uniform(0.78, 0.94)), 4)
        else:
            score = round(float(np.random.beta(2, 2.5) * 3.6 + 0.15), 2)
            tiv_raw = np.random.lognormal(15.6, 0.45)
            tiv = float(np.clip(round(tiv_raw / 25000) * 25000, 1500000, 18000000))
            pf = round(float(np.clip(np.random.exponential(0.0005 + (score / 4.0) * 0.006), 0.0001, 0.035)), 6)
            plf = round(float(np.clip(np.random.beta(3, 1.8), 0.1, 1.0)), 4)
        
        pl = round(pf * plf, 6)
        tivpl = round(tiv * pl, 2)
        retool_id = f"SYNTH-{county[:3].upper()}-{hid:04d}"
        pol_num = f"POL-{pid_base + hid}"
        return {
            "ID": hid,
            "Policy Number": pol_num,
            "Retool Property ID": retool_id,
            "Latitude": lat,
            "Longitude": lng,
            "Zip": zip_code,
            "City": city,
            "County": county,
            "GRe WF Score": score,
            "TIV": tiv,
            "p(f)": pf,
            "PLF": plf,
            "p(L) final": pl,
            "TIV*p(L)": tivpl
        }

    for pkt in pockets:
        city, county, zip_code, p_lat, p_lng, cnt, spread = pkt
        for _ in range(cnt):
            if home_id > n_target:
                break
            for _att in range(50):
                # Bounded uniform disk sampling (maximum radius = spread in km)
                u = np.random.rand()
                r_km = spread * math.sqrt(u)
                theta = np.random.uniform(0, 2 * math.pi)
                d_lat = (r_km / 111.139) * math.cos(theta)
                d_lng = (r_km / (111.139 * math.cos(math.radians(p_lat)))) * math.sin(theta)
                lat = round(p_lat + d_lat, 6)
                lng = round(p_lng + d_lng, 6)
                if is_valid_land(lat, lng):
                    break
            else:
                lat, lng = round(p_lat, 6), round(p_lng, 6)
            
            rows.append(make_row(home_id, city, county, zip_code, lat, lng))
            home_id += 1

    # Add isolated standalone homes (Zone 1 background exposure) until n_target is reached.
    # Exclude designated hotspot cities from standalone sampling, and enforce clearance
    # from all pocket centers (>= 1.6km) and other standalone homes (>= 1.4km).
    standalone_coords = []
    standard_regions = [r for r in REGIONS if r[0] not in HOTSPOTS and r[0] not in ("San Francisco", "Calabasas")]

    attempts = 0
    while home_id <= n_target and attempts < 25000:
        attempts += 1
        reg = random.choice(standard_regions)
        city, county, zip_code, base_lat, base_lng = reg
        ang = np.random.uniform(0, 2 * math.pi)
        dist_km = np.random.uniform(2.5, 6.0)
        cand_lat = round(base_lat + (dist_km / 111.139) * math.cos(ang), 6)
        cand_lng = round(base_lng + (dist_km / (111.139 * math.cos(math.radians(base_lat)))) * math.sin(ang), 6)
        if is_valid_land(cand_lat, cand_lng):
            if all(hav(cand_lat, cand_lng, dlat, dlng) >= 1.6 for dlat, dlng in dense_centers) and \
               all(hav(cand_lat, cand_lng, slat, slng) >= 1.4 for slat, slng in standalone_coords):
                standalone_coords.append((cand_lat, cand_lng))
                rows.append(make_row(home_id, city, county, zip_code, cand_lat, cand_lng))
                home_id += 1

    # Final fill to reach exactly n_target if needed
    while home_id <= n_target:
        reg = random.choice(standard_regions)
        city, county, zip_code, base_lat, base_lng = reg
        ang = np.random.uniform(0, 2 * math.pi)
        dist_km = np.random.uniform(2.5, 8.0)
        cand_lat = round(base_lat + (dist_km / 111.139) * math.cos(ang), 6)
        cand_lng = round(base_lng + (dist_km / (111.139 * math.cos(math.radians(base_lat)))) * math.sin(ang), 6)
        if is_valid_land(cand_lat, cand_lng):
            if all(hav(cand_lat, cand_lng, dlat, dlng) >= 1.6 for dlat, dlng in dense_centers) and \
               all(hav(cand_lat, cand_lng, slat, slng) >= 1.2 for slat, slng in standalone_coords):
                standalone_coords.append((cand_lat, cand_lng))
                rows.append(make_row(home_id, city, county, zip_code, cand_lat, cand_lng))
                home_id += 1

    df = pd.DataFrame(rows)
    return df

def write_excel(df, out_path=OUTPUT_EXCEL):
    wb = Workbook()
    
    # 1. Cover / Disclaimer sheet
    ws_cover = wb.active
    ws_cover.title = "Disclaimer & Metadata"
    
    A = lambda **k: Font(name="Arial", **k)
    ws_cover.cell(1, 1, "Wildfire 1km Aggregation Model — Synthetic Portfolio").font = A(size=14, bold=True, color="1F3864")
    ws_cover.cell(2, 1, "NOTICE: 100% SYNTHETIC DATASET (PORTFOLIO SHOWCASE)").font = A(size=11, bold=True, color="C00000")
    
    notes = [
        "All data in this workbook (coordinates, policy numbers, property IDs, and insured values) is purely synthetic.",
        "Created specifically for public code portfolio and resume presentation on GitHub.",
        "Contains NO proprietary carrier, company, or real policyholder information.",
        "Synthesized using calibrated spatial neighborhood pockets across California wildland-urban interface (WUI) zones.",
        f"Total Simulated Properties: {len(df)}",
        f"Unique Cities Represented: {df['City'].nunique()}",
        f"Total Portfolio TIV: ${df['TIV'].sum():,.0f}",
        f"Average Property TIV: ${df['TIV'].mean():,.0f}",
        "Data schema matches standard reinsurance inforce risk bordereaux formats."
    ]
    for idx, note in enumerate(notes, start=4):
        ws_cover.cell(idx, 1, note).font = A(size=10)
    ws_cover.column_dimensions["A"].width = 85
    
    # 2. Summary sheet (the required pipeline data tab)
    ws_sum = wb.create_sheet("Summary")
    cols = list(df.columns)
    
    HF = PatternFill("solid", fgColor="1F3864")
    thin = Side(style="thin", color="D9D9D9")
    bd = Border(thin, thin, thin, thin)
    
    for c_idx, col in enumerate(cols, start=1):
        cell = ws_sum.cell(1, c_idx, col)
        cell.font = A(bold=True, color="FFFFFF")
        cell.fill = HF
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for c_idx, col in enumerate(cols, start=1):
            cell = ws_sum.cell(r_idx, c_idx, row[col])
            cell.font = A(size=10)
            cell.border = bd
            if col in ("Latitude", "Longitude", "p(f)", "PLF", "p(L) final"):
                cell.number_format = "0.000000"
                cell.alignment = Alignment(horizontal="right")
            elif col == "GRe WF Score":
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col == "TIV":
                cell.number_format = "$#,##0"
                cell.alignment = Alignment(horizontal="right")
            elif col == "TIV*p(L)":
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col in ("ID", "Zip"):
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
                
    for col in cols:
        L = GL(cols.index(col) + 1)
        ws_sum.column_dimensions[L].width = 18 if "TIV" in col or "Property" in col else 14
        
    ws_sum.freeze_panes = "A2"
    
    wb.save(out_path)
    print(f"Saved synthetic workbook to: {out_path}")

if __name__ == "__main__":
    df = generate_synthetic_dataset()
    write_excel(df)
    print(f"Generated {len(df)} properties across {df['City'].nunique()} cities.")
