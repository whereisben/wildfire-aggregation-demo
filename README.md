# Wildfire 1 km Geospatial Exposure Aggregation Model

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![MapLibre GL](https://img.shields.io/badge/MapLibre-GL%20JS%20v5-blueviolet.svg)](https://maplibre.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Dataset](https://img.shields.io/badge/Data-100%25%20Synthetic-orange.svg)](#disclaimer--synthetic-dataset)

A high-performance geospatial insurance-accumulation pipeline. For every insured property in a portfolio, this model identifies the **optimal 1 km-radius circle (with continuous, free-center placement within 1 km of the property)** that maximizes cumulative financial exposure. The analysis evaluates exposure across a **6-version risk-score filter matrix**, rolling up **worst-case accumulation zones**, **equal-count quintiles**, and **canonical hotspot clusters** across both **Total Insured Value (TIV)** and **Risk-Adjusted Expected Loss ($TIV \times p(L)$)**.

Accompanied by an interactive **3D MapLibre WebGL Globe** and a **responsive dynamic pivot dashboard**.

---

> ### ⚠️ Disclaimer & Synthetic Dataset Notice
> **All data in this repository is 100% synthetically generated.** 
> Property locations, geographic coordinates, policy numbers, property IDs, and insured values are mathematically simulated across California wildland-urban interface (WUI) clusters. **This repository contains NO proprietary company, insurer, or real policyholder information.** It is authored strictly for technical portfolio showcase, algorithm demonstration, and open-source review.

---

## 🌟 Key Technical Highlights

1. **Continuous Free-Center Optimization (Spherical Geometry):**
   - Traditional accumulation buffers anchor circles rigidly on property locations. This pipeline solves for the optimal continuous center $(lat, lng)$ within 1 km of each home by computing spherical circle-circle intersection points (Haversine, forward bearings, and great-circle destination points), finding the true mathematical maximum exposure.
2. **Multi-Version Risk Filtering:**
   - Evaluates a $3 \times 3$ matrix (6 active versions) of wildfire hazard score thresholds for home and neighbor inclusion, with dynamic zone limits ($150M, $100M, $50M).
3. **Dual Value Bases:**
   - Optimizes against both raw capital exposure (**TIV**) and catastrophe risk-adjusted expected loss (**$TIV \times p(L)$**).
4. **Interactive 3D WebGL Visualization:**
   - Self-contained MapLibre GL JS globe (`aggregation_hotspots_globe.html`) with 3D terrain/hillshade, interactive circle pins, member inspection, burn-rate chloropleths, and an interactive spatial probe tool.
5. **Interactive Dynamic Dashboard:**
   - Browser-based pivot dashboard (`aggregation_pivots.html`) enabling instantaneous grouping, zone filtering, and sorting without requiring Microsoft Excel.
6. **Executive Excel Delivery:**
   - Generates an 9-sheet formatted workbook (`Wildfire_1km_Aggregation_FULL.xlsx`) with custom styling, freeze panes, number formatting, worst-case extracts, and limit breakouts.

---

## 📐 The Optimization Problem

For an anchor home $H = (\text{lat}_H, \text{lng}_H)$ with neighbors $\mathcal{N}$:

$$\max_{C \in \mathbb{R}^2} \sum_{i \in \mathcal{N}} \text{TIV}_i \cdot \mathbb{I}(\text{dist}(P_i, C) \le 1.0\text{ km})$$

$$\text{subject to } \text{dist}(H, C) \le 1.0\text{ km}$$

Where:
- $\text{dist}(A, B)$ is the spherical Haversine distance on WGS-84 ($R = 6371.0088\text{ km}$).
- $C$ is a continuous candidate center derived from property locations and pairwise chord intersections $\text{dest}(M, \theta \pm 90^\circ, h_c)$.

```
                      1 km Maximum Exposure Circle
                           .---''''---.
                        .-'      * P2  '-.
                       /    * P1          \
                      |             * P3   |
                      |        + C         |  <-- Free Center (optimal)
                      |   * H              |
                       \      * P4        /
                        '-.            .-'
                           '---....---'
                           |<-- 1km -->|
```

---

## 📊 Analytical Outputs

When executed, the pipeline generates:

### 1. Multi-Tab Formatted Excel Workbook (`Wildfire_1km_Aggregation_FULL.xlsx`)
- **`Disclaimer & Overview`** — Project metadata, mathematical parameters, and synthetic data notice.
- **`Aggregation Results (TIV)`** — Source properties alongside 6 version circle IDs, maximum TIV, zone designations, worst-case rollup, quintile zones, and optimal circle center coordinates.
- **`Aggregation Results (TIVxp(L))`** — Single-view Expected Loss ($TIV \times p(L)$) optimization results.
- **`Circle Summary (TIVxpL)`** — Canonical unique circles with member counts, average wildfire metrics ($p(f), PLF, p(L)$), aggregate TIV, and expected loss.
- **`WC Zone 5 Circles`** & **`WC Zone 5 Properties`** — Dedicated extracts for catastrophic accumulation zones exceeding underwriting limits.
- **`Version Key & Heat Map`** — Parametric quartile breakdown and threshold documentation.
- **`WC Zone 5 by Limit`** & **`WC Top Quintile by Limit`** — Limit-stratified summaries ($50M, $100M, $150M).

### 2. 3D MapLibre Globe (`aggregation_hotspots_globe.html`)
- Standalone HTML/JS application powered by MapLibre GL JS v5.
- Renderable directly in any browser (no local webserver required).
- Features:
  - 3D Globe Projection with hillshade elevation toggle.
  - Hover inspection revealing all individual homes and TIVs enclosed in any hotspot circle.
  - Dual-circle pinning and comparison view.
  - **Risk Contrast Showcase**: One-click side-by-side comparison of two high-TIV accumulation clusters ($70M–$100M+ TIV) illustrating how catastrophe probability $p(L)$ drives portfolio risk: an urban cluster (San Francisco: low $p(L)$, burn rate 0.003%, RA-TIV ~$2k) versus an extreme WUI interface (Calabasas: high $p(L)$, burn rate ~1.87%, RA-TIV ~$1.35M).
  - Interactive **Spatial Probe Tool**: click anywhere on the globe to calculate arbitrary 1 km exposure on the fly.
  - Burn-rate color interpolation ($TIV \times p(L) / TIV$).

### 3. Interactive Pivot Dashboard (`aggregation_pivots.html`)
- Filter and slice canonical accumulation circles.
- Groupable by Circle, City, or Zone.
- Toggle dynamically between Worst-Case TIV and Risk-Adjusted Expected Loss.

---

## 🚀 Quickstart

### Installation
```bash
# Clone the repository
git clone https://github.com/your-username/wildfire-aggregation-demo.git
cd wildfire-aggregation-demo

# Install dependencies (pandas, openpyxl, numpy)
pip install -r requirements.txt
```

### Run the Pipeline
```bash
python3 run_aggregation_full.py
```

Expected output:
```text
Data: Synthetic_Inforce_Risk_Portfolio.xlsx
Rows: 895
Saved: Wildfire_1km_Aggregation_FULL.xlsx
  globe updated: aggregation_hotspots_globe.html (tiv=434, tivpl=438, homes=895)
  pivots dashboard written: aggregation_pivots.html (wc=434, tivpl=438)
```

### View Visualizations
Open either HTML file in your web browser:
```bash
# macOS
open aggregation_hotspots_globe.html
open aggregation_pivots.html

# Linux
xdg-open aggregation_hotspots_globe.html
```

---

## 🔬 Regenerating the Synthetic Dataset

To generate a fresh synthetic portfolio with different seeds or distributions:
```bash
python3 generate_synthetic_data.py
python3 run_aggregation_full.py
```

The generator synthesizes Gaussian property clusters across 70+ wildfire-prone California communities (e.g., Malibu, Pacific Palisades, Carmel, Montecito, Lake Tahoe, Napa Valley, Berkeley Hills) with realistic log-normal TIV distributions and wildfire probabilities.

---

## 📁 Repository Structure

```text
wildfire_aggregation_demo/
├── README.md                                # Project documentation & methodology
├── Aggregation_Runbook.md                   # Full mathematical & technical specification
├── requirements.txt                         # Python dependencies
├── generate_synthetic_data.py               # Synthetic property generator (100% simulated)
├── run_aggregation_full.py                  # End-to-end accumulation & export pipeline
├── Synthetic_Inforce_Risk_Portfolio.xlsx    # Input synthetic property portfolio
├── Wildfire_1km_Aggregation_FULL.xlsx       # Output 9-sheet executive workbook
├── aggregation_hotspots_globe.html          # Interactive 3D WebGL Globe visualization
└── aggregation_pivots.html                  # Interactive dynamic pivot table dashboard
```

---

## 📄 License
Distributed under the MIT License. See `LICENSE` for more information.
