# Geospatial 1 km Wildfire Exposure Aggregation — Technical Runbook

A complete specification of the spatial optimization algorithms, mathematical formulas, multi-version filtering matrices, and tie-breaking rules used in this insurance accumulation pipeline.

---

> ### ⚠️ Portfolio Notice: 100% Synthetic Dataset
> All data referenced and included in this demonstration repository is synthetically generated using spatial clustering models across California wildfire-exposed interfaces. No proprietary company, insurer, or real policyholder information is contained herein.

---

## 1. Pipeline Overview & Architecture

The pipeline processes a portfolio of property insurance risks, identifying maximum exposure accumulation within 1 km circles across multiple risk-filtered views:

```
+------------------------------------------+
|  Synthetic_Inforce_Risk_Portfolio.xlsx   |
|  (895 simulated properties across CA)    |
+------------------------------------------+
                    |
                    v
+------------------------------------------+
|       1 km Circle Optimization Engine    |
|   - Pairwise spherical intersections     |
|   - Continuous free-center placement     |
|   - 6 Wildfire Score Filter Versions     |
|   - TIV & TIV * p(L) value metrics       |
+------------------------------------------+
                    |
      +-------------+-------------+
      |                           |
      v                           v
+-----------------------+   +-------------------------------+
| Multi-Tab Workbook    |   | Interactive Visualizations    |
| - 9 Analysis Sheets   |   | - 3D MapLibre Globe (HTML)    |
| - Worst-Case rollups  |   | - Dynamic Pivot Tables (HTML) |
| - Canonical summaries |   +-------------------------------+
+-----------------------+
```

---

## 2. Mathematical Formulation: 1 km Free-Center Circle Optimization

Unlike naive grid aggregation or fixed-center radial buffers (which anchor circle centers exclusively on individual properties), this algorithm employs a **continuous free-center spatial optimization**.

### 2.1 The Optimization Problem
For each property $H = (\text{lat}_H, \text{lng}_H)$ with Total Insured Value $TIV_H$ and wildfire score $S_H$:

$$\max_{C \in \mathbb{R}^2} \sum_{i \in \mathcal{N}} TIV_i \cdot \mathbb{I}(\text{dist}(P_i, C) \le 1.0\text{ km})$$

$$\text{subject to } \text{dist}(H, C) \le 1.0\text{ km}$$

Where:
- $C$ is the center of the 1 km circle.
- $\mathcal{N}$ is the set of eligible neighbor properties.
- $\text{dist}(A, B)$ is the great-circle haversine distance on WGS-84 sphere ($R = 6,371.0088\text{ km}$).

### 2.2 Candidate Center Generation (Spherical Geometry)
By duality in 2D Euclidean geometry (and spherical equivalent over small angles), the maximum subset of points enclosed by a circle of radius $R$ is achieved when the circle boundary passes through at least two points or a single point when no other valid points exist:
1. **1-point centers:** The property itself, plus each candidate neighbor.
2. **2-point intersection centers:** For every pair of candidate points within $2R$ of each other, compute the two circle centers whose boundaries pass through both points:
   - Compute spherical distance $d = \text{hav}(P_1, P_2)$ and forward bearing $\theta_{12} = \text{brg}(P_1, P_2)$.
   - Compute midpoint $M = \text{dest}(P_1, \theta_{12}, d/2)$.
   - Compute chord offset $h_c = \sqrt{R^2 - (d/2)^2}$.
   - The two intersection centers are $C_{1,2} = \text{dest}(M, \theta_{12} \pm 90^\circ, h_c)$.
3. **Filtering & Deduplication:** Candidate centers are pruned if their distance to the anchor home exceeds $1.0\text{ km} + \epsilon$ ($\epsilon = 0.01\text{ km}$).

---

## 3. Multi-Version Wildfire Risk Filter Matrix

To evaluate accumulation across varying underwriting appetites and wildfire severities, the pipeline optimizes across a 6-version filter matrix based on Wildfire Scores:

| Version | Home Row Filter | Neighbor Include Filter | Zone-5 Threshold Limit |
|---|---|---|---|
| **V1** | All | All | $150,000,000 |
| **V4** | Score > 0.5 | All | $150,000,000 |
| **V5** | Score > 0.5 | Score > 0.5 | $100,000,000 |
| **V7** | Score > 1.5 | All | $150,000,000 |
| **V8** | Score > 1.5 | Score > 0.5 | $100,000,000 |
| **V9** | Score > 1.5 | Score > 1.5 | $50,000,000 |

### 3.1 Equal-Width Zone Calibration
For each version with limit $L$, exposure is classified into parametric zones:
- **Zone 0:** Filter failed (row skipped).
- **Zone 1:** $(0, 0.25 L]$
- **Zone 2:** $(0.25 L, 0.50 L]$
- **Zone 3:** $(0.50 L, 0.75 L]$
- **Zone 4:** $(0.75 L, L)$
- **Zone 5:** $\ge L$ (Exceeds accumulation threshold limit)

---

## 4. Worst-Case Rollup & Quintile Ranking

1. **Worst-Case Selection:** For each property across all 6 versions, the worst-case scenario is determined by:
   - Primary: Highest Zone ($Z \in [1, 5]$).
   - Secondary: Highest Maximum Total TIV.
   - Tertiary: Lowest Version Number (preferring broader inclusion).
2. **Equal-Count Quintiles:** Properties are partitioned into 5 equal-frequency buckets using the rank method across both:
   - Total Insured Value ($TIV$)
   - Expected Loss ($TIV \times p(L)$)
3. **Zone Differential:** $\Delta Z = Z_{TIV \times p(L)} - Z_{TIV}$ identifies properties where high burn rates elevate expected loss above nominal exposure.

---

## 5. Canonical Circle Deduplication

When multiple nearby homes optimize to the exact same cluster of neighbor IDs, they represent the same physical hotspot.
- The pipeline canonicalizes circle IDs by sorting member IDs: `"12, 14, 18"`.
- Circle summaries deduplicate canonical sets so reports reflect unique accumulation events without double-counting.
