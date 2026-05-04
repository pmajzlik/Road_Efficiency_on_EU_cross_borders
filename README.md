[README.md](https://github.com/user-attachments/files/27362467/README.md)
# Cross-Border Network Efficiency Ratio (NER) Analysis

A set of ArcGIS Python toolbox scripts for analyzing road network efficiency in cross-border areas. The workflow downloads and preprocesses OpenStreetMap road data, builds a routable network, and calculates the **Network Efficiency Ratio (NER)** — a population-weighted measure of how efficiently the road network connects settlements across national borders compared to straight-line (theoretical) travel.

**Author:** Bc. Pavel MAJZLIK
**Institution:** Department of Geoinformatics, Faculty of Science, Palacký University Olomouc  
**Year:** 2026

---

## Requirements

- **ArcGIS Pro** with the **Network Analyst** extension
- Python packages: `arcpy`, `requests`, `pandas`, `openpyxl`
- Internet access for OSM data download (Overpass API)

---

## Workflow Overview

The scripts are numbered and must be run in order:

```
01 → 02 → 03 → 04 → 05 → [manual step] → 06
```

---

## Scripts

### 01 — Define Cross-Border Area
**`01_Define_cross_border_area.py`**

Defines the cross-border study area from a country polygon layer. Extracts shared international borders and creates a 25 km buffer around them, clipped to the country boundaries.

**Inputs:**
- Country polygon layer (must contain a `CNTR_CODE` field)
- Output folder and geodatabase name

**Outputs:**
- `<layer>_border` — polyline feature class of international borders
- `<layer>_border_buffer25km` — 25 km buffer polygon (the study area)

---

### 02 — Download OSM Data
**`02_Download_OSM_data.py`**

Downloads road network and settlement data from the **Overpass API** for the defined study area. Only road types relevant for routing are downloaded. Settlements are filtered to those with a recorded population. Country codes are assigned to settlements via a spatial join.

**Inputs:**
- Output geodatabase
- Clip polygon (the study area from step 01)
- Output name prefix
- Timeout (seconds, default 180)
- Country polygons (for assigning country codes to settlements)

**Outputs:**
- `<name>_roads` — polyline feature class of roads
- `<name>_places` — point feature class of settlements (with population and country)

**Downloaded road types:** motorway, motorway_link, trunk, trunk_link, primary, primary_link, secondary, secondary_link, tertiary, tertiary_link, unclassified, residential, living_street

---

### 03 — Roundabout Flow Improvement
**`03_Roundabout_flow_improvement.py`**

Improves topological correctness of the road network at roundabouts. Selects all roundabout features (tagged `junction = roundabout` or `circuit`), adds connecting roads, and planarizes the geometry using **Feature To Line**. Attributes are transferred via a spatial join.

**Inputs:**
- Road feature class (from step 02)
- Workspace (geodatabase)

**Outputs:**
- `<roads>_edited` — road feature class with corrected roundabout topology

---

### 04 — Geoprocessing Data
**`04_Geoprocessing_data.py`**

Prepares road attributes for network routing and snaps settlement points onto the nearest road.

- Adds a `speed_kmh` field with default speeds per road type (with a 10% safety reduction applied to all values)
- Calculates `length_km` (geodesic) and `TravelTime` (in minutes) for each road segment
- Hierarchically snaps settlement points to the nearest road (priority: primary → secondary → tertiary → residential) within a 250 m search radius and a maximum allowed shift of 500 m

**Inputs:**
- Settlements layer (from step 02)
- Road network layer (from step 03)

**Outputs:** Fields `speed_kmh`, `length_km`, and `TravelTime` added in place; settlement geometries updated.

**Default speeds (after 10% reduction):**

| Road type | Base speed | Used speed |
|---|---|---|
| motorway | 130 km/h | 117 km/h |
| trunk | 110 km/h | 99 km/h |
| primary / secondary | 90 km/h | 81 km/h |
| tertiary | 80 km/h | 72 km/h |
| unclassified / residential | 50 km/h | 45 km/h |

---

### 05 — Create Network Dataset
**`05_Create_network_dataset.py`**

Creates and builds an ArcGIS **Network Dataset** from the preprocessed road layer. Validates the input geometry and spatial reference (reprojecting to WGS84 / EPSG:4326 if necessary). Checks for existing datasets to prevent conflicts.

**Inputs:**
- Road layer (from step 04)
- Geodatabase path
- Network dataset name

**Outputs:**
- A feature dataset containing the roads and a built Network Dataset

> ⚠️ **Manual configuration required after this step:**  
> The `TravelTime` field must be registered as a cost attribute in the Network Dataset properties before solving. See the script's output messages for step-by-step instructions.

---

### 06 — Calculate NER
**`06_Calculate_NER.py`**

Calculates the **Network Efficiency Ratio (NER)** for each settlement using an OD Cost Matrix. Only cross-border settlement pairs (different countries) are included in the calculation.

**NER formula:**

$$NE_i = \frac{\sum (t_{actual} / t_{theoretical}) \cdot P_j}{\sum P_j}$$

$$NER_i = \frac{1}{NE_i}$$

Where $t_{theoretical}$ is the straight-line travel time (Euclidean distance / 130 km/h) and $P_j$ is the population of the destination settlement.

**Inputs:**
- Settlements layer
- Network Dataset (from step 05)
- Geodatabase path
- Output folder

**Outputs:**
- `ODLines_<network>` — feature class with solved OD matrix lines
- `Actual_times_<network>.xlsx` — actual network travel times
- `Theoretical_times_<network>.xlsx` — straight-line theoretical times (cross-border pairs only)
- `NER_results_<network>.xlsx` — final NER values per settlement with country and population

---


