#------------------------------------
# Name: Calculate_NER.py
# Description: The script is used to calculate network efficiency ratio for the settlements.
# Author: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------

import arcpy
import math
import csv
import pandas as pd
import os
from openpyxl import Workbook

arcpy.env.overwriteOutput = True


def create_od_cost_matrix(network_dataset, gdb_path, impedance_attribute="TravelTime",
                           max_destinations=100000, cutoff=None):
    """
    Creates an OD Cost Matrix layer.

    Parameters:
    network_dataset (str): Path to the network dataset
    gdb_path (str): Path to the geodatabase for saving output
    impedance_attribute (str): Impedance attribute (default "TravelTime")
    max_destinations (int): Maximum number of destinations (default 100)
    cutoff (float): Maximum distance/time cutoff (default None)

    Returns:
    tuple: (od_layer, sub_layers dictionary)
    """
    od_layer = arcpy.na.MakeODCostMatrixLayer(
        in_network_dataset=network_dataset,
        out_network_analysis_layer="OD_CostMatrix",
        impedance_attribute=impedance_attribute,
        default_cutoff=cutoff,
        default_number_destinations_to_find=max_destinations,
        output_path_shape="STRAIGHT_LINES"
    ).getOutput(0)

    arcpy.AddMessage("OD Cost Matrix layer created successfully.")

    sub_layers = arcpy.na.GetNAClassNames(od_layer)
    return od_layer, sub_layers


def add_locations_to_od_matrix(od_layer, sub_layers, settlements_layer,
                                origin_tolerance="5000 Meters",
                                destination_tolerance="20000 Meters"):
    """
    Adds locations (origins and destinations) to the OD Cost Matrix.

    Parameters:
    od_layer: OD Cost Matrix layer
    sub_layers (dict): Dictionary of sub-layers
    settlements_layer (str): Path to the settlements layer
    origin_tolerance (str): Search tolerance for origins
    destination_tolerance (str): Search tolerance for destinations
    """
    origins_layer      = sub_layers["Origins"]
    destinations_layer = sub_layers["Destinations"]

    arcpy.na.AddLocations(
        od_layer, origins_layer, settlements_layer,
        search_tolerance=origin_tolerance,
        search_criteria=[["Roads", "SHAPE"]],
        match_type="MATCH_TO_CLOSEST"
    )

    arcpy.na.AddLocations(
        od_layer, destinations_layer, settlements_layer,
        search_tolerance=destination_tolerance,
        search_criteria=[["Roads", "SHAPE"]],
        match_type="MATCH_TO_CLOSEST"
    )

    arcpy.AddMessage("Locations added successfully.")
    


def solve_od_matrix(od_layer):
    """
    Solves the OD Cost Matrix.

    Parameters:
    od_layer: OD Cost Matrix layer
    """
    arcpy.na.Solve(od_layer)
    arcpy.AddMessage("OD Cost Matrix solved successfully.")


def analyze_od_matrix_success(od_lines_fc, settlements_data):
    """
    Analyzes the success rate of the OD Cost Matrix and exports failed pairs.

    Parameters:
    od_lines_fc (str): Path to the ODLines feature class
    settlements_data (list): List of dictionaries with settlement names
    output_failed_csv (str): Path to the output CSV with failures

    Returns:
    dict: Statistics dictionary
    """
    arcpy.AddMessage("=" * 60)
    arcpy.AddMessage("OD COST MATRIX SUCCESS ANALYSIS")
    arcpy.AddMessage("=" * 60)

    calculated_pairs = set()
    successful_count = 0
    failed_count     = 0

    fields = ["Name", "Total_TravelTime"]

    with arcpy.da.SearchCursor(od_lines_fc, fields) as cursor:
        for row in cursor:
            name, travel_time = row
            if " - " in name:
                parts = name.split(" - ", 1)
                if len(parts) == 2:
                    city1, city2 = parts[0].strip(), parts[1].strip()
                    calculated_pairs.add((city1, city2))
                    if city1 != city2:  
                        if travel_time is not None and travel_time > 0:
                            successful_count += 1
        else:
            failed_count += 1

    all_cities     = [s["name"] for s in settlements_data]
    expected_pairs = set()

    for city1 in all_cities:
        for city2 in all_cities:
            if city1 != city2:
                expected_pairs.add((city1, city2))

    total_expected   = len(expected_pairs)
    total_calculated = len(calculated_pairs)
    missing_pairs    = expected_pairs - calculated_pairs

    arcpy.AddMessage(f"\n STATISTICS:")
    arcpy.AddMessage(f"  Expected pairs:          {total_expected}")
    arcpy.AddMessage(f"  Calculated pairs:        {total_calculated}")
    arcpy.AddMessage(f"  Successful (with value): {successful_count}")
    arcpy.AddMessage(f"  Failed (NULL):           {failed_count}")
    arcpy.AddMessage(f"  Missing pairs:           {len(missing_pairs)}")
    arcpy.AddMessage(f"  Success rate:            {(successful_count / total_expected) * 100:.1f}%")

    if failed_count > 0 or len(missing_pairs) > 0:
        with open('failed_pairs.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Origin', 'Destination', 'Status'])

            with arcpy.da.SearchCursor(od_lines_fc, fields) as cursor:
                for row in cursor:
                    name, travel_time = row
                    if " - " in name and (travel_time is None or travel_time == 0):
                        parts = name.split(" - ", 1)
                        if len(parts) == 2:
                            writer.writerow([parts[0].strip(), parts[1].strip(), 'NULL_VALUE'])

            for pair in missing_pairs:
                writer.writerow([pair[0], pair[1], 'MISSING'])

        arcpy.AddMessage("\nFirst 10 failed pairs:")
        for pair in list(missing_pairs)[:10]:
            arcpy.AddMessage(f"  {pair[0]} -> {pair[1]}")
    else:
        arcpy.AddMessage("\n All pairs calculated successfully!")

    arcpy.AddMessage("=" * 60)

    return {
        'total_expected': total_expected,
        'successful':     successful_count,
        'failed':         failed_count,
        'missing':        len(missing_pairs),
        'success_rate':   (successful_count / total_expected) * 100
    }


def save_od_lines(od_layer, gdb_path, network_dataset_name=None):
    """
    Saves the output Lines layer to the geodatabase, prepending ODLines_ to the network dataset name.

    Parameters:
    od_layer: OD Cost Matrix layer
    gdb_path (str): Path to the geodatabase
    network_dataset_name (str): Name of the network dataset to include in output

    Returns:
    str: Path to the saved feature class
    """
    od_lines_temp = f"{od_layer}\\Lines"

    if network_dataset_name:
        output_name = f"ODLines_{os.path.basename(network_dataset_name)}"
    else:
        output_name = "ODLines"

    od_lines_fc = os.path.join(gdb_path, output_name)

    if arcpy.Exists(od_lines_fc):
        arcpy.management.Delete(od_lines_fc)

    arcpy.management.CopyFeatures(od_lines_temp, od_lines_fc)
    arcpy.AddMessage(f"Output layer {output_name} saved to: {od_lines_fc}")

    return od_lines_fc


def export_actual_times_to_excel(od_lines_fc, output_excel_path):
    """
    Exports actual travel times from ODLines to Excel.

    Parameters:
    od_lines_fc (str): Path to the ODLines feature class
    output_excel_path (str): Path to the output Excel file
    """
    data_actual = []
    fields      = ["Name", "Total_TravelTime"]

    with arcpy.da.SearchCursor(od_lines_fc, fields) as cursor:
        for row in cursor:
            name, travel_time = row
            if " - " in name:
                parts = name.split(" - ", 1)
                if len(parts) == 2:
                    city1, city2 = parts[0].strip(), parts[1].strip()
                    if city1 != city2:
                        data_actual.append({
                            "City1":  city1,
                            "City2":  city2,
                            "Actual": travel_time
                        })

    df_actual = pd.DataFrame(data_actual, columns=["City1", "City2", "Actual"])
    df_actual.to_excel(output_excel_path, index=False)
    arcpy.AddMessage(f"Excel with actual travel times created: {output_excel_path}")


def get_settlement_coordinates(input_fc, name_field="name"):
    """
    Retrieves coordinates and country code of all settlements.
    Coordinates are always returned in EPSG:3035.

    Parameters:
    input_fc (str): Path to the settlements feature class
    name_field (str): Name of the field containing settlement names

    Returns:
    list: List of dictionaries with name, coordinates, and country
    """
    sr = arcpy.Describe(input_fc).spatialReference
    arcpy.AddMessage(f"Layer coordinate system: {sr.name}")

    
    sr_3035 = arcpy.SpatialReference(3035)

    fields = [name_field, "SHAPE@X", "SHAPE@Y", "country"]
    data = []

    with arcpy.da.SearchCursor(input_fc, fields, spatial_reference=sr_3035) as cursor:
        for row in cursor:
            name, x, y, country = row
            if None in (x, y):
                arcpy.AddMessage(f"Missing coordinates for {name}, skipping.")
                continue

            data.append({
                "name":    name,
                "x":       x,
                "y":       y,
                "country": country if country else ""
            })

    return data


def calculate_theoretical_times(settlements_data, speed_kmh=120):
    """
    Calculates theoretical travel times between all pairs of settlements
    based on straight-line (Euclidean) distance at a given speed.

    Parameters:
    settlements_data (list): List of dictionaries with settlement names and coordinates
    speed_kmh (float): Theoretical speed in km/h (default 120)

    Returns:
    list: Matrix of theoretical times (nested lists)
    """
    results = []
    for origin in settlements_data:
        row_result = []
        for dest in settlements_data:
            dx          = origin["x"] - dest["x"]
            dy          = origin["y"] - dest["y"]
            distance_m  = math.sqrt(dx ** 2 + dy ** 2)
            distance_km = distance_m / 1000.0
            time_min    = (distance_km / speed_kmh) * 60
            row_result.append(round(time_min, 1))
        results.append(row_result)

    return results

def export_theoretical_times(settlements_data, results, output_excel_path):
    """
    Exports theoretical times to Excel in long format (City1, City2, Theoretical).
    Only exports cross-border pairs (different countries).

    Parameters:
    settlements_data (list): List of dictionaries with settlement names and country codes
    results (list): Matrix of theoretical times
    output_excel_path (str): Path to the output Excel file
    """
    wb_long = Workbook()
    ws_long = wb_long.active
    ws_long.title = "Theoretical times"

    ws_long.append(["City1", "City2", "Theoretical"])

    cross_border_count = 0
    skipped_count = 0

    for i, origin in enumerate(settlements_data):
        for j, dest in enumerate(settlements_data):
            if i == j:
                continue
            # Export only cross-border pairs
            if origin["country"] and dest["country"] and origin["country"] == dest["country"]:
                skipped_count += 1
                continue
            ws_long.append([origin["name"], dest["name"], results[i][j]])
            cross_border_count += 1

    wb_long.save(output_excel_path)
    arcpy.AddMessage(
        f"Theoretical times (long format) file created: {output_excel_path} "
        f"| Cross-border pairs: {cross_border_count} | Skipped (same country): {skipped_count}"
    )


def get_population_data(shapefile_path, name_field="name", pop_field="population"):
    """
    Loads population and country data for settlements from a shapefile.

    Parameters:
    shapefile_path (str): Path to the shapefile/feature class
    name_field (str): Name of the field containing city names
    pop_field (str): Name of the field containing population values

    Returns:
    dict: Dictionary {city name: {"population": int, "country": str}}
    """
    arcpy.AddMessage(f"Loading settlements shapefile from {shapefile_path}...")
    city_layer = arcpy.MakeFeatureLayer_management(shapefile_path, "city_layer")

    city_data = {}
    with arcpy.da.SearchCursor(city_layer, [name_field, pop_field, "country"]) as cursor:
        for row in cursor:
            city_data[row[0]] = {
                "population": row[1],
                "country":    row[2] if row[2] else ""
            }

    arcpy.AddMessage("City population and country data loaded successfully.")
    return city_data


def calculate_ner(theoretical_excel, actual_excel, population_dict):
    """
    Calculates the Network Efficiency Ratio (NER) for all cities.
    Only cross-border pairs (different countries) are included in the calculation.

    NER formula:
        NE_i = sum(t_actual / t_theoretical * P_j) / sum(P_j)
        NER_i = 1 / NE_i

    Parameters:
    theoretical_excel (str): Path to Excel file with theoretical times
    actual_excel (str): Path to Excel file with actual travel times
    population_dict (dict): Dictionary with population and country data per city

    Returns:
    dict: Dictionary {city: NER value}
    """
    arcpy.AddMessage("Loading theoretical times from Excel...")
    theoretical_times = pd.read_excel(theoretical_excel)
    arcpy.AddMessage("Theoretical times loaded successfully.")

    arcpy.AddMessage("Loading actual times from Excel...")
    actual_times = pd.read_excel(actual_excel)
    arcpy.AddMessage("Actual times loaded successfully.")

    ner_results = {}

    for city_from in theoretical_times['City1'].unique():

        country_from    = population_dict.get(city_from, {}).get("country", "")
        sum_numerator   = 0
        sum_denominator = 0
        pair_count      = 0

        for city_to in theoretical_times['City2'].unique():
            if city_from == city_to:
                continue

            country_to = population_dict.get(city_to, {}).get("country", "")

            # Key condition - skip pairs from the same country
            if country_from and country_to and country_from == country_to:
                continue

            row_theoretical = theoretical_times[
                (theoretical_times['City1'] == city_from) &
                (theoretical_times['City2'] == city_to)
            ]
            row_actual = actual_times[
                (actual_times['City1'] == city_from) &
                (actual_times['City2'] == city_to)
            ]

            if not row_theoretical.empty and not row_actual.empty:
                t_theoretical   = row_theoretical.iloc[0]['Theoretical']
                t_actual        = row_actual.iloc[0]['Actual']
                pop_destination = population_dict.get(city_to, {}).get("population", 0)

                if t_theoretical > 0 and pop_destination > 0:
                    ratio            = t_actual / t_theoretical
                    sum_numerator   += ratio * pop_destination
                    sum_denominator += pop_destination
                    pair_count      += 1

                    arcpy.AddMessage(
                        f"{city_from} ({country_from}) -> {city_to} ({country_to}) | "
                        f"theoretical: {t_theoretical}, actual: {t_actual}, "
                        f"pop_dest: {pop_destination}, ratio: {ratio:.4f}"
                    )

        if sum_denominator > 0:
            ner = 1 / (sum_numerator / sum_denominator)
            ner_results[city_from] = ner
            arcpy.AddMessage(
                f"{city_from} ({country_from}): NER = {ner:.4f} "
                f"from {pair_count} cross-border pairs"
            )
        else:
            ner_results[city_from] = None
            arcpy.AddMessage(
                f"{city_from}: NER cannot be calculated (no valid cross-border data)"
            )

    return ner_results


def export_ner_results(ner_dict, population_dict, output_excel_path):
    """
    Exports NER results to Excel including country and population columns.

    Parameters:
    ner_dict (dict): Dictionary with NER values per city
    population_dict (dict): Dictionary with population and country data per city
    output_excel_path (str): Path to the output Excel file
    """
    rows = []
    for city, ner in ner_dict.items():
        country    = population_dict.get(city, {}).get("country", "")
        population = population_dict.get(city, {}).get("population", "")
        rows.append({
            'City':       city,
            'Country':    country,
            'Population': population,
            'NER':        ner
        })

    ner_df = pd.DataFrame(rows, columns=['City', 'Country', 'Population', 'NER'])

    arcpy.AddMessage(f"Saving NER results to Excel: {output_excel_path}...")
    ner_df.to_excel(output_excel_path, index=False)
    arcpy.AddMessage("NER results saved successfully.")


def main():
    """
    Main function running the full OD Cost Matrix analysis and NER calculation workflow.
    """
    # Toolbox parameters
    settlements_path = arcpy.GetParameterAsText(0)
    network_dataset = arcpy.GetParameterAsText(1)
    gdb_path = arcpy.GetParameterAsText(2)
    output_folder = arcpy.GetParameterAsText(3)

    # Paths 
    output_layer   = settlements_path
    settlements_fc = os.path.join(gdb_path, output_layer)

    # Output Excel file paths
    output_excel_actual = os.path.join(output_folder, f"Actual_times_{os.path.basename(network_dataset)}.xlsx")
    output_excel_theor_long = os.path.join(output_folder, f"Theoretical_times_{network_dataset}.xlsx")
    output_excel_ner = os.path.join(output_folder, f"NER_results_{network_dataset}.xlsx")

    # Create OD Cost Matrix
    od_layer, sub_layers = create_od_cost_matrix(network_dataset, gdb_path)

    # Add locations
    add_locations_to_od_matrix(od_layer, sub_layers, settlements_fc)
 
    # Solve OD matrix 
    solve_od_matrix(od_layer)

    # Save output lines
    od_lines_fc = save_od_lines(od_layer, gdb_path, network_dataset)

    # Export actual travel times 
    export_actual_times_to_excel(od_lines_fc, output_excel_actual)

    # Get settlement coordinates and countries
    settlements_data = get_settlement_coordinates(settlements_fc)

    # Calculate theoretical times 
    theoretical_results = calculate_theoretical_times(settlements_data)

    # Export theoretical times (cross-border pairs only) 
    export_theoretical_times(settlements_data, theoretical_results, output_excel_theor_long)

    # Load population and country data
    population_dict = get_population_data(settlements_fc)

    # Calculate NER
    ner_results = calculate_ner(
        output_excel_theor_long,
        output_excel_actual,
        population_dict
    )

    settlements_data_temp = get_settlement_coordinates(settlements_fc)
    stats = analyze_od_matrix_success(od_lines_fc, settlements_data_temp)

    # Warning if OD matrix success rate is low
    if stats['success_rate'] < 80:
        arcpy.AddWarning(
            f"WARNING: Success rate only {stats['success_rate']:.1f}%! "
            f"Check network connectivity."
        )

    # Export NER results
    export_ner_results(ner_results, population_dict, output_excel_ner)

    arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
    arcpy.AddMessage("=" * 60)


if __name__ == "__main__":
    main()
