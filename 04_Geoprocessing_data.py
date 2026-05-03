#------------------------------------
# Name: Geoprocessing_data.py
# Description: The script is used to calculate speed limits, travel times for a road network and snap settlements directly on the road.
# Author: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------

import arcpy
import math
import csv 
from pyproj import Transformer
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import unicodedata

arcpy.env.overwriteOutput = True

def speed_assignment(highway_type):
    """
    Assigns a speed limit based on road type.

    Parameters:
    highway_type (str): Road type (e.g. 'motorway', 'trunk', 'primary', etc.).

    Returns:
    float: Speed limit for the given road type in km/h.
    """
    default_speeds = {
        'motorway':          130 * 0.9,
        'motorway_link':     80  * 0.9,
        'trunk':             110 * 0.9,
        'trunk_link':        80  * 0.9,
        'primary':           90  * 0.9,
        'primary_link':      60  * 0.9,
        'secondary':         90  * 0.9,
        'secondary_link':    60  * 0.9,
        'tertiary':          80  * 0.9,
        'tertiary_link':     60  * 0.9,
        'unclassified':      50  * 0.9,
        'residential':       50  * 0.9,
    }
    return float(default_speeds.get(highway_type, 50))


def add_speed_field(network_layer):
    """
    Adds a 'speed_kmh' field to the road layer and populates it based on road type.
    
    Parameters:
    network_layer (str): Path to the road network layer.
    """
    network_layer = arcpy.MakeFeatureLayer_management(network_layer, "network_layer")
    
    # Add new speed field if it does not already exist
    fields = [f.name for f in arcpy.ListFields(network_layer)]
    if "speed_kmh" not in fields:
        arcpy.AddField_management(network_layer, "speed_kmh", "DOUBLE")
    
    # Update values in the speed_kmh field
    with arcpy.da.UpdateCursor(network_layer, ['highway', 'speed_kmh']) as cursor:
        for row in cursor:
            highway_type = row[0]
            speed = speed_assignment(highway_type)
            row[1] = float(speed)
            cursor.updateRow(row)
    
    arcpy.AddMessage("Field 'speed_kmh' has been successfully added and populated by road type.")
    return network_layer


def add_length_and_travel_time_fields(network_layer):
    """
    Adds 'length_km' and 'TravelTime' fields to the network layer if they do not exist.
    
    Parameters:
    network_layer (str): Path to the network layer.
    """
    travel_time_field = "TravelTime"
    length_field = "length_km"
    
    fields = [f.name for f in arcpy.ListFields(network_layer)]
    
    if travel_time_field not in fields:
        arcpy.AddField_management(network_layer, travel_time_field, "DOUBLE")
        arcpy.AddMessage(f"Field '{travel_time_field}' has been added to the layer.")
    
    if length_field not in fields:
        arcpy.AddField_management(network_layer, length_field, "DOUBLE")
        arcpy.AddMessage(f"Field '{length_field}' has been added to the layer.")


def calculate_travel_time(network_layer):
    """
    Calculates the length in km and travel time based on speed for each edge in the network.
    
    Parameters:
    network_layer (str): Path to the network layer.
    """
    with arcpy.da.UpdateCursor(network_layer, ["SHAPE@", "speed_kmh", "length_km", "TravelTime"]) as cursor:
        for row in cursor:
            shape = row[0]
            speed = row[1]
            
            # Skip records without geometry
            if shape is None:
                row[2] = 0
                row[3] = 0
                cursor.updateRow(row)
                continue
            
            length_m = shape.getLength("GEODESIC", "METERS")
            length_km = length_m / 1000.0
            
            # Write length to the 'length_km' attribute
            row[2] = length_km
            
            # Calculate travel time (in minutes)
            if speed and speed > 0:
                travel_time = (length_km / speed) * 60.0
                row[3] = travel_time
            else:
                row[3] = 0
            
            cursor.updateRow(row)
    
    arcpy.AddMessage("TravelTime calculation based on geometry and speed completed successfully.")


def point_distance(p1, p2):
    return math.hypot(p1[0] - p2[0], p1[1] - p2[1])


def snap_settlements_to_roads_safe(settlements_layer, network_layer, snap_distance=250, max_allowed_shift=500):
    """
    Hierarchical snapping of points to roads by priority.
    Searches first for primary, then secondary, tertiary, residential, motorway.
    """

    arcpy.AddMessage("=" * 50)
    arcpy.AddMessage("Hierarchical snapping of settlements to roads...")

    # Road type priority order
    road_hierarchy = ['primary', 'secondary', 'tertiary', 'residential']

    # Store original positions
    original_positions = {}
    with arcpy.da.SearchCursor(settlements_layer, ["OID@", "SHAPE@XY"]) as cursor:
        for oid, xy in cursor:
            original_positions[oid] = xy

    # Dictionary to store the best snap position found for each point
    best_snap = {}  

    # Iterate through road type hierarchy
    for road_type in road_hierarchy:
        arcpy.AddMessage(f"Searching for road type: {road_type}")
        
        roads_lyr = "roads_lyr_temp"
        arcpy.MakeFeatureLayer_management(
            network_layer,
            roads_lyr,
            f"highway = '{road_type}'"
        )

        # NEAR analysis for this road type
        arcpy.Near_analysis(
            settlements_layer, 
            roads_lyr, 
            search_radius=f"{snap_distance} Meters",
            location="LOCATION",
            angle="NO_ANGLE"
        )

        # Process results and save if no better solution exists yet
        with arcpy.da.SearchCursor(settlements_layer, 
                                   ["OID@", "NEAR_DIST", "NEAR_X", "NEAR_Y"]) as cursor:
            for oid, near_dist, near_x, near_y in cursor:
                # If no snap found yet for this point and a road was found
                if oid not in best_snap and near_dist != -1 and near_x is not None:
                    orig_xy = original_positions[oid]
                    shift = point_distance(orig_xy, (near_x, near_y))
                    
                    # Check maximum allowed shift
                    if shift <= max_allowed_shift:
                        best_snap[oid] = (near_x, near_y, near_dist, road_type)

        arcpy.Delete_management(roads_lyr)

    # Apply the best snap for each point
    moved = 0
    not_snapped = 0
    
    with arcpy.da.UpdateCursor(settlements_layer, ["OID@", "SHAPE@XY"]) as cursor:
        for oid, xy in cursor:
            if oid in best_snap:
                near_x, near_y, distance, road_type = best_snap[oid]
                cursor.updateRow([oid, (near_x, near_y)])
                moved += 1
                arcpy.AddMessage(f"  OID {oid} → {road_type} (distance: {distance:.1f}m)")
            else:
                not_snapped += 1

    # Clean up NEAR attribute fields
    fields_to_delete = ["NEAR_FID", "NEAR_DIST", "NEAR_X", "NEAR_Y"]
    existing_fields = [f.name for f in arcpy.ListFields(settlements_layer)]
    for field in fields_to_delete:
        if field in existing_fields:
            try:
                arcpy.DeleteField_management(settlements_layer, field)
            except:
                pass

    arcpy.AddMessage(f"Successfully snapped: {moved}")
    arcpy.AddMessage(f"Not snapped (out of range or shift too large): {not_snapped}")
    arcpy.AddMessage("=" * 50)


def main():
    """
    Main function running the entire workflow.
    """
    # Inputs from toolbox
    settlements_path = arcpy.GetParameterAsText(0)
    network_layer = arcpy.GetParameterAsText(1)
    
    # Add speed field to roads
    add_speed_field(network_layer)
    
    # Add length and travel time fields
    add_length_and_travel_time_fields(network_layer)
    
    # Calculate travel time
    calculate_travel_time(network_layer)
    
    # Smart snap of settlements to roads with hierarchical priority 
    snap_settlements_to_roads_safe(
        settlements_path,
        network_layer,
        snap_distance=250,   
        max_allowed_shift=500  
    )

    arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
    arcpy.AddMessage("=" * 60)



if __name__ == "__main__":
    main()
