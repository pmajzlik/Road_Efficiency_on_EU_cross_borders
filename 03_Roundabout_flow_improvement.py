#------------------------------------
# Name: Roundabout_flow_improvement.py
# Description: The script is used to improve the flow of traffic at roundabouts by analyzing their geometry and surrounding infrastructure.
# Author: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------
import arcpy
import os

# Fixed parameters
junction_field = "junction"
search_distance = "0.001 Meters"
planarize_tolerance = "0.001 Meters"

# Toolbox parameters
road_fc = arcpy.GetParameterAsText(0)
workspace = arcpy.GetParameterAsText(1)    

arcpy.env.workspace = workspace
arcpy.env.overwriteOutput = True

def make_layer(fc):
    """
    Creates a feature layer from the input feature class and returns its name and OID field.
    """
    arcpy.management.MakeFeatureLayer(fc, "roads_lyr")
    oid_field = arcpy.Describe("roads_lyr").OIDFieldName
    return "roads_lyr", oid_field


def select_roundabouts(layer):
    """
    Selects all roundabout features from the layer based on the junction field value.
    Returns the WHERE clause used for the selection.
    """
    arcpy.AddMessage("Selecting roundabouts")
    where_clause = f"{junction_field} IN ('roundabout', 'circuit')"
    arcpy.management.SelectLayerByAttribute(layer, "NEW_SELECTION", where_clause)
    count = int(arcpy.management.GetCount(layer)[0])
    arcpy.AddMessage(f"  Roundabouts found: {count}")
    return where_clause


def add_touching_roads(layer):
    """
    Expands the current selection by adding all roads that spatially intersect the selected roundabouts. 
    Returns a set of OIDs of all selected features.
    """
    arcpy.AddMessage("Adding roads touching roundabouts")
    arcpy.management.SelectLayerByLocation(layer, "INTERSECT", layer, search_distance, "ADD_TO_SELECTION")
    total = int(arcpy.management.GetCount(layer)[0])
    arcpy.AddMessage(f"  Total features selected for planarization: {total}")
    oid_field = arcpy.Describe(layer).OIDFieldName
    return set(r[0] for r in arcpy.da.SearchCursor(layer, [oid_field]))


def feature_to_line_selected(layer):
    """
    Exports the selected features, converts them to planarized lines using Feature To Line and transfers attributes through Spatial Join. 
    Returns paths to all three temporary datasets.
    """
    arcpy.AddMessage("Exporting selected features")
    temp_input = os.path.join(workspace, "temp_input_all")
    arcpy.management.CopyFeatures(layer, temp_input)

    arcpy.AddMessage("Feature To Line")
    temp_output = os.path.join(workspace, "temp_output_all")
    arcpy.management.FeatureToLine(temp_input, temp_output, planarize_tolerance, "ATTRIBUTES")

    arcpy.AddMessage("Spatial Join - transferring attributes")
    temp_joined = os.path.join(workspace, "temp_joined_all")
    arcpy.analysis.SpatialJoin(
        target_features=temp_output,
        join_features=temp_input,
        out_feature_class=temp_joined,
        join_operation="JOIN_ONE_TO_ONE",
        join_type="KEEP_ALL",
        match_option="HAVE_THEIR_CENTER_IN",
        search_radius=search_distance
    )
    return temp_input, temp_output, temp_joined


def select_untouched(layer, where_clause):
    """
    Selects all roads that do not intersect any roundabout by inverting the roundabout selection. 
    Exports and returns the path to the temporary untouched roads dataset.
    """
    arcpy.AddMessage("Selecting untouched roads (outside roundabouts)")
    arcpy.management.SelectLayerByAttribute(layer, "CLEAR_SELECTION")
    arcpy.management.SelectLayerByAttribute(layer, "NEW_SELECTION", where_clause)
    arcpy.management.SelectLayerByLocation(layer, "INTERSECT", layer, search_distance, "ADD_TO_SELECTION")
    arcpy.management.SelectLayerByAttribute(layer, "SWITCH_SELECTION")
    untouched_count = int(arcpy.management.GetCount(layer)[0])
    arcpy.AddMessage(f"  Untouched roads: {untouched_count}")

    temp_untouched = os.path.join(workspace, "temp_untouched_all")
    arcpy.management.CopyFeatures(layer, temp_untouched)
    return temp_untouched


def create_result(temp_untouched, temp_joined):
    """
    Creates the final output feature class by combining untouched roads with the planarized roundabout features. 
    Field mapping ensures only relevant attributes (osm_id, highway, junction) are transferred.
    """
    arcpy.AddMessage("Creating the result layer...")
    result_fc = os.path.join(workspace, road_fc + "_edited")
    if arcpy.Exists(result_fc):
        arcpy.management.Delete(result_fc)

    arcpy.management.CopyFeatures(temp_untouched, result_fc)

    arcpy.AddMessage("Starting to append planarized features")
    field_mappings = arcpy.FieldMappings()
    field_map_pairs = {"osm_id": "osm_id", "highway": "highway", "junction": "junction"}

    for target_field, source_field in field_map_pairs.items():
        fm = arcpy.FieldMap()
        fm.addInputField(temp_joined, source_field)
        out_field = fm.outputField
        out_field.name = target_field
        fm.outputField = out_field
        field_mappings.addFieldMap(fm)

    arcpy.management.Append(temp_joined, result_fc, "NO_TEST", field_mappings)
    return result_fc


def cleanup_temp(temp_files, layer):
    """
    Deletes all temporary datasets created during processing and clears any remaining selection on the input layer.
    """
    arcpy.AddMessage("Cleaning up temporary layers")
    for temp in temp_files:
        arcpy.management.Delete(temp)
    arcpy.management.SelectLayerByAttribute(layer, "CLEAR_SELECTION")



# Create a feature layer from the road dataset
lyr, oid_field = make_layer(road_fc)

# Select all roundabouts
where_clause = select_roundabouts(lyr)

# Add roads touching the selected roundabouts
selected_oids = add_touching_roads(lyr)

# Planarize selected features and transfer attributes
temp_input, temp_output, temp_joined = feature_to_line_selected(lyr)

# Select roads outside roundabouts
temp_untouched = select_untouched(lyr, where_clause)

# Merge planarized roundabouts with untouched roads
result_fc = create_result(temp_untouched, temp_joined)

# Delete temporary datasets and clear selection
cleanup_temp([temp_input, temp_output, temp_joined, temp_untouched], lyr)

# Count features in the final result
final_count = int(arcpy.management.GetCount(result_fc)[0])


final_count = int(arcpy.management.GetCount(result_fc)[0])
arcpy.AddMessage("=" * 50)
arcpy.AddMessage(f"Result saved to: {result_fc}")
arcpy.AddMessage(f"   Original feature count:  {int(arcpy.management.GetCount(road_fc)[0])}")
arcpy.AddMessage(f"   New feature count:       {final_count}")
arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
arcpy.AddMessage("=" * 50)
