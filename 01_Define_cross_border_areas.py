#------------------------------------
# Name: Define_cross_border_areas.py
# Description: The script is used to define cross-border areas based on the intersection of country boundaries and a buffer around international borders.
# Author: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------

import arcpy
import os

arcpy.env.overwriteOutput = True

# Field name containing the country code in the input polygon layer
CNTR_FIELD = "CNTR_CODE"

# Temporary in-memory workspace for intermediate outputs
TEMP_WS = "in_memory"


def get_workspace_info(input_polygons, output_gdb):
    """
    Returns (workspace, layer_name, spatial_reference, border_fc, buffer_fc).
    """

    desc = arcpy.Describe(input_polygons)
    layer_name = os.path.splitext(desc.name)[0]
    sr = desc.spatialReference

    workspace = output_gdb
    border_fc = os.path.join(workspace, f"{layer_name}_border")
    buffer_fc = os.path.join(workspace, f"{layer_name}_border_buffer25km")

    return workspace, layer_name, sr, border_fc, buffer_fc


def polygons_to_lines(input_polygons):
    """
    Converts polygons to lines with neighbor identification.
    Returns the path to a temporary feature class.
    """

    arcpy.AddMessage("Converting polygons to lines...")
    poly_to_line = os.path.join(TEMP_WS, "poly_to_line")

    arcpy.management.PolygonToLine(
        in_features=input_polygons,
        out_feature_class=poly_to_line,
        neighbor_option="IDENTIFY_NEIGHBORS",
    )
    return poly_to_line


def build_fid_to_cntr(input_polygons):
    """
    Builds a dictionary {OID: CNTR_CODE} from the input polygons.
    """

    fid_to_cntr = {}

    with arcpy.da.SearchCursor(input_polygons, ["OID@", CNTR_FIELD]) as cursor:
        for oid, cntr in cursor:
            fid_to_cntr[oid] = cntr

    return fid_to_cntr


def extract_international_borders(poly_to_line, fid_to_cntr, border_fc, sr):
    """
    Selects lines where LEFT and RIGHT belong to different countries
    and saves them to border_fc.
    """

    arcpy.AddMessage("Selecting international borders...")
    workspace = os.path.dirname(border_fc)
    border_name = os.path.basename(border_fc)

    arcpy.management.CreateFeatureclass(
        out_path=workspace,
        out_name=border_name,
        geometry_type="POLYLINE",
        spatial_reference=sr,
    )

    insert_cursor = arcpy.da.InsertCursor(border_fc, ["SHAPE@"])
    with arcpy.da.SearchCursor(poly_to_line, ["LEFT_FID", "RIGHT_FID", "SHAPE@"]) as cursor:

        for left_fid, right_fid, shape in cursor:

            if left_fid == -1 or right_fid == -1:
                continue

            left_cntr = fid_to_cntr.get(left_fid)
            right_cntr = fid_to_cntr.get(right_fid)

            if left_cntr and right_cntr and left_cntr != right_cntr:
                insert_cursor.insertRow([shape])

    del insert_cursor

    count = int(arcpy.management.GetCount(border_fc)[0])

    if count == 0:
        arcpy.AddWarning("No shared borders between countries were found!")
    else:
        arcpy.AddMessage(f"Found {count} border segments")

    return count

def dissolve_borders(border_fc):
    """Merges all border segments into a single geometry."""
    arcpy.AddMessage("Dissolving borders...")
    dissolved = os.path.join(TEMP_WS, "dissolved_border")
    arcpy.management.Dissolve(border_fc, dissolved)

    arcpy.management.Delete(border_fc)
    arcpy.management.CopyFeatures(dissolved, border_fc)


def create_clipped_buffer(border_fc, input_polygons, buffer_fc):
    """
    Creates a 25 km buffer around borders
    and clips it to country boundaries.
    """
    arcpy.AddMessage("Creating 25 km buffer...")
    buffer_temp = os.path.join(TEMP_WS, "buffer_temp")
    arcpy.analysis.Buffer(
        in_features=border_fc,
        out_feature_class=buffer_temp,
        buffer_distance_or_field="25 Kilometers",
        line_side="FULL",
        line_end_type="ROUND",
        dissolve_option="ALL",
    )

    arcpy.AddMessage("Clipping buffer to country boundaries...")
    dissolved_polygons = os.path.join(TEMP_WS, "dissolved_polygons")
    arcpy.management.Dissolve(input_polygons, dissolved_polygons)
    arcpy.analysis.Clip(
        in_features=buffer_temp,
        clip_features=dissolved_polygons,
        out_feature_class=buffer_fc,
    )
    arcpy.AddMessage("Buffer clipped to country boundaries")

def report_results(border_fc, buffer_fc, count):
    total_length = 0
    with arcpy.da.SearchCursor(border_fc, ["SHAPE@LENGTH"]) as cursor:
        for (length,) in cursor:
            total_length += length

    arcpy.AddMessage("--------------------------------------------")
    arcpy.AddMessage(f"Number of segments: {count}")
    arcpy.AddMessage(f"Output lines:  {os.path.basename(border_fc)}")
    arcpy.AddMessage(f"Output buffer: {os.path.basename(buffer_fc)}")

def main():

    arcpy.management.Delete("in_memory")

    input_polygons = arcpy.GetParameterAsText(0)
    output_folder = arcpy.GetParameterAsText(1)
    gdb_name = arcpy.GetParameterAsText(2)

    # Ensure GDB name ends with .gdb
    if not gdb_name.endswith(".gdb"):
        gdb_name += ".gdb"

    # GDB path
    gdb_path = os.path.join(output_folder, gdb_name)

    # Create GDB if it doesn't exist
    if not arcpy.Exists(gdb_path):
        arcpy.management.CreateFileGDB(output_folder, gdb_name)

    workspace, layer_name, sr, border_fc, buffer_fc = get_workspace_info(
        input_polygons, gdb_path
    )

    poly_to_line = polygons_to_lines(input_polygons)
    fid_to_cntr = build_fid_to_cntr(input_polygons)
    count = extract_international_borders(poly_to_line, fid_to_cntr, border_fc, sr)

    if count > 0:
        dissolve_borders(border_fc)
        create_clipped_buffer(border_fc, input_polygons, buffer_fc)
        report_results(border_fc, buffer_fc, count)

arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
arcpy.AddMessage("=" * 50)


if __name__ == "__main__":
    main()