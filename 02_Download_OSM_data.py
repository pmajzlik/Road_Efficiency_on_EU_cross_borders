#------------------------------------
# Name: Download_OSM_data.py
# Description: The script is used to download road and place data from the Overpass API for a specified bounding box.
# Author: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------

import arcpy
import requests
import json
import os
import traceback


# Highway types allowed in the NER analysis
ALLOWED_HIGHWAY = {
    "motorway", "motorway_link",
    "trunk", "trunk_link",
    "primary", "primary_link",
    "secondary", "secondary_link",
    "tertiary", "tertiary_link",
    "unclassified", "residential",
    "living_street"
}


def get_parameters():
    """
    Loads and returns parameters from the ArcGIS toolbox.
    """
    out_gdb         = arcpy.GetParameterAsText(0)
    clip_feature    = arcpy.GetParameterAsText(1)
    output_name     = arcpy.GetParameterAsText(2)
    timeout         = arcpy.GetParameter(3)
    country_polygons = arcpy.GetParameterAsText(4)  

    timeout = int(timeout) if timeout else 180

    arcpy.AddMessage(f"Output geodatabase: {out_gdb}")
    arcpy.AddMessage(f"Clip polygon: {clip_feature}")
    arcpy.AddMessage(f"Country polygons: {country_polygons}")

    return out_gdb, clip_feature, output_name, timeout, country_polygons


def get_bbox(clip_feature):
    """
    Returns the bounding box (bbox string) from a polygon in Overpass API format.
    """
    desc   = arcpy.Describe(clip_feature)
    extent = desc.extent

    bbox = f"{extent.YMin},{extent.XMin},{extent.YMax},{extent.XMax}"
    arcpy.AddMessage(
        f"Bounding box for download: "
        f"{extent.XMin}, {extent.YMin}, {extent.XMax}, {extent.YMax}"
    )
    return bbox


def download_osm_data(bbox, timeout):
    """
    Downloads roads and places from the Overpass API for the given bbox.
    Only downloads highway types relevant for NER analysis.
    Returns a list of elements (list of dicts).
    """
    # Build highway filter string from ALLOWED_HIGHWAY set
    highway_filter = "|".join(sorted(ALLOWED_HIGHWAY))

    overpass_query = f"""
    [out:json][timeout:{timeout}];
    (
      way["highway"~"{highway_filter}"]({bbox});
      way["junction"="roundabout"]({bbox});
      node["place"~"city|town|village"]["population"]({bbox});
      way["place"~"city|town|village"]["population"]({bbox});
    );
    out body;
    >;
    out skel qt;
    """

    arcpy.AddMessage("Sending query to Overpass API...")
    arcpy.AddMessage("Please wait, download may take several minutes...")
    arcpy.AddMessage(f"Downloading only these highway types: {', '.join(sorted(ALLOWED_HIGHWAY))}")

    response = requests.post(
        "http://overpass-api.de/api/interpreter",
        data={"data": overpass_query},
        timeout=timeout,
    )
    response.raise_for_status()

    elements    = response.json().get("elements", [])
    way_count   = sum(1 for e in elements if e["type"] == "way" and e.get("tags", {}).get("highway"))
    place_count = sum(1 for e in elements if e.get("tags", {}).get("place"))

    arcpy.AddMessage(f"Retrieved {way_count} roads and {place_count} places from OSM")

    if way_count == 0 and place_count == 0:
        arcpy.AddWarning("No roads or places were found in the specified area!")
        arcpy.AddWarning("Please check that the extent is correctly defined.")

    return elements


def create_roads_fc(out_gdb, output_name):
    """Creates (or overwrites) a feature class for roads. Returns the path to the FC."""
    spatial_ref = arcpy.SpatialReference(4326)
    output_fc   = os.path.join(out_gdb, output_name)

    if arcpy.Exists(output_fc):
        arcpy.AddMessage(f"Deleting existing feature class: {output_name}")
        arcpy.Delete_management(output_fc)

    arcpy.AddMessage("Creating feature class for roads...")
    arcpy.CreateFeatureclass_management(out_gdb, output_name, "POLYLINE",
                                        spatial_reference=spatial_ref)

    arcpy.AddMessage("Adding attribute fields for roads...")
    arcpy.AddField_management(output_fc, "osm_id",  "TEXT", field_length=50)
    arcpy.AddField_management(output_fc, "highway", "TEXT", field_length=50)
    arcpy.AddField_management(output_fc, "junction", "TEXT", field_length=50)

    return output_fc


def create_places_fc(out_gdb, output_name):
    """
    Creates (or overwrites) a feature class for places. Returns the path to the FC.
    """
    spatial_ref   = arcpy.SpatialReference(4326)
    places_name   = output_name + "_places"
    output_places = os.path.join(out_gdb, places_name)

    if arcpy.Exists(output_places):
        arcpy.Delete_management(output_places)

    arcpy.AddMessage("Creating feature class for settlements...")
    arcpy.CreateFeatureclass_management(out_gdb, places_name, "POINT",
                                        spatial_reference=spatial_ref)

    arcpy.AddMessage("Adding attribute fields for settlements...")
    arcpy.AddField_management(output_places, "osm_id",     "TEXT", field_length=50)
    arcpy.AddField_management(output_places, "place",      "TEXT", field_length=50)
    arcpy.AddField_management(output_places, "name",       "TEXT", field_length=200)
    arcpy.AddField_management(output_places, "population", "LONG")
    arcpy.AddField_management(output_places, "country",    "TEXT", field_length=10)

    return output_places


def build_nodes_dict(elements):
    """
    Builds a dictionary {node_id: (lon, lat)} from the downloaded elements.
    """
    nodes = {
        elem["id"]: (elem["lon"], elem["lat"])
        for elem in elements
        if elem["type"] == "node"
    }
    arcpy.AddMessage(f"Processed {len(nodes)} nodes")
    return nodes


def insert_roads(output_fc, elements, nodes):
    """
    Inserts roads into the feature class.
    Only roads matching ALLOWED_HIGHWAY types are inserted.
    Returns (inserted count, skipped count).
    """
    spatial_ref = arcpy.SpatialReference(4326)
    count = skipped = 0

    arcpy.AddMessage("Inserting roads into feature class...")

    with arcpy.da.InsertCursor(output_fc, ["SHAPE@", "osm_id", "highway", "junction"]) as cursor:
        for elem in elements:
            if elem["type"] != "way" or "nodes" not in elem:
                continue
            tags = elem.get("tags", {})
            if "highway" not in tags:
                continue

            # Filter: only allowed highway types
            if tags.get("highway") not in ALLOWED_HIGHWAY:
                skipped += 1
                continue

            coords = [nodes[nid] for nid in elem["nodes"] if nid in nodes]

            if len(coords) < 2:
                skipped += 1
                continue

            try:
                array = arcpy.Array([arcpy.Point(x, y) for x, y in coords])
                geom  = arcpy.Polyline(array, spatial_ref)
                cursor.insertRow([geom, str(elem["id"]), tags.get("highway", ""), tags.get("junction", "")])
                count += 1
                if count % 100 == 0:
                    arcpy.AddMessage(f"Processed {count} roads...")
            except Exception as e:
                skipped += 1
                arcpy.AddWarning(f"Error processing way {elem['id']}: {e}")

    return count, skipped


def insert_places(output_places, elements):
    """
    Inserts places (population > 5000) into the feature class. Returns (inserted count, skipped count).
    """
    spatial_ref  = arcpy.SpatialReference(4326)
    places_count = places_skipped = 0
    place_fields = ["SHAPE@", "osm_id", "place", "name", "population", "country"]

    arcpy.AddMessage("Inserting settlements into feature class...")

    with arcpy.da.InsertCursor(output_places, place_fields) as cursor:
        for elem in elements:
            tags = elem.get("tags", {})
            if "place" not in tags:
                continue

            try:
                pop_str = "".join(filter(str.isdigit, tags.get("population", "0")))
                if not pop_str:
                    continue
                population = int(pop_str)
                if population <= 5000:
                    continue
            except (ValueError, TypeError):
                places_skipped += 1
                continue

            if elem["type"] != "node":
                continue

            try:
                geom = arcpy.PointGeometry(
                    arcpy.Point(elem["lon"], elem["lat"]), spatial_ref
                )
                cursor.insertRow([
                    geom,
                    str(elem["id"]),
                    tags.get("place", ""),
                    tags.get("name", ""),
                    population,
                    "",   
                ])
                places_count += 1
            except Exception as e:
                places_skipped += 1
                arcpy.AddWarning(f"Error processing place {elem['id']}: {e}")

    return places_count, places_skipped


def assign_country_to_places(output_places, country_polygons, out_gdb):
    """
    Assigns a country code to each settlement using a spatial join with a country polygon layer containing the CNTR_CODE field.

    Parameters:
    output_places (str): Path to the settlement feature class
    country_polygons (str): Path to the country polygon feature class with the CNTR_CODE field
    out_gdb (str): Path to the geodatabase for temporary outputs
    """
    arcpy.AddMessage("Přiřazuji kódy zemí sídlům pomocí prostorového joinu...")

    joined = os.path.join(out_gdb, "places_joined_temp")

    if arcpy.Exists(joined):
        arcpy.Delete_management(joined)

    # Spatial join – each settlement receives the attributes of the polygon it falls into
    arcpy.SpatialJoin_analysis(
        target_features=output_places,
        join_features=country_polygons,
        out_feature_class=joined,
        join_operation="JOIN_ONE_TO_ONE",
        join_type="KEEP_ALL",
        match_option="INTERSECT"
    )

    # Read CNTR_CODE from the join result and store in a dictionary by the location name
    join_dict = {}
    with arcpy.da.SearchCursor(joined, ["name", "CNTR_CODE"]) as cursor:
        for row in cursor:
            if row[0] and row[1]:
                join_dict[row[0]] = row[1]

    arcpy.AddMessage(f"Found {len(join_dict)} settlements with assigned country codes.")

    # Write country codes back to the original feature class
    updated = 0
    missing = 0
    with arcpy.da.UpdateCursor(output_places, ["name", "country"]) as cursor:
        for row in cursor:
            name = row[0]
            if name in join_dict:
                row[1] = join_dict[name]
                cursor.updateRow(row)
                updated += 1
            else:
                missing += 1
                arcpy.AddWarning(f"Code not found for place: {name}")

    arcpy.AddMessage(f"Updated: {updated} settlements | No assignment: {missing} settlements")
    arcpy.Delete_management(joined)
    arcpy.AddMessage("Country codes were successfully assigned.")


def prepare_clip_feature(clip_feature, out_gdb):
    """
    If clip_feature is not in WGS84, reprojects it to a temporary layer.
    Returns (path to clip feature for use, flag indicating whether a temp file was created).
    """
    clip_sr = arcpy.Describe(clip_feature).spatialReference

    if clip_sr.factoryCode != 4326:
        arcpy.AddMessage("Reprojecting clip polygon to WGS84...")
        temp_clip = os.path.join(out_gdb, "temp_clip_wgs84")
        if arcpy.Exists(temp_clip):
            arcpy.Delete_management(temp_clip)
        arcpy.Project_management(clip_feature, temp_clip, arcpy.SpatialReference(4326))
        return temp_clip, True

    return clip_feature, False


def clip_layer(source_fc, clip_feature, out_gdb, result_name):
    """
    Clips source_fc by clip_feature and saves the result as result_name in out_gdb.
    Returns (path to result, feature count).
    """
    result_path = os.path.join(out_gdb, result_name)
    if arcpy.Exists(result_path):
        arcpy.Delete_management(result_path)

    arcpy.Clip_analysis(source_fc, clip_feature, result_path)

    count = int(arcpy.GetCount_management(result_path)[0])
    return result_path, count


def finalize_outputs(out_gdb, output_name, clipped_fc, clipped_places, temp_clip=None):
    """
    Renames clipped layers to their final names.
    Returns (path to roads, path to places).
    """
    final_fc     = os.path.join(out_gdb, output_name)
    final_places = os.path.join(out_gdb, output_name + "_places")

    for path in (final_fc, final_places):
        if arcpy.Exists(path):
            arcpy.Delete_management(path)

    arcpy.Rename_management(clipped_fc,     output_name)
    arcpy.Rename_management(clipped_places, output_name + "_places")

    if temp_clip and arcpy.Exists(temp_clip):
        arcpy.Delete_management(temp_clip)

    return final_fc, final_places


def main():
    # Parameters
    out_gdb, clip_feature, output_name, timeout, country_polygons = get_parameters()

    arcpy.AddMessage("Downloading roads from OSM (filtered highway types only)...")

    # Download raw OSM data within the bounding box of the clip polygon
    bbox     = get_bbox(clip_feature)
    elements = download_osm_data(bbox, timeout)

    # Create empty feature classes for roads and settlements
    output_fc     = create_roads_fc(out_gdb, output_name)
    output_places = create_places_fc(out_gdb, output_name)

    # Build node lookup and insert geometries into feature classes
    nodes                        = build_nodes_dict(elements)
    count,        skipped        = insert_roads(output_fc, elements, nodes)
    places_count, places_skipped = insert_places(output_places, elements)

    arcpy.AddMessage("=" * 50)
    arcpy.AddMessage(f"Downloaded {count} roads from bounding box")
    if skipped:
        arcpy.AddMessage(f"Skipped {skipped} roads (excluded highway type, missing nodes, or errors)")
    arcpy.AddMessage(f"Downloaded {places_count} settlements with population > 5000")
    if places_skipped:
        arcpy.AddMessage(f"Skipped {places_skipped} settlements (missing population or errors)")

    # Reproject clip polygon if needed, then clip both layers to the study area
    arcpy.AddMessage("=" * 50)
    clip_feature_to_use, temp_created = prepare_clip_feature(clip_feature, out_gdb)

    arcpy.AddMessage("Clipping roads by the specified polygon...")
    clipped_fc, clipped_count = clip_layer(
        output_fc, clip_feature_to_use, out_gdb, output_name + "_clipped"
    )
    arcpy.AddMessage(f"After clipping: {clipped_count} roads remain inside the polygon")

    arcpy.AddMessage("Clipping settlements by the specified polygon...")
    clipped_places, clipped_places_count = clip_layer(
        output_places, clip_feature_to_use, out_gdb, output_name + "_places_clipped"
    )
    arcpy.AddMessage(f"After clipping: {clipped_places_count} settlements remain inside the polygon")

    arcpy.Delete_management(output_fc)
    arcpy.Delete_management(output_places)

    # Rename clipped layers to final output names and clean up temp files
    temp_clip_path = clip_feature_to_use if temp_created else None
    final_fc, final_places = finalize_outputs(
        out_gdb, output_name, clipped_fc, clipped_places, temp_clip_path
    )

    # Assign country codes to settlements if country polygons were provided
    arcpy.AddMessage("=" * 50)
    if country_polygons:
        assign_country_to_places(final_places, country_polygons, out_gdb)
    else:
        arcpy.AddWarning("Country polygons were not provided – the 'country' field will remain empty.")


    arcpy.AddMessage(f"Final road count:       {clipped_count}")
    arcpy.AddMessage(f"Final settlement count: {clipped_places_count}")
    arcpy.AddMessage(f"Roads saved to:         {final_fc}")
    arcpy.AddMessage(f"Settlements saved to:   {final_places}")
    arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
    arcpy.AddMessage("=" * 50)

    arcpy.SetParameterAsText(0, final_fc)


if __name__ == "__main__":
    timeout = 180

    try:
        main()

    except requests.exceptions.Timeout:
        arcpy.AddError(f"Query exceeded the time limit of {timeout} seconds!")
        arcpy.AddError("Try increasing the timeout or reducing the area of interest.")

    except requests.exceptions.RequestException as e:
        arcpy.AddError(f"Error downloading data from OSM: {e}")
        arcpy.AddError("Check your internet connection and try again.")

    except Exception as e:
        arcpy.AddError(f"Unexpected error: {e}")
        arcpy.AddError(traceback.format_exc())
        arcpy.AddError("Check the parameters and try again.")