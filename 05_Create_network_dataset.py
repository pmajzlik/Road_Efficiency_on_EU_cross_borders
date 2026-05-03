#------------------------------------
# Name: Create_network_dataset.py
# Description: The script is used to create a network dataset from a road network layer.
# Author: Petr MIKESKA, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2025
# Update: Bc. Pavel MAJZLIK, Department of Geoinformatics, Faculty of Science, Palacky University Olomouc, 2026
#------------------------------------

import arcpy
import os
import sys

arcpy.env.overwriteOutput = True

def check_network_extension():
    if arcpy.CheckExtension("Network") != "Available":
        arcpy.AddError("Network Analyst extension is not available.")
        sys.exit(1)
    arcpy.CheckOutExtension("Network")
    arcpy.AddMessage("Network Analyst extension OK.")

def validate_input(roads_layer, feat_dataset):
    desc = arcpy.Describe(roads_layer)
    if desc.shapeType != "Polyline":
        arcpy.AddError("Input layer must be of type Polyline.")
        sys.exit(1)

    sr = desc.spatialReference
    wgs84 = arcpy.SpatialReference(4326)

    if sr.factoryCode != 4326:
        arcpy.AddWarning(f"Input layer uses EPSG:{sr.factoryCode}. Reprojecting to EPSG:4326 (WGS 1984)...")
        reprojected = os.path.join(feat_dataset, "Roads_reprojected")
        arcpy.management.Project(roads_layer, reprojected, wgs84)
        arcpy.AddMessage("Reprojection completed successfully.")
        return reprojected, wgs84

    arcpy.AddMessage("Coordinate system check OK (EPSG:4326).")
    return roads_layer, sr

def network_dataset_exists(gdb_path, feature_dataset_name, network_name):
    """
    Checks the existence of a Network Dataset by Feature Dataset name and ND name.
    """
    fd_path = os.path.join(gdb_path, feature_dataset_name)
    if arcpy.Exists(fd_path):
        # Check if ND already exists under this name
        nd_path = os.path.join(fd_path, network_name)
        if arcpy.Exists(nd_path):
            return True
    # Check root GDB (ND may be directly in root)
    nd_path_root = os.path.join(gdb_path, network_name)
    if arcpy.Exists(nd_path_root):
        return True
    return False

def build_network_dataset(roads_layer, gdb_path, network_name):
    arcpy.AddMessage("=== START: Network Dataset Creation ===")

    check_network_extension()

    base_name = os.path.basename(roads_layer).replace(" ", "_")
    feature_dataset_name = f"{base_name}_FD"

    # --- Check if ND already exists ---
    if network_dataset_exists(gdb_path, feature_dataset_name, network_name):
        arcpy.AddError(f"Network Dataset '{network_name}' already exists in database '{gdb_path}'. Script terminated.")
        sys.exit(1)

    # --- Feature Dataset ---
    # Check if any Feature Dataset already exists in the GDB
    arcpy.env.workspace = gdb_path
    existing_datasets = arcpy.ListDatasets("*", "Feature")
    if existing_datasets:
        arcpy.AddError(f"Geodatabase '{gdb_path}' already contains a Feature Dataset: '{existing_datasets[0]}'. Please use a different geodatabase.")
        sys.exit(1)

    # Feature Dataset must exist before validate_input, because reprojected layer is saved into it
    feat_dataset = os.path.join(gdb_path, feature_dataset_name)
    temp_sr = arcpy.Describe(roads_layer).spatialReference
    arcpy.CreateFeatureDataset_management(gdb_path, feature_dataset_name, temp_sr)
    arcpy.AddMessage(f"Feature Dataset '{feature_dataset_name}' created.")

    # --- Validate input and reproject if necessary ---
    roads_layer, sr = validate_input(roads_layer, feat_dataset)

    # --- If reprojection happened, recreate Feature Dataset with correct WGS84 SR ---
    if sr.factoryCode == 4326:
        existing_fd_sr = arcpy.Describe(feat_dataset).spatialReference
        if existing_fd_sr.factoryCode != 4326:
            arcpy.AddMessage("Recreating Feature Dataset with EPSG:4326 spatial reference...")
            arcpy.management.Delete(feat_dataset)
            arcpy.CreateFeatureDataset_management(gdb_path, feature_dataset_name, sr)
            arcpy.AddMessage(f"Feature Dataset '{feature_dataset_name}' recreated with EPSG:4326.")

    # --- Copy roads ---
    roads_name = "Roads"
    roads_fd_path = os.path.join(feat_dataset, roads_name)
    if not arcpy.Exists(roads_fd_path):
        arcpy.FeatureClassToFeatureClass_conversion(roads_layer, feat_dataset, roads_name)
        arcpy.AddMessage(f"Roads copied as '{roads_name}'.")
    else:
        arcpy.AddMessage(f"Roads '{roads_name}' already exist, copying skipped.")

    # --- Delete temporary reprojected layer if it exists ---
    reprojected_path = os.path.join(feat_dataset, "Roads_reprojected")
    if arcpy.Exists(reprojected_path):
        arcpy.management.Delete(reprojected_path)
        arcpy.AddMessage("Temporary reprojected layer deleted.")

    # --- Create ND ---
    nd_path = os.path.join(feat_dataset, network_name)
    arcpy.AddMessage("Creating Network Dataset...")
    try:
        arcpy.na.CreateNetworkDataset(
            feature_dataset=feat_dataset,
            out_name=network_name,
            source_feature_class_names=roads_name,
            elevation_model="ELEVATION_FIELDS"
        )
        arcpy.na.BuildNetwork(nd_path)
        arcpy.AddMessage("Network Dataset successfully created and built.")
    except arcpy.ExecuteError:
        arcpy.AddError("Failed to create ND - a Network Dataset with the same road source likely already exists.")
        sys.exit(1)

    arcpy.AddMessage(f"GDB:     {gdb_path}")
    arcpy.AddMessage(f"ND Path: {nd_path}")
    arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
    arcpy.AddMessage("IMPORTANT – Manual Configuration Required The road features in the database already contain a pre-calculated travel duration for each segment (the TravelTime attribute). However, the Network Dataset does not use this information automatically — it must be explicitly registered as a cost input before any routing calculations can use it. You must complete the following steps manually in ArcGIS Pro:"
    "1. Open the Network Dataset Properties and navigate to Travel Attributes → Costs. Add a new cost attribute representing travel time. " \
    "2. Configure the evaluator for this attribute: select Field Script as the source type for edges, and enter !TravelTime! as the field expression. Confirm and save." \
    "3. Run the Build Network geoprocessing tool to rebuild the updated network." \
    "Do not proceed to the next step until the network has been successfully rebuilt.")

    arcpy.AddMessage("=" * 60)
    return nd_path

if __name__ == "__main__":
    roads_layer  = arcpy.GetParameterAsText(0)
    gdb_path     = arcpy.GetParameterAsText(1)
    network_name = arcpy.GetParameterAsText(2)

    build_network_dataset(roads_layer, gdb_path, network_name)
