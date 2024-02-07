
############################################################
# INPUT FOLDER PARAMETER - replace as appropriate
############################################################

# Replace the full folder path below with the path to the folder that contains
# one or more Excel files that include Primary county FIPS codes.
# (Do not include the name of an Excel file itself, ONLY THE CONTAINING FOLDER)
# The folder may have one OR MORE Excel files, and the files may contain multiple sheets.
# One map will be generated per sheet per Excel file (i.e., a folder containing two
# Excel files with three sheets each will produce six maps).
# Be sure to include the leading and trailing double quotes!

input_folder = r"C:\Users\misti.wudtke\OneDrive - USDA\PROJECTS\FSA_AutoMap_v1_2023Sept\Input_XLSXs"


# If you want to script to delete all files in the output folders
# before the script is re-run, set this to True; otherwise to False
# Note that even if this is set to False, the script
# WILL STILL OVERWRITE files in the output locations with the same name!
# Also note that this deletes BOTH OUTPUT SUBFOLDERS (before recreating)
# So don't store other files you need in the output folders!

del_prior_outputs = True


################################################################################
# DO THE WORK
################################################################################

import arcpy
import os
import shutil
import pandas as pd

############################################################

# Print a message and also use addmessage method for toolbox use
def print_message(message):
    
    print(message)
    arcpy.AddMessage(message)


############################################################

# Make an output folder to store output file types
def make_folders(output_folder):

    # Define the full path string
    output_path = os.path.join(input_folder, output_folder)
    
    # If the output folder already exist, do some checking:
    if os.path.exists(output_path):
        
        # If we are supposed to delete prior outputs, do that,
        # then re-make the folder
        if del_prior_outputs:
            shutil.rmtree(output_path)
            os.mkdir(output_path)
    
    # Otherwise if folder do NOT already exist, just make it!
    else:
        os.mkdir(output_path)
        
    # Ship the full output path back to the rest of the script
    return output_path

        
############################################################

# In the current map in the current project open in ArcGIS Pro,
# Get the service URL of the layer named "US Counties"
def get_counties():

    # Get the current project
    current_project = arcpy.mp.ArcGISProject("CURRENT")

    # Get the map in the project named "Main Map"
    main_map = current_project.listMaps("Main Map")[0]

    # Get the layer in the map named "US Counties"
    counties_layer = main_map.listLayers("US Counties")[0]

    # Print status message
    print_message("Found US Counties layer in Main Map")

    # Return the variable "counties_layer" for use
    # in the rest of the script
    return counties_layer
        

############################################################

# Build queries to select either Primary, Contiguous or Both counties
def build_queries(counties_layer):

    # Add the appropriate field delimiters to the field name
    # Proper delimiter depends on file format so we do it this way
    delimited = arcpy.AddFieldDelimiters(counties_layer, "CLASS")

    primary_query = f"{delimited} = 'Primary'"

    contiguous_query = f"{delimited} = 'Contiguous'"

    both_query = f"{delimited} IN ('Primary', 'Contiguous')"
    
    return [primary_query, contiguous_query, both_query]


############################################################

# This should only need to be called once.
# The output will be a dictionary with map names as the key
# and a list of the primary county fips codes as the values.
def get_fips(input_folder):
    
    # Get a list of all files in the folder
    file_list = os.listdir(input_folder)

    # Filter the list to include only Excel files
    excel_files = [file for file in file_list if file.endswith('.xlsx')]
    
    # Dictionary to store the data
    fips_dict = {}

    # Iterate through Excel files
    for excel_file in excel_files:
        file_path = os.path.join(input_folder, excel_file)

        # Read all sheets in the Excel file
        excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)

        # Iterate through sheets
        for sheet_name, sheet_data in excel_data.items():

            # Ignore empty sheets
            if sheet_data.empty:
                continue

            #  Extract the first column header
            column_header = sheet_data.columns[0]

            # Extract non-blank values from the first column
            values = sheet_data[column_header].dropna().tolist()

            # Format excel document name for use in map title
            # Chop off the .xlsx extension and replace underscores with spaces
            excel_title = excel_file.rsplit(".", 1)[0].replace("_", " ")
            sheet_title = sheet_name.replace("_", " ")
            
            # Use both title of Excel document and sheet name
            # to build unique map title
            key = excel_title + " " + str(sheet_title)

            # Add key-value pair to the dictionary
            fips_dict[key] = values
    
    # Print status message
    print_message("\nFinished generating dictionary of map titles and county fips code lists")
    
    return fips_dict

    
############################################################

#  This should be called once per map layout.
#  After the map is exported, flip all the counties back to "Not Selected"
#  We don't want selected counties from previous calls impinging on current call.

#  Update cursor loops through all counties;
#  For any row with a matching value in the current get_fips key/value pair,
#  flip the Class attribute to "Primary"
def code_primary(counties_layer, fips_list):

    #  Fields for update cursor
    cursor_fields = ["FIPS_C", "CLASS", "ObjctID"]
    
    # Run through the cursor and code all counties in the fips list as Primary
    with arcpy.da.UpdateCursor(counties_layer, cursor_fields) as update_cursor:
        for row in update_cursor:
            if row[0] in fips_list:
                row[1] = "Primary"
                row[2] = 1
                update_cursor.updateRow(row)
        
    # Print status message
    print_message("Finished coding Primary counties")
        

############################################################

#  This should also be called once per map layout.
#  Select all primary counties; select all counties adjacent to primary;
#  Flip the class attribute to "Contiguous"
#  Additional senanigans: apply 5-mile buffer; select all additional counties;
#  Flip Class attribute to "Countiguous"; may need to check for any counties
#  selected that are within 5 miles but NOT contiguous, NOT across water 
def code_contiguous(counties_layer, key, queries, use_buffer, maps_to_check):
    
    # Name for new Primary counties only layer
    primary_counties = "Primary Counties"
    
    # Make Feature Layer for use with intersect GP tool
    arcpy.management.MakeFeatureLayer(counties_layer, primary_counties, queries[0])
    
    # Select all counties from original all counties layer that intersect primary counties
    # If use_buffer is True, the 5-mile rule will be implemented
    if use_buffer:
        arcpy.management.SelectLayerByLocation(counties_layer, "WITHIN_A_DISTANCE_GEODESIC", primary_counties, "5 Miles")
    
    # If use_buffer is False, NO 5-mile rule will be implemented
    else:
        arcpy.management.SelectLayerByLocation(counties_layer, "INTERSECT", primary_counties)
    
    # All intersecting counties, including primary counties, are selected;
    # We do not want to re-code Primary counties so remove them from selection
    arcpy.management.SelectLayerByAttribute(counties_layer, "REMOVE_FROM_SELECTION", queries[0])
    
    # Calculate all selected counties to "Contiguous"
    arcpy.management.CalculateField(counties_layer, "CLASS", "'Contiguous'")
    arcpy.management.CalculateField(counties_layer, "ObjctID", "0")
    
    # Count the number of counties coded as contiguous
    contiguous_count = arcpy.management.GetCount(counties_layer)[0]
    
    # TO-DO
    maps_to_check[key].append(int(contiguous_count))
    
    # Delete temp layer for only Primary counties
    arcpy.management.Delete(primary_counties)
    
    # Print status message
    if use_buffer:
        print_message("Finished coding Contiguous counties with 5mi buffer")
    else:
        print_message("Finished coding Contiguous counties")

    
############################################################

#  Not sure if this is going to make use of map series or not
#  Literally just zoom to map extent in layout template, populate title, export
def export_map(counties_layer, key, queries, output_folders, use_buffer):
    
    arcpy.env.workspace = output_folders[0]
    
    # Get the current project
    current_project = arcpy.mp.ArcGISProject("CURRENT")

    # Get map layout
    layout = current_project.listLayouts("Layout")[0]
    
    # Get map frame layout element
    map_frame = layout.listElements("MAPFRAME_ELEMENT", "Main Map")[0]    
    
    # Select counties currently coded as contiguous
    arcpy.management.SelectLayerByAttribute(counties_layer, "NEW_SELECTION", queries[2])
    
    # Get extent of selected features to use
    contiguous_extent = map_frame.getLayerExtent(counties_layer, True, True)
    
    # Apply extent to layout map frame
    map_frame.camera.setExtent(contiguous_extent)
    
    # Clear selected features
    arcpy.management.SelectLayerByAttribute(counties_layer, "CLEAR_SELECTION")
    
    # Zoom out a little so contiguous counties are not right up to the edge of the map frame
    map_frame.camera.scale = map_frame.camera.scale * 1.2
    
    # from the current Excel map sheet or wherever it comes from
    layout.listElements("TEXT_ELEMENT", "Title")[0].text = key
    layout.listElements("TEXT_ELEMENT", "Subtitle")[0].text = key
    
    # **************************************************
    # TO-DO (OPTIONAL): Add option in toolbox to choose between JPG or PDF export
    # **************************************************
    
    # Export layout
    # Add buffer to file output name if buffer has been implemented in script
    if use_buffer:
        layout.exportToPDF(os.path.join(output_folders[0], key + " 5mi buffer.pdf"), resolution=200)
    
    else:
        layout.exportToPDF(os.path.join(output_folders[0], key + ".pdf"), resolution=200)
    
    # Print status message
    if use_buffer:
        print_message("Finished exporting the current map with 5mi buffer")
    else:
        print_message("Finished exporting the current map")
    
    
############################################################

# Export xlsx files to send wherever they need to go downstream
# With future iterations of the workflow this will not be required
def export_excel(counties_layer, key, queries, output_folders, use_buffer):
    
    arcpy.env.workspace = output_folders[1]
    
    # Select counties currently coded as either primary or contiguous
    arcpy.management.SelectLayerByAttribute(counties_layer, "NEW_SELECTION", queries[2])
    
    # Construct full file name for output dpf files
    # Add buffer to file output name if buffer has been implemented in script
    if use_buffer:
        output_excel = key + " 5mi buffer.xlsx"
        
    else:
        output_excel = key + ".xlsx"
    
    # Do the work - send exels to output folder
    arcpy.conversion.TableToExcel(counties_layer, output_excel)
    
    # Clear selected features
    arcpy.management.SelectLayerByAttribute(counties_layer, "CLEAR_SELECTION")
    
    # Print status message
    if use_buffer:
        print_message("Finished exporting the Excel files with 5mi buffer")
    else:
        print_message("Finished exporting the Excel files")
    

############################################################

#  Use field calc to reset all values in counties back to "Not Selected"
#  Update: Calculate Field was throwing an error for which Esri had no solution,
#  So switched to UpdateCursor instead. Still plenty zippy.
def reset_class(counties_layer, queries, use_buffer):
    
    if use_buffer:
        current_query = queries[1]
    else:
        current_query = queries[2]
    
    with arcpy.da.UpdateCursor(counties_layer, ["CLASS", "ObjctID"], current_query) as update_cursor:
        for row in update_cursor:
            row[0] = "Not Selected"
            row[1] = 0
            update_cursor.updateRow(row)
        
    # Print status message
    if use_buffer:
        print_message("Finished resetting CLASS attribute with 5mi buffer")
    else:
        print_message("Finished resetting CLASS attribute")
    

############################################################

def contiguous_subcall(counties_layer, key, output_folders, use_buffer, maps_to_check):
    
    # Call to function to build the required SQL queries
    queries = build_queries(counties_layer)

    # Call to function to code all counties contiguous to Primary counties as such
    code_contiguous(counties_layer, key, queries, use_buffer, maps_to_check)

    # Call to function to export map
    export_map(counties_layer, key, queries, output_folders, use_buffer)

    # Call to function to export xlsx files
    export_excel(counties_layer, key, queries, output_folders, use_buffer)

    # Call to function to reset class
    reset_class(counties_layer, queries, use_buffer)
    

############################################################

#  This should mostly consist of iterating through our FIPS dictionary
#  And calling code_primary, code_contiguous and export_map for each key/value pair
def iterate_maps(counties_layer, fips_dict):
    
    maps_to_check = {}
    
    # Provide string names for our two output folders
    output_pdfs = make_folders("Output PDFs")
    output_excels = make_folders("Output Excels")
    
    # Stuff both in a list for easier passing to various functions
    output_folders = (output_pdfs, output_excels)
    
    # "key" is every Excel sheet (which will be a map) and "value" is list of 5-digit FIPS codes
    for key, value in fips_dict.items():
        
        maps_to_check[key] = []

        # Print status message
        print_message(f"\nWorking on map title '{key}'")

        # Call to function to iterate through counties layer and flip Class attr of all Primary counties
        code_primary(counties_layer, value)

        # First call to contiguous sub-fuction, iterates through sub-processes WITH 5mi buffer
        contiguous_subcall(counties_layer, key, output_folders, True, maps_to_check)
        
        # Second call to contiguous sub-fuction, iterates through sub-processes WITHOUT 5mi buffer
        contiguous_subcall(counties_layer, key, output_folders, False, maps_to_check)
    
        # Print status message
        print_message(f"Finished map title '{key}'")
        
    return maps_to_check
    

############################################################

# Run through the dictionary we generated to check for maps
# that may need to be adjusted manually.
# For all these maps, the number of contiguous counties WITH
# the 5mi boundary and the number contiguous counties WITHOUT
# the buffer is 2 or more...so the buffer may apply to one
# but not the other. Script cannot fix automatically yet.
def find_manual_maps(maps_to_check):

    print_message("\n\n****************************")
    print_message("**** MAP OUTPUT REPORT: ****")
    print_message("****************************\n")
    
    buffer_on = "\tContiguous counties with buffer ON:"
    buffer_off = "\tContiguous counties with buffer OFF:"
    
    for k, v in maps_to_check.items():
        if v[0] - v[1] >= 2:
            print_message(f"** POSSIBLE MANUAL MAP: **\n{k}")
            print_message(f"\tDifference buffered vs unbuffered")
            print_message(f"\tcontiguous counties is 2 or more")
            print_message(f"{buffer_on} {v[0]}")
            print_message(f"{buffer_off} {v[1]}")
            print_message(f"\tTotal difference: {v[0]-v[1]}\n")
            
        elif v[0] - v[1] == 1:
            print_message(f"CHOOSE BUFFERED OR UNBUFFERED MAP:\n{k}")
            print_message(f"{buffer_on} {v[0]}")
            print_message(f"{buffer_off} {v[1]}\n")

        else:
            print_message(f"IDENTICAL MAPS:\n{k}")
            print_message(f"{buffer_on} {v[0]}")
            print_message(f"{buffer_off} {v[1]}\n")

            
############################################################

def do_the_work():
    
    # Call to function to get the counties layer
    counties_layer = get_counties()

    # Call to function to get all the fips codes from all the sheets in all the excel files in a folder
    fips_dict = get_fips(input_folder)

    # Call to function to iterate fips_dict and export a single map
    # Various per-map functions are called from within this function
    maps_to_check = iterate_maps(counties_layer, fips_dict)

    # Print status message
    print_message("\nFinished generating all maps successfully")
    
    # Call to function to see whether any maps may need to be run semi-manually
    find_manual_maps(maps_to_check)
    
    
################################################################################
# DO THE WORK
################################################################################

# Call to the MAIN OVERARCHING function from which all other functions are called
# One call does it all...should make it easier to stuff in a toolbox later
do_the_work()

