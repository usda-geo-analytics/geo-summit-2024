
############################################################
# INPUT FOLDER PARAMETER - replace as appropriate
############################################################

# Replace the full folder path below with the path to the folder that contains
# one or more Excel files that include Primary county FIPS or Tribal codes.
# (Do not include the name of an Excel file itself, ONLY THE CONTAINING FOLDER)
# The folder may have one OR MORE Excel files, and the files may contain multiple sheets.
# One map will be generated per sheet per Excel file (i.e., a folder containing two
# Excel files with three sheets each will produce six maps).
# Be sure to include the leading and trailing double quotes!
# Update 4/24/24: If TRIBAL maps are being generated, the script produced one map 
# PER EXCEL FILE rather than one map per sheet in each excel file.

input_folder = r"C:\Users\misti.wudtke\OneDrive - USDA\PROJECTS\FSA_AutoMap_v3_100Automated\Input_XLSXs"


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
import sys
import pandas as pd

############################################################

# Print a message and also use addmessage method for toolbox use
def print_message(message):
    
    print(message)
    arcpy.AddMessage(message)


############################################################

# In the current map in the current project open in ArcGIS Pro,
# get the layer objects for US Counties, Counties including
# water areas (direct from Census) and Tribal Lands
def get_layers():

    # Get the current project
    current_project = arcpy.mp.ArcGISProject("CURRENT")

    # Get the map in the project named "Main Map"
    main_map = current_project.listMaps("Main Map")[0]

    # Get the layer in the map named "US Counties"
    counties_layer = main_map.listLayers("US Counties")[0]
    
    # Get the layer in the map named "Tribal Lands"
    tribal_layer = main_map.listLayers("Tribal Lands")[0]
    
    # Get the layer in the map named "US Counties Water"
    water_layer = main_map.listLayers("US Counties Water")[0]
    
    print_message("Found the layers in the map")

    # Send both layer objects back to the rest of the script
    return [counties_layer, tribal_layer, water_layer]
        

############################################################

# Just in case anything has been randomly selected in any of the layers,
# Clear anything currently selected
def clear_all(map_layers):
    
    for map_layer in map_layers:
        
        arcpy.management.SelectLayerByAttribute(map_layer, "CLEAR_SELECTION")
        

############################################################

# Build queries to select either Primary, Contiguous or Both counties
def build_queries(in_layer):

    # Add the appropriate field delimiters to the field name
    # Proper delimiter depends on file format so we do it this way
    # (When this process transitions to using services things won't break)
    delimited = arcpy.AddFieldDelimiters(in_layer, "CLASS")

    primary_query = f"{delimited} = 'Primary'"

    contiguous_query = f"{delimited} = 'Contiguous'"

    both_query = f"{delimited} IN ('Primary', 'Contiguous')"
    
    return [primary_query, contiguous_query, both_query]


############################################################

# Have reconfigured the script to work to generate both tribal
# and non-tribal maps. Data for tribal maps is usually received in 
# a slightly different format: 1 map per excel file vs 1 map per excel sheet.
# Here we check whether the current script run is for tribal maps;
# if it is, reprocess the excel dictionary to work on a per/excel basis.
def check_for_tribal(fips_dict, excel_files):
    
    # Get list of excel files with the word "tribal" in the file name
    # (This is how we make the determination of whether this is a tribal run)
    tribal_files = [ef for ef in excel_files if "tribal" in ef.lower()]

    # If there's at least one file with the word "tribal" in the name,
    # continue with additional logic
    if tribal_files:

        # The script isn't set up to process a mix of tribal and non-tribal maps
        # So, confirm that, if there's ONE tribal file, ALL the files are tribal
        if not len(excel_files) == len(tribal_files):

            # If it's a mix, give the user a warning to divy up the files
            # and start the script over; exit script early after message.
            print("""
            \n\tAwe crap; it looks like there's a mix of tribal 
            and non-tribal excel files in the input folder.
            Since tribal and non-tribal are processed differently by the script,
            There needs to be only ALL tribal or ALL non-tribal excel files 
            in the folder for a given script run. Please fix, then run again!\n
            (Note: Tribal excel files must have the word "tribal" in the file name;
            non-tribal excel files must not have it.)
            """)
            print("\tExiting now! :)\n")
            sys.exit()

        # Assuming it's all-tribal-all-the-time, go ahead and reprocess dictionary
        tribal_dict = {}
        
        # Key values in original dict are composed of excel file name + sheet name;
        # since the map is generated per excel we only need the excel file name,
        # so split the current key where it was concatenated previously by "__"
        for k, val in fips_dict.items():
            tk = k.split("__")[0]
            
            # Then for all values in the dict where the key excel file name is the same,
            # smash all the values together in one total list:
            # Add the key/val pair to the new tribal dict if it isn't already in it...
            if not tk in tribal_dict:
                tribal_dict[tk] = val
            # If the key has already been added...
            else:
                # ...loop through values and append to current list of values
                for v in val:
                    tribal_dict[tk].append(v)
                    
        # Return the reprocessed dictionary and "True" for whether this is a tribal run
        return [tribal_dict, True]
    
    # If this isn't a tribal run, just return the original dict pluse "False" for tribal run
    else:
        return [fips_dict, False]


############################################################

# This is called once per script run.
# This function looks through the input folder (the only script parameter),
# assembles a list of the excel files in it, then extract the first column of data
# in each sheet of the excel file and smashes it into a list as the value of a dictionary.
# (first column of data is usually county fips codes but may be tribal codes for tribal maps)
# One key/value pair in the dictionary equals one output map.
def get_fips():
    
    # Get a list of all files in the folder
    file_list = os.listdir(input_folder)

    # Dictionary to store the data
    fips_dict = {}
    
    # Filter the list to include only Excel files
    excel_files = [file for file in file_list if file.endswith('.xlsx')]

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
            key = excel_title + "__" + str(sheet_title)

            # Add key-value pair to the dictionary
            fips_dict[key] = values
            
    # Check if this is a tribal session
    # And reprocess dict to convert to tribal if it is
    fips_stuff = check_for_tribal(fips_dict, excel_files)

    print_message("\nFinished generating dictionary of map titles and values lists")

    # Send the stuff back whence it was called
    return fips_stuff

    
############################################################

# Make an output folder to store output file types
def make_folders(folder_names):

    # Empty list to store all full file paths (currently only 2)
    full_paths = []
    
    # Iterate through input folder names
    for folder_name in folder_names:

        # Define the full path string
        full_path = os.path.join(input_folder, folder_name)

        # If the output folder already exist, do some checking:
        if os.path.exists(full_path):

            # If we are supposed to delete prior outputs, do that,
            # then re-make the folder
            if del_prior_outputs:
                shutil.rmtree(full_path)
                os.mkdir(full_path)

        # Otherwise if folder do NOT already exist, just make it!
        else:
            os.mkdir(full_path)
            
        # Add new full path to list to be returned
        full_paths.append(full_path)

    # Ship the full output path back to the rest of the script
    return full_paths

        
############################################################

#  This should be called once per map layout
#  Update cursor loops through all counties;
#  For any row with a matching value in the current get_fips key/value pair,
#  flip the Class attribute to "Primary"
def code_primary(map_layers, fips_list, id_field):

    for map_layer in map_layers:
        
        #  Fields for update cursor
        cursor_fields = [id_field, "CLASS"]

        # Run through the cursor and code all counties in the fips list as Primary
        with arcpy.da.UpdateCursor(map_layer, cursor_fields) as update_cursor:
            for row in update_cursor:
                if row[0] in fips_list:
                    row[1] = "Primary"
                    update_cursor.updateRow(row)

    print_message("\tFinished coding Primary counties")
        
        
############################################################

# Called once per map layout. Select all primary counties; 
# select all counties adjacent to primary; flip the class attribute to "Contiguous"
def code_contiguous(primary, contiguous, query):

    # Name for new Primary counties only layer
    primary_features = "Primary Features (Temp)"

    # Make Feature Layer for use with intersect GP tool
    arcpy.management.MakeFeatureLayer(primary, primary_features, query)

    # Use Select by Location to select counties from the counties layer
    # that are contiguous to primary features (whether primary is in county or tribal layer)
    # Current policy does not provide for the possibility of "contiguous tribal areas";
    # If that changes, this section will need to be reworked
    arcpy.management.SelectLayerByLocation(contiguous, "INTERSECT", primary_features)

    # If primary features are also counties, we do not want to re-code them as contiguous
    # in next step; remove anything currently coded as "primary" from selection
    arcpy.management.SelectLayerByAttribute(contiguous, "REMOVE_FROM_SELECTION", query)

    # Calculate all selected counties to "Contiguous"
    arcpy.management.CalculateField(contiguous, "CLASS", "'Contiguous'")

    # Clear selected features from permanent counties layer
    arcpy.management.SelectLayerByAttribute(contiguous, "CLEAR_SELECTION")

    # Delete temp primary-only counties layer
    arcpy.management.Delete(primary_features)

    print_message("\tFinished coding Contiguous counties")
    
    
############################################################

# We have done it! The holy grail of Secretarial Disaster Designation maps!!!
# At this point, contiguous counties are already coded in the Census (with water) layer;
# Now we need to transfer that coding to the counties with water removed,
# THEN remove the "contiguous" coding from any counties where the non-water portions
# are separated by x distance (here assumed to be 5 miles)
def modify_contiguous(map_layers, queries):
    
    water = map_layers[2]
    non_water = map_layers[0]

    # STEP 1: We already coded the non-water PRIMARY counties for cartographic purposes,
    # but the non-water layer has no counties coded as contiguous yet.
    # We need them coded as contiguous. Compare FIPS and code everything that is
    # CURRENTLY coded as contiguous in the water layer as contiguous in the non-water layer
    
    # Get the list of FIPS values for all the contiguous counties in water layer
    contiguous_fips = [row[0] for row in arcpy.da.SearchCursor(water, ["FIPS_C"], queries[1])]

    # Use that list of FIPS to update the counties in the non-water layer
    with arcpy.da.UpdateCursor(non_water, ["FIPS_C", "CLASS"]) as update_cursor:
        for row in update_cursor:
            if row[0] in contiguous_fips:
                row[1] = "Contiguous"
                update_cursor.updateRow(row)
                
    # STEP 2: We now have ALL the correct counties coded as Contiguous, plus (possibly)
    # ridiculous counties like something across a Great Lake from a primary county.
    # Not a problem! We just need to remove some counties from Contiguous.

    # Make a layer of all non-water Primary counties
    prime_carto = "Primary Carto Counties (Temp)"
    arcpy.management.MakeFeatureLayer(map_layers[0], prime_carto, queries[0])

    # Make a layer of all non-water Primary Tribal areas
    prime_tribal = "Primary Tribal Counties (Temp)"
    arcpy.management.MakeFeatureLayer(map_layers[1], prime_tribal, queries[0])

    # Make a layer of all non-water Contiguous counties
    contig_carto = "Contiguous Carto Counties (Temp)"
    arcpy.management.MakeFeatureLayer(map_layers[0], contig_carto, queries[1])
    
    # Intersect non-water Primary with non-water Contigous, using "WITHIN_A_DISTANCE_GEODESIC"
    # Value used will depend on outcome of meeting tomorrow. Maybe 5 miles? 10 miles?
    arcpy.management.SelectLayerByLocation(contig_carto, "WITHIN_A_DISTANCE_GEODESIC", prime_carto, "5 Miles")

    arcpy.management.SelectLayerByLocation(contig_carto, "WITHIN_A_DISTANCE_GEODESIC", prime_tribal, "5 Miles", "ADD_TO_SELECTION")

    # Then just invert the selection... 
    arcpy.management.SelectLayerByLocation(contig_carto, selection_type="SWITCH_SELECTION")
    
    # I previously just went straight from inverting the selection to updating features
    # to "Not selected". However, if there is no difference, when you switch the selection,
    # you are left with nothing at all selected, and the following update cursor iterates
    # over THE WHOLE FEATURE SET. Takes forever and kinda dumb, since there is nothing to update.
    
    # So now we check whether there is anything left in this layer after switching selection...
    still_selected = [row for row in arcpy.da.SearchCursor(contig_carto, ["CLASS"])]
    
    # ...if there is, go ahead and reset it to "Not Selected"
    if still_selected:
        
        with arcpy.da.UpdateCursor(contig_carto, ["CLASS"]) as update_cursor:
            for row in update_cursor:
                row[0] = "Not Selected"
                update_cursor.updateRow(row)
            
    # Delete all three temp layers w/queries
    arcpy.management.Delete(prime_carto)
    arcpy.management.Delete(prime_tribal)
    arcpy.management.Delete(contig_carto)


############################################################

#  Literally just zoom to map extent in layout template, populate title, export
def export_map(map_layer, map_title, query, pdf_folder):
    
    arcpy.env.workspace = pdf_folder
    
    # Get the current project
    current_project = arcpy.mp.ArcGISProject("CURRENT")

    # Get map layout
    layout = current_project.listLayouts("Layout")[0]
    
    # Get map frame layout element
    map_frame = layout.listElements("MAPFRAME_ELEMENT", "Main Map")[0]    
    
    # Get map frame layout element
    overview_frame = layout.listElements("MAPFRAME_ELEMENT", "Overview Map")[0]    
    
    # Select counties currently coded as either primary or contiguous
    arcpy.management.SelectLayerByAttribute(map_layer, "NEW_SELECTION", query)
    
    # Get extent of selected features to use
    selected_extent = map_frame.getLayerExtent(map_layer, True, True)
    
    # Apply extent to layout map frame
    map_frame.camera.setExtent(selected_extent)
    
    # Clear selected features
    arcpy.management.SelectLayerByAttribute(map_layer, "CLEAR_SELECTION")
    
    # Zoom out a little so extent counties are not right up to the edge of the map frame
    map_frame.camera.scale = map_frame.camera.scale * 1.1
    
    # Zoom out a little so extent counties are not right up to the edge of the map frame
    overview_frame.camera.scale = map_frame.camera.scale * 8
    
    # Populate the map title dynamic text using Excel file name or file name/sheet combo
    layout.listElements("TEXT_ELEMENT", "Title")[0].text = map_title.split(".")[0]
    layout.listElements("TEXT_ELEMENT", "Subtitle")[0].text = map_title.split(".")[0]
    
    # Export layout - can adjust resolution here manually if desired
    layout.exportToPDF(os.path.join(pdf_folder, map_title), resolution=200)
    
    print_message("\tFinished exporting the current map")
    
    
############################################################

# Export excel files (previously dbf was used; just wtaf);
# these are ingested downstream by some xlsm file to populate
# excel reports, some word doc letter...someday, when we near our goal
# of taking over the world, all of that will go away...
def export_excel(map_layer, file_name, queries, excel_folder):
    
    arcpy.env.workspace = excel_folder

    # Select counties currently coded as either primary or contiguous
    arcpy.management.SelectLayerByAttribute(map_layer, "NEW_SELECTION", queries[2])

    # Use Table to Excel GP tool to export
    arcpy.conversion.TableToExcel(map_layer, file_name)

    # Clear selected features
    arcpy.management.SelectLayerByAttribute(map_layer, "CLEAR_SELECTION")

    print_message("\tFinished exporting the Excel files")
    

############################################################

# Once both buffered and unbuffered maps and excel documents are exported,
# reset the "CLASS" attribute to "Not Selected" for both county and tribal layers.
# (This effectively resets symbology and labels for the entire map)
def reset_class(map_layers, query):
    
    # Loop happens 2x, once for counties once for tribal areas
    for map_layer in map_layers:
        
        # Tried just calcing this field but it was throwing an error
        # for which Esri had no help page, so just switched to cursor; still plenty zippy.
        with arcpy.da.UpdateCursor(map_layer, ["CLASS"], query) as update_cursor:
            for row in update_cursor:
                row[0] = "Not Selected"
                update_cursor.updateRow(row)

    print_message("\tFinished resetting CLASS attribute")
    

############################################################

# The meat and potatoes of the script--iterate through the dictionary of fips codes
# (And possibly tribal codes) and create 2 maps + 2 excels 
# (1 set buffered, 1 set not buffered) for each iteration.
def iterate_maps(map_layers, fips_dict, tribal, queries, output_folders):
    
    # For non-tribal maps, "key" is every excel sheet and "value" is list of 5-digit FIPS codes
    # for tribal maps, "key" is every excel file title and "value" is list of FIPS + tribal codes
    for key, value in fips_dict.items():
        
        print_message(f"\nWorking on map title '{key}'")

        # If this is a tribal map, 2nd call to code Primary tribal areas as well
        if tribal:
            
            code_primary([map_layers[1]], value, "AIANNH")

            # Call to function to code contiguous counties
            code_contiguous(map_layers[1], map_layers[2], queries[0])

            export_excel(map_layers[1], key + " TRIBAL.xlsx", queries, output_folders[1])

        # Call to function to iterate through counties and code "CLASS" attribute as Primary
        code_primary([map_layers[0], map_layers[2]], value, "FIPS_C")

        # Call to function to code contiguous counties
        code_contiguous(map_layers[2], map_layers[2], queries[0])

        # Call to function to modify contiguous counties if necessary
        modify_contiguous(map_layers, queries)

        # Call to function to export the map layout as .pdf
        export_map(map_layers[0], key + ".pdf", queries[2], output_folders[0])

        # Call to function to export the associated excel file
        export_excel(map_layers[0], key + ".xlsx", queries, output_folders[1])

        # Call to function to reset CLASS attribute for both layers
        reset_class(map_layers, queries[2])

        print_message(f"Finished map title '{key}'")
    

############################################################

def do_the_work():
    
    # Call to function to get the counties layer
    map_layers = get_layers()
    
    # Call to function to build the required SQL queries
    queries = build_queries(map_layers[0])

    # Call to function to clear any and all selections
    clear_all(map_layers)

    # Call to function to reset CLASS attribute for both layers
    reset_class(map_layers, queries[2])
    
    # Call to function to get all the fips codes from all the sheets in all the excel files in a folder
    fips_stuff = get_fips()
    
    # Provide string names for our two output folders
    output_folders = make_folders(["Output PDFs", "Output Excels"])

    # Call to function to iterate fips_dict and export a single map
    # Various per-map functions are called from within this function
    iterate_maps(map_layers, fips_stuff[0], fips_stuff[1], queries, output_folders)

    # Print status message
    print_message("\nFinished generating all maps successfully")
    
    
################################################################################
# DO THE WORK
################################################################################

# Call to the MAIN OVERARCHING function from which all other functions are called
# One call does it all...should make it easier to stuff in a toolbox later
do_the_work()
