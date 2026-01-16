from itertools import count

import pandas as pd
from io import StringIO
import xlwings as xw
import numpy as np
import subprocess

first_column_name = "Run Name"

def generate_output_file():
    """
    Output File
    """

    pass

def import_spacegass_script(master_excel):
    print("Importing SPACEGASS script into SPACEGASS")

    # Define some typical naming conventions
    df_properties = import_section_properties(master_excel)

    import_spacegass_model = 'ACTION OPEN "File=C:\\Users\\k146666\\PyCharmMiscProject\\SET059 - Sutherland Piled Extension_V4.SG"\n'

    default_header = "SPACE GASS Script File\n" "VERSION 14000000\n" "SHOW Normal\n"
    default_output_name = "ACTION EXPORT_TXT " + '"' + "File=C:\\Users\\k146666\\PyCharmMiscProject\\"

    default_grabs = '"Stations=1" "ND=No" "MA=No" "ID=No" "IA=Yes" "PA=No" "PS=No" "NR=No" "BF=No" "BL=No" "DF=No" "DM=No" "SD=No" "MS=No"'

    default_name = ""

    # For writing the inputs to the SPACEGASS script file
    count_rows = df_properties[first_column_name].notna().sum()
    for i in range(count_rows):
        default_name += default_output_name
        default_name += str(df_properties.iloc[i][first_column_name]) + str(df_properties.iloc[i]['Load Cases'])  + '.txt" '
        default_name += '"Cases=' + str(df_properties.iloc[i]['Load Cases']) + '" '
        default_name += '"Filter=' + str(int(df_properties.iloc[i]['Section Filter Number'])) + '" '
        default_name += default_grabs
        default_name += "\n"

    script_text = default_header + import_spacegass_model + default_name

        # Then create the SPACEGASS script file
    with open("script.TXT", "w") as f:
        f.write(script_text)

    # TODO Run SPACEGASS script file via SPACEGASS script mode
    sg_exe_dir = 'C:\Program Files\SPACE GASS 14.2\SGCore.exe'
    sg_script = r'C:\Users\k146666\PyCharmMiscProject\script.txt'

    sg_exe = fr'"{sg_exe_dir}" -n -s "{sg_script}"'


    # TODO add wait in subprocess
   # subprocess.run(sg_exe, check=True)

    return

def import_section_properties(section_properties_file):
    """
    Imports section properties from excel that has all design property data
    :param section_properties_file:
    :return:
    """

    # TODO
    df_properties = pd.read_excel(section_properties_file) # pd.read_excel or something

    # This is how you get the ROW with Index 0
    # df_properties_row = df_properties.iloc[0]
    # This is how you get the Name WITHIN the row, for example I want the Section Name

    return df_properties

def import_sg_output(master_excel):

    # 1) Create the SPACE GASS script
    import_spacegass_script(master_excel)

    # 2) Import the master Excel
    print("Importing master Excel file")
    df_properties = import_section_properties(master_excel)

    # 3) Create an array of text_files which SPACEGASS has exported
    spacegass_output_texts = []
    count_rows = df_properties[first_column_name].notna().sum()

    for i in range(count_rows):
        spacegass_output_texts.append(str(df_properties.iloc[i][first_column_name]) + str(df_properties.iloc[i]['Load Cases'])  + '.txt')

    # Define some column names
    column_names = ["Load Case", "Member ID", "Segment Number", "Segment Length", "Fx", "Fy", "Fz", "Mx", "My", "Mz"]
    additional_columns = ["Ultimate Strength", "Ultimate Utilisation", "Bar Stress", "Serviceability Pass/Fail", "Max or Min", "Depth", "Width", "Added Moment", "Mem 1", "Mem 2" ,"Mem 1 Max/Min", "Mem 2 Max/Min"]

    # Create a dataframe to store the results
    df_results = pd.DataFrame(columns=column_names + additional_columns)

    # 4) Iterate over the text files created AND the rows in the Excel (which are the SAME length)
    for i, text_files in zip(range(count_rows), spacegass_output_texts):
        print("Currently importing: ", text_files)

        print(df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'], df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Btm Bar Layer 1'])

        with open(text_files) as f:
            lines = f.readlines()

        # Add a new line in the results Excel to separate different text files
        first_col = df_results.columns[0]
        df_results.loc[len(df_results)] = {first_col: text_files}

        # Grab the relevant intermediate forces in the SPACEGASS output text file
        idx_forces_start = lines.index("MEMBER INTERMEDIATE FORCES AND MOMENTS\n")
        try:
            idx_forces_end =  lines.index("STEELMEMBERS\n")
        except:
            idx_forces_end = len(lines)-1

        line_forces = lines[idx_forces_start + 1:idx_forces_end]

        # Use Pandas to read the "CSV" file and create column names
        df = pd.read_csv(StringIO("".join(line_forces)),
                         names=column_names)

        # Iterate over maximum forces in Fz, Mx, My, Mz columns for each row in the Master Excel
        for max_forces in ["Fz", "Mx", "My", "Mz"]:
            idx_max = df[[max_forces]].idxmax().values[0]
            idx_min = df[[max_forces]].idxmin().values[0]

            max_row = df.loc[idx_max]
            min_row = df.loc[idx_min]

            # Find the two closest members using the average_moment function
            max_mem_1, max_mem_2 = average_moment(text_files, max_row['Member ID'])
            min_mem_1, min_mem_2 = average_moment(text_files, min_row['Member ID'])

            max_mem_1_filtered = df[(df['Member ID']==max_mem_1) & (df['Load Case']==max_row['Load Case'])]
            max_mem_2_filtered = df[(df['Member ID']==max_mem_2) & (df['Load Case']==max_row['Load Case'])]

            min_mem_1_filtered = df[(df['Member ID']==min_mem_1) & (df['Load Case']==min_row['Load Case'])]
            min_mem_2_filtered = df[(df['Member ID']==min_mem_2) & (df['Load Case']==min_row['Load Case'])]

            '''
            Get the maximum
            mem_1_max_Mz = mem_1_filtered["Mz"].max()
            mem_2_max_Mz = mem_2_filtered["Mz"].max()
            '''

            max_avg_moment = max_row.Mz + max_mem_1_filtered["Mz"].max() + max_mem_2_filtered["Mz"].max()
            min_avg_moment = min_row.Mz + min_mem_1_filtered["Mz"].min() + min_mem_2_filtered["Mz"].min()

            # TODO find the moments of the two closest members with the same load case
            print(max_mem_1, max_mem_2)
            print(min_mem_1, min_mem_2)

            # Calculate maximum forces for "i" row in the Master Excel.
            # TODO CLEAN UP THE CALL UPS TO FUNCTION - PROBABLY CAN BE SIMPLIFIED
            ultimate_utilisation, ultimate_strength, serviceability_btm_stress, serviceability_utilisation = calculate_utilisation(
                df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'], df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Top Bar Layer 1 Spacing'],
                df_properties.iloc[i]['Top Bar Layer 2 Spacing'], df_properties.iloc[i]['Btm Bar Layer 1'], df_properties.iloc[i]['Btm Bar Layer 1 Spacing'], df_properties.iloc[i]['Btm Bar Layer 2 Spacing'],
                max_row.Fx, max_row.Fy, max_row.Fz, max_row.Mx, max_row.My, max_row.Mz)

            df_results.loc[len(df_results)] = list(max_row.values) + [ultimate_strength] + [ultimate_utilisation] + [serviceability_btm_stress] + [serviceability_utilisation] + ["max " + max_forces] + [df_properties.iloc[i]['Depth']] + [df_properties.iloc[i]['Width']] + [max_avg_moment] + [max_mem_1] + [max_mem_2] + [max_mem_1_filtered["Mz"].max()] + [max_mem_2_filtered["Mz"].max()]

            # Calculate minimum forces for "i" row in the Master Excel
            ultimate_utilisation, ultimate_strength, serviceability_btm_stress, serviceability_utilisation = calculate_utilisation(
                df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'], df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Top Bar Layer 1 Spacing'],
                df_properties.iloc[i]['Top Bar Layer 2 Spacing'], df_properties.iloc[i]['Btm Bar Layer 1'], df_properties.iloc[i]['Btm Bar Layer 1 Spacing'], df_properties.iloc[i]['Btm Bar Layer 2 Spacing'],
                min_row.Fx, min_row.Fy, min_row.Fz, min_row.Mx, min_row.My, min_row.Mz)

            df_results.loc[len(df_results)] = list(min_row.values) + [ultimate_strength] + [ultimate_utilisation] + [serviceability_btm_stress] + [serviceability_utilisation] + ["min " + max_forces] + [df_properties.iloc[i]['Depth']] + [df_properties.iloc[i]['Width']] + [min_avg_moment] + [min_mem_1] + [min_mem_2] + [min_mem_1_filtered["Mz"].min()] + [min_mem_2_filtered["Mz"].min()]

        print(str(text_files) + " Done")

    # LAST: Compile results and print results file
    print("Printing output file")
    output_file = "results_file.xlsx"
    df_results.to_excel(output_file, index=False)

    pass

def average_moment(text_file, ref_member):
    # Define column names
    member_column_names = ["Member ID", "PH0", "PH1", "PH2", "Y/N", "Node 1", "Node 2", "Section Number",
                           "Material Number", "Fixity 1", "Fixity 2", "PH3", "PH4", "PH5", "PH6", "PH7", "PH8", "PH9",
                           "PH10", "PH11"]
    node_column_names = ["Node ID", "x", "y", "z"]

    # Open text files
    with open(text_file) as f:
        lines = f.readlines()

    # Index of member start lines
    idx_members_start = lines.index("MEMBERS\n")
    idx_members_end = lines.index("RESTRAINTS\n")

    member_data = lines[idx_members_start + 1:idx_members_end]

    df_members = pd.read_csv(StringIO("".join(member_data)), names=member_column_names)

    # Index of node start lines
    idx_node_start = lines.index("NODES\n")
    idx_node_end = lines.index("MEMBERS\n")

    node_data = lines[idx_node_start + 1:idx_node_end]

    df_nodes = pd.read_csv(StringIO("".join(node_data)), names=node_column_names)

    # Find the row where the reference member is, where ref_member is an input to the function
    ref_member_row = df_members[df_members["Member ID"] == ref_member]

    # Get the connecting nodes of the reference member
    ref_node_1 = int(ref_member_row["Node 1"].iloc[0])
    ref_node_2 = int(ref_member_row["Node 2"].iloc[0])

        # Then get the coordinates of both connecting nodes
    ref_node_1_coordinates = [float(df_nodes[df_nodes["Node ID"]==ref_node_1]['x']), float(df_nodes[df_nodes["Node ID"]==ref_node_1]['y']), float(df_nodes[df_nodes["Node ID"]==ref_node_1]['z'])]
    ref_node_2_coordinates = [float(df_nodes[df_nodes["Node ID"]==ref_node_2]['x']), float(df_nodes[df_nodes["Node ID"]==ref_node_2]['y']), float(df_nodes[df_nodes["Node ID"]==ref_node_2]['z'])]

    # Find the mid-point of the reference member
    mid_point_ref = [(ref_node_2_coordinates[0] + ref_node_1_coordinates[0])/2, (ref_node_2_coordinates[1] + ref_node_1_coordinates[1])/2, (ref_node_2_coordinates[2] + ref_node_1_coordinates[2])/2]

    # Find the vector value of the reference member
    ref_vector = [ref_node_1_coordinates[0] - ref_node_2_coordinates[0], ref_node_1_coordinates[1] - ref_node_2_coordinates[1], ref_node_1_coordinates[2] - ref_node_2_coordinates[2]]
    ref_vector_magnitude = (ref_vector[0]**2 + ref_vector[1]**2 + ref_vector[2]**2)**0.5


    closest_above = [None, 1000]
    closest_below = [None, 1000]

    # TODO iterate through all members and do the same as the same of reference member and find the midpoint of the member, then compare to mid_point_ref
    for i in range(len(df_members)):
        node_1 = df_members.iloc[i]["Node 1"]
        node_2 = df_members.iloc[i]["Node 2"]

        node_1_coordinates = [float(df_nodes[df_nodes["Node ID"]==node_1]['x']), float(df_nodes[df_nodes["Node ID"]==node_1]['y']), float(df_nodes[df_nodes["Node ID"]==node_1]['z'])]
        node_2_coordinates = [float(df_nodes[df_nodes["Node ID"]==node_2]['x']), float(df_nodes[df_nodes["Node ID"]==node_2]['y']), float(df_nodes[df_nodes["Node ID"]==node_2]['z'])]

        direction_vector = [node_1_coordinates[0] - node_2_coordinates[0], node_1_coordinates[1] - node_2_coordinates[1], node_1_coordinates[2] - node_2_coordinates[2]]
        direction_vector_magnitude = (direction_vector[0]**2 + direction_vector[1]**2 + direction_vector[2]**2)**0.5

        mid_point_member = [(node_2_coordinates[0] + node_1_coordinates[0]) / 2,
                         (node_2_coordinates[1] + node_1_coordinates[1]) / 2,
                         (node_2_coordinates[2] + node_1_coordinates[2]) / 2]

        # Check Dot Product
        dot_product = ref_vector[0]*direction_vector[0] + ref_vector[1]*direction_vector[1] + ref_vector[2]*direction_vector[2]
        # print(dot_product)
            # Check value
        cos_theta = dot_product/(ref_vector_magnitude*direction_vector_magnitude)

        # Check distance from mid-point of reference member to mid-point of member
        distance = (((mid_point_ref[0] - mid_point_member[0]))**2 + ((mid_point_ref[1] - mid_point_member[1]))**2 + ((mid_point_ref[2] - mid_point_member[2]))**2)**0.5
        # print(distance)
        mid_point_dif = [float(mid_point_ref[0] - mid_point_member[0]), float(mid_point_ref[1] - mid_point_member[1]), float(mid_point_ref[2] - mid_point_member[2])]
        # print(mid_point_dif)

        moving_x = False
        moving_z = False

        if ref_vector[0] !=0:
            moving_x = True
        elif ref_vector[2] != 0:
            moving_z = True

        # TODO check if it is parallel or not
        if abs(cos_theta) != 1.0:
            continue

        # TODO add an if to check similar nodes
        if node_1 == ref_node_1 or node_2 == ref_node_2 or node_1 == ref_node_2 or node_2 == ref_node_1:
            continue

        # TODO function to change closest variable

        if moving_x == True:
            if mid_point_dif[2] > 0:
                if distance < closest_above[1]:
                    closest_above[0] = df_members.iloc[i]["Member ID"]
                    closest_above[1] = distance
            elif mid_point_dif[2] < 0:
                if distance < closest_below[1]:
                    closest_below[0] = df_members.iloc[i]["Member ID"]
                    closest_below[1] = distance

        elif moving_z == True:
            if mid_point_dif[0] > 0:
                if distance < closest_above[1]:
                    closest_above[0] = df_members.iloc[i]["Member ID"]
                    closest_above[1] = distance
            elif mid_point_dif[0] < 0:
                if distance < closest_below[1]:
                    closest_below[0] = df_members.iloc[i]["Member ID"]
                    closest_below[1] = distance


        # mid_point = [(node_1_coordinates[0]+node_2_coordinates[0])/2, (node_1_coordinates[1]+node_2_coordinates[1])/2, (node_1_coordinates[2]+node_2_coordinates[2])/2]

    # print(mid_point_ref)

    return closest_above[0], closest_below[0]


def calculate_utilisation(depth, width, top_bar_1, top_bar_spacing_1, top_bar_spacing_2, btm_bar_1, btm_bar_spacing_1, btm_bar_spacing_2, Fx,Fy,Fz,Mx,My,Mz):
    print("Extracting forces and calculating capacities")

    # Import the Concrete Capacity Excel
    workbook = xw.Book("RC Beam Design to AS3600 - 2018.xlsm", visible=False)
    sheet = workbook.sheets[0]

    # Change BAR SIZES
    sheet["D15"].value = top_bar_1
    sheet["D16"].value = btm_bar_1

    # Change BAR SPACING
    sheet["K27"].value = top_bar_spacing_1
    sheet["K28"].value = top_bar_spacing_2
    sheet["K34"].value = btm_bar_spacing_1
    sheet["K35"].value = btm_bar_spacing_2

    # Hold bar information
    top_bar = sheet["D15"].value
    btm_bar = sheet["D16"].value

    top_bar_cells = ["K27", "K28", "K29", "K30", "K31"]
    btm_bar_cells = ["K34", "K35", "K36", "K37", "K38"]

    top_bar_amount = [sheet["K27"].value, sheet["K28"].value, sheet["K29"].value, sheet["K30"].value, sheet["K31"].value]
    btm_bar_amount = [sheet["K34"].value, sheet["K35"].value, sheet["K36"].value, sheet["K37"].value, sheet["K38"].value]

    # Flip the size of bars
    if Mz < 0:
        sheet["D15"].value = btm_bar
        sheet["D16"].value = top_bar

        # Flip the number of bars
        for i, amount in zip(top_bar_cells, range(0, 5)):
            sheet[i].value = btm_bar_amount[amount]

        for i, amount in zip(btm_bar_cells, range(0, 5)):
            sheet[i].value = top_bar_amount[amount]

    # Change the values in the capacity Excel for ultimate positive and negative bending moments
    sheet["D33"].value = abs(Mz)

    # Change the values in the capacity Excel for serviceability positive and negative bending moments
    sheet["D38"].value = abs(Mz)

    # Change the value in the capacity Excel for axial forces
    sheet["D35"].value = abs(Fx)

    # Change the value in the capacity Excel for shear forces
    sheet["D36"].value = abs(Fz)

    # Change the value in the capacity Excel for torsion forces
    sheet["D37"].value = abs(Mx)

    # Change SECTION PROPERTIES
    sheet["D8"].value = depth
    sheet["D9"].value = width

    # Run macro in workbook
    workbook.macro("Solvefordn")()

    # Grab values from the Excel
    ultimate_utilisation = sheet["L8"].value
    ultimate_strength = sheet["J8"].value
    serviceability_btm_stress = sheet["J19"].value
    serviceability_utilisation = sheet["L19"].value

    # Unflip the bars
    sheet["D15"].value = top_bar
    sheet["D16"].value = btm_bar

    for i, amount in zip(top_bar_cells, range(0, 5)):
        sheet[i].value = top_bar_amount[amount]

    for i, amount in zip(btm_bar_cells, range(0, 5)):
        sheet[i].value = btm_bar_amount[amount]

    return ultimate_utilisation, ultimate_strength, serviceability_btm_stress, serviceability_utilisation

if __name__ == "__main__":
    # For checking the MAIN code
    import_sg_output("MASTER.xlsx")


    # average_moment("Deck Long ULS1000-1999.TXT", 1732)
    # push_values("MASTER.xlsx", "")

    # For checking importing SPACEGASS script file
    # import_spacegass_script("MASTER.xlsx")

    # For checking importing section properties Excel file
    # import_section_properties("MASTER.xlsx")

