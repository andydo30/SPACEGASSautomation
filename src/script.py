from itertools import count

import pandas as pd
from io import StringIO
import xlwings as xw
from sg_results import SGResults
from pathlib import Path
import subprocess
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk
import os

first_column_name = "Run Name"

# gui_spacegass_selector.py
# Requires: customtkinter (pip install customtkinter)
# Python 3.8+

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk


# ------------------------------------------- Co-pilot customtkinter GUI -----------------------------------------------
class SpaceGassSelectorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.result = None
        # ...
        self.result = None  # <-- holds (model_file, output_dir) when confirmed

        # ---------- Window ----------
        self.title("SpaceGass Analysis – File & Output Folder")
        self.geometry("700x340")
        self.minsize(680, 320)
        ctk.set_appearance_mode("System")  # "Light" or "Dark" or "System"
        ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

        # Remember last paths in this session
        self.last_model_dir = os.getcwd()
        self.last_output_dir = os.getcwd()

        # ---------- Variables ----------
        self.model_file_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Please select a SpaceGass file and an output folder.")

        # ---------- Layout ----------
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        header = ctk.CTkLabel(self, text="SpaceGass Analysis Setup", font=ctk.CTkFont(size=18, weight="bold"))
        header.grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")

        subheader = ctk.CTkLabel(self, text="Choose your SpaceGass model file and where to store results.")
        subheader.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

        # ----- Model file frame -----
        file_frame = ctk.CTkFrame(self, corner_radius=8)
        file_frame.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        lbl_model = ctk.CTkLabel(file_frame, text="SpaceGass file:")
        lbl_model.grid(row=0, column=0, padx=12, pady=12, sticky="w")

        self.ent_model = ctk.CTkEntry(file_frame, textvariable=self.model_file_var)
        self.ent_model.grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")
        self.ent_model.bind("<KeyRelease>", lambda e: self.validate_ready())

        btn_browse_file = ctk.CTkButton(file_frame, text="Browse…", command=self.browse_model_file, width=110)
        btn_browse_file.grid(row=0, column=2, padx=12, pady=12)

        # ----- Output dir frame -----
        out_frame = ctk.CTkFrame(self, corner_radius=8)
        out_frame.grid(row=3, column=0, padx=16, pady=(0, 12), sticky="ew")
        out_frame.grid_columnconfigure(1, weight=1)

        lbl_out = ctk.CTkLabel(out_frame, text="Output folder:")
        lbl_out.grid(row=0, column=0, padx=12, pady=12, sticky="w")

        self.ent_out = ctk.CTkEntry(out_frame, textvariable=self.output_dir_var)
        self.ent_out.grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")
        self.ent_out.bind("<KeyRelease>", lambda e: self.validate_ready())

        btn_browse_out = ctk.CTkButton(out_frame, text="Browse…", command=self.browse_output_dir, width=110)
        btn_browse_out.grid(row=0, column=2, padx=12, pady=12)

        # ----- Status + Actions -----
        bottom_frame = ctk.CTkFrame(self, corner_radius=8)
        bottom_frame.grid(row=4, column=0, padx=16, pady=(0, 16), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        self.status_lbl = ctk.CTkLabel(bottom_frame, textvariable=self.status_var, wraplength=620)
        self.status_lbl.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        action_row = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        action_row.grid(row=1, column=0, padx=8, pady=(4, 12), sticky="e")

        self.btn_run = ctk.CTkButton(action_row, text="Analyse", command=self.on_run_clicked, state="disabled",
                                     width=130)
        self.btn_run.grid(row=0, column=0, padx=6)

        self.btn_theme = ctk.CTkSegmentedButton(
            action_row,
            values=["Light", "Dark", "System"],
            command=self.on_theme_change,
            selected_color=("gray20", "gray90"),
            text_color=("white", "black")
        )
        self.btn_theme.set("System")
        self.btn_theme.grid(row=0, column=1, padx=6)

        # Initial validation
        self.validate_ready()

    # ---------- UI Callbacks ----------
    def browse_model_file(self):
        """
        Open a file dialog for selecting a SpaceGass model file.
        Adjust filetypes as needed for your workflow.
        """
        # Common SpaceGass-related possibilities: .sg, .sgs, .txt (exports)
        filetypes = [
            ("SpaceGass files", "*.sg *.SG *.sgs *.SGS"),
            ("Text files", "*.txt *.TXT"),
            ("All files", "*.*")
        ]
        path = filedialog.askopenfilename(
            title="Select SpaceGass file",
            initialdir=self.last_model_dir,
            filetypes=filetypes
        )
        if path:
            self.model_file_var.set(path)
            self.last_model_dir = os.path.dirname(path)
            self.validate_ready()

    def browse_output_dir(self):
        """
        Open a directory picker for selecting where results will be written.
        """
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self.last_output_dir,
            mustexist=True
        )
        if path:
            self.output_dir_var.set(path)
            self.last_output_dir = path
            self.validate_ready()

    def on_theme_change(self, value: str):
        ctk.set_appearance_mode(value)

    def validate_ready(self):
        """
        Validate that the inputs are usable. Enables the Analyse button if OK.
        """
        model = self.model_file_var.get().strip()
        outdir = self.output_dir_var.get().strip()

        errors = []
        if not model:
            errors.append("No SpaceGass file selected.")
        elif not os.path.isfile(model):
            errors.append("Selected SpaceGass file does not exist.")
        else:
            # Optional: enforce extension(s)
            valid_exts = {".sg", ".sgs", ".txt"}
            _, ext = os.path.splitext(model)
            if ext and ext.lower() not in valid_exts:
                # Not fatal; just warn
                errors.append(f"Warning: Unexpected file extension '{ext}'. (Proceed if intentional.)")

        if not outdir:
            errors.append("No output folder selected.")
        elif not os.path.isdir(outdir):
            errors.append("Selected output folder does not exist.")

        if errors:
            msg = " • " + "\n • ".join(errors)
            self.status_var.set(msg)
            self.btn_run.configure(state="disabled")
        else:
            self.status_var.set("Ready. Click ‘Analyse’ to proceed.")
            self.btn_run.configure(state="normal")

    def on_run_clicked(self):
        from pathlib import Path
        model = self.model_file_var.get().strip()
        outdir = self.output_dir_var.get().strip()

        # (optional) normalize & validate before returning
        model_path = Path(model).expanduser().resolve()
        outdir_path = Path(outdir).expanduser().resolve()
        if not model_path.exists():
            from tkinter import messagebox
            messagebox.showerror("Error", f"Model not found: {model_path}")
            return
        if not outdir_path.exists():
            from tkinter import messagebox
            messagebox.showerror("Error", f"Output folder not found: {outdir_path}")
            return

        # Return raw paths (and a pre-quoted variant if you like)
        sg_model_quoted = f'"{str(model_path)}"'
        self.result = (model_path, outdir_path, sg_model_quoted)
        self.destroy()

    # ---------- Your analysis function ----------
    def run_analysis(self, model_file: str, output_dir: str) -> str:
        """
        Stub for your analysis logic.
        Implement your SpaceGass parsing/automation here.
        Return a human-friendly message.
        """
        # Example: simulate writing a results file
        results_path = os.path.join(output_dir, "analysis_results.txt")
        with open(results_path, "w", encoding="utf-8") as f:
            f.write("SpaceGass Analysis Results\n")
            f.write("==========================\n")
            f.write(f"Model: {model_file}\n")
            f.write(f"Output folder: {output_dir}\n")
            f.write("Status: Success\n")
        return f"Analysis complete. Results saved to:\n{results_path}"

def pick_spacegass_inputs():
    app = SpaceGassSelectorApp()
    app.mainloop()
    return app.result if app.result else (None, None, None)


from pathlib import Path

def sg_quote_windows_path(p: str | Path) -> str:
    """
    Return a Windows-style absolute path wrapped in quotes, suitable for
    SpaceGass OPEN commands or command scripts.
    Ensures backslashes (\\), no mixed slashes, and existence check.
    """
    path = Path(p).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"SpaceGass model not found: {path}")
    # SpaceGass prefers Windows-style backslashes, but paths must be quoted if they have spaces
    win = path.as_posix()  # start from POSIX to avoid accidental escapes
    # Convert to backslashes:
    win = str(path)  # on Windows this is already backslashes, on other OS you usually won't be running SG
    return f'"{win}"'


def assert_sg_extension(p: str | Path, allowed=(".sg", ".sgs")):
    ext = Path(p).suffix.lower()
    if ext not in allowed:
        raise ValueError(f"Unexpected SpaceGass file extension {ext!r}. Allowed: {allowed}")


def analyse_spacegass(model_file: str, output_dir: str) -> str:
    # Your existing function
    # ... do the real work here ...
    return f"Analysed '{model_file}' → results in '{output_dir}'."


# ----------------------------------------------------------------------------------------------------------------------

def generate_output_file():
    """
    Output File
    """

    pass


def import_spacegass_script(master_excel):
    print("Importing SPACEGASS script into SPACEGASS")

    # Define some typical naming conventions
    df_properties = import_section_properties(master_excel)

    # Use custom gui to determine pathing of SPACEGASS models
    model, outdir, sg_model_quoted = pick_spacegass_inputs()

    import_spacegass_model = 'ACTION OPEN ' '"' "File=" + str(model) + '"' + "\n"

    default_header = "SPACE GASS Script File\n" "VERSION 14000000\n" "SHOW Normal\n" "SILENT\n" "CLOSE_AT_END\n"
    default_output_name = "ACTION EXPORT_TXT " + '"' "File=" + str(outdir) + '\\'

    default_grabs = '"Stations=1" "ND=No" "MA=No" "ID=No" "IA=Yes" "PA=No" "PS=No" "NR=No" "BF=No" "BL=No" "DF=No" "DM=No" "SD=No" "MS=No"'

    default_name = ""

    # For writing the inputs to the SPACEGASS script file
    count_rows = df_properties[first_column_name].notna().sum()
    for i in range(count_rows):
        default_name += default_output_name
        default_name += str(df_properties.iloc[i][first_column_name]) + str(
            df_properties.iloc[i]['Load Cases']) + '.txt" '
        default_name += '"Cases=' + str(df_properties.iloc[i]['Load Cases']) + '" '
        default_name += '"Filter=' + str(int(df_properties.iloc[i]['Section Filter Number'])) + '" '
        default_name += default_grabs
        default_name += "\n"

    script_text = default_header + import_spacegass_model + default_name

    # Then create the SPACEGASS script file
    with open("script.TXT", "w") as f:
        f.write(script_text)

    # TODO Since the SPACE GASS script will always be created in the same directory as this script, simplify sg_script
    sg_exe_dir = 'C:\Program Files\SPACE GASS 14.2\SGCore.exe'
    sg_script = r'C:\Users\k146666\PyCharmMiscProject\script.txt'

    sg_exe = fr'"{sg_exe_dir}" -n -s "{sg_script}"'

    # TODO add wait in subprocess
    subprocess.call(sg_exe)

    return

def import_section_properties(section_properties_file):
    """
    Imports section properties from excel that has all design property data
    :param section_properties_file:
    :return:
    """

    # TODO
    df_properties = pd.read_excel(section_properties_file)  # pd.read_excel or something

    # This is how you get the ROW with Index 0
    # df_properties_row = df_properties.iloc[0]
    # This is how you get the Name WITHIN the row, for example I want the Section Name

    return df_properties


def import_sg_output(master_excel):
    # 1) Create the SPACE GASS script
    # import_spacegass_script(master_excel)

    # 2) Import the master Excel
    print("Importing master Excel file")
    df_properties = import_section_properties(master_excel)

    # 3) Create an array of text_files which SPACEGASS has exported
    spacegass_output_texts = []
    count_rows = df_properties[first_column_name].notna().sum()

    '''
    TODO need to update this so it work's with output texts from any folder, currently only works if the output texts are in the same folder as this script
    '''
    for i in range(count_rows):
        spacegass_output_texts.append(
            str(df_properties.iloc[i][first_column_name]) + str(df_properties.iloc[i]['Load Cases']) + '.txt')

    # Define some column names
    column_names = ["Load Case", "Member ID", "Segment Number", "Segment Length", "Fx", "Fy", "Fz", "Mx", "My", "Mz"]
    additional_columns = ["Ultimate Strength", "Ultimate Utilisation", "Bar Stress", "Serviceability Pass/Fail",
                          "Max or Min", "Depth", "Width", "Added Moment", "Mem 1", "Mem 2", "Mem 1 Max/Min",
                          "Mem 2 Max/Min"]

    # Create a dataframe to store the results
    df_results = pd.DataFrame(columns=column_names + additional_columns)

    # 4) Iterate over the text files created AND the rows in the Excel (which are the SAME length)
    for i, text_files in zip(range(count_rows), spacegass_output_texts):
        print("Currently importing: ", text_files)

        # print(df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'], df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Btm Bar Layer 1'])

        with open(text_files) as f:
            lines = f.readlines()

        # Add a new line in the results Excel to separate different text files
        first_col = df_results.columns[0]
        df_results.loc[len(df_results)] = {first_col: text_files}

        # Grab the relevant intermediate forces in the SPACE GASS output text file
        '''
        Using SGResults Class by Chris Leaman to create a dataframe for:
        Member Intermediate Forces and Moments stored into forces_df
        '''
        results = SGResults(text_files)
        forces_df = results.member_int_forces_moments
        section_df = results.sections
        member_df = results.members

        # Iterate over maximum forces in Fz, Mx, My, Mz columns for each row in the Master Excel
        for max_forces in ["fy", "fz", "mx", "my", "mz"]:
            idx_max = forces_df[[max_forces]].idxmax().values[0]
            idx_min = forces_df[[max_forces]].idxmin().values[0]

            max_row = forces_df.loc[idx_max]
            min_row = forces_df.loc[idx_min]

            # Find the two closest members using the average_moment function
            '''
            max_row['Member ID'] retrieves the member of the maximum row
            max_mem_1 retrieves the member above the maximum member
            max_mem_2 retrieves the member below the maximum member
            '''
            max_mem_1, max_mem_2 = average_moment(text_files, max_row['member_id'])
            min_mem_1, min_mem_2 = average_moment(text_files, min_row['member_id'])

            # Filter the two closest members for 1. member, 2. the load case
            # This for maximum load case
            max_mem_1_filtered = forces_df[
                (forces_df['member_id'] == max_mem_1) & (forces_df['load_case_id'] == max_row['load_case_id'])]
            max_mem_2_filtered = forces_df[
                (forces_df['member_id'] == max_mem_2) & (forces_df['load_case_id'] == max_row['load_case_id'])]

            # This for minimum load case
            min_mem_1_filtered = forces_df[
                (forces_df['member_id'] == min_mem_1) & (forces_df['load_case_id'] == min_row['load_case_id'])]
            min_mem_2_filtered = forces_df[
                (forces_df['member_id'] == min_mem_2) & (forces_df['load_case_id'] == min_row['load_case_id'])]

            # TODO Retrieve section properties using SGResults, match all 3 members with the property "width"
            '''
            Retrieve 3 section properties:
            1. The maximum member
            2. The member ABOVE the maximum member
            3. The member BELOW the maximum member
            '''
            # Retrieve the Section ID of the maximum member, and the two adjacent members
            max_mem_section, max_mem_section_adj_1, max_mem_section_adj_2 = member_df[
                member_df['member_id'] == max_row['member_id']]['section_id'], \
                member_df[member_df['member_id'] == max_mem_1]['section_id'], \
                member_df[member_df['member_id'] == max_mem_2]['section_id']
            min_mem_section, min_mem_section_adj_1, min_mem_section_adj_2 = member_df[
                member_df['member_id'] == min_row['member_id']]['section_id'], \
                member_df[member_df['member_id'] == min_mem_1]['section_id'], \
                member_df[member_df['member_id'] == min_mem_2]['section_id']

            # Retrieve the section width using the Section ID
            max_mem_width = section_df[section_df['section_id'] == max_mem_section.item()]['col_19']
            if max_mem_1 is not None:
                max_mem_width_adj_1 = section_df[section_df['section_id'] == max_mem_section_adj_1.item()]['col_19']
            if max_mem_2 is not None:
                max_mem_width_adj_2 = section_df[section_df['section_id'] == max_mem_section_adj_2.item()]['col_19']

            min_mem_width = section_df[section_df['section_id'] == min_mem_section.item()]['col_19']
            if min_mem_1 is not None:
                min_mem_width_adj_1 = section_df[section_df['section_id'] == min_mem_section_adj_1.item()]['col_19']
            if min_mem_2 is not None:
                min_mem_width_adj_2 = section_df[section_df['section_id'] == min_mem_section_adj_2.item()]['col_19']

            # Average the moment across the 3 members
            max_avg_moment = (max_row.mz + max_mem_1_filtered["mz"].max() + max_mem_2_filtered["mz"].max()) / ((
                        max_mem_width.item() + max_mem_width_adj_1.item() + max_mem_width_adj_2.item())/1000)
            min_avg_moment = (min_row.mz + min_mem_1_filtered["mz"].min() + min_mem_2_filtered["mz"].min()) / ((
                        min_mem_width.item() + min_mem_width_adj_1.item() + min_mem_width_adj_2.item())/1000)

            # Calculate maximum forces for "i" row in the Master Excel.
            # TODO CLEAN UP THE CALL UPS TO FUNCTION - PROBABLY CAN BE SIMPLIFIED
            ultimate_utilisation, ultimate_strength, serviceability_btm_stress, serviceability_utilisation = calculate_utilisation(df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'],df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Top Bar Layer 1 Spacing'],df_properties.iloc[i]['Top Bar Layer 2 Spacing'], df_properties.iloc[i]['Btm Bar Layer 1'],df_properties.iloc[i]['Btm Bar Layer 1 Spacing'], df_properties.iloc[i]['Btm Bar Layer 2 Spacing'],max_row.fx, max_row.fy, max_row.fz, max_row.mx, max_row.my, max_avg_moment)

            # TODO need to clean up the max_row list as it's reading NaN values and creating too many columns, below code is a placeholder
            max_list = []

            for a in range(0, 10):
                max_list.append(list(max_row.values)[a])

            df_results.loc[len(df_results)] = max_list + [ultimate_strength] + [ultimate_utilisation] + [serviceability_btm_stress] + [serviceability_utilisation] + ["max " + max_forces] + [df_properties.iloc[i]['Depth']] + [df_properties.iloc[i]['Width']] + [max_avg_moment] + [max_mem_1] + [max_mem_2] + [max_mem_1_filtered["mz"].max()] + [max_mem_2_filtered["mz"].max()]

            # Calculate minimum forces for "i" row in the Master Excel
            ultimate_utilisation, ultimate_strength, serviceability_btm_stress, serviceability_utilisation = calculate_utilisation(df_properties.iloc[i]['Depth'], df_properties.iloc[i]['Width'],df_properties.iloc[i]['Top Bar Layer 1'], df_properties.iloc[i]['Top Bar Layer 1 Spacing'],df_properties.iloc[i]['Top Bar Layer 2 Spacing'], df_properties.iloc[i]['Btm Bar Layer 1'],df_properties.iloc[i]['Btm Bar Layer 1 Spacing'], df_properties.iloc[i]['Btm Bar Layer 2 Spacing'],min_row.fx, min_row.fy, min_row.fz, min_row.mx, min_row.my, min_avg_moment)

            min_list = []

            for b in range(0, 10):
                min_list.append(list(min_row.values)[b])

            df_results.loc[len(df_results)] = min_list + [ultimate_strength] + [ultimate_utilisation] + [serviceability_btm_stress] + [serviceability_utilisation] + ["min " + max_forces] + [df_properties.iloc[i]['Depth']] + [df_properties.iloc[i]['Width']] + [min_avg_moment] + [min_mem_1] + [min_mem_2] + [min_mem_1_filtered["mz"].min()] + [min_mem_2_filtered["mz"].min()]

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
    # TODO Andy to replace below code with Chris' updates
    idx_members_start = lines.index("MEMBERS\n")
    idx_members_end = lines.index("PLATES\n")

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
    ref_node_1 = ref_member_row["Node 1"].item()
    ref_node_2 = ref_member_row["Node 2"].item()

    # Then get the coordinates of both connecting nodes
    ref_node_1_coordinates = [df_nodes[df_nodes["Node ID"] == ref_node_1]['x'].item(),
                              df_nodes[df_nodes["Node ID"] == ref_node_1]['y'].item(),
                              df_nodes[df_nodes["Node ID"] == ref_node_1]['z'].item()]
    ref_node_2_coordinates = [df_nodes[df_nodes["Node ID"] == ref_node_2]['x'].item(),
                              df_nodes[df_nodes["Node ID"] == ref_node_2]['y'].item(),
                              df_nodes[df_nodes["Node ID"] == ref_node_2]['z'].item()]

    # Find the mid-point of the reference member
    mid_point_ref = [(ref_node_2_coordinates[0] + ref_node_1_coordinates[0]) / 2,
                     (ref_node_2_coordinates[1] + ref_node_1_coordinates[1]) / 2,
                     (ref_node_2_coordinates[2] + ref_node_1_coordinates[2]) / 2]

    # Find the vector value of the reference member
    ref_vector = [ref_node_1_coordinates[0] - ref_node_2_coordinates[0],
                  ref_node_1_coordinates[1] - ref_node_2_coordinates[1],
                  ref_node_1_coordinates[2] - ref_node_2_coordinates[2]]
    ref_vector_magnitude = (ref_vector[0] ** 2 + ref_vector[1] ** 2 + ref_vector[2] ** 2) ** 0.5

    closest_above = [None, 1000]
    closest_below = [None, 1000]

    # TODO iterate through all members and do the same as the same of reference member and find the midpoint of the member, then compare to mid_point_ref
    for i in range(len(df_members)):
        node_1 = df_members.iloc[i]["Node 1"]
        node_2 = df_members.iloc[i]["Node 2"]

        node_1_coordinates = [float(df_nodes[df_nodes["Node ID"] == node_1]['x'].item()),
                              float(df_nodes[df_nodes["Node ID"] == node_1]['y'].item()),
                              float(df_nodes[df_nodes["Node ID"] == node_1]['z'].item())]
        node_2_coordinates = [float(df_nodes[df_nodes["Node ID"] == node_2]['x'].item()),
                              float(df_nodes[df_nodes["Node ID"] == node_2]['y'].item()),
                              float(df_nodes[df_nodes["Node ID"] == node_2]['z'].item())]

        direction_vector = [node_1_coordinates[0] - node_2_coordinates[0],
                            node_1_coordinates[1] - node_2_coordinates[1],
                            node_1_coordinates[2] - node_2_coordinates[2]]
        direction_vector_magnitude = (direction_vector[0] ** 2 + direction_vector[1] ** 2 + direction_vector[
            2] ** 2) ** 0.5

        mid_point_member = [(node_2_coordinates[0] + node_1_coordinates[0]) / 2,
                            (node_2_coordinates[1] + node_1_coordinates[1]) / 2,
                            (node_2_coordinates[2] + node_1_coordinates[2]) / 2]

        # Check Dot Product
        dot_product = ref_vector[0] * direction_vector[0] + ref_vector[1] * direction_vector[1] + ref_vector[2] * \
                      direction_vector[2]
        # print(dot_product)
        # Check value
        cos_theta = dot_product / (ref_vector_magnitude * direction_vector_magnitude)

        # Check distance from mid-point of reference member to mid-point of member
        distance = (((mid_point_ref[0] - mid_point_member[0])) ** 2 + (
            (mid_point_ref[1] - mid_point_member[1])) ** 2 + ((mid_point_ref[2] - mid_point_member[2])) ** 2) ** 0.5
        # print(distance)
        mid_point_dif = [float(mid_point_ref[0] - mid_point_member[0]), float(mid_point_ref[1] - mid_point_member[1]),
                         float(mid_point_ref[2] - mid_point_member[2])]
        # print(mid_point_dif)

        moving_x = False
        moving_z = False

        if ref_vector[0] != 0:
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
        # Check if the member is going in the x-direction and if it is:
        if moving_x == True:
            # Check whether it's going up
            if mid_point_dif[2] > 0:
                # Check whether the member in question is closer to the previous member
                if distance < closest_above[1]:
                    closest_above[0] = df_members.iloc[i]["Member ID"]
                    closest_above[1] = distance
            # Or down
            elif mid_point_dif[2] < 0:
                # Check whether the member in question is closer to the previous member
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


def calculate_utilisation(depth, width, top_bar_1, top_bar_spacing_1, top_bar_spacing_2, btm_bar_1, btm_bar_spacing_1,
                          btm_bar_spacing_2, Fx, Fy, Fz, Mx, My, Mz):
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

    top_bar_amount = [sheet["K27"].value, sheet["K28"].value, sheet["K29"].value, sheet["K30"].value,
                      sheet["K31"].value]
    btm_bar_amount = [sheet["K34"].value, sheet["K35"].value, sheet["K36"].value, sheet["K37"].value,
                      sheet["K38"].value]

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

    # average_moment("Deck Longs ULS1000-1999.TXT", 1733)
    # push_values("MASTER.xlsx", "")

    # For checking importing SPACEGASS script file
    # import_spacegass_script("MASTER.xlsx")

    # For checking importing section properties Excel file
    # import_section_properties("MASTER.xlsx")
