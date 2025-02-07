import streamlit as st
import pandas as pd
import numpy as np
import io

# Hardcoded table data
ceiling_reflectance_levels = [90, 80, 70, 60, 50]
wall_reflectance_levels = [90, 80, 70, 60, 50, 40, 30, 20, 10, 0]
rcr_values = [0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6, 1.8, 2.0, 2.2, 2.4, 2.6, 2.8, 3.0, 3.2, 3.4, 3.6, 3.8, 4.0, 4.2, 4.4, 4.6, 4.8, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]

utilization_factors = {
        (90, 90): [89, 88, 87, 87, 86, 85, 85, 84, 83, 83, 82, 82, 81, 81, 80, 79, 79, 78, 78, 77, 77, 76, 76, 75, 75, 73, 70, 68, 68, 65],
        (90, 80): [88, 87, 86, 85, 83, 82, 80, 79, 78, 77, 76, 75, 74, 73, 72, 71, 70, 69, 69, 69, 62, 61, 60, 59, 59, 61, 58, 55, 52, 51],
        (90, 70): [88, 86, 84, 82, 80, 78, 77, 75, 73, 72, 70, 69, 67, 66, 64, 63, 62, 61, 60, 58, 57, 56, 55, 54, 53, 49, 45, 42, 38, 36],
        (90, 60): [87, 85, 82, 80, 77, 75, 73, 71, 69, 67, 65, 64, 62, 60, 58, 56, 54, 53, 51, 51, 50, 49, 47, 46, 45, 41, 38, 35, 31, 29],
        (90, 50): [86, 84, 80, 77, 75, 72, 69, 67, 64, 62, 59, 58, 56, 54, 52, 50, 48, 47, 45, 44, 43, 42, 40, 39, 38, 34, 30, 27, 25, 22], 
    
        (80, 70): [78, 76, 75, 73, 72, 70, 68, 67, 66, 64, 63, 61, 60, 59, 58, 57, 56, 54, 53, 53, 52, 51, 50, 49, 48, 44, 41, 38, 36, 33],
}
        

def find_utilization_factor(ceiling_reflectance, wall_reflectance, rcr):
    key = (ceiling_reflectance, wall_reflectance)
    if key in utilization_factors:
        if rcr in rcr_values:
            return utilization_factors[key][rcr_values.index(rcr)]
        lower_rcr = max([r for r in rcr_values if r <= rcr], default=None)
        upper_rcr = min([r for r in rcr_values if r >= rcr], default=None)
        if lower_rcr is None or upper_rcr is None:
            return None
        lower_index = rcr_values.index(lower_rcr)
        upper_index = rcr_values.index(upper_rcr)
        return utilization_factors[key][lower_index] + ((rcr - lower_rcr) / (upper_rcr - lower_rcr)) * (utilization_factors[key][upper_index] - utilization_factors[key][lower_index])
    return None

st.title("Lamp Requirement & Utilization Factor Calculator")

# Initialize session state for storing rooms
if "rooms" not in st.session_state:
    st.session_state.rooms = []

# Sidebar for RCR Calculation
st.sidebar.header("Room Cavity Ratio (RCR) Calculation")
hrc = st.sidebar.number_input("Height from lighting to work area (m):", min_value=0.0)
perimeter = st.sidebar.number_input("Perimeter of the room (m):", min_value=0.0)
area = st.sidebar.number_input("Area of the room (m²):", min_value=0.0)

RCR = (2.5 * hrc * perimeter) / area if perimeter > 0 and area > 0 else None
if RCR is not None:
    st.sidebar.write(f"Calculated RCR: {RCR:.2f}")
else:
    st.sidebar.write("Please enter valid room dimensions.")

# User input for reflectance values
ceiling_reflectance = st.selectbox("Select Ceiling Reflectance", ceiling_reflectance_levels)
wall_reflectance = st.selectbox("Select Wall Reflectance", wall_reflectance_levels)

# Calculate Utilization Factor
utilization_factor = find_utilization_factor(ceiling_reflectance, wall_reflectance, RCR) if RCR is not None else None
if utilization_factor is not None:
    st.success(f"Utilization Factor: {utilization_factor:.2f}")
else:
    st.error("No matching reflectance values found in the table.")

# Lamp Requirement Calculation
st.header("Lamp Requirement Calculator")
area_name = st.text_input("Enter Area Name:")
description = st.text_input("Enter Description of Fitting:")
watts = st.number_input("Enter Watts of the Fitting:", min_value=0.0)
E = st.number_input("Enter the illuminance level required (lux):", min_value=0.0)
A = st.number_input("Enter the area at working plane height (m²):", min_value=0.0)
F = st.number_input("Enter the average luminous flux from each lamp (lm):", min_value=0.0)
MF = 0.80  # Fixed maintenance factor

# Compute N (number of lamps)
N = (E * A) / (F * (utilization_factor / 100) * MF) if F > 0 and utilization_factor else None
if N is not None:
    st.write(f"Number of lamps required: {N:.2f}")
else:
    st.write("Please enter valid values for luminous flux and illuminance.")

# Add Room to List
if st.button("Add Room"):
    if area_name and description and F > 0 and utilization_factor:
        st.session_state.rooms.append({
            "Area Name": area_name,
            "Description": description,
            "Watts": watts,
            "E": E,
            "A": A,
            "F": F,
            "U.F": utilization_factor if utilization_factor is not None else 0,
            "M.F": MF,
            "N": N if N is not None else 0
        })
        st.success(f"Room '{area_name}' added successfully!")
    else:
        st.error("Please fill in all required fields before adding.")

# Display added rooms
if st.session_state.rooms:
    st.subheader("Added Rooms")
    df = pd.DataFrame(st.session_state.rooms)
    st.dataframe(df)

    # Export to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Lamp Calculation")
    output.seek(0)

    st.download_button(label="Download Excel File",
                       data=output,
                       file_name="Lamp_Calculation.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
