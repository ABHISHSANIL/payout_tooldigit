import os
import pandas as pd
from datetime import datetime
import streamlit as st

# Set the file path relative to the script's location
current_dir = os.path.dirname(os.path.abspath(__file__))  # Get the directory where the script is located
file_path = os.path.join(current_dir, "Viaansh_Insurance_Brokers.xlsx")  # Excel file in the same directory as the script

# Load Excel file
@st.cache_data
def load_data():
    # Load the Excel file
    try:
        # Load the sheets
        satp_rto_data = pd.read_excel(file_path, sheet_name='4W SATP RTO')
        satp_data = pd.read_excel(file_path, sheet_name='4W SATP', skiprows=1)  # Skip first row for proper headers
    except FileNotFoundError:
        st.error("The Excel file could not be found. Ensure it's in the correct location.")
        st.stop()
    except ValueError as e:
        st.error(f"Error loading Excel sheets: {e}")
        st.stop()

    # Clean and adjust `4W SATP RTO`
    if 'RTO' in satp_rto_data.columns and 'New Cluster' in satp_rto_data.columns:
        satp_rto_data = satp_rto_data.rename(columns={"RTO": "RTO Code", "New Cluster": "Cluster"})
        satp_rto_data = satp_rto_data[["RTO Code", "Cluster"]].dropna()
    else:
        st.error("The '4W SATP RTO' sheet is missing required columns: 'RTO' and 'New Cluster'.")
        st.stop()

    # Clean and adjust `4W SATP`
    if len(satp_data.columns) >= 5:
        satp_data.columns = ["Cluster", "Segment Mapping", "Age Band", "Max CD2", "Avg CD2"]
        satp_data = satp_data.dropna(subset=["Cluster", "Segment Mapping", "Avg CD2"])
    else:
        st.error("The '4W SATP' sheet does not contain the expected number of columns.")
        st.stop()

    return satp_rto_data, satp_data

satp_rto_data, satp_data = load_data()

# Function to calculate age band based on registration month/year
def get_age_band(reg_month_year):
    reg_date = datetime.strptime(reg_month_year, "%m/%Y")
    current_date = datetime.now()
    age_years = (current_date.year - reg_date.year) + (current_date.month - reg_date.month) / 12.0
    return ">10" if age_years > 10 else "<10"

# Function to determine segment mapping based on fuel type and engine capacity
def refined_get_segment_mapping(fuel_type, engine_capacity, available_segments):
    if fuel_type == "petrol":
        if engine_capacity < 1000:
            return "Petrol<1000"
        elif 1000 <= engine_capacity <= 1499:
            return "Petrol1000-1500" if "Petrol1000-1500" in available_segments else "Petrol>1000"
        else:
            return "Petrol>1500" if "Petrol>1500" in available_segments else "Petrol>1000"
    elif fuel_type == "cng":
        if engine_capacity < 1000:
            return "CNG<1000"
        elif 1000 <= engine_capacity <= 1499:
            return "CNG1000-1500" if "CNG1000-1500" in available_segments else "CNG>1000"
        else:
            return "CNG>1500" if "CNG>1500" in available_segments else "CNG>1000"
    elif fuel_type == "diesel":
        if engine_capacity < 1500:
            return "Diesel<1500"
        else:
            return "Diesel>1500"
    return None

# Function to find Avg CD2 and retrieve relevant segment data
def refined_find_avg_cd2(rto_code, fuel_type, engine_capacity, reg_month_year):
    # Step 1: Match RTO to cluster
    cluster = satp_rto_data.loc[satp_rto_data['RTO Code'] == rto_code, 'Cluster']
    if cluster.empty:
        return f"Cluster not found for RTO: {rto_code}", None
    cluster = cluster.iloc[0]

    # Step 2: Get available segments for this cluster
    available_segments = satp_data[satp_data['Cluster'] == cluster]['Segment Mapping'].unique()

    # Step 3: Determine segment mapping and age band
    segment_mapping = refined_get_segment_mapping(fuel_type, engine_capacity, available_segments)
    age_band = get_age_band(reg_month_year)

    # Step 4: Find matching row in "4W SATP"
    relevant_data = satp_data[
        (satp_data['Cluster'] == cluster) &
        (satp_data['Segment Mapping'] == segment_mapping) &
        ((satp_data['Age Band'] == age_band) | (satp_data['Age Band'] == "All"))
    ]

    if relevant_data.empty:
        return f"No matching data for Cluster: {cluster}, Segment: {segment_mapping}, Age Band: {age_band}", None

    avg_cd2 = relevant_data['Avg CD2'].iloc[0]
    return avg_cd2, relevant_data

# Streamlit Web Interface
st.title("Payout Automation Tool")

# Input Form
with st.form("Input Form"):
    rto_code = st.text_input("Enter RTO Code (e.g., MH04)").upper()
    fuel_type = st.selectbox("Select Fuel Type", ["petrol", "diesel", "cng"]).lower()
    engine_capacity = st.number_input("Enter Engine Capacity (CC)", min_value=1, step=1)
    reg_month_year = st.text_input("Enter Registration Month/Year (MM/YYYY)")
    submit = st.form_submit_button("Calculate Payout")

# Process Input and Display Results
if submit:
    if rto_code and fuel_type and engine_capacity and reg_month_year:
        result, relevant_data = refined_find_avg_cd2(rto_code, fuel_type, engine_capacity, reg_month_year)
        if isinstance(result, str):
            st.error(result)
        else:
            st.success(f"The payout for the given criteria is: {result * 100:.1f}%")
        if relevant_data is not None:
            st.subheader("Relevant Segment Data:")
            st.dataframe(relevant_data)
