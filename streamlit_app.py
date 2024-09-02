import streamlit as st
import pandas as pd
from natsort import natsorted
import re
import plotly.graph_objects as go

from GraphFunctions import create_fig_pie
from GraphFunctions import plot_power_distribution
from GraphFunctions import plot_peak_values
from GraphFunctions import plot_off_peak_values
from GraphFunctions import plot_power_values
from GraphFunctions import plot_combined_power_values
from GraphFunctions import plot_peak_power_values
from GraphFunctions import plot_electrical_cost
from GraphFunctions import plot_discount_values
from GraphFunctions import plot_net_electric_cost
from GraphFunctions import plot_combined_cost
from GraphFunctions import plot_discount_percentage

st.set_page_config(
    page_title="PTTOR Solar Dashboard",
    page_icon="P_Ter_NoBG.png",
    layout="wide"
)

dark_mode_css = """
<style>
body {
    background-color: #0e1117;
    color: #c9d1d9;
}
header, .css-1d391kg, .css-1v3fvcr {
    background-color: #161b22;
}
</style>
"""

st.markdown(dark_mode_css, unsafe_allow_html=True)

logo_url = "ptt logo 2.png"

col1, col2 = st.columns([1, 3])

with col1:
    st.image(logo_url, width=200)

st.title("Solar Project Dashboard")

class InvalidExcelFormatException(Exception):
    pass

def load_data(file, sheet_name):
    data = pd.read_excel(file, sheet_name=sheet_name)
    required_columns = ['Unnamed: 3']
    missing_columns = [column for column in required_columns if column not in data.columns]
    
    if missing_columns:
        raise InvalidExcelFormatException(
            f"The file does not have the required format. Missing columns: {', '.join(missing_columns)}"
        )
    
    return data

uploaded_files = st.file_uploader("Choose a file", accept_multiple_files=True)

if uploaded_files:
    options = ["Pie Charts", "Total Power Distribution", "Cost", "Discount Forecasting", "Minimum Guarantee"]
    selected_option = st.selectbox("Select Attributes", options)

    excel_file = pd.ExcelFile(uploaded_files[0])
    sheet_names = excel_file.sheet_names
    selected_sheet = st.selectbox("Select sheet for all files", sheet_names)

if not uploaded_files:
    st.info("Upload a file through config")
    st.stop()

def extract_date_from_excel(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=1, usecols="A:B")
    cell_content = df.iloc[0, 0]
    date_match = re.search(r"\d{4}-\d{2}-\d{2}", cell_content)
    if date_match:
        date = pd.to_datetime(date_match.group(0))
        formatted_date = date.strftime('%Y-%m-%d')
        return formatted_date
    else:
        return None

file_dates = []

for file in uploaded_files:
    try:
        data = load_data(file, selected_sheet)
        date = extract_date_from_excel(file, selected_sheet)
        if date:
            file_dates.append((file, date))
    except InvalidExcelFormatException as e:
        st.sidebar.header(file.name)
        st.sidebar.error(f"{file.name} will be excluded because the format does not match.")
        continue

file_dates = natsorted(file_dates, key=lambda x: x[1])

dates = [pd.to_datetime(date) for _, date in file_dates]
years = sorted(set(date.year for date in dates))
months = list(range(1, 13))

start_year = st.selectbox("Select Start Year", years)
start_month = st.selectbox("Select Start Month", months)
end_year = st.selectbox("Select End Year", years, index=len(years)-1)
end_month = st.selectbox("Select End Month", months, index=11)

start_date = pd.Timestamp(year=start_year, month=start_month, day=1)
end_date = pd.Timestamp(year=end_year, month=end_month, day=1) + pd.offsets.MonthEnd(0)

filtered_file_dates = [(file, date) for file, date in file_dates if start_date <= pd.to_datetime(date) <= end_date]

total_peak = 0
total_off_peak = 0
total_power = 0
total_electric_cost = 0
total_discount = 0
total_net_electric_cost = 0
total_peak_baht = 0
total_off_peak_baht = 0

peak_values = []
off_peak_values = []
power_values = []
peak_power_values = []
e_cost_values = []
discount_values = []
net_e_cost_values = []
file_names = []
peak_baht_values = []
off_peak_baht_values = []
discount_percentages = []

for i, (file, date) in enumerate(filtered_file_dates):

    try:
        df = load_data(file, selected_sheet)
    except InvalidExcelFormatException as e:
        st.sidebar.header(file.name)
        st.sidebar.error(str(e))
        continue

    st.sidebar.title(file.name)
    with st.sidebar.expander(f"Data Preview: {file.name}"):
        st.sidebar.dataframe(df)

    st.sidebar.write(f"**Date**: {date}")

    def check_type(value, expected_type):
        if not isinstance(value, expected_type):
            raise ValueError(f"Expected {expected_type}, but got {type(value)}")

    def convert_to_number(value):
        try:
            return float(value.replace(',', ''))
        except (ValueError, AttributeError):
            return value

    try:
        peak = convert_to_number(df.iloc[32]['Unnamed: 1'])
        baht_peak = convert_to_number(df.iloc[32]['Unnamed: 3'])
        check_type(peak, (int, float))
        check_type(baht_peak, (int, float))
        st.sidebar.write(f"**Peak**: {peak:,.2f} (kWh) ({baht_peak:,.2f} Baht)")
        total_peak += peak
        total_peak_baht += baht_peak
        peak_values.append(peak)
        peak_baht_values.append(baht_peak)
        file_names.append(file.name) 
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Peak**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{str(e)}")
        continue

    try:
        off_peak = convert_to_number(df.iloc[33]['Unnamed: 1'])
        baht_off_peak = convert_to_number(df.iloc[33]['Unnamed: 3'])
        check_type(off_peak, (int, float))
        check_type(baht_off_peak, (int, float))
        st.sidebar.write(f"**Off-Peak**: {off_peak:,.2f} (kWh) ({baht_off_peak:,.2f} Baht)")
        total_off_peak += off_peak
        total_off_peak_baht += baht_off_peak
        off_peak_values.append(off_peak)
        off_peak_baht_values.append(baht_off_peak)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Off-Peak**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    try:
        power = peak + off_peak
        check_type(power, (int, float))
        st.sidebar.write(f"**Power**: {power:,.2f} (kWh)")
        total_power += power
        power_values.append(power)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Power**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    try:
        peak_power = convert_to_number(df.iloc[36]['Unnamed: 1'])
        check_type(peak_power, (int, float))
        st.sidebar.write(f"**Peak Power**: {peak_power:,.2f} (kW)")
        peak_power_values.append(peak_power)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Peak Power**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    try:
        electic_cost = convert_to_number(df.iloc[39]['Unnamed: 3'])
        check_type(electic_cost, (int, float))
        st.sidebar.write(f"**Total Electric Cost**: {electic_cost:,.2f} Baht")
        total_electric_cost += electic_cost
        e_cost_values.append(electic_cost)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Total Electric Cost**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    try:
        discount = convert_to_number(df.iloc[40]['Unnamed: 3'])
        discount_percent = convert_to_number(df.iloc[40]['Unnamed: 1'])
        check_type(discount, (int, float))
        check_type(discount_percent, (int, float))
        discount_percent = discount_percent * 100
        st.sidebar.write(f"**Discount**: {discount:,.2f} Baht ({discount_percent:,.0f}%)")
        total_discount += discount
        discount_values.append(discount)
        discount_percentages.append(discount_percent)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Discount**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    try:
        net_electric_cost = electic_cost - discount
        check_type(net_electric_cost, (int, float))
        st.sidebar.write(f"**Net Electrical Cost (Baht) (ไม่รวมภาษี 7 %)**: {net_electric_cost:,.2f} Baht")
        total_net_electric_cost += net_electric_cost
        net_e_cost_values.append(net_electric_cost)
    except (KeyError, IndexError, ValueError) as e:
        st.sidebar.write(f"**Net Electrical Cost (Baht) (ไม่รวมภาษี 7 %)**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
        st.sidebar.error(f"{file.name} will be excluded because it has missing or invalid type information.")
        continue

    if selected_option == "Pie Charts":
        if i % 5 == 0:
            cols = st.columns(5)

        fig_pie = create_fig_pie(peak, off_peak, file.name, date)
        cols[i % 5].plotly_chart(fig_pie)

if uploaded_files:

    if selected_option == "Pie Charts":
        st.header("Total Power Distribution")
        fig_pie = plot_power_distribution(total_peak, total_off_peak)
        st.plotly_chart(fig_pie)

    if selected_option == "Total Power Distribution":

        st.header(f"Total Peak: {total_peak:,.2f} (kWh)")
        st.write(f"Total Peak in Baht: {total_peak_baht:,.2f} Baht")
        fig_peak = plot_peak_values(file_names, peak_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_peak)

        st.header(f"Total Off-Peak: {total_off_peak:,.2f} (kWh)")
        st.write(f"Total Off-Peak in Baht: {total_off_peak_baht:,.2f} Baht")
        fig_off_peak = plot_off_peak_values(file_names, off_peak_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_off_peak)
        
        st.header(f"Total Power: {total_power:,.2f} (kWh)")
        st.write(f"Total Power in Baht: {total_off_peak_baht+total_peak_baht:,.2f} Baht")
        fig_power = plot_power_values(file_names, power_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_power)

        st.header(f"Peak, Off-Peak, Power")
        fig_combined = plot_combined_power_values(file_names, power_values, peak_values, off_peak_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_combined)

        st.header("Peak Power (kW)")
        fig_peak_power = plot_peak_power_values(file_names, peak_power_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_peak_power)

    if selected_option == "Cost":

        st.header(f"Total Electric Cost: {total_electric_cost:,.2f} (Baht)")
        fig_e_cost = plot_electrical_cost(file_names, e_cost_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_e_cost)

        st.header(f"Total Discount: {total_discount:,.2f} (Baht)")
        fig_discount = plot_discount_values(file_names, discount_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_discount)

        st.header(f"Discount Percentage")
        fig_discount_percentage = plot_discount_percentage(file_names, discount_percentages, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_discount_percentage)

        st.header(f"Total Net Electric Cost: {total_net_electric_cost:,.2f} (Baht)")
        fig_net_e_cost = plot_net_electric_cost(file_names, net_e_cost_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_net_e_cost)

        st.header("Electric Cost, Discount, and Net Electric Cost")
        fig_combined_cost = plot_combined_cost(file_names, e_cost_values, discount_values, net_e_cost_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_combined_cost)

    if selected_option == "Discount Forecasting":
        st.title("Discount Forecasting")

        dates = []
        power_values = []  

        for file, date in filtered_file_dates:
            df = pd.read_excel(file)
            try:
                if pd.notna(df.iloc[32]['Unnamed: 3']) and pd.notna(df.iloc[33]['Unnamed: 3']):
                    power_32 = df.iloc[32]['Unnamed: 3']
                    power_33 = df.iloc[33]['Unnamed: 3']
                    if isinstance(power_32, (int, float)) and isinstance(power_33, (int, float)):
                        dates.append(date)
                        power = power_32 + power_33
                        power_values.append(power)
                    else:
                        st.write(f"**:red[File {file.name} has non-numeric values in required cells.]**")
                else:
                    st.write(f"**:red[File {file.name} has missing values in required cells.]**")
            except (KeyError, IndexError) as e:
                st.write(f"**:red[Error processing file {file.name}: {e}]**")

        dates = pd.to_datetime(dates)

        if len(dates) == 0:
            st.write("**:red[No data available for the selected period. Please select a different period.]**")
        else:
            power_series = pd.Series(power_values, index=dates)

            total_power = power_series.sum()
            num_months = len(power_series.resample('M').mean())
            average_power = total_power / num_months

            average_power_series = pd.Series([average_power] * num_months, index=power_series.resample('M').mean().index)

            fig = go.Figure()

            fig.add_trace(go.Scatter(
                x=power_series.index, 
                y=power_series, 
                mode='lines', 
                name='Power',
                hovertemplate='%{x|%Y-%m-%d}: %{y:,.0f}'  
            ))
            fig.add_trace(go.Scatter(
                x=average_power_series.index, 
                y=average_power_series, 
                mode='lines', 
                name='Average Power', 
                line=dict(dash='dash'),
                hovertemplate='%{y:,.0f}'
            ))

            st.plotly_chart(fig)

        power_sum = power_series.sum()

        if 0 <= power_sum <= 951700:
            discount = 13
            next_target = 951700
            next_discount = 100
            additional_power_needed = next_target - power_sum
            level = 1
        elif 951700 < power_sum <= 1163100:
            discount = 100
            next_target = 1163100
            next_discount = 50
            additional_power_needed = next_target - power_sum
            level = 2
        elif 1163100 < power_sum <= 1353000:
            discount = 50
            next_target = 1353000
            next_discount = 0
            additional_power_needed = next_target - power_sum
            level = 3
        else:
            discount = 0
            next_target = None
            next_discount = None
            additional_power_needed = None
            level = 0

        st.header(f"The total power consumption is: ")
        st.title(f"{power_sum:,.2f} kWh")
        st.header(f"The customer is eligible for: ")
        st.title(f"Level {level} ({discount})%")

        if power_sum <= 951700:
            st.title("Additional power needed for all discount levels:")
            st.markdown(f"<h2>Remaining power until Level 2 (100%): <span style='font-size:40px; font-weight: bold;'>{951700 - power_sum:,.2f} kWh</span></h2>", unsafe_allow_html=True)
            st.markdown(f"<h2>Remaining power until Level 3 (50%): <span style='font-size:40px; font-weight: bold;'>{1163100 - power_sum:,.2f} kWh</span></h2>", unsafe_allow_html=True)
        elif 951700 < power_sum <= 1163100:
            st.title("Additional power needed for all discount levels:")
            st.markdown(f"<h2>Remaining power until Level 3 (50%): <span style='font-size:40px; font-weight: bold;'>{1163100 - power_sum:,.2f} kWh</span></h2>", unsafe_allow_html=True)

    if selected_option == "Minimum Guarantee":
        st.title("Minimum Guarantee")

        meter_data = {}
        user_minimum_guarantees = {}

        for file in uploaded_files:
            excel_file = pd.ExcelFile(file)
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name)
                
                meter_name_match = re.search(r'Meter\s?\d+', sheet_name)
                if meter_name_match:
                    meter_name = meter_name_match.group(0)
                    
                    try:
                        peak = df.iloc[32]['Unnamed: 1']
                        off_peak = df.iloc[33]['Unnamed: 1']
                        check_type(peak, (int, float))
                        check_type(off_peak, (int, float))
                        power_value = peak + off_peak
                        
                        if meter_name not in meter_data:
                            meter_data[meter_name] = 0
                        meter_data[meter_name] += power_value
                    except (KeyError, IndexError, ValueError) as e:
                        st.write(f"**{sheet_name}**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
                        st.error(f"{file.name} - {sheet_name} will be excluded because it has missing or invalid type information. Error: {str(e)}")
                        continue

        if meter_data:
            for meter_name, total_power in meter_data.items():
                st.title(f"{meter_name}")
                st.header(f"Total Power: {total_power:,.2f} (kWh)")
                
                minimum_guarantee = st.number_input(
                    f"Enter Minimum Guarantee for {meter_name}",
                    min_value=0.0,
                    format="%.2f"
                )
                user_minimum_guarantees[meter_name] = minimum_guarantee
                
                if minimum_guarantee > 0:
                    missing_power = max(0, minimum_guarantee - total_power)
                    st.write(f"Missing to complete Minimum Guarantee: <span style='font-size:20px; font-weight:bold;'>{missing_power:,.2f} (kWh)</span>", unsafe_allow_html=True)
                    
                    if total_power >= minimum_guarantee:
                        st.write(f"Goal has been reached for {meter_name}!")
                        progress = 1.0
                    else:
                        progress = total_power / minimum_guarantee
                    
                    st.write(f"{meter_name} Progress: {total_power:,.2f} / {minimum_guarantee:,.2f} kWh")
                    st.progress(progress)
                else:
                    st.write(f"Please enter a valid minimum guarantee for {meter_name}.")
        else:
            st.write("No valid meter data found.")







