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
    options = ["Pie Charts", "Total Power Distribution", "Cost", "Discount", "Minimum Guarantee", "All Meters"]
    selected_option = st.selectbox("Select Attributes", options)

    excel_file = pd.ExcelFile(uploaded_files[0])
    sheet_names = excel_file.sheet_names

    selected_sheet = st.selectbox("Select sheet for all files", sheet_names)

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
    date_tracker = {}

    for file in uploaded_files:
        try:
            data = load_data(file, selected_sheet)
            date = extract_date_from_excel(file, selected_sheet)
            if date:
                file_dates.append((file, date))
                if date in date_tracker:
                    date_tracker[date].append(file.name)
                else:
                    date_tracker[date] = [file.name]
        except InvalidExcelFormatException as e:
            st.sidebar.header(file.name)
            st.sidebar.error(f"{file.name} will be excluded because the format does not match.")
            continue

    file_dates = natsorted(file_dates, key=lambda x: x[1])

    dates = [pd.to_datetime(date) for _, date in file_dates]
    years = sorted(set(date.year for date in dates))
    months = list(range(1, 13))

    # Calculate missing months and display full months with missing ones highlighted in red
    for year in years:
        year_dates = [date for date in dates if date.year == year]
        year_months = sorted(set(date.month for date in year_dates))
        missing_months = [month for month in months if month not in year_months]
        
        st.sidebar.header(f"Months for {year}")
        cols = st.sidebar.columns(12)
        for i, month in enumerate(months):
            if month in missing_months:
                cols[i].markdown(f"<span style='color:red'>{month}</span>", unsafe_allow_html=True)
            else:
                cols[i].write(month)

else:
    st.info("Upload a file through config")
    st.stop()

start_year = st.selectbox("Select Start Year", years)
start_month = st.selectbox("Select Start Month", months)
end_year = st.selectbox("Select End Year", years, index=len(years)-1)
end_month = st.selectbox("Select End Month", months, index=11)

start_date = pd.Timestamp(year=start_year, month=start_month, day=1)
end_date = pd.Timestamp(year=end_year, month=end_month, day=1) + pd.offsets.MonthEnd(0)

filtered_file_dates = [(file, date) for file, date in file_dates if start_date <= pd.to_datetime(date) <= end_date]

# Check for duplicate dates and display warnings after the date selection
duplicate_dates = {date: files for date, files in date_tracker.items() if len(files) > 1}

if duplicate_dates:
    st.warning("Warning: Duplicate dates found in the uploaded files!")
    for date, files in duplicate_dates.items():
        st.write(f"Date: {pd.to_datetime(date).strftime('%Y-%m-%d')}")
        st.write("Files:")
        for file in files:
            st.write(f"- {file}")

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

        st.header(f"Total Net Electric Cost: {total_net_electric_cost:,.2f} (Baht)")
        st.write("(Electric Cost - Discount)")
        fig_net_e_cost = plot_net_electric_cost(file_names, net_e_cost_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_net_e_cost)

        st.header("Electric Cost, Discount, and Net Electric Cost")
        fig_combined_cost = plot_combined_cost(file_names, e_cost_values, discount_values, net_e_cost_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_combined_cost)

    if selected_option == "Discount":

        st.header(f"Total Discount: {total_discount:,.2f} (Baht)")
        fig_discount = plot_discount_values(file_names, discount_values, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_discount)

        st.header(f"Discount Percentage")
        fig_discount_percentage = plot_discount_percentage(file_names, discount_percentages, [date for _, date in filtered_file_dates])
        st.plotly_chart(fig_discount_percentage)


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

        total_power_all_meters = sum(meter_data.values())

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
            
            st.title("All Meters")
            st.header(f"Total Power: {total_power_all_meters:,.2f} (kWh)")
            
            all_meters_minimum_guarantee = st.number_input(
                "Enter Minimum Guarantee for All Meters",
                min_value=0.0,
                format="%.2f"
            )
            
            if all_meters_minimum_guarantee > 0:
                missing_power_all_meters = max(0, all_meters_minimum_guarantee - total_power_all_meters)
                st.write(f"Missing to complete Minimum Guarantee: <span style='font-size:20px; font-weight:bold;'>{missing_power_all_meters:,.2f} (kWh)</span>", unsafe_allow_html=True)
                
                if total_power_all_meters >= all_meters_minimum_guarantee:
                    st.write("Goal has been reached for All Meters!")
                    progress_all_meters = 1.0
                else:
                    progress_all_meters = total_power_all_meters / all_meters_minimum_guarantee
                
                st.write(f"All Meters Progress: {total_power_all_meters:,.2f} / {all_meters_minimum_guarantee:,.2f} kWh")
                st.progress(progress_all_meters)
            else:
                st.write("Please enter a valid minimum guarantee for All Meters.")
        else:
            st.write("No valid meter data found.")

    if selected_option == "All Meters":
        meter_summaries = {}
        total_peak = 0
        total_off_peak = 0
        total_power = 0
        total_electric_cost = 0
        total_discount = 0
        total_net_electric_cost = 0
        total_peak_baht = 0
        total_off_peak_baht = 0

        for file, date in filtered_file_dates:  
            excel_file = pd.ExcelFile(file)
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    meter_name_match = re.search(r'Meter\s?\d+', sheet_name)
                    if meter_name_match:
                        meter_name = meter_name_match.group(0)
                        if meter_name not in meter_summaries:
                            meter_summaries[meter_name] = {
                                'peak': 0,
                                'off_peak': 0,
                                'power': 0,
                                'electric_cost': 0,
                                'discount': 0,
                                'net_electric_cost': 0,
                                'peak_baht': 0,
                                'off_peak_baht': 0
                            }

                        peak = convert_to_number(df.iloc[32]['Unnamed: 1'])
                        baht_peak = convert_to_number(df.iloc[32]['Unnamed: 3'])
                        check_type(peak, (int, float))
                        check_type(baht_peak, (int, float))
                        meter_summaries[meter_name]['peak'] += peak
                        meter_summaries[meter_name]['peak_baht'] += baht_peak

                        off_peak = convert_to_number(df.iloc[33]['Unnamed: 1'])
                        baht_off_peak = convert_to_number(df.iloc[33]['Unnamed: 3'])
                        check_type(off_peak, (int, float))
                        check_type(baht_off_peak, (int, float))
                        meter_summaries[meter_name]['off_peak'] += off_peak
                        meter_summaries[meter_name]['off_peak_baht'] += baht_off_peak

                        power = peak + off_peak
                        check_type(power, (int, float))
                        meter_summaries[meter_name]['power'] += power

                        electic_cost = convert_to_number(df.iloc[39]['Unnamed: 3'])
                        check_type(electic_cost, (int, float))
                        meter_summaries[meter_name]['electric_cost'] += electic_cost

                        discount = convert_to_number(df.iloc[40]['Unnamed: 3'])
                        check_type(discount, (int, float))
                        meter_summaries[meter_name]['discount'] += discount

                        net_electric_cost = electic_cost - discount
                        check_type(net_electric_cost, (int, float))
                        meter_summaries[meter_name]['net_electric_cost'] += net_electric_cost

                        total_peak += peak
                        total_peak_baht += baht_peak
                        total_off_peak += off_peak
                        total_off_peak_baht += baht_off_peak
                        total_power += power
                        total_electric_cost += electic_cost
                        total_discount += discount
                        total_net_electric_cost += net_electric_cost

                except (KeyError, IndexError, ValueError) as e:
                    st.sidebar.write(f"**{sheet_name}**: <span style='color:red'>Missing or Invalid Type</span>", unsafe_allow_html=True)
                    st.sidebar.error(f"{file.name} - {sheet_name} will be excluded because it has missing or invalid type information. Error: {str(e)}")
                    continue

        st.header("All Meters Summary")
        st.write(f"**Total Peak**: {total_peak:,.2f} kWh ({total_peak_baht:,.2f} Baht)")                            
        st.write(f"**Total Off-Peak**: {total_off_peak:,.2f} kWh ({total_off_peak_baht:,.2f} Baht)")
        st.write(f"**Total Power**: {total_power:,.2f} kWh")
        st.write(f"**Total Electric Cost**: {total_electric_cost:,.2f} Baht")
        st.write(f"**Total Discount**: {total_discount:,.2f} Baht")
        st.write(f"**Total Net Electric Cost**: {total_net_electric_cost:,.2f} Baht")

        for meter_name, summary in meter_summaries.items():
            st.header(f"{meter_name} Summary")
            st.write(f"**Peak**: {summary['peak']:,.2f} kWh ({summary['peak_baht']:,.2f} Baht)")
            st.write(f"**Off-Peak**: {summary['off_peak']:,.2f} kWh ({summary['off_peak_baht']:,.2f} Baht)")
            st.write(f"**Power**: {summary['power']:,.2f} kWh")
            st.write(f"**Electric Cost**: {summary['electric_cost']:,.2f} Baht")
            st.write(f"**Discount**: {summary['discount']:,.2f} Baht")
            st.write(f"**Net Electric Cost**: {summary['net_electric_cost']:,.2f} Baht")

        categories = ['Peak', 'Off Peak', 'Power', 'Electric Cost', 'Discount', 'Net Electric Cost']
        fig = go.Figure()

        for meter_name, summary in meter_summaries.items():
            fig.add_trace(go.Bar(
                x=categories,
                y=[summary['peak'], summary['off_peak'], summary['power'], summary['electric_cost'], summary['discount'], summary['net_electric_cost']],
                name=meter_name,
                hovertemplate='%{y:,}'
            ))

        fig.update_layout(
            title='Comparison of Meters',
            xaxis_title='Categories',
            yaxis_title='Values',
            barmode='group'
        )

        st.plotly_chart(fig)

