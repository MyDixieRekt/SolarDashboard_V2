import plotly.graph_objects as go

def create_fig_pie(peak, off_peak, file_name, file_date):
    fig_pie = go.Figure(data=[go.Pie(
        labels=['Peak Power', 'Off-Peak Power'],
        values=[peak, off_peak],
        textinfo='label+percent',
        insidetextorientation='radial',
        textposition='inside'
    )])

    fig_pie.update_layout(
        title=f"{file_name}",
        annotations=[
            dict(
                x=0.5,
                y=-0.1,
                xref='paper',
                yref='paper',
                text=f"Date: {file_date}",
                showarrow=False
            )
        ]
    )
    return fig_pie

def plot_power_distribution(total_peak, total_off_peak):
    fig_pie = go.Figure(data=[go.Pie(
        labels=['Peak Power', 'Off-Peak Power'],
        values=[total_peak, total_off_peak],
        textinfo='label+percent',
        insidetextorientation='radial'
    )])

    fig_pie.update_layout(
        title="Power Distribution",
    )
    return fig_pie

import plotly.graph_objects as go

def plot_peak_values(file_names, peak_values, dates):
    fig_peak = go.Figure()
    fig_peak.add_trace(go.Scatter(
        x=file_names,
        y=peak_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(peak_values, dates)],
        hoverinfo='text'
    ))

    fig_peak.update_layout(
        title="Peak Values Over Time",
        xaxis_title="File Names",
        yaxis_title="Peak (kWh)",
        xaxis=dict(tickangle=-45)
    )

    return fig_peak

def plot_peak_values_baht(file_names, peak_baht_values, dates):
    fig_peak_baht = go.Figure()
    fig_peak_baht.add_trace(go.Scatter(
        x=file_names,
        y=peak_baht_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} Baht<br>Date: {date}' for val, date in zip(peak_baht_values, dates)],
        hoverinfo='text'
    ))

    fig_peak_baht.update_layout(
        title="Peak Values in Baht Over Time",
        xaxis_title="File Names",
        yaxis_title="Peak (Baht)",
        xaxis=dict(tickangle=-45)
    )

    return fig_peak_baht

def plot_off_peak_values(file_names, off_peak_values, dates):
    fig_off_peak = go.Figure()
    fig_off_peak.add_trace(go.Scatter(
        x=file_names,
        y=off_peak_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(off_peak_values, dates)],
        hoverinfo='text'
    ))

    fig_off_peak.update_layout(
        title="Off-Peak Values Over Time",
        xaxis_title="File Names",
        yaxis_title="Off-Peak (kWh)",
        xaxis=dict(tickangle=-45)
    )

    return fig_off_peak

def plot_off_peak_values_baht(file_names, off_peak_baht_values, dates):
    fig_off_peak_baht = go.Figure()
    fig_off_peak_baht.add_trace(go.Scatter(
        x=file_names,
        y=off_peak_baht_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} Baht<br>Date: {date}' for val, date in zip(off_peak_baht_values, dates)],
        hoverinfo='text'
    ))

    fig_off_peak_baht.update_layout(
        title="Off-Peak Values in Baht Over Time",
        xaxis_title="File Names",
        yaxis_title="Off-Peak (Baht)",
        xaxis=dict(tickangle=-45)
    )

    return fig_off_peak_baht

def plot_power_values(file_names, power_values, dates):
    fig_power = go.Figure()
    fig_power.add_trace(go.Scatter(
        x=file_names,
        y=power_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(power_values, dates)],
        hoverinfo='text'
    ))

    fig_power.update_layout(
        title="Power Values Over Time",
        xaxis_title="File Names",
        yaxis_title="Power (kWh)",
        xaxis=dict(tickangle=-45)
    )

    return fig_power

def plot_combined_power_values(file_names, power_values, peak_values, off_peak_values, dates):
    fig_combined = go.Figure()

    fig_combined.add_trace(go.Bar(
        x=file_names,
        y=power_values,
        name='Total Power',
        marker=dict(color='rgba(55, 83, 109, 0.7)'),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(power_values, dates)],
        hoverinfo='text',
        textposition='none'
    ))

    fig_combined.add_trace(go.Bar(
        x=file_names,
        y=peak_values,
        name='Peak Power',
        marker=dict(color='rgba(26, 118, 255, 0.7)'),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(peak_values, dates)],
        hoverinfo='text',
        textposition='none'  
    ))
    
    fig_combined.add_trace(go.Bar(
        x=file_names,
        y=off_peak_values,
        name='Off-Peak Power',
        marker=dict(color='rgba(50, 171, 96, 0.7)'),
        text=[f'{val:,.2f} kWh<br>Date: {date}' for val, date in zip(off_peak_values, dates)],
        hoverinfo='text',
        textposition='none' 
    ))

    fig_combined.update_layout(
        title="Peak, Off-Peak, and Total Power Values Over Time",
        xaxis_title="File Names",
        yaxis_title="Values (kWh)",
        xaxis=dict(tickangle=-45),
        barmode='group'  
    )

    return fig_combined

def plot_peak_power_values(file_names, peak_power_values, dates):
    fig_peak_power = go.Figure()
    fig_peak_power.add_trace(go.Scatter(
        x=file_names,
        y=peak_power_values,
        mode='markers+lines',
        marker=dict(size=10), 
       text=[f'{val:,.2f} kW<br>Date: {date}' for val, date in zip(peak_power_values, dates)],
        hoverinfo='text'
    ))

    fig_peak_power.update_layout(
        title="Peak Power Values Over Time",
        xaxis_title="File Names",
        yaxis_title="Peak Power (kW)",
        xaxis=dict(tickangle=-45)
    )

    return fig_peak_power

def plot_electrical_cost(file_names, e_cost_values, dates):
    fig_e_cost = go.Figure()
    fig_e_cost.add_trace(go.Scatter(
        x=file_names,
        y=e_cost_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(e_cost_values, dates)],
        hoverinfo='text'
    ))

    fig_e_cost.update_layout(
        title="Electrical Cost Over Time",
        xaxis_title="File Names",
        yaxis_title="Baht(B)",
        xaxis=dict(tickangle=-45)
    )

    return fig_e_cost

def plot_discount_values(file_names, discount_values, dates):
    fig_discount = go.Figure()
    fig_discount.add_trace(go.Scatter(
        x=file_names,
        y=discount_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(discount_values, dates)],
        hoverinfo='text'
    ))

    fig_discount.update_layout(
        title="Total Discount Over Time",
        xaxis_title="File Names",
        yaxis_title="Baht(B)",
        xaxis=dict(tickangle=-45)
    )

    return fig_discount

def plot_discount_percentage(file_names, discount_values, dates):
    fig_discount = go.Figure()
    fig_discount.add_trace(go.Scatter(
        x=file_names,
        y=discount_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:.0f}%<br>Date: {date}' for val, date in zip(discount_values, dates)],
        hoverinfo='text'
    ))

    fig_discount.update_layout(
        title="Discount Percentage Over Time",
        xaxis_title="File Names",
        yaxis_title="Discount Percentage (%)",
        xaxis=dict(tickangle=-45)
    )

    return fig_discount

def plot_net_electric_cost(file_names, net_e_cost_values, dates):
    fig_net_e_cost = go.Figure()
    fig_net_e_cost.add_trace(go.Scatter(
        x=file_names,
        y=net_e_cost_values,
        mode='markers+lines',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(net_e_cost_values, dates)],
        hoverinfo='text'
    ))

    fig_net_e_cost.update_layout(
        title="Total Net Electric Cost Over Time",
        xaxis_title="File Names",
        yaxis_title="Baht(B)",
        xaxis=dict(tickangle=-45)
    )

    return fig_net_e_cost

def plot_combined_cost(file_names, e_cost_values, discount_values, net_e_cost_values, dates):
    fig_combined_cost = go.Figure()

    fig_combined_cost.add_trace(go.Scatter(
        x=file_names,
        y=e_cost_values,
        mode='markers+lines',
        name='Electric Cost',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(e_cost_values, dates)],
        hoverinfo='text'
    ))

    fig_combined_cost.add_trace(go.Scatter(
        x=file_names,
        y=discount_values,
        mode='markers+lines',
        name='Discount',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(discount_values, dates)],
        hoverinfo='text'
    ))

    fig_combined_cost.add_trace(go.Scatter(
        x=file_names,
        y=net_e_cost_values,
        mode='markers+lines',
        name='Net Electric Cost',
        marker=dict(size=10),
        text=[f'{val:,.2f} (Baht)<br>Date: {date}' for val, date in zip(net_e_cost_values, dates)],
        hoverinfo='text'
    ))

    fig_combined_cost.update_layout(
        title="Electric Cost, Discount, and Net Electric Cost Over Time",
        xaxis_title="File Names",
        yaxis_title="Baht(B)",
        xaxis=dict(tickangle=-45)
    )

    return fig_combined_cost