import pandas as pd
import sys
import re
import numpy as np
import datetime
import math
from ast import literal_eval
from scipy.optimize import fsolve
from bokeh.plotting import figure, output_file, show, ColumnDataSource
from bokeh.models import HoverTool, DatetimeTickFormatter

# -----------------------------------------------------------------------

# yearly discount rate applied in the deal model
YEARLY_DISC = 0.2

# -----------------------------------------------------------------------

# start the clock
start_time = datetime.datetime.now()
print('Starting clock: {}'.format(start_time))

td = pd.read_excel('Tidy Data - F - 18.12.03.xlsx',
                   sheet_name='Tidy Data',
                   converters={'T': eval})
# Color coding
colors = {'2016Q1': 'orangered',
          '2016Q2': 'indianred',
          '2016Q3': 'red',
          '2016Q4': 'darkred',
          '2017Q1': 'greenyellow',
          '2017Q2': 'lawngreen',
          '2017Q3': 'green',
          '2017Q4': 'darkseagreen',
          '2018Q1': 'lightskyblue',
          '2018Q2': 'royalblue',
          '2018Q3': 'blue',
          '2018Q4': 'darkslateblue'}

# Create a new plot with a title and axis labels
p = figure(title="Deal Data Analysis - Deal Model",
           x_axis_label='Time',
           y_axis_label='Monthly Revenue',
           plot_width=1200,
           plot_height=630,
           x_axis_type='datetime')

# -----------------------------------------------------------------------

# Plot all deals, in deal model, one line each
# Initialise lists of lists
ts_x_list_of_list = []
vals_y_list_of_list = []
cols_to_use = []
legend_to_use = []
customers_list = []
tcv_list = []
duration_list = []
for x, row in td.iterrows():
    # load years
    if td.loc[x, 'years'] != 0:
        years = literal_eval(td.loc[x, 'years'])
    else:
        years = []

    deal_start = td.loc[x, 'Starting Date']
    deal_date_point = deal_start
    date_portions = []
    revenue_portions = []

    # Starting year's monthly revenue
    if years == []:
        current_year_monthly_revenue = 0
    else:
        current_year_monthly_revenue = years[0][0]/12
    td.loc[x, 'SYMR'] = current_year_monthly_revenue

    for tcv_portion, duration_portion in years:
        date_portions.append([deal_date_point])
        revenue_portions.append([current_year_monthly_revenue])
        deal_date_point += pd.DateOffset(months=np.ceil(duration_portion)-1)

        date_portions.append([deal_date_point])
        revenue_portions.append([current_year_monthly_revenue])
        deal_date_point += pd.DateOffset(months=1)

        current_year_monthly_revenue *= (1-YEARLY_DISC)

    # Compile xs & ys for multi_line
    ts_x_list_of_list.append(date_portions)
    vals_y_list_of_list.append(revenue_portions)

    # Color by quarter & add hover data
    cols_to_use.append([colors[str(td.loc[x, 'Qtr'])]])
    legend_to_use.append(td.loc[x, 'Qtr'])
    customers_list.append([td.loc[x, 'Customer']])
    tcv_list.append([td.loc[x, 'TCV']])
    duration_list.append([td.loc[x, 'Duration Max']])

# Plot lines for the deals
source = ColumnDataSource(data=dict(xs=ts_x_list_of_list,
                                    ys=vals_y_list_of_list,
                                    colors=cols_to_use,
                                    legends=legend_to_use,
                                    Customer=customers_list,
                                    TCV=tcv_list,
                                    Duration=duration_list))

deal_lines = p.multi_line(xs='xs',
                          ys='ys',
                          source=source,
                          line_color='colors',
                          legend='legends')

p.add_tools(HoverTool(renderers=[deal_lines],
                      tooltips=[('Customer', '@Customer'),
                                ('TCV', '@TCV'),
                                ('Duration', '@Duration')]))

# Plot cumulative revenue line
# Get all columms right from 'years'
cr = td.iloc[:, td.columns.get_loc('years')+1:]

dates = cr.columns
total_revenues = cr.sum()

# Format hover
hover_date = []
for x in str(dates.format()).replace(' 00:00:00', '').replace('[', '')\
        .replace(']', '').replace("'", '').replace(' ', '').split(','):
    hover_date.append(x)
revenues = ColumnDataSource(data=dict(xs=dates,
                                      ys=total_revenues,
                                      xs_hover=hover_date))
revenue_line = p.line(x='xs',
                      y='ys',
                      source=revenues,
                      color='black',
                      legend='Total Monthly Revenue')

p.line(x=[datetime.datetime(2018, 9, 30), datetime.datetime(2018, 9, 30)],
       y=[0, max(total_revenues)],
       color='magenta',
       line_width=1)
p.add_tools(HoverTool(renderers=[revenue_line],
                      tooltips={'Starting Date': '@xs_hover',
                                'Monthly Revenue': '$y'}))

# Output to static HTML file
output_file(sys.argv[0].replace('.py', '.html'))

# Show the results
show(p)
