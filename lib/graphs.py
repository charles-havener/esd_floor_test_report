# plotly and pandas
# open up files 1-x and pull data from them to create the graphs
# can get the temp/rh/date/time save as variables by reading side cells

import pandas as pd
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import win32com.client
import os

def correct_corrupt_files(path):
    # The esd meter creates old excel files that pandas thinks are currupt
    # need to convert them if not alreday converted
    # returns path to the new uncorrupt file
    if not os.path.exists(path + ".xlsx"):
        print(f"Correcting corrupt file: {path}.xls")
        p = path + ".xls"
        o = win32com.client.Dispatch("Excel.Application")
        o.Visible = False

        filename = p
        output = path + ".xlsx"
        wb = o.Workbooks.Open(filename)
        wb.ActiveSheet.SaveAs(output,51)
        wb.Close(True)
        o.quit()
    return path + ".xlsx"


def get_data(path, n):
    p = correct_corrupt_files(path + "\\" + str(n))
    df = pd.read_excel(p, names=["Electric Potential (V)","Time (ms)"], usecols=[0,1], skiprows=1)

    # Retrieve the environmental/datetime data and fix it's messed up formatting issues
    env_data = pd.read_excel(p, header=None, usecols="C:F", skiprows=1, nrows=1, names=["Temp", "RH", "Time", "Date"])
    temp = env_data["Temp"][0].replace("_x0000_", "").replace(" ", "")
    hum = env_data["RH"][0].replace("_x0000_", "").replace(" ", "").replace("RH","")
    time = env_data["Time"][0].replace("_x0000_", "").replace(" ", "")
    date = env_data["Date"][0].replace("_x0000_", "").replace(" ", "")

    return df, temp, hum, time, date


def generate_charts(path, num):

    PLOTS_PER_PAGE = 5

    x_axis_title = "Time (ms)"
    y_axis_title = "Electric Potential (V)"

    blue_100_color = 'rgba(0,101,159,1)'
    axis_color = 'rgba(204,204,204,1)'
    axis_grid_color = 'rgba(204,204,204,0.5)'
    text_color = 'rgba(119,119,119,1)'
    background_color = 'rgba(255,255,255,1)'

    file_names = []

    count = 1
    page_count = 1
    for n in range(1, num+1):

        df, temp, hum, time, date = get_data(path, n)

        # Create new figure when starting a new page
        if count==1:
            title_map = {str(i+1):"" for i in range(PLOTS_PER_PAGE)}
            fig = make_subplots(
                rows=PLOTS_PER_PAGE, 
                cols=1,
                subplot_titles = [str(i+1) for i in range(PLOTS_PER_PAGE)],
            )

        title_map[str(count)] = f"Location {n} - {date} - {time} - {temp} - {hum}"

        # Create the graph at this location and add to page
        fig.append_trace(go.Scatter(
            x = df['Time (ms)'],
            y = df['Electric Potential (V)'],
            line = {
                "color": blue_100_color,
                "width": 1
            },
        ), row=count, col=1)

        # Create the page if on last element to be added to the pag
        count+=1
        if count>=PLOTS_PER_PAGE+1 or n==num:

            # Get title for page (can have less than plots_per_page if not evenly divisible)
            if count>=PLOTS_PER_PAGE:
                title = f"Locations {n-(PLOTS_PER_PAGE-1)}-{n}"
            else:
                first = num-(num%PLOTS_PER_PAGE)+1
                last = num
                if first == last:
                    title = f"Location {first}"
                else:
                    title = f"Locations {first}-{last}"

            # Update the layout of the stacked_subplots
            fig.update_layout(
                height=990, 
                width=765, 
                title = dict(
                    font = dict(
                        family = 'Open Sans, Arial, Helvetica, sans-serif',
                        size = 28,
                        color = text_color
                    ),
                    text = title
                ),
                showlegend=False,
                paper_bgcolor = background_color,
                plot_bgcolor = background_color,
            )

            # Update the xaxes layout of the subplots
            fig.update_xaxes(
                gridcolor = axis_grid_color,
                linecolor = axis_color,
                linewidth=2,
                tickfont = dict(
                    family = 'Open Sans, Arial, Helvetica, sans-serif',
                    size = 14,
                    color = text_color,
                ),
                title = dict(
                    font = dict(
                        family = 'Open Sans, Arial, Helvetica, sans-serif',
                        size = 14,
                        color = text_color,
                    ),
                    standoff = 4,
                    text = x_axis_title
                ),
            )

            # Update the yaxes layout of the subplots
            fig.update_yaxes(
                gridcolor = axis_grid_color,
                linecolor = axis_color,
                linewidth=2,
                tickfont = dict(
                    family = 'Open Sans, Arial, Helvetica, sans-serif',
                    size = 14,
                    color = text_color,
                ),
                title = dict(
                    font = dict(
                        family = 'Open Sans, Arial, Helvetica, sans-serif',
                        size = 14,
                        color = text_color,
                    ),
                    standoff = 4,
                    text = y_axis_title
                ),
            )

            # Adjust the titles of the subplots to match the data pulled during the loop
            fig.for_each_annotation(
                lambda a: a.update(
                    text = title_map[a.text], 
                    font = dict(
                        size=12,
                        color = text_color,
                    )
                )
            )
            file_name = f"{page_count}.pdf"
            fig.write_image(file_name)
            file_names.append(file_name)
            count = 1
            page_count+=1

    return file_names
