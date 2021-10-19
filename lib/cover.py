import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import date

blue_100_color = 'rgba(0,101,159,1)'
blue_50_color = 'rgba(0,101,159,0.5)'
white_font_color = 'rgba(255,255,255,1)'
rowA_bg_color = 'rgba(255,255,255,1)'
rowB_bg_color = 'rgba(242,242,242,1)'
rowA_font_color = 'rgba(0,0,0,1)'
rowB_font_color = 'rgba(0,0,0,1)'

text_color = 'rgba(119,119,119,1)'
background_color = 'rgba(255,255,255,1)'

#TODO: need to get area/asset/freq added to title of report

def create_header_table(area, asset, freq):
    fig = go.Table(
        header=dict(
            values=["ANSI/ESD STM97.2-2016 Annex B1"],
            fill_color=white_font_color,
            align='center',
            font=dict(
                color=text_color,
                size=12,
            ),
        ),
        cells=dict(
            values=[
                ["ANSI/ESD STM97.1-2015 Annex B",
                area,
                asset,
                freq,
                date.today().strftime("%B %d, %Y")],
            ],
            fill_color=white_font_color,
            align='center',
            font=dict(
                color=text_color,
                size=12,
            ),
        )
    )

    return fig

def create_sub_table(labels, values, title):

    fig = go.Table(
        header=dict(
            values=[title,""],
            fill_color=blue_100_color,
            align='left',
            font=dict(
                color=white_font_color,
                size=12,
            ),
            line = dict(
                color = blue_100_color
            )
        ),
        cells=dict(
            align='left',
            values=[
                [i for i in labels],
                [j for j in values],
            ],
        )
    )

    return fig


def create_cover_page(labels, values, titles, area, asset, freq):

    fig = make_subplots(
        rows = len(labels)+1,
        cols = 1,
        horizontal_spacing=0.03,
        vertical_spacing=0.01,
        specs = [[{'type':'table'}] for _ in range(len(labels)+1)]
    )

    fig.add_trace(
        create_header_table(area, asset, freq),
        row=1,
        col=1
    )

    for i in range(len(labels)):

        fig.add_trace(
            create_sub_table(labels[i], values[i], titles[i]),
            row = i+2,
            col = 1
        )

    fig.update_layout(
        height=990,
        width=765,
        title = dict(
            font = dict(
                family = 'Open Sans, Arial, Helvetica, sans-serif',
                size = 28,
                color = text_color
            ),
            text = "Footwear/Flooring System Test Record"
        ),
    )

    file_name = "tmp/cover.pdf"
    fig.write_image(file_name)
    return [file_name]