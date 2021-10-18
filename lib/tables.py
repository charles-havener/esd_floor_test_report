import plotly.graph_objects as go
from plotly.subplots import make_subplots


blue_100_color = 'rgba(0,101,159,1)'
white_font_color = 'rgba(255,255,255,1)'
rowA_bg_color = 'rgba(255,255,255,1)'
rowB_bg_color = 'rgba(242,242,242,1)'
rowA_font_color = 'rgba(0,0,0,1)'
rowB_font_color = 'rgba(0,0,0,1)'


text_color = 'rgba(119,119,119,1)'
background_color = 'rgba(255,255,255,1)'

def draw_summary_table(nums):
    nums.sort()
    low = nums[0]
    high = nums[4]
    avg = sum(nums)/5
    med = nums[2]

    fig = go.Table(
        header=dict(
            values=["Statistic", "Value"],
            fill_color=blue_100_color,
            align='center',
            font=dict(
                color=white_font_color,
                size=12,
            )
        ),
        cells=dict(
            align='center',
            values=[
                ["Min", "Max", "Average", "Median"],
                ["{:.2e}".format(v) for v in [low, high, avg, med]]
            ]
        )
    )

    return fig

def draw_data_table(nums, idx):

    table_start = idx//5 * 5 + 1
    table_end = table_start+4

    fig = go.Table(
        header=dict(
            values=["Location","Measurement"],
            fill_color=blue_100_color,
            align='center',
            font=dict(
                color=white_font_color,
                size=12,
            )
        ),
        cells=dict(
            align='center',
            values=[
                [i for i in range(table_start, table_end+1)],
                ["{:.2e}".format(n) for n in nums],
            ]
        )
    )

    return fig


def create_tables(base, power):
    if len(base) != len(power):
        print(f"Length mismatch: {len(base)} != {len(power)}")

    fig = make_subplots(
        rows = len(base)//5,
        cols = 2,
        horizontal_spacing=0.03,
        vertical_spacing=0.01,
        specs = [[{'type':'table'}, {'type':'table'}] for _ in range(len(base)//5)]
    )

    count = 0
    nums = [0 for _ in range(5)]
    for idx in range(len(base)):
        nums[idx%5] = base[idx] * 10**power[idx]
        count+=1

        if count == 5:
            r = idx//5+1

            fig.add_trace(
                draw_data_table(nums,idx),
                row=r, 
                col=1
            )

            fig.add_trace(
                draw_summary_table(nums),
                row=r,
                col=2
            )

            nums = [0 for _ in range(5)]
            count = 0
    
    fig.update_layout(
        height=990,
        width=765,
        title = dict(
            font = dict(
                family = 'Open Sans, Arial, Helvetica, sans-serif',
                size = 28,
                color = text_color
            ),
            text = "Floor Test Summary"
        ),
    )
    
    file_name = "tables.pdf"
    fig.write_image(file_name)
    return [file_name]