import tkinter as tk
from tkinter import filedialog
from tkinter.messagebox import showerror
from tkinter import ttk
from lib.graphs import generate_charts
from lib.tables import create_tables
from lib.cover import create_cover_page
import os

class App(ttk.Frame):
    def __init__(self, parent):
        ttk.Frame.__init__(self)

        self.parent = parent

        for index in range(2):
            self.columnconfigure(index=index, weight=1)

        # Create value lists
        self.location_count_list = [5, 10, 15, 20, 25, 30]
        self.freq_list = ["day", "week", "month", "quarter", "year"]

        self.folder_path = "N:\\"

        # Create widgets :)
        self.setup_widgets()

    def setup_widgets(self):

        # ==================================================
        # ================= Path + Buttons =================
        # ==================================================

        self.path_frame = ttk.LabelFrame(self, text="Path Information", padding=(20,10))
        self.path_frame.grid(
            row=0, column=0, padx=(20,10), pady=(20,5), sticky="nsew", columnspan=2
        )
        self.path_frame.columnconfigure(index=0, weight=0)
        self.path_frame.columnconfigure(index=1, weight=1)
        self.path_frame.columnconfigure(index=2, weight=0)

        #TODO: add click functionality
        self.path_update_button = ttk.Button(self.path_frame, text="Update Path", command=self.folder_select)
        self.path_update_button.grid(row=0, column=0, padx=(20,10), pady=(5,10), sticky="nsew")

        self.path_label = ttk.Label(
            self.path_frame, text=self.folder_path, justify="left", width=70
        )
        self.path_label.grid(row=0, column=1, padx=(10,0), sticky="nsew")

        # ----- Button -----
        #TODO: rename self.button1 to something better
        #TODO: buttons should be static size rather than matching full column
            #todo: maybe that means putting them in their own frame
        #TODO: make buttons work
        self.button1 = ttk.Button(
            self.path_frame, text="Create Report", padding=(20,10), command=self.create_report
        )
        self.button1.grid(
            row=0, column=3, padx=(20,10), pady=(5,10), sticky="nsew"
        )
        self.button2 = ttk.Button(
            self.path_frame, text="Light", padding=(20,10), command=self.change_theme
        )
        self.button2.grid(
            row=0, column=4, padx=(20,10), pady=(5,10), sticky="nsew"
        )


        # ===================================================
        # ================ Resistance Values ================
        # ===================================================

        self.res_frame = ttk.LabelFrame(self, text="Resistance Values")
        self.res_frame.grid(row=1, column=0, padx=(20,10), pady=(0,15), sticky="nsew")

        for index in range(3):
            self.res_frame.columnconfigure(index=index, weight=1)
        
        for index in range(3):
            self.res_frame.rowconfigure(index=index, weight=1)

        # General information about the test location
        self.location_subframe = ttk.Frame(self.res_frame)
        self.location_subframe.grid(row=0, column=0, columnspan=3)

        self.area_label = ttk.Label(self.location_subframe, text="Area:")
        self.area_label.grid(row=0, column=0, padx=(0,15), pady=(5,5), sticky="w")

        self.area = ttk.Entry(self.location_subframe)
        #self.area.insert(0, "Area")
        self.area.grid(row=0, column=1, columnspan=3, pady=(5,5))

        self.asset_label = ttk.Label(self.location_subframe, text="Asset #")
        self.asset_label.grid(row=1, column=0, padx=(0,15), sticky="w")

        self.asset = ttk.Entry(self.location_subframe)
        #self.asset.insert(0, "Asset #")
        self.asset.grid(row=1, column=1, columnspan=3, pady=(5,5))

        self.frequency_label = ttk.Label(self.location_subframe, text="Frequency:")
        self.frequency_label.grid(row=2, column=0, padx=(0,15), pady=(5,5))

        self.frequency = ttk.Entry(self.location_subframe, width=2)
        #self.frequency.insert(0, "4")
        self.frequency.grid(row=2, column=1, pady=(5,5))

        self.frequency_per_label = ttk.Label(self.location_subframe, text="per")
        self.frequency_per_label.grid(row=2, column=2, padx=(2,2), pady=(5,5))

        self.frequency_time = ttk.Combobox(
            self.location_subframe, state="readonly", values=self.freq_list, width=10
        )
        self.frequency_time.current(4)
        self.frequency_time.grid(row=2, column=3)

        # Separator
        self.separator = ttk.Separator(self.res_frame)
        self.separator.grid(row=1, column=0, padx=(10, 10), pady=10, columnspan=3, sticky="ew")

        # ----- 1-10 -----
        self.value_subframe1 = ttk.LabelFrame(self.res_frame, text="1-10")
        self.value_subframe1.grid(row=2, column=0, padx=(5,5))

        self.location1_label = ttk.Label(self.value_subframe1, text="1.")
        self.location1_label.grid(row=0, column=0, padx=(5,5), pady=(5,5))
        self.location1_base = ttk.Entry(self.value_subframe1, width=5)
        self.location1_base.grid(row=0, column=1, padx=(5,5), pady=(5,5))
        self.location1_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location1_x10.grid(row=0, column=2, pady=(5,5))
        self.location1_power = ttk.Entry(self.value_subframe1, width=2)
        self.location1_power.grid(row=0, column=3, padx=(5,5), pady=(5,5))

        self.location2_label = ttk.Label(self.value_subframe1, text="2.")
        self.location2_label.grid(row=1, column=0, padx=(5,5), pady=(5,5))
        self.location2_base = ttk.Entry(self.value_subframe1, width=5)
        self.location2_base.grid(row=1, column=1, padx=(5,5), pady=(5,5))
        self.location2_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location2_x10.grid(row=1, column=2, pady=(5,5))
        self.location2_power = ttk.Entry(self.value_subframe1, width=2)
        self.location2_power.grid(row=1, column=3, padx=(5,5), pady=(5,5))

        self.location3_label = ttk.Label(self.value_subframe1, text="3.")
        self.location3_label.grid(row=2, column=0, padx=(5,5), pady=(5,5))
        self.location3_base = ttk.Entry(self.value_subframe1, width=5)
        self.location3_base.grid(row=2, column=1, padx=(5,5), pady=(5,5))
        self.location3_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location3_x10.grid(row=2, column=2, pady=(5,5))
        self.location3_power = ttk.Entry(self.value_subframe1, width=2)
        self.location3_power.grid(row=2, column=3, padx=(5,5), pady=(5,5))

        self.location4_label = ttk.Label(self.value_subframe1, text="4.")
        self.location4_label.grid(row=3, column=0, padx=(5,5), pady=(5,5))
        self.location4_base = ttk.Entry(self.value_subframe1, width=5)
        self.location4_base.grid(row=3, column=1, padx=(5,5), pady=(5,5))
        self.location4_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location4_x10.grid(row=3, column=2, pady=(5,5))
        self.location4_power = ttk.Entry(self.value_subframe1, width=2)
        self.location4_power.grid(row=3, column=3, padx=(5,5), pady=(5,5))

        self.location5_label = ttk.Label(self.value_subframe1, text="5.")
        self.location5_label.grid(row=4, column=0, padx=(5,5), pady=(5,5))
        self.location5_base = ttk.Entry(self.value_subframe1, width=5)
        self.location5_base.grid(row=4, column=1, padx=(5,5), pady=(5,5))
        self.location5_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location5_x10.grid(row=4, column=2, pady=(5,5))
        self.location5_power = ttk.Entry(self.value_subframe1, width=2)
        self.location5_power.grid(row=4, column=3, padx=(5,5), pady=(5,5))

        self.location6_label = ttk.Label(self.value_subframe1, text="6.")
        self.location6_label.grid(row=5, column=0, padx=(5,5), pady=(5,5))
        self.location6_base = ttk.Entry(self.value_subframe1, width=5)
        self.location6_base.grid(row=5, column=1, padx=(5,5), pady=(5,5))
        self.location6_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location6_x10.grid(row=5, column=2, pady=(5,5))
        self.location6_power = ttk.Entry(self.value_subframe1, width=2)
        self.location6_power.grid(row=5, column=3, padx=(5,5), pady=(5,5))

        self.location7_label = ttk.Label(self.value_subframe1, text="7.")
        self.location7_label.grid(row=6, column=0, padx=(5,5), pady=(5,5))
        self.location7_base = ttk.Entry(self.value_subframe1, width=5)
        self.location7_base.grid(row=6, column=1, padx=(5,5), pady=(5,5))
        self.location7_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location7_x10.grid(row=6, column=2, pady=(5,5))
        self.location7_power = ttk.Entry(self.value_subframe1, width=2)
        self.location7_power.grid(row=6, column=3, padx=(5,5), pady=(5,5))

        self.location8_label = ttk.Label(self.value_subframe1, text="8.")
        self.location8_label.grid(row=7, column=0, padx=(5,5), pady=(5,5))
        self.location8_base = ttk.Entry(self.value_subframe1, width=5)
        self.location8_base.grid(row=7, column=1, padx=(5,5), pady=(5,5))
        self.location8_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location8_x10.grid(row=7, column=2, pady=(5,5))
        self.location8_power = ttk.Entry(self.value_subframe1, width=2)
        self.location8_power.grid(row=7, column=3, padx=(5,5), pady=(5,5))

        self.location9_label = ttk.Label(self.value_subframe1, text="9.")
        self.location9_label.grid(row=8, column=0, padx=(5,5), pady=(5,5))
        self.location9_base = ttk.Entry(self.value_subframe1, width=5)
        self.location9_base.grid(row=8, column=1, padx=(5,5), pady=(5,5))
        self.location9_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location9_x10.grid(row=8, column=2, pady=(5,5))
        self.location9_power = ttk.Entry(self.value_subframe1, width=2)
        self.location9_power.grid(row=8, column=3, padx=(5,5), pady=(5,5))

        self.location10_label = ttk.Label(self.value_subframe1, text="10.")
        self.location10_label.grid(row=9, column=0, padx=(5,5), pady=(5,5))
        self.location10_base = ttk.Entry(self.value_subframe1, width=5)
        self.location10_base.grid(row=9, column=1, padx=(5,5), pady=(5,5))
        self.location10_x10 = ttk.Label(self.value_subframe1, text="*10^")
        self.location10_x10.grid(row=9, column=2, pady=(5,5))
        self.location10_power = ttk.Entry(self.value_subframe1, width=2)
        self.location10_power.grid(row=9, column=3, padx=(5,5), pady=(5,5))

        # ----- 11-20 -----
        self.value_subframe2 = ttk.LabelFrame(self.res_frame, text="11-20")
        self.value_subframe2.grid(row=2, column=1, padx=(5,5))

        self.location11_label = ttk.Label(self.value_subframe2, text="11.")
        self.location11_label.grid(row=0, column=0, padx=(5,5), pady=(5,5))
        self.location11_base = ttk.Entry(self.value_subframe2, width=5)
        self.location11_base.grid(row=0, column=1, padx=(5,5), pady=(5,5))
        self.location11_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location11_x10.grid(row=0, column=2, pady=(5,5))
        self.location11_power = ttk.Entry(self.value_subframe2, width=2)
        self.location11_power.grid(row=0, column=3, padx=(5,5), pady=(5,5))

        self.location12_label = ttk.Label(self.value_subframe2, text="12.")
        self.location12_label.grid(row=1, column=0, padx=(5,5), pady=(5,5))
        self.location12_base = ttk.Entry(self.value_subframe2, width=5)
        self.location12_base.grid(row=1, column=1, padx=(5,5), pady=(5,5))
        self.location12_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location12_x10.grid(row=1, column=2, pady=(5,5))
        self.location12_power = ttk.Entry(self.value_subframe2, width=2)
        self.location12_power.grid(row=1, column=3, padx=(5,5), pady=(5,5))

        self.location13_label = ttk.Label(self.value_subframe2, text="13.")
        self.location13_label.grid(row=2, column=0, padx=(5,5), pady=(5,5))
        self.location13_base = ttk.Entry(self.value_subframe2, width=5)
        self.location13_base.grid(row=2, column=1, padx=(5,5), pady=(5,5))
        self.location13_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location13_x10.grid(row=2, column=2, pady=(5,5))
        self.location13_power = ttk.Entry(self.value_subframe2, width=2)
        self.location13_power.grid(row=2, column=3, padx=(5,5), pady=(5,5))

        self.location14_label = ttk.Label(self.value_subframe2, text="14.")
        self.location14_label.grid(row=3, column=0, padx=(5,5), pady=(5,5))
        self.location14_base = ttk.Entry(self.value_subframe2, width=5)
        self.location14_base.grid(row=3, column=1, padx=(5,5), pady=(5,5))
        self.location14_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location14_x10.grid(row=3, column=2, pady=(5,5))
        self.location14_power = ttk.Entry(self.value_subframe2, width=2)
        self.location14_power.grid(row=3, column=3, padx=(5,5), pady=(5,5))

        self.location15_label = ttk.Label(self.value_subframe2, text="15.")
        self.location15_label.grid(row=4, column=0, padx=(5,5), pady=(5,5))
        self.location15_base = ttk.Entry(self.value_subframe2, width=5)
        self.location15_base.grid(row=4, column=1, padx=(5,5), pady=(5,5))
        self.location15_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location15_x10.grid(row=4, column=2, pady=(5,5))
        self.location15_power = ttk.Entry(self.value_subframe2, width=2)
        self.location15_power.grid(row=4, column=3, padx=(5,5), pady=(5,5))

        self.location16_label = ttk.Label(self.value_subframe2, text="16.")
        self.location16_label.grid(row=5, column=0, padx=(5,5), pady=(5,5))
        self.location16_base = ttk.Entry(self.value_subframe2, width=5)
        self.location16_base.grid(row=5, column=1, padx=(5,5), pady=(5,5))
        self.location16_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location16_x10.grid(row=5, column=2, pady=(5,5))
        self.location16_power = ttk.Entry(self.value_subframe2, width=2)
        self.location16_power.grid(row=5, column=3, padx=(5,5), pady=(5,5))

        self.location17_label = ttk.Label(self.value_subframe2, text="17.")
        self.location17_label.grid(row=6, column=0, padx=(5,5), pady=(5,5))
        self.location17_base = ttk.Entry(self.value_subframe2, width=5)
        self.location17_base.grid(row=6, column=1, padx=(5,5), pady=(5,5))
        self.location17_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location17_x10.grid(row=6, column=2, pady=(5,5))
        self.location17_power = ttk.Entry(self.value_subframe2, width=2)
        self.location17_power.grid(row=6, column=3, padx=(5,5), pady=(5,5))

        self.location18_label = ttk.Label(self.value_subframe2, text="18.")
        self.location18_label.grid(row=7, column=0, padx=(5,5), pady=(5,5))
        self.location18_base = ttk.Entry(self.value_subframe2, width=5)
        self.location18_base.grid(row=7, column=1, padx=(5,5), pady=(5,5))
        self.location18_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location18_x10.grid(row=7, column=2, pady=(5,5))
        self.location18_power = ttk.Entry(self.value_subframe2, width=2)
        self.location18_power.grid(row=7, column=3, padx=(5,5), pady=(5,5))

        self.location19_label = ttk.Label(self.value_subframe2, text="19.")
        self.location19_label.grid(row=8, column=0, padx=(5,5), pady=(5,5))
        self.location19_base = ttk.Entry(self.value_subframe2, width=5)
        self.location19_base.grid(row=8, column=1, padx=(5,5), pady=(5,5))
        self.location19_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location19_x10.grid(row=8, column=2, pady=(5,5))
        self.location19_power = ttk.Entry(self.value_subframe2, width=2)
        self.location19_power.grid(row=8, column=3, padx=(5,5), pady=(5,5))

        self.location20_label = ttk.Label(self.value_subframe2, text="20.")
        self.location20_label.grid(row=9, column=0, padx=(5,5), pady=(5,5))
        self.location20_base = ttk.Entry(self.value_subframe2, width=5)
        self.location20_base.grid(row=9, column=1, padx=(5,5), pady=(5,5))
        self.location20_x10 = ttk.Label(self.value_subframe2, text="*10^")
        self.location20_x10.grid(row=9, column=2, pady=(5,5))
        self.location20_power = ttk.Entry(self.value_subframe2, width=2)
        self.location20_power.grid(row=9, column=3, padx=(5,5), pady=(5,5))


        # ----- 21-30 -----
        self.value_subframe3 = ttk.LabelFrame(self.res_frame, text="21-30")
        self.value_subframe3.grid(row=2, column=2, padx=(5,5))

        self.location21_label = ttk.Label(self.value_subframe3, text="21.")
        self.location21_label.grid(row=0, column=0, padx=(5,5), pady=(5,5))
        self.location21_base = ttk.Entry(self.value_subframe3, width=5)
        self.location21_base.grid(row=0, column=1, padx=(5,5), pady=(5,5))
        self.location21_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location21_x10.grid(row=0, column=2, pady=(5,5))
        self.location21_power = ttk.Entry(self.value_subframe3, width=2)
        self.location21_power.grid(row=0, column=3, padx=(5,5), pady=(5,5))

        self.location22_label = ttk.Label(self.value_subframe3, text="22.")
        self.location22_label.grid(row=1, column=0, padx=(5,5), pady=(5,5))
        self.location22_base = ttk.Entry(self.value_subframe3, width=5)
        self.location22_base.grid(row=1, column=1, padx=(5,5), pady=(5,5))
        self.location22_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location22_x10.grid(row=1, column=2, pady=(5,5))
        self.location22_power = ttk.Entry(self.value_subframe3, width=2)
        self.location22_power.grid(row=1, column=3, padx=(5,5), pady=(5,5))

        self.location23_label = ttk.Label(self.value_subframe3, text="23.")
        self.location23_label.grid(row=2, column=0, padx=(5,5), pady=(5,5))
        self.location23_base = ttk.Entry(self.value_subframe3, width=5)
        self.location23_base.grid(row=2, column=1, padx=(5,5), pady=(5,5))
        self.location23_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location23_x10.grid(row=2, column=2, pady=(5,5))
        self.location23_power = ttk.Entry(self.value_subframe3, width=2)
        self.location23_power.grid(row=2, column=3, padx=(5,5), pady=(5,5))

        self.location24_label = ttk.Label(self.value_subframe3, text="24.")
        self.location24_label.grid(row=3, column=0, padx=(5,5), pady=(5,5))
        self.location24_base = ttk.Entry(self.value_subframe3, width=5)
        self.location24_base.grid(row=3, column=1, padx=(5,5), pady=(5,5))
        self.location24_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location24_x10.grid(row=3, column=2, pady=(5,5))
        self.location24_power = ttk.Entry(self.value_subframe3, width=2)
        self.location24_power.grid(row=3, column=3, padx=(5,5), pady=(5,5))

        self.location25_label = ttk.Label(self.value_subframe3, text="25.")
        self.location25_label.grid(row=4, column=0, padx=(5,5), pady=(5,5))
        self.location25_base = ttk.Entry(self.value_subframe3, width=5)
        self.location25_base.grid(row=4, column=1, padx=(5,5), pady=(5,5))
        self.location25_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location25_x10.grid(row=4, column=2, pady=(5,5))
        self.location25_power = ttk.Entry(self.value_subframe3, width=2)
        self.location25_power.grid(row=4, column=3, padx=(5,5), pady=(5,5))

        self.location26_label = ttk.Label(self.value_subframe3, text="26.")
        self.location26_label.grid(row=5, column=0, padx=(5,5), pady=(5,5))
        self.location26_base = ttk.Entry(self.value_subframe3, width=5)
        self.location26_base.grid(row=5, column=1, padx=(5,5), pady=(5,5))
        self.location26_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location26_x10.grid(row=5, column=2, pady=(5,5))
        self.location26_power = ttk.Entry(self.value_subframe3, width=2)
        self.location26_power.grid(row=5, column=3, padx=(5,5), pady=(5,5))

        self.location27_label = ttk.Label(self.value_subframe3, text="27.")
        self.location27_label.grid(row=6, column=0, padx=(5,5), pady=(5,5))
        self.location27_base = ttk.Entry(self.value_subframe3, width=5)
        self.location27_base.grid(row=6, column=1, padx=(5,5), pady=(5,5))
        self.location27_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location27_x10.grid(row=6, column=2, pady=(5,5))
        self.location27_power = ttk.Entry(self.value_subframe3, width=2)
        self.location27_power.grid(row=6, column=3, padx=(5,5), pady=(5,5))

        self.location28_label = ttk.Label(self.value_subframe3, text="28.")
        self.location28_label.grid(row=7, column=0, padx=(5,5), pady=(5,5))
        self.location28_base = ttk.Entry(self.value_subframe3, width=5)
        self.location28_base.grid(row=7, column=1, padx=(5,5), pady=(5,5))
        self.location28_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location28_x10.grid(row=7, column=2, pady=(5,5))
        self.location28_power = ttk.Entry(self.value_subframe3, width=2)
        self.location28_power.grid(row=7, column=3, padx=(5,5), pady=(5,5))

        self.location29_label = ttk.Label(self.value_subframe3, text="29.")
        self.location29_label.grid(row=8, column=0, padx=(5,5), pady=(5,5))
        self.location29_base = ttk.Entry(self.value_subframe3, width=5)
        self.location29_base.grid(row=8, column=1, padx=(5,5), pady=(5,5))
        self.location29_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location29_x10.grid(row=8, column=2, pady=(5,5))
        self.location29_power = ttk.Entry(self.value_subframe3, width=2)
        self.location29_power.grid(row=8, column=3, padx=(5,5), pady=(5,5))

        self.location30_label = ttk.Label(self.value_subframe3, text="30.")
        self.location30_label.grid(row=9, column=0, padx=(5,5), pady=(5,5))
        self.location30_base = ttk.Entry(self.value_subframe3, width=5)
        self.location30_base.grid(row=9, column=1, padx=(5,5), pady=(5,5))
        self.location30_x10 = ttk.Label(self.value_subframe3, text="*10^")
        self.location30_x10.grid(row=9, column=2, pady=(5,5))
        self.location30_power = ttk.Entry(self.value_subframe3, width=2)
        self.location30_power.grid(row=9, column=3, padx=(5,5), pady=(5,5))


        # ===================================================
        # ================ Cover Page Values ================
        # ===================================================

        self.input_frame = ttk.LabelFrame(self, text="Cover Page Info", padding=(20,10))
        self.input_frame.grid(
            row=1, column=1, padx=(0,10), pady=(0,15), sticky="nsew"
        )

        # ----- Description of Floor -----
        self.floor_subframe = ttk.LabelFrame(self.input_frame, text="Description of Floor")
        self.floor_subframe.grid(row=0, column=0, padx=(5,5), pady=(0,5))

        self.e1_1_label = ttk.Label(self.floor_subframe, text="Manufacturer & Product ID:", width=30)
        self.e1_1_label.grid(row=0, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e1_1 = ttk.Entry(self.floor_subframe, width=40)
        self.e1_1.insert(0, "Staticworx 1-1912")
        self.e1_1.grid(row=0, column=1, pady=(0,5))

        self.e1_2_label = ttk.Label(self.floor_subframe, text="Type (hard or soft tile, carpet, etc):", width=30)
        self.e1_2_label.grid(row=1, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e1_2 = ttk.Entry(self.floor_subframe, width=40)
        self.e1_2.insert(0, "Soft Tile")
        self.e1_2.grid(row=1, column=1, pady=(0,5))

        self.e1_3_label = ttk.Label(self.floor_subframe, text="Material (vinal, rubber, etc):", width=30)
        self.e1_3_label.grid(row=2, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e1_3 = ttk.Entry(self.floor_subframe, width=40)
        self.e1_3.insert(0, "Vinyl")
        self.e1_3.grid(row=2, column=1, pady=(0,5))

        self.e1_4_label = ttk.Label(self.floor_subframe, text="Substrate Material:", width=30)
        self.e1_4_label.grid(row=3, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e1_4 = ttk.Entry(self.floor_subframe, width=40)
        self.e1_4.insert(0, "Excelsior AW-510 Acrylic Wet Set Adhesive")
        self.e1_4.grid(row=3, column=1, pady=(0,5))


        # ----- Description of Footwear -----
        self.footwear_subframe = ttk.LabelFrame(self.input_frame, text="Description of Footwear")
        self.footwear_subframe.grid(row=1, column=0, padx=(5,5), pady=(5,5))

        self.e2_1_label = ttk.Label(self.footwear_subframe, text="Manufacturer:", width=30)
        self.e2_1_label.grid(row=0, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e2_1 = ttk.Entry(self.footwear_subframe, width=40)
        self.e2_1.insert(0, "Botron B7530")
        self.e2_1.grid(row=0, column=1, pady=(0,5))

        self.e2_2_label = ttk.Label(self.footwear_subframe, text="Type of Footwear:", width=30)
        self.e2_2_label.grid(row=1, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e2_2 = ttk.Entry(self.footwear_subframe, width=40)
        self.e2_2.insert(0, "ESD Safe Heel Grounder")
        self.e2_2.grid(row=1, column=1, pady=(0,5))

        self.e2_3_label = ttk.Label(self.footwear_subframe, text="Composition of Soles:", width=30)
        self.e2_3_label.grid(row=2, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e2_3 = ttk.Entry(self.footwear_subframe, width=40)
        self.e2_3.insert(0, "Conductive Nylon w/ 1Meg Ohm Resistor")
        self.e2_3.grid(row=2, column=1, pady=(0,5))

        self.e2_4_label = ttk.Label(self.footwear_subframe, text="Composition of Socks:", width=30)
        self.e2_4_label.grid(row=3, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e2_4 = ttk.Entry(self.footwear_subframe, width=40)
        self.e2_4.insert(0, "Cotton")
        self.e2_4.grid(row=3, column=1, pady=(0,5))


        # ----- Description of Test Hardware -----
        self.test_hardware_subframe = ttk.LabelFrame(self.input_frame, text="Description of Test Hardware:")
        self.test_hardware_subframe.grid(row=2, column=0, padx=(5,5), pady=(5,5))

        self.e3_1_label = ttk.Label(self.test_hardware_subframe, text="Charge Plate Monitor", width=30)
        self.e3_1_label.grid(row=0, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_1 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_1.insert(0, "DESCO 19290 - Asset # PA0497")
        self.e3_1.grid(row=0, column=1, pady=(0,5))

        self.e3_2_label = ttk.Label(self.test_hardware_subframe, text="Manufacturer", width=30)
        self.e3_2_label.grid(row=1, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_2 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_2.insert(0, "DESCO")
        self.e3_2.grid(row=1, column=1, pady=(0,5))

        self.e3_3_label = ttk.Label(self.test_hardware_subframe, text="Description:", width=30)
        self.e3_3_label.grid(row=2, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_3 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_3.insert(0, "Digital Surface Resistance Meter")
        self.e3_3.grid(row=2, column=1, pady=(0,5))

        self.e3_4_label = ttk.Label(self.test_hardware_subframe, text="Graphical Recording Device:", width=30)
        self.e3_4_label.grid(row=3, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_4 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_4.insert(0, "DESCO 19431 - Asset # PA0881")
        self.e3_4.grid(row=3, column=1, pady=(0,5))

        self.e3_5_label = ttk.Label(self.test_hardware_subframe, text="Manufacturer:", width=30)
        self.e3_5_label.grid(row=4, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_5 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_5.insert(0, "DESCO")
        self.e3_5.grid(row=4, column=1, pady=(0,5))

        self.e3_6_label = ttk.Label(self.test_hardware_subframe, text="Description:", width=30)
        self.e3_6_label.grid(row=5, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e3_6 = ttk.Entry(self.test_hardware_subframe, width=40)
        self.e3_6.insert(0, "Body Voltage Meter")
        self.e3_6.grid(row=5, column=1, pady=(0,5))


        # ----- Other Information -----
        self.tester_subframe = ttk.LabelFrame(self.input_frame, text="Other Information")
        self.tester_subframe.grid(row=3, column=0, padx=(5,5), pady=(5,5))

        self.e4_1_label = ttk.Label(self.tester_subframe, text="Test Perfomed By", width=30)
        self.e4_1_label.grid(row=0, column=0, padx=(5,15), pady=(0,5), sticky="w")
        self.e4_1 = ttk.Entry(self.tester_subframe, width=40)
        self.e4_1.insert(0, "")
        self.e4_1.grid(row=0, column=1, pady=(0,5))

    def change_theme(self):
        if self.parent.tk.call("ttk::style", "theme", "use") == "sun-valley-dark":
            self.parent.tk.call("set_theme", "light")
            self.button2['text'] = "Dark"
        else:
            self.parent.tk.call("set_theme", "dark")
            self.button2['text'] = "Light"

    def folder_select(self):
        self.path_label['text'] = filedialog.askdirectory()

    def validate_inputs(self):
        self.text_inputs_to_validate = [
            self.area, self.asset, self.frequency, self.frequency_time,
            self.e1_1, self.e1_2, self.e1_3, self.e1_4,
            self.e2_1, self.e2_2, self.e2_3, self.e2_4,
            self.e3_1, self.e3_2, self.e3_3, self.e3_4, self.e3_5, self.e3_6,
            self.e4_1
        ]

        self.float_inputs_to_validate = [
            self.location1_base, self.location2_base, self.location3_base,
            self.location4_base, self.location5_base, self.location6_base,
            self.location7_base, self.location8_base, self.location9_base,
            self.location10_base, self.location11_base, self.location12_base,
            self.location13_base, self.location14_base, self.location15_base,
            self.location16_base, self.location17_base, self.location18_base,
            self.location19_base, self.location20_base, self.location21_base,
            self.location22_base, self.location23_base, self.location24_base,
            self.location25_base, self.location26_base, self.location27_base,
            self.location28_base, self.location29_base, self.location30_base,
        ]

        self.int_inputs_to_validate = [
            self.location1_power, self.location2_power, self.location3_power,
            self.location4_power, self.location5_power, self.location6_power,
            self.location7_power, self.location8_power, self.location9_power,
            self.location10_power, self.location11_power, self.location12_power,
            self.location13_power, self.location14_power, self.location15_power,
            self.location16_power, self.location17_power, self.location18_power,
            self.location19_power, self.location20_power, self.location21_power,
            self.location22_power, self.location23_power, self.location24_power,
            self.location25_power, self.location26_power, self.location27_power,
            self.location28_power, self.location29_power, self.location30_power,
        ]
        
        min_index = 0
        min_index1, min_index2 = 0, 0
        
        # must be able to cast certain inputs to int
        for idx, item_i in enumerate(self.int_inputs_to_validate):
            try:
                int(item_i.get())
                if idx==len(self.int_inputs_to_validate)-1: min_index1 = len(self.int_inputs_to_validate)
            except ValueError:
                min_index1 = idx
                break
        print(min_index1)

        # must be able to cast certain inputs to float
        for idx, item_f in enumerate(self.float_inputs_to_validate):
            try:
                float(item_f.get())
                if idx==len(self.float_inputs_to_validate)-1: min_index2 = len(self.float_inputs_to_validate)
            except ValueError:
                min_index2 = idx
                break
        print(min_index2)
        
        if min_index1 != min_index2:
            l = item_i if min_index1 < min_index2 else item_f
            h = item_i if min_index1 > min_index2 else item_f
            l.state(["invalid"])
            min_index = 0
        elif (min_index1 == 0 and min_index2 == 0) or min_index1%5 != 0:
            item_i.state(["invalid"])
            item_f.state(["invalid"])
        else:
            min_index = min_index1

        # string inputs cannot be empty
        for item in self.text_inputs_to_validate:
            if item.get() == "":
                item.state(["invalid"])
                min_index = 0
        
        return min_index

    def check_files_exist(self, number):
        for n in range(1, number+1):
            file_name = self.path_label['text'] + "\\" +  str(n) + ".xls"
            if not os.path.exists(file_name):
                print(f"file {n} not found in path provided")
                return False
        return True

    def create_report(self):

        # Retrieves the number of recordings entered and proceeds if no errors were found
        m = self.validate_inputs()
        if m == 0: return

        # Ensures that an excel file is associated with each entered data point
        if not self.check_files_exist(m):
            showerror("Error", "Could not find all excel files in the path provided")
            return

        labels = [
            [self.e1_1_label, self.e1_2_label, self.e1_3_label, self.e1_4_label],
            [self.e2_1_label, self.e2_2_label, self.e2_3_label, self.e2_4_label],
            [self.e3_1_label, self.e3_2_label, self.e3_3_label, self.e3_4_label, self.e3_5_label, self.e3_6_label],
            [self.e4_1_label]
        ]

        values = [
            [self.e1_1, self.e1_2, self.e1_3, self.e1_4],
            [self.e2_1, self.e2_2, self.e2_3, self.e2_4],
            [self.e3_1, self.e3_2, self.e3_3, self.e3_4,self.e3_5, self.e3_6],
            [self.e4_1]
        ]

        titles = [
            "Description of Floor",
            "Description of Footwear",
            "Description of Test Hardware",
            "Other Information"
        ]

        area = self.area.get()
        asset = self.asset.get()
        freq = str(self.frequency.get()) + "/" + str(self.frequency_time.get())

        labels = [[i['text'] for i in labels[j]] for j in range(len(labels))]
        values = [[i.get() for i in values[j]] for j in range(len(values))]

        b = [float(i.get()) for i in self.float_inputs_to_validate[:m]]
        p = [int(i.get()) for i in self.int_inputs_to_validate[:m]]

        create_cover_page(labels, values, titles, area, asset, freq)
        create_tables(b,p)
        generate_charts(self.path_label['text'].replace("/", "\\"), m)