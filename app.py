import os, sys
from pathlib import Path

import customtkinter
import customtkinter as ctk
from tkcalendar import Calendar
import pandas as pd
from PIL import Image

from service_card import create_service_card


if sys.platform == "darwin":
    data_dir = Path.home() / "Personal" / "PythonProject"/ 'data'
    # data_dir = Path.cwd() / 'data'
elif sys.platform == "win32":
    data_dir = Path(os.environ["APPDATA"]) / "MyApp"
else:
    data_dir = Path.cwd() / 'data'

if sys.platform == "darwin":
    assets_dir = Path.home() / "Personal" / "PythonProject"/ 'assets'
elif sys.platform == "win32":
    assets_dir = Path(os.environ["APPDATA"]) / "MyApp"
else:
    assets_dir = Path.cwd() / 'data'

data_dir.mkdir(parents=True, exist_ok=True)
assets_dir.mkdir(parents=True, exist_ok=True)

services_path = data_dir / "service.xlsx"
inventory_path = data_dir / "produkty.xlsx"

san_sherif = assets_dir / "DejaVuSans.ttf"
img_path = assets_dir / "Qmobile_transparent.png"
icon_dir = assets_dir / "ico.ico"

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class DateEntry(ctk.CTkToplevel):
    def __init__(self, parent, entry_widget):
        super().__init__(parent)
        self.entry = entry_widget

        self.overrideredirect(True)  # remove title bar
        self.attributes("-topmost", True)

        # Position BELOW the entry
        x = entry_widget.winfo_rootx()
        y = entry_widget.winfo_rooty() + entry_widget.winfo_height()
        self.geometry(f"+{x}+{y}")

        self.calendar = Calendar(
            self,
            selectmode="day",
            date_pattern="yyyy-mm-dd",

            # 🎨 colors
            background="#1f6aa5",
            foreground="white",
            headersbackground="#144870",
            headersforeground="white",
            selectbackground="#2fa4ff",
            selectforeground="white",
            normalbackground="#2b2b2b",
            normalforeground="white",
            weekendbackground="#2b2b2b",
            weekendforeground="#ff6b6b",
            othermonthforeground="#777777"
        )

        self.calendar.pack(padx=5, pady=5)

        self.calendar.bind("<<CalendarSelected>>", self.set_date)

    def set_date(self, event):
        self.entry.delete(0, "end")
        self.entry.insert(0, self.calendar.get_date())
        self.destroy()

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        # configure window
        self.title("Qmobile test program")
        self.geometry(f"{1450}x{850}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure((0, 1, 2,3,4), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = MenuFrame(self, self.show_frame,width=140,corner_radius=0 )
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.frames = {}

        for Page in (InventoryFrame, ServicesFrame):
            page = Page(self)
            page.grid(row=0, column=1, padx=(20, 10), pady=(20, 10), sticky="nsew")
            self.frames[Page.__name__] = page

        self.show_frame("InventoryFrame")

    def show_frame(self, page_class):
        self.frames[page_class].tkraise()

        # create main entry and button
            # self.entry = customtkinter.CTkEntry(self, placeholder_text="CTkEntry")
            # self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
            #
            # self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
            # self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
            #
            # # create textbox
            # self.textbox = customtkinter.CTkTextbox(self, width=250)
            # self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
            #
            # # create tabview
            # self.tabview = customtkinter.CTkTabview(self, width=250)
            # self.tabview.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
            # self.tabview.add("CTkTabview")
            # self.tabview.add("Tab 2")
            # self.tabview.add("Tab 3")
            # self.tabview.tab("CTkTabview").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
            # self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)
            #
            # self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("CTkTabview"), dynamic_resizing=False,
            #                                                 values=["Value 1", "Value 2", "Value Long Long Long"])
            # self.optionmenu_1.grid(row=0, column=0, padx=20, pady=(20, 10))
            # self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("CTkTabview"),
            #                                             values=["Value 1", "Value 2", "Value Long....."])
            # self.combobox_1.grid(row=1, column=0, padx=20, pady=(10, 10))
            # self.string_input_button = customtkinter.CTkButton(self.tabview.tab("CTkTabview"), text="Open CTkInputDialog",
            #                                                    command=self.open_input_dialog_event)
            # self.string_input_button.grid(row=2, column=0, padx=20, pady=(10, 10))
            # self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("Tab 2"), text="CTkLabel on Tab 2")
            # self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)
            #
            # # create radiobutton frame
            # self.radiobutton_frame = customtkinter.CTkFrame(self)
            # self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
            # self.radio_var = tkinter.IntVar(value=0)
            # self.label_radio_group = customtkinter.CTkLabel(master=self.radiobutton_frame, text="CTkRadioButton Group:")
            # self.label_radio_group.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky="")
            # self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=0)
            # self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")
            # self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=1)
            # self.radio_button_2.grid(row=2, column=2, pady=10, padx=20, sticky="n")
            # self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=2)
            # self.radio_button_3.grid(row=3, column=2, pady=10, padx=20, sticky="n")
            #
            # # create slider and progressbar frame
            # self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
            # self.slider_progressbar_frame.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
            # self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
            # self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
            # self.seg_button_1 = customtkinter.CTkSegmentedButton(self.slider_progressbar_frame)
            # self.seg_button_1.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
            # self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
            # self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
            # self.progressbar_2 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
            # self.progressbar_2.grid(row=2, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
            # self.slider_1 = customtkinter.CTkSlider(self.slider_progressbar_frame, from_=0, to=1, number_of_steps=4)
            # self.slider_1.grid(row=3, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
            # self.slider_2 = customtkinter.CTkSlider(self.slider_progressbar_frame, orientation="vertical")
            # self.slider_2.grid(row=0, column=1, rowspan=5, padx=(10, 10), pady=(10, 10), sticky="ns")
            # self.progressbar_3 = customtkinter.CTkProgressBar(self.slider_progressbar_frame, orientation="vertical")
            # self.progressbar_3.grid(row=0, column=2, rowspan=5, padx=(10, 20), pady=(10, 10), sticky="ns")
            #
            # # create scrollable frame
            # self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text="CTkScrollableFrame")
            # self.scrollable_frame.grid(row=1, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
            # self.scrollable_frame.grid_columnconfigure(0, weight=1)
            # self.scrollable_frame_switches = []
            # for i in range(100):
            #     switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
            #     switch.grid(row=i, column=0, padx=10, pady=(0, 20))
            #     self.scrollable_frame_switches.append(switch)
            #
            # # create checkbox and switch frame
            # self.checkbox_slider_frame = customtkinter.CTkFrame(self)
            # self.checkbox_slider_frame.grid(row=1, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
            # self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
            # self.checkbox_1.grid(row=1, column=0, pady=(20, 0), padx=20, sticky="n")
            # self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
            # self.checkbox_2.grid(row=2, column=0, pady=(20, 0), padx=20, sticky="n")
            # self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
            # self.checkbox_3.grid(row=3, column=0, pady=20, padx=20, sticky="n")

            # set default values
            # self.sidebar_button_3.configure(state="disabled", text="Disabled CTkButton")
            # self.checkbox_3.configure(state="disabled")
            # self.checkbox_1.select()
            # self.scrollable_frame_switches[0].select()
            # self.scrollable_frame_switches[4].select()
            # self.radio_button_3.configure(state="disabled")
            # self.appearance_mode_optionemenu.set("Dark")
            # self.scaling_optionemenu.set("100%")
            # self.optionmenu_1.set("CTkOptionmenu")
            # self.combobox_1.set("CTkComboBox")
            # self.slider_1.configure(command=self.progressbar_2.set)
            # self.slider_2.configure(command=self.progressbar_3.set)
            # self.progressbar_1.configure(mode="indeterminnate")
            # self.progressbar_1.start()
            # self.textbox.insert("0.0", "CTkTextbox\n\n" + "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.\n\n" * 20)
            # self.seg_button_1.configure(values=["CTkSegmentedButton", "Value 2", "Value 3"])
            # self.seg_button_1.set("Value 2")


class MenuFrame(customtkinter.CTkFrame):
    def __init__(self, master, switch_callback,**kwargs):
        super().__init__(master, **kwargs)
        self.bg_image = ctk.CTkImage(
            light_image=Image.open(img_path),
            dark_image=Image.open(img_path),
            size=(200, 130)
        )

        # Background label
        self.bg_label = ctk.CTkLabel(
            self,
            image=self.bg_image,
            text=""
        )
        self.bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.menu_label = customtkinter.CTkLabel(self, text="МЕНЮ",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.menu_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_invoices = customtkinter.CTkButton(self,text='Склад',command=lambda:  switch_callback("InventoryFrame"))
        self.sidebar_button_invoices.grid(row=1, column=0, padx=20, pady=10)

        self.sidebar_button_repair = customtkinter.CTkButton(self,text="Сервизна история",command=lambda:  switch_callback("ServicesFrame"))
        self.sidebar_button_repair.grid(row=2, column=0, padx=20, pady=10)

        self.appearance_mode_label = customtkinter.CTkLabel(self, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self,
                                                                       values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=7, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=8, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self,
                                                               values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=9, column=0, padx=20, pady=(10, 20))

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)


class InventoryFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.logo_label = customtkinter.CTkLabel(self, text="Складова наличност",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0,column=0)
        self.textbox = customtkinter.CTkTextbox(self,width=900)
        self.textbox.grid(row=2, column=0, columnspan=5, rowspan=9,  padx=(20, 0), pady=(20, 20),sticky="nsew")
        self.textbox.configure(font=("Courier New", 15))

        #buttons
        self.entry = customtkinter.CTkEntry(self, placeholder_text="Потърси ... ")
        self.entry.grid(row=1, column=0, columnspan=4, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text_color=("gray10", "#DCE4EE"), text='Търсене',
                                                     command=self.__search_item_in_db)
        self.main_button_1.grid(row=1, column=4,padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.check_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                      text_color=("gray10", "#DCE4EE"), text='Провери наличности',
                                                      command=self.__check_availability)
        self.check_button_1.grid(row=1, column=6, padx=(20, 20), pady=(20, 0), sticky="nsew")

        self.limit_label=customtkinter.CTkLabel(self,text='Минимална граница')
        self.limit_label.grid(row=2, column=6,pady=(0, 0))

        self.optionmenu_limit = customtkinter.CTkComboBox(self,
                                                    values=["1", "2", "3", "4", "5","6", "7", "8", "9","10","11","12","13","14","15","16","17","18","19","20"])
        self.optionmenu_limit.grid(row=3, column=6,pady=(0, 0))

        self.help_button = customtkinter.CTkButton(master=self, fg_color="transparent",
                                                   border_width=2, text_color=("gray10", "#DCE4EE"),
                                                   text='Упътване', command=self.__help_text)
        self.help_button.grid(row=18, column=6, padx=(20, 20), pady=(20, 20), sticky="nsew")

        # create tabview
        self.tabview = customtkinter.CTkTabview(self,border_width=2)
        self.tabview.grid(row=12, columnspan=5,rowspan=7, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.add("Склад")
        self.tabview.add("Exports")
        self.tabview.tab("Склад").grid_columnconfigure(1, weight=1)  # configure grid of individual tabs

        self.label_entry_v0 = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Парт. Номер",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v0.grid(row=1,column=0,sticky="nsew")
        self.entry_v0 = customtkinter.CTkEntry(self.tabview.tab("Склад"), placeholder_text="Парт. Номер: ... ")
        self.entry_v0.grid(row=1, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v1 = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Модел",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v1.grid(row=2,column=0,sticky="nsew")
        self.entry_v1 = customtkinter.CTkEntry(self.tabview.tab("Склад"), placeholder_text="Модел ... ")
        self.entry_v1.grid(row=2, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v2 = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Количество",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v2.grid(row=3,column=0,sticky="nsew")
        self.entry_v2 = customtkinter.CTkEntry(self.tabview.tab("Склад"), placeholder_text="Количество ... ")
        self.entry_v2.grid(row=3, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v3 = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Описание",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v3.grid(row=4,column=0,sticky="nsew")
        self.entry_v3 = customtkinter.CTkEntry(self.tabview.tab("Склад"), placeholder_text="Описание ... ")
        self.entry_v3.grid(row=4, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v4 = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Цена",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v4.grid(row=5,column=0,sticky="nsew")
        self.entry_v4 = customtkinter.CTkEntry(self.tabview.tab("Склад"), placeholder_text="Цена ... ")
        self.entry_v4.grid(row=5, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.sidebar_button_remove = customtkinter.CTkButton(
            self.tabview.tab("Склад"),
            text='Премахни',
            fg_color="transparent",
            border_width=2,
            text_color=("gray10", "#DCE4EE"),hover_color="#953333",command=self.__remove_item_in_db)
        self.sidebar_button_remove.grid(row=7, column=2, padx=(20, 20), pady=(20, 20))

        self.sidebar_button_add = customtkinter.CTkButton(
            self.tabview.tab("Склад"),
            text='Добави',
            fg_color="transparent",
            border_width=2,
            text_color=("gray10", "#DCE4EE"),command=self.__add_item_in_db)
        self.sidebar_button_add.grid(row=7, column=1, padx=(20, 20), pady=(20, 20))

        self.option_entry = customtkinter.CTkLabel(self.tabview.tab("Склад"), text="Категория",
                                                 font=customtkinter.CTkFont(size=12))
        self.option_entry.grid(row=0,column=0,sticky="nsew")
        self.optionmenu_1 = customtkinter.CTkComboBox(self.tabview.tab("Склад"),
                                                    values=["Телефони", "Калъфи", "Дисплеи", "Протектори", "Резервни части"])
        self.optionmenu_1.grid(row=0, column=1, sticky="nsw")

    def __help_text(self):
        self.textbox.insert("0.0", "CTkTextbox\n\n" + "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.\n\n" * 20)

    def __check_availability(self):
        df = pd.read_excel(inventory_path)
        min_limit = int(self.optionmenu_limit.get())
        mask = df["Количество"] <= min_limit

        item_repr = f"Няма намерени ниски количества\n\n"
        if mask.any():
            item_repr = df.loc[mask]
            item_repr = self.__dataframe_to_pretty_text(df,item_repr)
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0",text=item_repr)

    def __get_data(self):
        productID,model,qty,description, category, price = (self.entry_v0.get(),self.entry_v1.get(),
                                                     self.entry_v2.get(),self.entry_v3.get(),
                                                     self.optionmenu_1.get(), self.entry_v4.get())

        data_to_write = {"Категория": category, "Парт. Номер": productID, "Модел": model, "Количество": qty,
                         "Описание": description, "Цена":price}
        return  productID, model, qty, description, category, price, data_to_write

    def __add_item_in_db(self):
        df = pd.read_excel(inventory_path)
        productID, model, qty, description, category, price, data_to_write = self.__get_data()
        if not any((productID,model,qty,description)):
            text_repr = f"За да бъде добавен артикул трябва да има попълнено поле."
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0",text=text_repr)
            return

        if not qty.isdigit() and not qty == "0":
            qty = "1"

        if productID:
            # Search condition
            mask = df["Парт. Номер"].str.lower() == productID.lower()

            if mask.any():
                df.loc[mask, "Количество"] += int(qty)  # increase quantity
                if description:
                    df.loc[mask, "Описание"] = description
                if model:
                    df.loc[mask, "Модел"] = model
                if price:
                    df.loc[mask, "Цена"] = price
                if category:
                    df.loc[mask, "Категория"] = category
                #text that will be printed in textbox
                item_repr = df.loc[mask]
                text_repr = f"Намерен е артикул с номер: {productID}. Добавено са: {qty} бройки.\n\n"
                text = self.__dataframe_to_pretty_text(df,item_repr)
            else:
                # Append row
                df.loc[len(df)] = data_to_write

                # text that will be printed in textbox
                item_repr = df.tail(1)
                text_repr = f"Артикул с номер: {productID} не е намерен. Добавено са: {qty} бройки.\n\n"
                text = self.__dataframe_to_pretty_text(df,item_repr)
        else:
            df.loc[len(df)] = data_to_write

            # text that will be printed in textbox
            item_repr = df.tail(1)
            text_repr = f"Артикул с номер: {productID} не е намерен. Добавено са: {qty} бройки.\n\n"
            text = self.__dataframe_to_pretty_text(df, item_repr)

        df.to_excel(inventory_path, index=False)

        #clear textbox and add modified data to textbox
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0",text=text_repr)
        self.textbox.insert("end",text=text)
        #clear all entries
        self.entry_v0.delete(0,'end')
        self.entry_v1.delete(0,'end')
        self.entry_v2.delete(0,'end')
        self.entry_v3.delete(0,'end')

    def __remove_item_in_db(self):
        df = pd.read_excel(inventory_path)
        productID, model, qty, description, category, price, data_to_write = self.__get_data()

        if not qty or not qty.isdigit():
            qty = "1"
        item = False

        if productID:
            # Search condition
            mask = df["Парт. Номер"].str.lower() == productID.lower()

            if mask.any():
                df.loc[mask, "Количество"] -= int(qty) or 1  # increase quantity

                #text that will be printed in textbox
                item_repr = df.loc[mask]
                text_repr = f"Намерен е артикул с номер: {productID}. Извадени са: {qty} бройки.\n\n"
                text = self.__dataframe_to_pretty_text(df,item_repr)
                df.to_excel(inventory_path, index=False)
                item = True
            else:
                text_repr = f"Артикул с номер: {productID} не е намерен. Проверете пак номера в полето 'Парт.Номер'\n\n"
        else:
            text_repr = f"Моля добавете 'Парт.Номер' за да бъде премахнат артикул.\n\n"


        #clear textbox and add modified data to textbox
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0",text=text_repr)
        if item:
            self.textbox.insert("end",text=text)
        #clear all entries
        self.entry_v0.delete(0,'end')
        self.entry_v1.delete(0,'end')
        self.entry_v2.delete(0,'end')
        self.entry_v3.delete(0,'end')
        self.entry_v4.delete(0,'end')

    @staticmethod
    def __dataframe_to_pretty_text(df,item):
        col_widths = [
            max(len(str(val)) for val in df[col].astype(str).tolist() + [col])
            for col in df.columns
        ]

        lines = []

        # Header
        header = " | ".join(col.ljust(col_widths[i]) for i, col in enumerate(df.columns))
        separator = "-+-".join("-" * w for w in col_widths)

        lines.append(header)
        lines.append(separator)
        # Rows
        for _, row in item.iterrows():
            line = " | ".join(
                str(row[col]).ljust(col_widths[i])
                for i, col in enumerate(df.columns)
            )
            lines.append(line)

        return "\n".join(lines)

    def __search_item_in_db(self):
        searched_word = self.entry.get()

        df = pd.read_excel(inventory_path)
        # if not df["Парт. Номер"]
        mask_number = df["Парт. Номер"].str.lower() == searched_word.lower()
        mask_model = df["Модел"].str.lower() == searched_word.lower()
        mask_description = df["Описание"].str.lower() == searched_word.lower()
        mask_category = df["Категория"].str.lower() == searched_word.lower()
        is_find = False
        self.textbox.delete("1.0", "end")
        #TODO Reformat this code. too many reps
        if mask_number.any():
            item_number = df.loc[mask_number]
            text = self.__dataframe_to_pretty_text(df, item_number)
            self.textbox.insert("end", text=text)
            is_find = True

        if mask_model.any():
            item_model = df.loc[mask_model]
            text = self.__dataframe_to_pretty_text(df, item_model)
            self.textbox.insert("end", text=text)
            is_find = True

        if mask_description.any():
            item_description = df.loc[mask_description]
            text = self.__dataframe_to_pretty_text(df, item_description)
            self.textbox.insert("end", text=text)
            is_find = True

        if mask_category.any():
            item = df.loc[mask_category]
            text = self.__dataframe_to_pretty_text(df, item)
            self.textbox.insert("end", text=text)
            is_find = True

        if not is_find:
            text_repr = f"Няма намерен такъв артикул съдържаш '{searched_word}' .\n\n"
        else:
            text_repr = f"Това са артикулите съдържащи: {searched_word}.\n\n"

        #clear textbox and add modified data to textbox
        self.textbox.insert("1.0",text=text_repr)


class ServicesFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.order_number = self.__last_order_number()
        self.logo_label = customtkinter.CTkLabel(self, text="Сервизни дейности",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0,column=0)
        self.textbox = customtkinter.CTkTextbox(self,width=900)
        self.textbox.grid(row=2, column=0, columnspan=5, rowspan=9,  padx=(20, 0), pady=(20, 20),sticky="nsew")
        self.textbox.configure(font=("Courier New", 15))

        #buttons
        self.entry = customtkinter.CTkEntry(self, placeholder_text="Номер на поръчката ... ")
        self.entry.grid(row=1, column=0, columnspan=4, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text_color=("gray10", "#DCE4EE"), text='Търсене',
                                                     command=self.__find_order)
        self.main_button_1.grid(row=1, column=4,padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.save_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text_color=("gray10", "#DCE4EE"), text='Запазване',
                                                     command=self.__add_order)
        self.save_button.grid(row=19, column=1,padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.clear_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text_color=("gray10", "#DCE4EE"), text='Изчистване',
                                                    hover_color="#953333",command=self.__clear_values
                                                     )
        self.clear_button.grid(row=19, column=2,padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.delete_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2,
                                                     text_color=("gray10", "#DCE4EE"), text='Изтриване',
                                                    hover_color="#953333",command=self.__remove_order
                                                     )
        self.delete_button.grid(row=19, column=3,padx=(20, 20), pady=(20, 20), sticky="nsew")


        self.print_button = customtkinter.CTkButton(master=self, fg_color="transparent",
                                                   border_width=2, text_color=("gray10", "#DCE4EE"),state='disabled',
                                                   text='ПРИНТИРАЙ', command=self.__create_pdf_file)
        self.print_button.grid(row=17, column=6, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.help_button = customtkinter.CTkButton(master=self, fg_color="transparent",
                                                   border_width=2, text_color=("gray10", "#DCE4EE"),
                                                   text='Упътване', command=self.__help_text)
        self.help_button.grid(row=18, column=6, padx=(20, 20), pady=(20, 20), sticky="nsew")

        # create tabview
        self.tabview = customtkinter.CTkTabview(self,border_width=2)
        self.tabview.grid(row=12, columnspan=5,rowspan=7, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.add("Данни на у-во")
        self.tabview.add("Състояние")
        self.tabview.add("Данни на клиент")
        self.tabview.add("Карта")
        self.tabview.tab("Данни на у-во").grid_columnconfigure(1, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Данни на клиент").grid_columnconfigure(1, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Състояние").grid_columnconfigure(1, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Карта").grid_columnconfigure(1, weight=1)  # configure grid of individual tabs

        #TABVIEW - Данни на у-во
        self.label_entry_v0 = customtkinter.CTkLabel(self.tabview.tab("Данни на у-во"), text="Марка ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v0.grid(row=1,column=0,sticky="nsew")
        self.entry_v0 = customtkinter.CTkEntry(self.tabview.tab("Данни на у-во"), placeholder_text="Марка: ... ")
        self.entry_v0.grid(row=1, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v1= customtkinter.CTkLabel(self.tabview.tab("Данни на у-во"), text="Модел ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v1.grid(row=2,column=0)
        self.entry_v1 = customtkinter.CTkEntry(self.tabview.tab("Данни на у-во"), placeholder_text="Модел ... ")
        self.entry_v1.grid(row=2, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v1= customtkinter.CTkLabel(self.tabview.tab("Данни на у-во"), text="IMEI ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v1.grid(row=3,column=0)
        self.entry_v2 = customtkinter.CTkEntry(self.tabview.tab("Данни на у-во"), placeholder_text="IMEI ... ")
        self.entry_v2.grid(row=3, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v1= customtkinter.CTkLabel(self.tabview.tab("Данни на у-во"), text="Сериен Номер ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v1.grid(row=4,column=0)
        self.entry_v3 = customtkinter.CTkEntry(self.tabview.tab("Данни на у-во"), placeholder_text="Сериен Номер ... ")
        self.entry_v3.grid(row=4, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        #TABVIEW - Състояние
        self.label_entry_v4= customtkinter.CTkLabel(self.tabview.tab("Състояние"), text="Описание на дефекта ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v4.grid(row=1,column=0)
        self.entry_v4 = customtkinter.CTkEntry(self.tabview.tab("Състояние"), placeholder_text="Описание на дефекта: ... ")
        self.entry_v4.grid(row=1, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v5 = customtkinter.CTkLabel(self.tabview.tab("Състояние"), text="Външно състояние ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v5.grid(row=2,column=0)
        self.entry_v5 = customtkinter.CTkEntry(self.tabview.tab("Състояние"), placeholder_text="Външно състояние(Драскотини, удари) ... ")
        self.entry_v5.grid(row=2, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v6 = customtkinter.CTkLabel(self.tabview.tab("Състояние"), text="Код за отключване  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v6.grid(row=3,column=0)
        self.entry_v6 = customtkinter.CTkEntry(self.tabview.tab("Състояние"), placeholder_text="Код за отключване ... ")
        self.entry_v6.grid(row=3, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.good_condition = customtkinter.CTkCheckBox(self.tabview.tab("Състояние"),text='Работи',onvalue="Работи",
                                                        offvalue="Не работи")
        self.good_condition.grid(row=4, column=1, padx=(20, 20), pady=(0, 0), sticky="nsew")
        self.status_order = customtkinter.CTkComboBox(self.tabview.tab("Състояние"),
                                                    values=["Прието", "В процес", "Чака части", "Готово", "Взето"])
        self.status_order.grid(row=5, column=1,pady=(15, 15),padx=(20, 20),sticky="nsw")

        #TABVIEW - Данни на клиент
        self.label_entry_v7 = customtkinter.CTkLabel(self.tabview.tab("Данни на клиент"), text="Имена  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v7.grid(row=1,column=0)
        self.entry_v7 = customtkinter.CTkEntry(self.tabview.tab("Данни на клиент"), placeholder_text="Имена: ... ")
        self.entry_v7.grid(row=1, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v8 = customtkinter.CTkLabel(self.tabview.tab("Данни на клиент"), text="Телефон  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v8.grid(row=2,column=0)
        self.entry_v8 = customtkinter.CTkEntry(self.tabview.tab("Данни на клиент"), placeholder_text="Телефон ... ")
        self.entry_v8.grid(row=2, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v9 = customtkinter.CTkLabel(self.tabview.tab("Данни на клиент"), text="Е-майл  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v9.grid(row=3,column=0)
        self.entry_v9 = customtkinter.CTkEntry(self.tabview.tab("Данни на клиент"), placeholder_text="Е-майл ... ")
        self.entry_v9.grid(row=3, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        #TABVIEW - Карта
        self.label_entry_v10 = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="SN на протокол  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v10.grid(row=1,column=0)
        self.entry_v10 = customtkinter.CTkEntry(self.tabview.tab("Карта"), placeholder_text=f"{self.order_number}")
        self.entry_v10.grid(row=1, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")
        self.entry_v10.configure(state="readonly")

        self.label_entry_v11 = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="Използвани части ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v11.grid(row=2,column=0)
        self.entry_v11 = customtkinter.CTkEntry(self.tabview.tab("Карта"), placeholder_text="Използвани части(намалява позицията в склада) ... ")
        self.entry_v11.grid(row=2, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v12 = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="Гаранция  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v12.grid(row=3,column=0)
        self.entry_v12 = customtkinter.CTkEntry(self.tabview.tab("Карта"), placeholder_text="Гаранция ... ")
        self.entry_v12.grid(row=3, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_entry_v13 = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="Номер на гаранционен стикер  ",
                                                 font=customtkinter.CTkFont(size=12))
        self.label_entry_v13.grid(row=4,column=0)
        self.entry_v13 = customtkinter.CTkEntry(self.tabview.tab("Карта"), placeholder_text="Номер на гаранционен стикер ... ")
        self.entry_v13.grid(row=4, column=1, columnspan=6, padx=(5, 0), pady=(5, 5), sticky="nsew")

        self.label_service_man = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="Специалист  ",
                                                      font=customtkinter.CTkFont(size=12))
        self.label_service_man.grid(row=6, column=0)
        self.service_man_dropdown = customtkinter.CTkComboBox(self.tabview.tab("Карта"),
                                                    values=["Крум Ранков","Друг"])
        self.service_man_dropdown.grid(row=6, column=1,pady=(0, 0),padx=(20, 20),sticky="nsw")

        # Calendar Pop-up window
        self.label_entry_v14 = customtkinter.CTkLabel(self.tabview.tab("Карта"), text="Дата на ремонт  ",
                                                                                               font=customtkinter.CTkFont(size=12))
        self.label_entry_v14.grid(row=6,column=2)
        self.calendar_pop = ctk.CTkEntry(self.tabview.tab("Карта",),
                                       placeholder_text="Изберете дата",
                                       width=200)
        self.calendar_pop.bind("<Button-1>", self.__snow_calendar)
        self.calendar_pop.grid(row=6, column=3,pady=(0, 0),padx=(20, 20),sticky="nsw")

    def __snow_calendar(self,event):
        DateEntry(self,self.calendar_pop)

    def __last_order_number(self):
        df = pd.read_excel(services_path)

        current_number = df['Поръчка'].max() + 1
        if current_number.is_integer():
            return str(current_number)
        return "1"

    def __help_text(self):
        self.__clear_values()
        self.textbox.insert("0.0", "CTkTextbox\n\n" + "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.\n\n" * 20)

    def __get_data(self):
        #TabView - Данни на у-во
        brand,model,imei,ser_number = self.entry_v0.get(),self.entry_v1.get(),self.entry_v2.get(),self.entry_v3.get()

        #TabView - Състояние
        defect_description, condition, pin, status, is_working = (self.entry_v4.get(),self.entry_v5.get(),
                                                                   self.entry_v6.get(),self.status_order.get(),
                                                                   self.good_condition.get())
        # TabView - Данни на клиент
        names, phone, mail, = self.entry_v7.get(), self.entry_v8.get(),self.entry_v9.get()

        # TabView - Карта на поръзката
        order_number, parts, warranty,sticker_warranty, date_order,technic = (self.entry_v10.get(), self.entry_v11.get(),
                                                                       self.entry_v12.get(),self.entry_v13.get(),
                                                                       self.calendar_pop.get(),self.service_man_dropdown.get())
        data_to_write = {
            "Марка": brand, "IMEI": imei, "Модел": model, "Сериен Номер": ser_number,
            "Описание на дефекта": defect_description, "Външно състояние": condition,
            "PIN": pin, "Работи": is_working, "Статус": status,
            "Имена": names, "Телефон": phone, "Е-майл": mail,
            "Поръчка": order_number, "Части": parts, "Гаранция": warranty, "Стикер": sticker_warranty,
                         "Дата": date_order, "Специалист":technic}

        return  data_to_write

    def __add_order(self):
        data_to_write = self.__get_data()
        df = pd.read_excel(services_path)

        ser_number = data_to_write['Сериен Номер']
        imei = data_to_write['IMEI']
        sticker = data_to_write['Стикер']
        search_for = data_to_write['Поръчка']
        data_to_write['Поръчка'] = self.order_number


        find_order,text,found, df = self.__find_order_func(searched_number=search_for, df=df)
        if found:
            data_to_write['Поръчка'] = find_order['Поръчка']
            df = self.__remove_order(update_flag=True,df=df)
        else:
            self.order_number = int(self.order_number) + 1

        # df = pd.read_excel("service.xlsx")
        if not any((ser_number,sticker, imei)):
            text_repr = (f"За да бъде добавен сервизен протокол трябва да има попълнено поне едно от полетата:"
                         f"\n\t'Гаранционен стикер', 'IMEI', 'Сериен Номер'.")
            self.__print_message(text_repr)
            return

        self.entry_v10.configure(state="normal")

        df.loc[len(df)] = data_to_write
        df.to_excel(services_path, index=False)

        text = 'Направен е нов запис на протокол:\n\n'
        for k, v in data_to_write.items():
            text += f'\t{k}' + '_' * (30 - len(k)) + f'{v} \n\n'
        self.__clear_values()
        self.__print_message(text)

    def __remove_order(self,df=False, update_flag=False):
        if df is False:
            df = pd.read_excel(services_path)
        data, text, found, df = self.__find_order_func(df=df)
        if found:
            mask = df['Поръчка'] == data['Поръчка']
            df = df.drop(df.loc[mask].index)
            df.to_excel(services_path, index=False)
            self.__clear_values()
            if not update_flag:
                text = f'Протокол с номер {data['Поръчка']} бе изтрит от базата.\n\n'
                self.__print_message(text)
        return df

    def __create_pdf_file(self):
        data = self.__get_data()
        create_service_card(work_directory=data_dir,data=data)


    def __print_message(self,message:str):
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", text=message)

    def __find_order_func(self,searched_number=None,df=False):
        # df = pd.read_excel('service.xlsx')
        searched_number = searched_number or self.entry.get()
        text = (f"Не е намерен сервизен протокол с номер: {searched_number}")
        if searched_number:
            mask = df['Поръчка'] == int(searched_number)
            text = (f"Не е намерен сервизен протокол с номе: {searched_number}")
            data = df.loc[mask]
            if mask.any():
                data = data.iloc[0]
                text = (f"Протоколът с номер {data['Поръчка']} е открит.")
                return data,text, True,df
        if searched_number == '' or searched_number == ' ':
            data = df
            return data, text, False, df
        data = pd.DataFrame()
        return data, text, False, df

    def __find_order(self):
        df = pd.read_excel(services_path)
        data, text, found, _ = self.__find_order_func(df=df)
        self.__clear_values()
        if found:
            self.print_button.configure(state='normal')
            self.entry_v0.insert(0, data['Марка'])
            self.entry_v1.insert(0, data['Модел'])
            self.entry_v2.insert(0, data['IMEI'])
            self.entry_v3.insert(0, data['Сериен Номер'])
            self.entry_v4.insert(0, data['Описание на дефекта'])
            self.entry_v5.insert(0, data['Външно състояние'])
            self.entry_v6.insert(0, data['PIN'])
            self.entry_v7.insert(0, data['Имена'])
            self.entry_v8.insert(0, data['Телефон'])
            self.entry_v9.insert(0, data['Е-майл'])
            self.entry_v10.configure(state="normal")
            # self.entry_v10.delete(0,'end')
            self.entry_v10.insert(0, data['Поръчка'])
            self.entry_v10.configure(state="readonly")
            self.entry_v11.insert(0, data['Части'])
            self.entry_v12.insert(0, data['Гаранция'])
            self.entry_v13.insert(0, data['Стикер'])
            self.calendar_pop.insert(0, data['Дата'])
            self.service_man_dropdown.set(data['Специалист'])
            if data['Работи'].lower() == 'работи':
                self.good_condition.select()
            else:
                self.good_condition.deselect()
            self.status_order.set(data['Статус'])
            self.__print_message(text)
        elif not data.empty:
            data = df.loc[:, ['Поръчка','Марка', 'Модел', 'IMEI','Описание на дефекта','Гаранция','Дата']]
            data = data.to_string()

            self.__print_message(data)
            self.print_button.configure(state='disabled')
        else:
            self.print_button.configure(state='disabled')
            self.__print_message(text)

    def __clear_values(self):

        # self.entry.delete(0, "end") if self.entry.get() else None
        self.entry_v0.delete(0,"end") if self.entry_v0.get() else None
        self.entry_v1.delete(0,"end") if self.entry_v1.get() else None
        self.entry_v2.delete(0,"end") if self.entry_v2.get() else None
        self.entry_v3.delete(0,"end") if self.entry_v3.get() else None
        self.entry_v4.delete(0,"end") if self.entry_v4.get() else None
        self.entry_v5.delete(0,"end") if self.entry_v5.get() else None
        self.entry_v6.delete(0,"end") if self.entry_v6.get() else None
        self.entry_v7.delete(0,"end") if self.entry_v7.get() else None
        self.entry_v8.delete(0,"end") if self.entry_v8.get() else None
        self.entry_v9.delete(0,"end") if self.entry_v9.get() else None
        self.entry_v10.configure(state="normal")
        self.entry_v10.delete(0,"end") if self.entry_v10.get() else None
        self.entry_v10.configure(placeholder_text=str(self.__last_order_number()))
        self.entry_v10.configure(state="readonly")
        self.entry_v11.delete(0,"end") if self.entry_v11.get() else None
        self.entry_v12.delete(0,"end") if self.entry_v12.get() else None
        self.entry_v13.delete(0,"end") if self.entry_v13.get() else None
        self.calendar_pop.delete(0,"end") if self.calendar_pop.get() else None
        self.entry_v0.configure(placeholder_text="Марка: ...")
        self.entry_v1.configure(placeholder_text="Модел: ...")
        self.entry_v2.configure(placeholder_text="IMEI ... ")
        self.entry_v3.configure(placeholder_text="Сериен Номер ... ")
        self.entry_v4.configure(placeholder_text="Описание на дефекта: ...")
        self.entry_v5.configure(placeholder_text="Външно състояние(Драскотини, удари) ... ")
        self.entry_v6.configure(placeholder_text="Код за отключване ... ")
        self.entry_v7.configure(placeholder_text="Имена: ... ")
        self.entry_v8.configure(placeholder_text="Телефон ... ")
        self.entry_v9.configure(placeholder_text="Е-майл ...")
        self.entry_v11.configure(placeholder_text="Използвани части(намалява позицията в склада)")
        self.entry_v12.configure(placeholder_text="Гаранция ...")
        self.entry_v13.configure(placeholder_text="Номер на гаранционен стикер ... ")
        self.calendar_pop.configure(placeholder_text="Дата на извършен ремонт: ... ")
        self.good_condition.deselect()
        self.status_order.set('---')
        self.good_condition.deselect()


if __name__ == "__main__":
    app = App()
    app.iconbitmap(img_path)
    app.mainloop()
