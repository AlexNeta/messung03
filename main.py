# -*- coding: utf-8 -*-

from instrument import Instrument
import xlsxwriter
from csv import reader, writer, QUOTE_MINIMAL

from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.image import Image
from kivy.properties import NumericProperty, BooleanProperty, ObjectProperty, StringProperty, ListProperty
from kivy.clock import Clock
from kivy.uix.listview import ListView
from kivy.uix.switch import Switch
# _____________________________________________________________________________________________________________________
# Toolbar


class Exit(Popup):
    pass


class About(Popup):
    pass


class New(Popup):
    toolbar = ObjectProperty()

    tester_name = StringProperty()
    meas_number = StringProperty()
    number_light = StringProperty()

    def build(self, toolbar_parent):
        self.toolbar = toolbar_parent
        self.tester_name = self.toolbar.parent.tester_name
        self.meas_number = self.toolbar.parent.meas_number
        self.number_light = self.toolbar.parent.number_light

    def make_new_file(self, name, meas_number, number_light):
        self.toolbar.parent.tester_name = name
        self.toolbar.parent.meas_number = meas_number
        self.toolbar.parent.number_light = number_light
        self.dismiss()


class Save(Popup):
    toolbar = ObjectProperty()
    save_dire = StringProperty("")
    save_name = StringProperty("Untitled")

    def build(self, toolbar_parent):
        self.toolbar = toolbar_parent
        self.save_dire = self.toolbar.parent.save_dire
        self.save_name = self.toolbar.parent.save_name[:-5]

    def save_file_name(self, name, dire):
        self.toolbar.parent.save_dire = dire
        self.toolbar.parent.save_name = name
        self.toolbar.parent.save_settings()
        self.dismiss()


class Instr(ListView):
    def find_all_instruments(self, instr_list):
        for name in instr_list:
            self.adapter.data.extend([name])
        self.adapter.data.remove("")
        self._trigger_reset_populate()


class Loading(Image):

    def found(self):
        self.source = "icons/Tick.png"

    def not_found(self):
        self.source = "icons/Close.png"


class Set(Popup):
    instr_list_view = ObjectProperty()
    load_icon_view = ObjectProperty()
    toolbar = ObjectProperty()
    device_id = StringProperty("")

    def build(self, toolbar_parent):
        self.toolbar = toolbar_parent
        # Search all instruments
        instr_list = self.toolbar.parent.instrument.show_instr_list()
        self.instr_list_view.find_all_instruments(instr_list)

    def connect_to_instrument(self):
        if self.instr_list_view.adapter.selection:
            selection = self.instr_list_view.adapter.selection[0].text
            # Connect to Selected Instrument
            if self.toolbar.parent.instrument.connect_to_device(selection):
                try:
                    self.device_id = self.toolbar.parent.instrument.get_id()
                    self.toolbar.parent.instr_name = selection
                    self.toolbar.parent.save_settings()
                    self.load_icon_view.found()
                except Exception:
                    self.device_id = ""
                    self.load_icon_view.not_found()
            else:
                self.device_id = ""
                self.load_icon_view.not_found()
        else:
            self.device_id = ""
            self.load_icon_view.not_found()


class ToolBar(GridLayout):
    mode = StringProperty("")

    def set_instrument(self):
        s = Set()
        s.open()
        s.build(self)

    def new_file(self):
        n = New()
        n.open()
        n.build(self)

    def save_file(self):
        s = Save()
        s.open()
        s.build(self)

    @staticmethod
    def about():
        About().open()

    @staticmethod
    def exit():
        Exit().open()


# _____________________________________________________________________________________________________________________
# Main
class MainWindow(BoxLayout):
    toolbar = ObjectProperty()
    # Settings:
    instr_name = StringProperty("")  # Save in settings
    save_dire = StringProperty("")  # Save in settings
    save_name = StringProperty("Untitled.xlsx")  # Save in settings
    saved_as = StringProperty("")
    # Instrument:
    instrument = ObjectProperty()
    # Setup:
    tester_name = StringProperty()  # Name des Testers
    meas_number = StringProperty()  # Name der Messung
    number_light = StringProperty()  # Anzahl Leuchten (int)
    curr_light = NumericProperty()  # Aktuelle Leuchte (int)
    # Measurement-Switch:
    buttons_label = ObjectProperty()
    switch_start = ObjectProperty()
    # Messages:
    meas_message = StringProperty()

    def build(self):
        self.instrument = Instrument()
        self.get_measurement_data()
        Clock.schedule_interval(self.init_measurement, 1. / 10.)
        Clock.schedule_once(lambda dt: self.toolbar.new_file(), 0.2)

    def get_measurement_data(self):
        try:
            self.load_settings()
        except Exception:
            pass
    # _________________________________________________________________________________________________________________
    # Settings

    def load_settings(self):
        save_file = []
        with open("~/settings.csv", "r", newline="") as csv_file:
            r = reader(csv_file, delimiter=" ", quotechar="|")
            for x in r:
                save_file.append("".join(x))
        # Load data
        self.save_dire, self.save_name, self.instr_name = save_file
        self.instrument.connect_to_device(self.instr_name)  # Connect to instrument

    def save_settings(self):
        # Save data
        save_file = self.save_dire, self.save_name, self.instr_name

        with open("~/settings.csv", "w", newline="") as csv_file:
            r = writer(csv_file, delimiter=" ", quotechar="|", quoting=QUOTE_MINIMAL)
            for x in save_file:
                r.writerow(x)

    # _________________________________________________________________________________________________________________
    # Measurement

    def init_measurement(self, dt):
        if self.tester_name == "" or self.meas_number == "" or self.number_light == "":
            self.meas_message = "Bitte neue Messung einrichten\n(weißes Blatt oben anklicken)"
        else:
            self.switch_start = Switch(size_hint_y=None, height=35)
            self.buttons_label.add_widget(self.switch_start)
            self.curr_light = 1  # Start at 1
            Clock.schedule_interval(self.start_measurement, 1./60.)
            Clock.unschedule(self.init_measurement)

    def start_measurement(self, dt):
        if self.instrument.connected:
            self.meas_message = "Gerät erkannt\nMessung kann gestartet werden"
        if self.instrument.connected and self.switch_start.active:
            self.meas_message = "Messung wurde gestartet"
            self.buttons_label.remove_widget(self.switch_start)
            self.init_channel()
            Clock.unschedule(self.start_measurement)
        elif not self.instrument.connected:
            self.meas_message = "Kein Gerät eingerichtet\n(Oben links einstellen)"

    def init_channel(self):
        self.instrument.instr_on(0)
        self.instrument.gen_on()
        Clock.schedule_once(self.get_data, 2)

    def get_data(self, dt):
        self.meas_message = "Bitte Anschließen der {}/{} Leuchte".format(self.curr_light, self.number_light)
    
    def end_measurement(self):
        self.instrument.instr_off(0)
        self.instrument.gen_off()


class MeasurementApp(App):
    icon = './icons/icon.png'

    def build(self):
        w = MainWindow()
        w.build()
        return w


if __name__ == "__main__":
    MeasurementApp().run()

"""
            Switch:
                size_hint_y: None
                height: 35
                id: switch_start
                on_active: root.start_measurement(self.active)
         Label:
            text: "[color=ff0000]ALL current Measurements will be deleted![/color]"
            markup: True
"""