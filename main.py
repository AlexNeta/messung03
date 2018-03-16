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
# _____________________________________________________________________________________________________________________
# Toolbar


class Exit(Popup):
    pass


class About(Popup):
    pass


class New(Popup):
    toolbar = ObjectProperty()

    def build(self, toolbar_parent):
        self.toolbar = toolbar_parent

    def make_new_file(self):
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
                    print(self.toolbar.parent.instrument.get_id())
                    self.toolbar.parent.instr_name = selection
                    self.toolbar.parent.save_settings()
                    self.load_icon_view.found()
                except Exception:
                    self.load_icon_view.not_found()
        self.load_icon_view.not_found()


class ToolBar(GridLayout):
    safety_lock = BooleanProperty(False)
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
    instr_name = StringProperty("")  # Save in settings
    save_dire = StringProperty("")  # Save in settings
    save_name = StringProperty("Untitled.xlsx")  # Save in settings
    saved_as = StringProperty("")
    instrument = ObjectProperty()

    bulb = ObjectProperty()

    def build(self):
        self.instrument = Instrument()
        self.get_measurement_data()
        self.bulb.color = [1, 0, 0, 1]

    def get_measurement_data(self):
        try:
            self.load_settings()
        except Exception:
            pass
    # _________________________________________________________________________________________________________________
    # Settings

    def load_settings(self):
        save_file = []
        with open("settings.csv", "r", newline="") as csv_file:
            r = reader(csv_file, delimiter=" ", quotechar="|")
            for x in r:
                save_file.append("".join(x))
        # Load data
        self.save_dire, self.save_name, self.instr_name = save_file
        self.instrument.connect_to_device(self.instr_name)  # Connect to instrument

    def save_settings(self):
        # Save data
        save_file = self.save_dire, self.save_name, self.instr_name

        with open("settings.csv", "w", newline="") as csv_file:
            r = writer(csv_file, delimiter=" ", quotechar="|", quoting=QUOTE_MINIMAL)
            for x in save_file:
                r.writerow(x)

    # _________________________________________________________________________________________________________________
    # Measurement

    def start_measurement(self, active):
        if active:
            Clock.schedule_interval(self.get_data, 1./60.)
        else:
            Clock.unschedule(self.get_data)
            self.end_measurement()

    def get_data(self, dt):
        self.measurement()
    
    def end_measurement(self):

        self.instrument.instr_off(0)
        self.instrument.gen_off()
        self.bulb.color = [1, 0, 0, 1]

    def measurement(self):
        self.instrument.instr_on(0)
        self.instrument.gen_on()
        self.bulb.color = [0, 1, 0, 1]


class MeasurementApp(App):
    icon = './icons/icon.png'

    def build(self):
        w = MainWindow()
        w.build()
        return w


if __name__ == "__main__":
    MeasurementApp().run()
