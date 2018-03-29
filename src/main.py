# -*- coding: utf-8 -*-

from src.instrument import Instrument
from csv import reader, writer, QUOTE_MINIMAL
from os.path import realpath
import openpyxl as oxl

from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.image import Image
from kivy.properties import NumericProperty, ObjectProperty, StringProperty, DictProperty, ListProperty
from kivy.clock import Clock
from kivy.uix.listview import ListView
from kivy.uix.switch import Switch
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.label import Label
# _____________________________________________________________________________________________________________________
# Toolbar


class About(Popup):
    pass


class New(Popup):
    toolbar = ObjectProperty()

    # Prüfer:
    name_spinner = ObjectProperty()
    # Auftragsbezeichung:
    meas_number = StringProperty()
    # Leuchte:
    light_spinner = ObjectProperty()
    # Anzahl:
    number_light = StringProperty()

    def build(self, toolbar_parent):
        self.toolbar = toolbar_parent

        self.name_spinner.values = self.toolbar.parent.personal
        self.meas_number = self.toolbar.parent.meas_number
        self.light_spinner.values = self.toolbar.parent.leuchten["Referenznummer"]
        self.number_light = str(self.toolbar.parent.number_light)

    def make_new_file(self, name, meas_number, testing_light, number_light):
        if name != "Namen auswählen" and meas_number != "" and testing_light != "Leuchte auswählen" and number_light != 0:
            self.toolbar.parent.tester_name = name
            self.toolbar.parent.meas_number = meas_number
            self.toolbar.parent.testing_light = testing_light
            self.toolbar.parent.number_light = int(number_light)
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


# _____________________________________________________________________________________________________________________
# Main
class MainWindow(BoxLayout):
    toolbar = ObjectProperty()
    # Settings:
    instr_name = StringProperty("")  # Save in settings
    save_dire = StringProperty("")  # Save in settings
    save_name = StringProperty("Untitled.xlsx")  # Save in settings
    saved_as = StringProperty("")
    # From-Excel-Files
    personal = ListProperty()
    leuchten = DictProperty()
    # Instrument:
    instrument = ObjectProperty()
    # Setup:
    tester_name = StringProperty()      # Name des Testers
    meas_number = StringProperty()      # Name der Messung
    testing_light = StringProperty()   # Ausgewählte Leuchte aus Excel-Datei
    number_light = NumericProperty()    # Anzahl Leuchten (int)
    curr_light = NumericProperty()      # Aktuelle Leuchte (int)
    # Measurement-Switch:
    buttons_label = ObjectProperty()
    switch_start = ObjectProperty()
    test_widgets = DictProperty()
    # Messages:
    meas_message = StringProperty()
    meas_in_range = ObjectProperty()
    # Results:
    results = DictProperty()
    io_nio = ListProperty([0, 0])

    def build(self):
        # Startvorgang des Programms:
        self.instrument = Instrument()
        # Laden der Einstellungen:
        self.get_measurement_data()
        # Verbinden zum Gerät:
        if self.instrument.connect_to_device(self.instr_name):
            print("Verbindung zum gerät hergestellt")
        else:
            print("Keine Verbindung zum Gerät möglich")
        # Laden der Excel-Dateien:
        self.load_from_excel()
        # Sequenz zur Erkennung ob Daten ausgefüllt wurden:
        Clock.schedule_interval(self.init_measurement, 1. / 10.)
        # Starten des ersten Fensters zum ausfüllen der Daten:
        Clock.schedule_once(lambda dt: self.toolbar.new_file(), 0.2)

    def get_measurement_data(self):
        print("Laden der Einstellungen aus:")
        print(realpath("src/settings.csv"))
        try:
            self.load_settings()
            print("Laden erfolgreich!")
        except Exception:
            print("Laden war nicht möglich!")
            pass
    # _________________________________________________________________________________________________________________
    # Einstellungen

    def load_settings(self):
        save_file = []
        with open(realpath("src/settings.csv"), "r", newline="") as csv_file:
            r = reader(csv_file, delimiter=" ", quotechar="|")
            for x in r:
                save_file.append("".join(x))
        # Load data
        print("Einstellungen:", save_file)
        self.save_dire, self.save_name, self.instr_name = save_file

    def save_settings(self):
        # Save data
        save_file = self.save_dire, self.save_name, self.instr_name

        with open(realpath("src/settings.csv"), "w", newline="") as csv_file:
            r = writer(csv_file, delimiter=" ", quotechar="|", quoting=QUOTE_MINIMAL)
            for x in save_file:
                r.writerow(x)

    # _________________________________________________________________________________________________________________
    # Laden aus den Excel-Dateien

    def load_from_excel(self):
        # Laden der Leuchtendaten:
        wb = oxl.load_workbook(realpath("excel_datei_einstellungen/Leuchten.xlsx"))
        wsh = wb.active
        self.leuchten["Referenznummer"] = []
        self.leuchten["Spannung"] = []
        self.leuchten["LED1Minimalstrom"] = []
        self.leuchten["LED1Maximalstrom"] = []
        self.leuchten["LED2Minimalstrom"] = []
        self.leuchten["LED2Maximalstrom"] = []

        for row in wsh.iter_rows(row_offset=1):
            if row[0].value is not None:
                print(row[2].value)
                self.leuchten["Referenznummer"].append(row[0].value)
                self.leuchten["Spannung"].append(row[1].value)
                self.leuchten["LED1Minimalstrom"].append(row[2].value/1000)
                self.leuchten["LED1Maximalstrom"].append(row[3].value/1000)
                self.leuchten["LED2Minimalstrom"].append(row[4].value/1000)
                self.leuchten["LED2Maximalstrom"].append(row[5].value/1000)

        # Laden der Personalladen:
        wb = oxl.load_workbook(realpath("excel_datei_einstellungen/Personal.xlsx"))
        wsh = wb.active

        for row in wsh.iter_rows(row_offset=1):
            if row[0].value is not None:
                self.personal.append(row[0].value)

        # Laden der möglichen optischen Fehler:
        wb = oxl.load_workbook(realpath("excel_datei_einstellungen/optische_Fehler.xlsx"))
        wsh = wb.active
        self.leuchten["optischeFehler"] = []

        for row in wsh.iter_rows(row_offset=1):
            if row[0].value is not None:
                self.leuchten["optischeFehler"].append(row[0].value)

        print(self.leuchten)

    # _________________________________________________________________________________________________________________
    # Messung

    def init_measurement(self, dt):
        # Falls alle Daten eingegeben und ausgewählt:
        if self.tester_name == "" or \
                self.meas_number == "" or \
                self.number_light == "" or \
                self.testing_light == "Leuchte auswählen":
            self.meas_message = "Bitte neue Messung einrichten\n(weißes Blatt oben anklicken)."
        else:
            self.switch_start = Switch(size_hint_y=None, height=35)
            self.buttons_label.add_widget(self.switch_start)
            self.curr_light = 1  # Start at 1
            # Hinzufügen der Ergebnisse:
            self.results["Stromwerte"] = []
            self.results["Leuchten_iO"] = []
            self.results["Fehler"] = []
            Clock.schedule_interval(self.start_measurement, 1./60.)
            Clock.unschedule(self.init_measurement)

    def start_measurement(self, dt):
        if self.instrument.connected:
            self.meas_message = "Gerät erkannt.\n" \
                                "Messung kann gestartet werden."
        if self.instrument.connected and self.switch_start.active:
            self.meas_message = "Messung wurde gestartet."
            self.buttons_label.remove_widget(self.switch_start)
            self.init_channel()
            Clock.unschedule(self.start_measurement)
        elif not self.instrument.connected:
            self.meas_message = "Kein Gerät eingerichtet\n(Oben links einstellen)."

    def get_testing_light_nr(self):
        return [i for i, x in enumerate(self.leuchten["Referenznummer"]) if x == self.testing_light][0]

    def init_channel(self):
        # Nummer der ausgewählten Leuchte
        nr = self.get_testing_light_nr()
        # Spannung und Strom
        volt = self.leuchten["Spannung"][nr]
        curr = "MAX"
        # Kanäle einstellen
        self.instrument.ch_set(ch_nr=0, volt=volt, curr=curr)
        self.instrument.ch_set(ch_nr=1, volt=volt, curr=curr)
        # Kanäle einschalten
        self.instrument.instr_on(0)
        self.instrument.instr_on(1)
        self.instrument.gen_on()
        # Nach etwas Zeit kann erst mit der Messung begonnen werden
        Clock.schedule_once(lambda dt: self.connect_light(), 2)

    # Anschließen der Leuchten:
    def connect_light(self):
        self.meas_message = "Bitte Anschließen der {}/{} Leuchte".format(self.curr_light, self.number_light)
        Clock.schedule_interval(self.listen_channel_connected, 0.4)

    def disconnect_light(self):
        self.meas_message = "Bitte diese Leuchte entfernen"
        Clock.schedule_interval(self.listen_channel_disconnected, 0.4)

    def listen_channel_disconnected(self, dt):
        volt1, curr1 = self.instrument.ch_measure(0)
        volt2, curr2 = self.instrument.ch_measure(1)
        if curr1 < 0.001 and curr2 < 0.001:
            Clock.schedule_once(lambda dt: self.connect_light(), 1)
            Clock.unschedule(self.listen_channel_disconnected)

    def listen_channel_connected(self, dt):
        volt1, curr1 = self.instrument.ch_measure(0)
        volt2, curr2 = self.instrument.ch_measure(1)
        if curr1 > 0.001 and curr2 > 0.001:
            self.meas_message = "Bitte warten bis sich der Strom stabilisiert hat."
            Clock.schedule_once(lambda dt: self.measure_light(), 3)
            Clock.unschedule(self.listen_channel_connected)

    def measure_light(self):
        nr = self.get_testing_light_nr()
        # Messen des Stromes beider Kanäle:
        curr = self.instrument.ch_measure(ch_nr=0)[1], self.instrument.ch_measure(ch_nr=1)[1]
        self.results["Stromwerte"].append(curr)
        # Anfragen ob im Bereich
        if self.leuchten["LED1Maximalstrom"][nr] >= curr[0] >= self.leuchten["LED1Minimalstrom"][nr] \
                and self.leuchten["LED2Maximalstrom"][nr] >= curr[1] >= self.leuchten["LED2Minimalstrom"][nr]:
            self.meas_message = "[color=#268d0d]Messwerte in Ordnung[/color]"

            # Messung liegt im Bereich
            self.results["Leuchten_iO"].append(True)
            self.meas_in_range.found()
            self.optical_testing_init()
        else:
            # Messung liegt nicht im Bereich
            self.meas_in_range.not_found()
            self.add_buttons_measurement()

    def add_buttons_measurement(self):
        # Hinzufügen der Buttons um nochmal zu messen oder fortzufahren:
        self.test_widgets["Messung_fortfahren"] = Button(size_hint_y=None, height=35, text="Messung trotzdem forfahren")
        self.test_widgets["Messung_fortfahren"].bind(on_release=self.continue_meas)
        self.test_widgets["Messung_wiederholen"] = Button(size_hint_y=None, height=35, text="Messung wiederholen")
        self.test_widgets["Messung_wiederholen"].bind(on_release=self.remeasure)
        self.test_widgets["Box_Messung"] = BoxLayout(orientation="horizontal")
        self.test_widgets["Box_Messung"].add_widget(self.test_widgets["Messung_wiederholen"])
        self.test_widgets["Box_Messung"].add_widget(self.test_widgets["Messung_fortfahren"])
        self.buttons_label.add_widget(self.test_widgets["Box_Messung"])
        self.meas_message = "[color=ff0000]Messwerte liegen nicht im Bereich![/color]"

    # Messung wiederholen
    def remeasure(self, inst):
        self.meas_message = "Messung wird wiederholt!"
        # Buttons werden entfernt:
        self.buttons_label.remove_widget(self.test_widgets["Box_Messung"])
        # Altes Messergebnis entfernen
        del self.results["Stromwerte"][-1]
        # Messvorgang wieder beginnen
        Clock.schedule_interval(self.listen_channel_connected, 0.4)

    # Messung trotzdem forfahren
    def continue_meas(self, inst):
        # Buttons werden entfernt:
        self.buttons_label.remove_widget(self.test_widgets["Box_Messung"])
        # Fortfahren
        self.results["Leuchten_iO"].append(False)
        self.optical_testing_init()

    # Optisches testen
    def add_buttons_optical_test(self):
        self.test_widgets["Leuchte_ok"] = Button(size_hint_y=None, height=35, text="Leuchte ok")
        self.test_widgets["Leuchte_ok"].bind(on_release=self.light_works)
        self.test_widgets["Leuchte_fehlerhaft"] = Button(size_hint_y=None, height=35, text="Leuchte fehlerhaft")
        self.test_widgets["Leuchte_fehlerhaft"].bind(on_release=self.light_defect)
        self.test_widgets["Strom_umstellen_LED1"] = Switch(size_hint_y=None, height=35)
        self.test_widgets["Strom_umstellen_LED1"].bind(active=self.switch_light1)
        self.test_widgets["Strom_umstellen_LED2"] = Switch(size_hint_y=None, height=35)
        self.test_widgets["Strom_umstellen_LED2"].bind(active=self.switch_light2)

        self.test_widgets["Box_optisch"] = BoxLayout(orientation="vertical")
        box_ok = BoxLayout(orientation="horizontal")
        self.test_widgets["Box_optisch"].add_widget(box_ok)
        box_ok.add_widget(self.test_widgets["Leuchte_ok"])
        box_ok.add_widget(self.test_widgets["Leuchte_fehlerhaft"])
        self.test_widgets["Box_optisch"].add_widget(self.test_widgets["Strom_umstellen_LED1"])
        self.test_widgets["Box_optisch"].add_widget(self.test_widgets["Strom_umstellen_LED2"])
        self.buttons_label.add_widget(self.test_widgets["Box_optisch"])
        self.meas_message = "Mit dem Schaltern die LEDs ein-/ausschalten"

    def optical_testing_init(self):
        self.add_buttons_optical_test()
        self.instrument.instr_off(0)
        self.instrument.instr_off(1)

    def switch_light1(self, inst, value):
        # Strom umstellen falls Switch gedrückt wird:
        if value:
            self.instrument.instr_on(0)
        else:
            self.instrument.instr_off(0)

    def switch_light2(self, inst, value):
        # Strom umstellen falls Switch gedrückt wird:
        if value:
            self.instrument.instr_on(1)
        else:
            self.instrument.instr_off(1)

    def light_defect(self, inst):
        self.buttons_label.remove_widget(self.test_widgets["Box_optisch"])
        # Falls Leuchte defekt
        # Hinzufügen eines Spinners zur Auswahl des Defekts
        self.test_widgets["Fehler_spinner"] = Spinner(size_hint_y=None, height=35,
                                                      text="Auswahl des Fehlers",
                                                      values=self.leuchten["optischeFehler"])
        self.test_widgets["Auswahl"] = Button(size_hint_y=None, height=35,
                                              text="Fehler bestätigen",
                                              background_color=(1, 0, 0, 1))
        self.test_widgets["Auswahl"].bind(on_release=self.add_defect)

        self.test_widgets["Box_Error"] = BoxLayout(orientation="horizontal")
        self.test_widgets["Box_Error"].add_widget(self.test_widgets["Fehler_spinner"])
        self.test_widgets["Box_Error"].add_widget(self.test_widgets["Auswahl"])

        self.buttons_label.add_widget(self.test_widgets["Box_Error"])

    def add_defect(self, inst):
        defect = self.test_widgets["Fehler_spinner"].text
        # Nur forfahren falls ein Defekt ausgewählt wurde
        if defect != "Auswahl des Fehlers":
            self.buttons_label.remove_widget(self.test_widgets["Box_Error"])
            self.results["Fehler"].append(defect)
            # Ändern des letzten Ergebnisses falls optisch die Messung nicht in Ordnung ist
            self.results["Leuchten_iO"][-1] = False
            self.end_measurement()

    def light_works(self, inst):
        self.buttons_label.remove_widget(self.test_widgets["Box_optisch"])
        self.results["Fehler"].append("keine")
        self.end_measurement()

    def end_measurement(self):
        self.curr_light += 1

        liste = self.results["Leuchten_iO"]
        self.io_nio = sum(liste), len(liste) - sum(liste)
        print(self.results)

        if self.curr_light > self.number_light:
            # Kanäle ausschalten
            self.instrument.instr_off(0)
            self.instrument.instr_off(1)
            # Messung zu Ende neues Fesnster Öffnen
            self.tester_name = ""
            self.curr_light = 1
            # Sequenz zur Erkennung ob Daten ausgefüllt wurden:
            Clock.schedule_interval(self.init_measurement, 1. / 10.)
            # Starten des ersten Fensters zum ausfüllen der Daten:
            Clock.schedule_once(lambda dt: self.toolbar.new_file(), 0.2)
            self.save_result()
        else:
            self.instrument.instr_on(0)
            self.instrument.instr_on(1)
            self.disconnect_light()

    # Speichern aller Daten
    def save_result(self):
        # TODO Speichern
        pass


class MeasurementApp(App):
    icon = './icons/icon.png'

    def build(self):
        w = MainWindow()
        w.build()
        return w