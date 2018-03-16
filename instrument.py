import visa


class Instrument:
    # _____________________________________________________________________________________________________________________
    # Initialize instrument
    def __init__(self):
        self.rm = visa.ResourceManager('@py')
        self.hmp = None
        self.connected = False

    def show_instr_list(self):
        return self.rm.list_resources()

    def connect_to_device(self, instr_name):
        try:
            self.hmp = self.rm.open_resource(instr_name)
            self.hmp.timeout = 500
            # If the device returns id, then its connected:
            if self.get_id() != "":
                self.connected = True
                return True
            else:
                return False
        except visa.VisaIOError:
            return False

    def get_id(self):
        return self.hmp.query("*IDN?")
    # _____________________________________________________________________________________________________________________
    # Main Control Parts

    def ch_set(self, ch_nr, volt=None, curr=None):
        if self.connected:
            self.hmp.write("INST:NSEL " + str(ch_nr + 1))

            if volt is not None:
                self.hmp.write("VOLT " + str(volt))

            if curr == "MAX" or curr == "max":
                self.hmp.write("CURR MAX")
            elif curr is not None:
                self.hmp.write("CURR " + str(curr))

    def ch_measure(self, ch_nr):
        if self.connected:
            self.hmp.write("INST:NSEL " + str(ch_nr + 1))
            return float(self.hmp.ask("MEASure:VOLT[:DC]?")), float(self.hmp.ask("MEAS:CURR[:DC]?"))
        else:
            return None, None

    def all_off(self):
        self.gen_off()
        for i in range(4):
            self.instr_off(i)

    def instr_off(self, ch_nr):
        if self.connected:
            self.hmp.write("INST:NSEL " + str(ch_nr + 1))
            self.hmp.write("OUTP:SEL OFF")

    def gen_off(self):
        if self.connected:
            self.hmp.write("OUTP:GEN OFF")

    def instr_on(self, ch_nr):
        if self.connected:
            self.hmp.write("INST:NSEL " + str(ch_nr + 1))
            self.hmp.write("OUTP:SEL ON")

    def gen_on(self):
        if self.connected:
            self.hmp.write("OUTP:GEN ON")
