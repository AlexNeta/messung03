from csv import reader, writer, QUOTE_MINIMAL

def load_settings():
    save_file = []
    with open("settings.csv", newline="") as csv_file:
        r = reader(csv_file, delimiter=" ", quotechar="|")
        for x in r:
            save_file.append("".join(x))
    # Load data
    print(save_file)


def save_settings():
    # Save data
    save_file = "U:\\Arbeitsordner MA\\PRAKTIKANT\\Netaev", "messung03.xlsx", "ASRL5::INSTR"

    with open("settings.csv", "w", newline="") as csv_file:
        r = writer(csv_file, delimiter=" ", quotechar="|", quoting=QUOTE_MINIMAL)
        for x in save_file:
            r.writerow(x)


save_settings()
load_settings()
