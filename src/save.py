import xlsxwriter as xlsx
import datetime
from os.path import realpath


def save_as(data, results, usb_path=None):
    # _________________________________________________
    # Data:

    pruefer = data["tester_name"]
    pruefnummer = data["meas_number"]
    leuchte = data["testing_light"]
    anzahl = data["number_light"]

    moegliche_fehler = data["optischeFehler"]

    spannung = data["Spannung"]
    strombereich1 = data["Strombereich_LED1"]
    strombereich2 = data["Strombereich_LED2"]

    now = datetime.datetime.now()
    datum = "{}.{}.{}".format(str(now.day).zfill(2), str(now.month).zfill(2), now.year)
    save_datum = "{}-{}-{}".format(now.year,
                                   str(now.month).zfill(2),
                                   str(now.day).zfill(2))
    uhrzeit = "{}:{}".format(str(now.hour).zfill(2), str(now.minute).zfill(2))

    leuchte_werte = {"stromwerte": results["Stromwerte"], "opt_Fehler": results["opt_Fehler"]}

    name = "{}__{}__{}".format(save_datum, uhrzeit, pruefnummer)
    save_path = realpath("messergebnisse/{}.xlsx".format(name))
    if usb_path is not None:
        usb_path = realpath("{}/{}.xlsx".format(usb_path, name))
        save_path = (save_path, realpath(usb_path))

    # __________________________________________________
    # Save:
    # Save two times:
    for p in save_path:
        wb = xlsx.Workbook(p)
        ws = wb.add_worksheet()

        # Formatting:
        label = wb.add_format({'bold': True, "font_size": 18, "align": "center", "border": 2})
        side_border = wb.add_format({"left": 2, "right": 2})
        top_border = wb.add_format({"left": 2, "top": 2, "right": 2})
        bottom_border = wb.add_format({"left": 2, "bottom": 2, "right": 2})
        border = wb.add_format({"border": 2})

        red = wb.add_format({"bg_color": "red", "left": 2, "right": 2})
        green = wb.add_format({"bg_color": "green", "left": 2, "right": 2})

        # Adjust the column width
        ws.set_column(0, 4, 25)
        # Set side borders
        ws.set_column(0, 2 + len(moegliche_fehler), 25, side_border)

        # Start at:
        row = 0
        col = 0

        # Write file:
        # Write Label (with merged cells)
        ws.merge_range(0, 0,
                       0, 2 + len(moegliche_fehler), "Messergebnisse der Strommessungen", label)

        # Write input data:
        row += 2

        ws.write(row, col, "Prüfer", top_border)
        ws.write(row + 1, col, "Prüfnummer")
        ws.write(row + 2, col, "Leuchte")
        ws.write(row + 3, col, "Anzehl der Leuchten", bottom_border)

        ws.write(row, col + 1, pruefer, top_border)
        ws.write(row + 1, col + 1, pruefnummer)
        ws.write(row + 2, col + 1, leuchte)
        ws.write(row + 3, col + 1, anzahl, bottom_border)

        # Write measuring range:
        row += 5

        ws.write(row, col, "Spannung [V]", top_border)
        ws.write(row + 1, col, "Strombereich LED weiß [mA]")
        ws.write(row + 2, col, "Strombereich LED rot [mA]", bottom_border)

        ws.write(row, col + 1, spannung, top_border)
        ws.write(row, col + 2, "", top_border)
        ws.write(row + 1, col + 1, strombereich1[0]*1000)
        ws.write(row + 1, col + 2, strombereich1[1]*1000)
        ws.write(row + 2, col + 1, strombereich2[0]*1000, bottom_border)
        ws.write(row + 2, col + 2, strombereich2[1]*1000, bottom_border)

        # Write date and time:
        row -= 5
        col += 3

        ws.write(row, col, "Datum", top_border)
        ws.write(row + 1, col, "Uhrzeit", bottom_border)

        ws.write(row, col + 1, datum, top_border)
        ws.write(row + 1, col + 1, uhrzeit, bottom_border)

        # Write results:
        row += 9
        col -= 3

        ws.write(row, col, "Leuchte", border)
        ws.write(row, col + 1, "Stromwerte LED weiß [mA]", border)
        ws.write(row, col + 2, "Stromwerte LED rot [mA]", border)
        ws.write(row - 1, col + 3, "optische Fehler", border)
        # Add all possible optical errors:
        for i, err in enumerate(moegliche_fehler):
            ws.write(row, col + 3 + i, err, border)

        def get_color(min_curr, max_curr, c):
            if min_curr <= c <= max_curr:
                return green
            else:
                return red

        for i, curr in enumerate(leuchte_werte["stromwerte"]):
            ws.write(row + 1 + i, col, i + 1)

            if strombereich1[0] != 0 and strombereich1[1] != 0:
                color = get_color(strombereich1[0], strombereich1[1], curr[0])
                ws.write(row + 1 + i, col + 1, curr[0]*1000, color)
            if strombereich2[0] != 0 and strombereich2[1] != 0:
                color = get_color(strombereich2[0], strombereich2[1], curr[1])
                ws.write(row + 1 + i, col + 2, curr[1]*1000, color)

            opt_errors = leuchte_werte["opt_Fehler"][i]
            for j, err in enumerate(moegliche_fehler):
                if err in opt_errors:
                    ws.write(row + 1 + i, col + 3 + j, 1, red)
                else:
                    ws.write(row + 1 + i, col + 3 + j, 0, green)

        wb.close()
