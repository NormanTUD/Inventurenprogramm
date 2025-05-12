import sys
import openpyxl
from rich.console import Console
from openpyxl.styles import PatternFill
from pprint import pprint

def dier(msg):
    pprint(msg)
    sys.exit(10)

console = Console()

PRICES = {
    "Besprechungsstuhl Fiore Vierbeiner weiß/hellgrün mit Klapptisch": 241.45,
    "Rollcontainer 9 HE 1-2-3-3 Buche": 185.42,
    "Sitz-Steharbeitsplatz 180x80x64-125-Buche": 487.55,
    "Drehstuhl AJ5786 schwarz": 322.24,
    "Besprechungsstühle Cay Vierbeiner grau/hellgrün": 168.45,
    "Lenovo ThinkPad T14 AMD Gen 3": 2152.71,
    "Monitor Lenovo ThinkVision T27h-2L": 304.00
}

# TODO: Barcodes für Thinkpads mit einlesen
# Bugs:
# Fügt die an der falschen Stelle ein und dann grün
# Man kann den Namen nicht ändern

PREDEFINED_ITEM_TYPES = list(PRICES.keys())

current_person = ""
current_room = ""

def show_item_type_menu():
    console.print("[yellow]Gerätetyp auswählen:[/yellow]")
    for idx, name in enumerate(PREDEFINED_ITEM_TYPES, start=1):
        console.print(f"{idx}: {name}")
    choice = input("Nummer eingeben: ").strip()
    try:
        index = int(choice) - 1
        if 0 <= index < len(PREDEFINED_ITEM_TYPES):
            return PREDEFINED_ITEM_TYPES[index]
        else:
            console.print("[red]Ungültige Auswahl![/red]")
            return show_item_type_menu()
    except ValueError:
        console.print("[red]Bitte eine gültige Zahl eingeben![/red]")
        return show_item_type_menu()

def find_entry(sheet, anlagennummer):
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if str(row[0].value).strip() == anlagennummer:
            console.print(f"[green]Gefunden: Anlagennummer {anlagennummer} in Zeile {row[0].row}[/green]")
            return row
    console.print(f"[red]Nicht gefunden: Anlagennummer {anlagennummer}[/red]")
    return None

def find_first_matching_entry(sheet, item_type):
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[4].value == item_type:
            return row
    return None

def insert_sorted_row(sheet, anlagennummer, item_type, value, room_number, person):
    yellow_fill = PatternFill(
        fill_type="solid",
        start_color="FFFF00",  # RGB für Gelb
        end_color="FFFF00"
    )

    # Spalten: A = Inventarnummer, E = Bezeichnung, H = Währung, I = Standort, J = Raum, L = Inventurhinweis, M = Kostenstelle
    new_row = [anlagennummer, None, None, None, item_type, None, None, "EUR", "3331", room_number, None, person, "2340200G"]
    inserted = False

    for row_idx in range(2, sheet.max_row + 1):
        cell_value = sheet.cell(row=row_idx, column=1).value
        if cell_value is None or str(cell_value).strip() == "":
            continue
        try:
            if int(anlagennummer) < int(cell_value):
                sheet.insert_rows(row_idx)
                for col, val in enumerate(new_row, start=1):
                    sheet.cell(row=row_idx, column=col, value=val)
                sheet.cell(row=row_idx, column=7, value=value)  # Anschaffungswert in Spalte C
                inserted = True
                console.print(f"[blue]Neue Zeile sortiert eingefügt vor Zeile {row_idx}[/blue]")

                # Gelbes Füllmuster anwenden
                for col in range(1, len(new_row) + 1):
                    sheet.cell(row=row_idx, column=col).fill = yellow_fill

                # Einfärben der Zeile in grün (wie im ersten Codebeispiel)
                for col in range(1, len(new_row) + 1):
                    sheet.cell(row=row_idx, column=col).fill = yellow_fill

                break
        except ValueError:
            continue

    if not inserted:
        sheet.append(new_row)
        last_row = sheet.max_row
        sheet.cell(row=last_row, column=3, value=value)  # Anschaffungswert in Spalte C
        console.print(f"[blue]Neue Zeile ans Ende angehängt[/blue]")

        # Gelbes Füllmuster anwenden
        for col in range(1, len(new_row) + 1):
            sheet.cell(row=last_row, column=col).fill = yellow_fill

        # Einfärben der Zeile in grün (wie im ersten Codebeispiel)
        green_fill = PatternFill(
            fill_type="solid",
            start_color="00FF00",
            end_color="00FF00"
        )
        for col in range(1, len(new_row) + 1):
            sheet.cell(row=last_row, column=col).fill = green_fill

    console.print(f"[blue]Eintrag: {anlagennummer}, {item_type}, Wert: {value}, Währung: EUR, "
                  f"Standort: 3331, Raum: {room_number}, Person: {person}[/blue]")

def save_workbook(wb, file_name):
    try:
        wb.save(file_name)
        console.print(f"[green]Änderungen erfolgreich gespeichert in {file_name}[/green]")
    except Exception as e:
        console.print(f"[red]Fehler beim Speichern der Datei: {e}[/red]")



def mark_row_as_confirmed(sheet, row_idx):
    # Grüne Farbe im hex Format (Grün ohne Alpha)
    green_fill = PatternFill(
        fill_type="solid",
        start_color="00FF00",  # Hex für Grün
        end_color="00FF00"     # Hex für Grün
    )

    # Zelle in Spalte A und der Zeile row_idx auswählen
    cell = sheet.cell(row=row_idx, column=1)

    # Anwenden der grünen Füllung
    cell.fill = green_fill

    console.print(f"[green]Zelle A{row_idx} erfolgreich grün überschrieben.[/green]")

def main():
    global current_person, current_room

    if len(sys.argv) != 3:
        console.print("[red]Verwendung: python excel.py <excelfile> <worksheet>[/red]")
        sys.exit(1)

    excel_file = sys.argv[1]
    worksheet_name = sys.argv[2]

    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb[worksheet_name]
        console.print(f"[green]Arbeitsblatt '{worksheet_name}' wurde geöffnet.[/green]")
    except Exception as e:
        console.print(f"[red]Fehler beim Öffnen der Excel-Datei: {e}[/red]")
        sys.exit(1)

    while not current_room:
        current_room = input("Bitte geben Sie die Raumnummer ein: ").strip()
        if not current_room:
            console.print("[red]Raumnummer darf nicht leer sein.[/red]")

    while not current_person:
        current_person = input("Bitte Namen der zugeordneten Person eingeben: ").strip()
        if not current_person:
            console.print("[red]Name darf nicht leer sein.[/red]")

    while True:
        ask_for_input = "Bitte geben Sie die Anlagennummer ein (oder 'q'=Beenden, 'p'=Person ändern, 'r'=Raum ändern): "
        if current_person:
            ask_for_input = f"Name: {current_person}, {ask_for_input}"
        if current_room:
            ask_for_input = f"Raum: {current_room}, {ask_for_input}"

        anlagennummer_oder_kommando = input(ask_for_input).strip()

        if anlagennummer_oder_kommando.lower() == 'q':
            console.print("[yellow]Beenden...[/yellow]")
            break
        if anlagennummer_oder_kommando.lower() == 'p':
            current_person = input("Bitte neuen Namen der Person eingeben: ").strip()
            if current_person:
                console.print(f"[green]Person geändert zu: {current_person}[/green]")
            else:
                console.print("[red]Name darf nicht leer sein.[/red]")
            continue
        if anlagennummer_oder_kommando.lower() == 'r':
            current_room = input("Neue Raumnummer eingeben: ").strip()
            if current_room:
                console.print(f"[green]Raumnummer geändert zu: {current_room}[/green]")
            else:
                console.print("[red]Raumnummer darf nicht leer sein.[/red]")
            continue

        row = find_entry(sheet, anlagennummer_oder_kommando)

        if row:
            headers = []

            for cell in sheet[1]:
                headers.append(cell.value)

            console.print("Zeile enthält die folgenden Daten:")

            for i in range(len(headers)):
                header = f"Spalte {i + 1}"
                if i < len(headers):
                    header = headers[i]

                value = ""
                if i < len(row):
                    value = row[i].value

                console.print(f"[cyan]{header}[/cyan]: {value}")

            edit_msg = "Ist das korrekt? (Enter für Ja, 'e' zum Bearbeiten): "

            action = input(edit_msg).strip()

            is_valid_option = False

            while not is_valid_option:
                if action.lower() == "e":
                    console.print("[yellow]Welche Option möchtest du bearbeiten?[/yellow]")
                    print("p: Person")
                    print("r: Raum")
                    print("z: Zurück")
                    option = input("Gebe die Nummer der zu bearbeitenden Option ein: ").strip()

                    if option.lower() == "p":
                        row[11].value = current_person
                        console.print(f"[blue]Person geändert auf: {current_person}[/blue]")
                        mark_row_as_confirmed(sheet, row[0].row)
                        save_workbook(wb, excel_file)
                        is_valid_option = True

                    elif option.lower() == "r":
                        sheet.cell(row=row[0].row, column=10, value=current_room)
                        console.print(f"[blue]Raum (Spalte J) geändert auf: {current_room}[/blue]")
                        mark_row_as_confirmed(sheet, row[0].row)
                        save_workbook(wb, excel_file)
                        is_valid_option = True

                    elif option.lower() == "z":
                        console.print("[green]Ich änder doch nix[/green]")
                        is_valid_option = True

                    else:
                        console.print(f"[red]Ungültige Option! Option: {option}[/red]")

                elif action.lower() in ["", "y", "j"]:
                    console.print("[green]Bestätigt. Keine Änderungen.[/green]")
                    mark_row_as_confirmed(sheet, row[0].row)
                    save_workbook(wb, excel_file)
                    is_valid_option = True

                else:
                    console.print(f"[red]Ungültige Eingabe! Eingabe: '{action}'[/red]")
                    action = input(edit_msg).strip()
                    is_valid_option = False


        else:
            item_type = show_item_type_menu()
            matching_row = find_first_matching_entry(sheet, item_type)
            if matching_row and matching_row[2].value:
                value = matching_row[2].value
                console.print(f"[green]Gefundene Werte: {item_type} → Wert={value}, Währung=EUR[/green]")
            else:
                if item_type in PRICES:
                    value = PRICES[item_type]
                    console.print(f"[green]Gefundener Wert für {item_type}: {value} EUR[/green]")
                else:
                    console.print("[red]Kein Preis für das angegebene Item gefunden![/red]")
                    continue

            insert_sorted_row(sheet, anlagennummer_oder_kommando, item_type, value, current_room, current_person)
            save_workbook(wb, excel_file)

    print("das hier kommt nach der while schleife")

if __name__ == "__main__":
    try:
        main()
    except (EOFError, KeyboardInterrupt):
        console.print("\n[yellow]Du hast das Programm beendet[/yellow]")
