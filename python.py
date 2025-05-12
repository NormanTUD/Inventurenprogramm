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

def find_entry(sheet, inv_number):
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if str(row[0].value).strip() == inv_number:
            console.print(f"[green]Gefunden: Anlagennummer {inv_number} in Zeile {row[0].row}[/green]")
            return row
    console.print(f"[red]Nicht gefunden: Anlagennummer {inv_number}[/red]")
    return None

def find_first_matching_entry(sheet, item_type):
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[4].value == item_type:
            return row
    return None

def insert_sorted_row(sheet, inv_number, item_type, value, room_number, person):
    yellow_fill = PatternFill(
        fill_type="solid",
        start_color="FFFF00",  # RGB für Gelb
        end_color="FFFF00"
    )

    # Spalten: A = Inventarnummer, E = Bezeichnung, H = Währung, I = Standort, J = Raum, L = Inventurhinweis, M = Kostenstelle
    new_row = [inv_number, None, None, None, item_type, None, None, "EUR", "3331", room_number, None, person, "2340200G"]
    inserted = False

    for row_idx in range(2, sheet.max_row + 1):
        cell_value = sheet.cell(row=row_idx, column=1).value
        if cell_value is None or str(cell_value).strip() == "":
            continue
        try:
            if int(inv_number) < int(cell_value):
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

    console.print(f"[blue]Eintrag: {inv_number}, {item_type}, Wert: {value}, Währung: EUR, "
                  f"Standort: 3331, Raum: {room_number}, Person: {person}[/blue]")

def save_workbook(wb, file_name):
    try:
        wb.save(file_name)
        console.print(f"[green]Änderungen erfolgreich gespeichert in {file_name}[/green]")
    except Exception as e:
        console.print(f"[red]Fehler beim Speichern der Datei: {e}[/red]")



def mark_row_as_confirmed(sheet, row_idx: int):
    # Grüne Farbe im hex Format (Grün ohne Alpha)
    green_fill = PatternFill(
        fill_type="solid",
        start_color="00FF00",  # Hex für Grün
        end_color="00FF00"     # Hex für Grün
    )

    # Zelle in Spalte A und der Zeile row_idx auswählen
    cell = sheet.cell(row=row_idx, column=1)

    # DEBUG: Vorherige Füllung überprüfen
    console.print(f"[bold yellow]DEBUG: Vorherige Füllung von Zelle A{row_idx}:[/bold yellow]")
    console.print(f"  fill_type: {cell.fill.fill_type}")
    console.print(f"  start_color: {cell.fill.start_color.rgb or cell.fill.start_color.index}")
    console.print(f"  end_color: {cell.fill.end_color.rgb or cell.fill.end_color.index}")

    # Setze die Füllung auf eine transparente Standardfarbe (weiß, ohne Füllung)
    transparent_fill = PatternFill(fill_type="none")
    cell.fill = transparent_fill  # Setzt die Zelle auf Standardfarbe zurück

    # Anwenden der grünen Füllung
    cell.fill = green_fill

    # DEBUG: Neue Füllung überprüfen
    console.print(f"[bold green]DEBUG: Neue Füllung von Zelle A{row_idx}:[/bold green]")
    console.print(f"  fill_type: {cell.fill.fill_type}")
    console.print(f"  start_color: {cell.fill.start_color.rgb}")
    console.print(f"  end_color: {cell.fill.end_color.rgb}")

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
        inv_number = input("Bitte geben Sie die Anlagennummer ein (oder 'q'=Beenden, 'p'=Person ändern, 'r'=Raum ändern): ").strip()

        if inv_number.lower() == 'q':
            console.print("[yellow]Beenden...[/yellow]")
            break
        if inv_number.lower() == 'p':
            current_person = input("Bitte neuen Namen der Person eingeben: ").strip()
            if current_person:
                console.print(f"[green]Person geändert zu: {current_person}[/green]")
            else:
                console.print("[red]Name darf nicht leer sein.[/red]")
            continue
        if inv_number.lower() == 'r':
            current_room = input("Neue Raumnummer eingeben: ").strip()
            if current_room:
                console.print(f"[green]Raumnummer geändert zu: {current_room}[/green]")
            else:
                console.print("[red]Raumnummer darf nicht leer sein.[/red]")
            continue

        if not current_person:
            console.print("[red]Bitte zuerst eine Person mit 'p' setzen.[/red]")
            continue

        row = find_entry(sheet, inv_number)

        if row:
            headers = [cell.value for cell in sheet[1]]
            console.print("Zeile enthält die folgenden Daten:")

            for i in range(len(headers)):
                header = headers[i] if i < len(headers) else f"Spalte {i+1}"
                value = row[i].value if i < len(row) else ""
                console.print(f"[cyan]{header}[/cyan]: {value}")

            action = input("Ist das korrekt? (Enter für Ja, 'e' zum Bearbeiten): ").strip()

            if action == "e":
                console.print("[yellow]Welche Option möchten Sie bearbeiten?[/yellow]")
                print("1: Anschaffungswert")
                print("2: Raum")
                option = input("Geben Sie die Nummer der zu bearbeitenden Option ein: ").strip()

                if option == "1":
                    item_type = row[1].value
                    if item_type in PRICES:
                        new_value = PRICES[item_type]
                        row[2].value = new_value
                        console.print(f"[blue]Anschaffungswert geändert auf: {new_value}[/blue]")
                    else:
                        console.print(f"[red]Kein Preis für {item_type} gefunden.[/red]")
                elif option == "2":
                    new_room = input(f"Aktueller Raum: {current_room}. Neue Raumnummer eingeben (Enter für aktuellen Raum): ").strip()
                    if new_room:
                        current_room = new_room
                    sheet.cell(row=row[0].row, column=10, value=current_room)
                    console.print(f"[blue]Raum (Spalte J) geändert auf: {current_room}[/blue]")
                else:
                    console.print("[red]Ungültige Option![/red]")

                save_workbook(wb, excel_file)
                mark_row_as_confirmed(sheet, row[0].row)
                save_workbook(wb, excel_file)

            elif action == "":
                console.print("[green]Bestätigt. Keine Änderungen.[/green]")
                mark_row_as_confirmed(sheet, row[0].row)
                save_workbook(wb, excel_file)
            else:
                console.print("[red]Ungültige Eingabe![/red]")

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

            insert_sorted_row(sheet, inv_number, item_type, value, current_room, current_person)
            save_workbook(wb, excel_file)

if __name__ == "__main__":
    main()
