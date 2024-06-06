import win32print
import win32api

def get_printer_handle(printer_name):
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for printer in printers:
        if printer[2] == printer_name:
            return win32print.OpenPrinter(printer_name)
    return None

printer_name = "XP-58"
printer_handle = get_printer_handle(printer_name)

if printer_handle:
    print(f"Impresora '{printer_name}' encontrada y el handle obtenido.")
    printer_info = win32print.GetPrinter(printer_handle, 2)  
    print("Configuración actual de la impresora obtenida.")
    printer_info['pPortName'] = 'USB001'  
    win32print.SetPrinter(printer_handle, 2, printer_info, 0)
    print("El puerto de la impresora ha sido cambiado a USB001.")
    win32print.ClosePrinter(printer_handle)
else:
    print(f"No se encontró la impresora '{printer_name}'.")
