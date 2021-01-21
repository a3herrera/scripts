import os
import shutil
from os import listdir
from os.path import isfile, join
import win32print
import win32api
from time import sleep

# Ruta donde se van a estar viendo los archivos a imprimir
PATH_TO_WATCH = 'C:/Users/a3her/OneDrive/Documents/Libros/'

# Ruta donde se van a estar moviendo los archivos que se imprimieron
DESTINATION_PATH = 'C:/Users/a3her/OneDrive/Documents/Other/'

# Nombre de las impresoras a utilizar, pueden ser locales o en red
PRINTER_NAMES = [
    'EPSON L3110 Series',
    'EPSON L3110 Series',
]

# Nombre de los errores que se pueden obtener para una impresora
PRINTER_ERRORS = [
    win32print.PRINTER_STATUS_NO_TONER,
    win32print.PRINTER_STATUS_NOT_AVAILABLE,
    win32print.PRINTER_STATUS_OFFLINE,
    win32print.PRINTER_STATUS_OUT_OF_MEMORY,
    win32print.PRINTER_STATUS_OUTPUT_BIN_FULL,
    win32print.PRINTER_STATUS_PAGE_PUNT,
    win32print.PRINTER_STATUS_PAPER_JAM,
    win32print.PRINTER_STATUS_PAPER_OUT,
    win32print.PRINTER_STATUS_PAPER_PROBLEM
]


def exist_directory_to_watch() -> bool:
    # Indica Si/No si existe el directorio a estar vigilando
    return os.path.exists(PATH_TO_WATCH)


def get_files():
    # Regresa el listado de los archivos que se encuentran dentro de la carpeta que se esta vigilando
    return [join(PATH_TO_WATCH, file) for file in listdir(PATH_TO_WATCH) if isfile(join(PATH_TO_WATCH, file))]


def is_valid_printer_status(printer_name) -> bool:
    printer_status = win32print.GetPrinter(printer_name, 1)
    print(printer_status)
    # return printer_status not in PRINTER_ERRORS
    return True


def send_file_to_printer(file):
    printer_defaults = {
        "DesiredAccess": win32print.PRINTER_ALL_ACCESS
    }
    for printer_name in PRINTER_NAMES:
        printer = win32print.OpenPrinter(printer_name, printer_defaults)
        #
        # if not is_valid_printer_status(printer_name):
        #     continue
        printer_attrs = win32print.GetPrinter(printer, 2)
        win32print.SetPrinter(printer, 2, printer_attrs, 0)
        win32api.ShellExecute(0, "print", file, None, ".", 0)
        win32print.ClosePrinter(printer)


def move_file(file):
    shutil.move(file, DESTINATION_PATH)


def main():
    if not any(PRINTER_NAMES) or not exist_directory_to_watch():
        return

    files = get_files()
    if not any(files):
        return

    for file in files:
        send_file_to_printer(file)
        sleep(5) # Este es un tiempo para dormir ya que se tarda un tiempo para poder mandar el comando a imprimir
        move_file(file)


main()
