from openpyxl import load_workbook

# CONTROLs
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Agu", "Sep", "Oct", "Nov", "Dec"]
wb_month_ref = {
    "Jan": "C2",
    "Feb": "C11",
    "Mar": "C19",
    "Apr": "C27",
    "May": "C35",
    "Jun": "C19",
    "Jul": "C19",
    "Agu": "C19",
    "Sep": "C19",
    "Oct": "C19",
    "Nov": "C19",
    "Dec": "C19",
}

wk_to_take_values = None
wk_to_receive_values = None

def load_work_books():
    global wk_to_take_values
    global wk_to_receive_values

    wk_to_take_values = load_workbook(filename = 'planilha-rentabilidade-2022-Editavel.xlsm', data_only=True)
    wk_to_receive_values = load_workbook(filename = "planilha-rentabilidade-2022-Consolidado-Editavel.xlsm")

def init():
    load_work_books()


init()
