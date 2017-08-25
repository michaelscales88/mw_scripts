from pyexcel import get_sheet, Sheet
from datetime import datetime, timedelta

# from os.path import join

# report_directory = r'M:/Help Desk/Daily SLA Report/2017'
cari_clients = [
    '7506 AAP',
    '7509 Assoc Materials',
    '7516 BBA Aviation',
    '7515 BMW NoAm',
    '7579 Plains',
    '7541 Sonic Auto',
    '7589 Tokyo Electron'
]

perry_clients = [
    '7517 Humana',
    '7547 ULA',
    '7545 Danaher',
    '7512 Fortive',
    '7592 CocaColaCCR',
    '7591 CocaColaTCCC',
    '7593 CocaColaCanada',
    '7577 Oshkosh',
    '7548 Ecova',
    '7561 Netscout'
]


def collect_sheets(client):
    date_start = datetime.today().replace(year=2017, month=6, day=1)
    date_end = datetime.today().replace(year=2017, month=6, day=30)
    total_sheet = Sheet()
    total_sheet_colnames = None
    while date_start <= date_end:
        file_string = r'M:/Help Desk/Daily SLA Report/2017/{month}/{sla_rpt_fmt}_Incoming DID Summary.xlsx'.format(
            month=date_start.strftime('%B'), sla_rpt_fmt=date_start.strftime('%m%d%Y')
        )
        sheet = get_sheet(file_name=file_string, name_columns_by_row=0, name_rows_by_column=0)
        total_sheet_colnames = sheet.colnames

        try:
            new_row = [str(date_start)] + list(sheet.row[client])
            total_sheet.row += new_row
        except ValueError:
            pass
        date_start += timedelta(days=1)
    total_sheet.name_rows_by_column(0)
    total_sheet.colnames = total_sheet_colnames
    print(total_sheet)
    total_sheet.save_as('{mo}_{cli}_output.xlsx'.format(mo=date_end.strftime('%b'), cli=client))


if __name__ == '__main__':
    # clients = perry_clients
    # for a_client in clients:
    #     collect_sheets(a_client)
    collect_sheets('7506 AAP')
