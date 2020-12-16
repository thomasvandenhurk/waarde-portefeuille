import os

import pandas as pd


def format_header(wb, ws, year):
    format_header_r = wb.add_format({'bold': True, 'font_size': 24, 'align': 'right'})
    format_header_l = wb.add_format({'bold': True, 'font_size': 24, 'align': 'left'})
    ws.write(0, 0, year, format_header_r)
    ws.write(0, 1, "Waarde Portefeuille", format_header_l)


def set_cell_widths(ws, colnum):
    ws.set_column(0, 0, 34)
    ws.set_column(1, colnum, 17)
    ws.set_row(0, 32)


def format_portefeuille(wb, ws, df, start_row):
    # portefeuille header
    format_port_header = wb.add_format({'bold': True, 'bg_color': '#ECA359', 'align': 'center'})
    format_port_header_border = wb.add_format({'bold': True, 'bg_color': '#ECA359', 'align': 'center', 'left': 1})
    for col_num, value in enumerate(df.columns.values):
        if 'waarde' in value:
            value = value.replace(' (waarde)', '')
            ws.write(start_row, col_num, value, format_port_header)
        elif 'aantal' in value:
            value = ''
            ws.write(start_row, col_num, value, format_port_header_border)
        else:
            value = ''
            ws.write(start_row, col_num, value, format_port_header)

    # add numbers
    format_aantal_pos = wb.add_format({'font_color': '#796E63', 'bg_color': '#C6EFCE', 'left': 1})
    format_aantal_neg = wb.add_format({'font_color': '#796E63', 'bg_color': '#FFC7CE', 'left': 1})
    format_aantal_neutral = wb.add_format({'font_color': '#796E63', 'left': 1})
    format_waarde = wb.add_format({'num_format': '#,##0.00;-#,##0.00;—;@'})
    format_procent_pos = wb.add_format({'num_format': '0%', 'bg_color': '#C6EFCE'})
    format_procent_neg = wb.add_format({'num_format': '0%', 'bg_color': '#FFC7CE'})
    format_procent_neutral = wb.add_format({'num_format': '0%'})

    for i in range(len(df.index)):
        for j in range(len(df.columns)):
            if 'aantal' in df.columns[j]:
                if j > 2:
                    if df.iloc[i, j] > df.iloc[i, j - 3] and i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_pos)
                    elif df.iloc[i, j] < df.iloc[i, j - 3] and i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neg)
                    elif i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neutral)
                    else:
                        ws.write(start_row + i + 1, j, '', format_aantal_neutral)
                elif i > 0:
                    ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neutral)

            elif 'waarde' in df.columns[j]:
                ws.write(start_row + i + 1, j, df.iloc[i, j], format_waarde)
            elif 'procent' in df.columns[j]:
                if j > 2:
                    if df.iloc[i, j] > 0 and i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_pos)
                    elif df.iloc[i, j] < 0 and i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_neg)
                    elif i > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_neutral)
                elif i > 0:
                    ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_neutral)
            else:
                ws.write(start_row + i + 1, j, df.iloc[i, j])


def write_portefeuille(portefeuille, deposits, output_path='results'):
    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    wb = writer.book
    sheet_name = "Portefeuille"
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    start_row = 1
    portefeuille.reset_index(inplace=True)
    portefeuille.drop('Symbool/ISIN', axis=1, inplace=True)

    # format sheet
    format_portefeuille(wb, ws, portefeuille, start_row)
    format_header(wb, ws, 2020)
    set_cell_widths(ws, len(portefeuille.columns))
    ws.freeze_panes(2, 1)
    writer.save()
