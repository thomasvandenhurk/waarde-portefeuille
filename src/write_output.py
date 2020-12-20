import os

import pandas as pd

from settings import header_r, header_l, port_header, port_header_border, aantal_pos, aantal_neg, aantal_neutral, \
    waarde, procent_pos, procent_neg, procent_neutral, totaal_font, totaal_num, winstverlies_font, winstverlies_num


def format_header(wb, ws, year: int):
    """
    Format sheet header.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param year: Dataframe with totals overview.
    """

    format_header_r = wb.add_format(header_r)
    format_header_l = wb.add_format(header_l)
    ws.write(0, 0, year, format_header_r)
    ws.write(0, 1, "Waarde Portefeuille", format_header_l)


def set_cell_widths(ws, colnum: int):
    """
    Format column widths for the sheet.

    :param ws: xlsxwriter Worksheet object.
    :param colnum: Integer last col to set the width of.
    """

    ws.set_column(0, 0, 34)
    ws.set_column(1, colnum, 17)
    ws.set_row(0, 32)


def format_portefeuille(wb, ws, df: pd.DataFrame, start_row: int):
    """
    Format portefeuille overview and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param df: Dataframe with portefeuille overview.
    :param start_row: Integer to start writing at.
    """

    # portefeuille header
    format_port_header = wb.add_format(port_header)
    format_port_header_border = wb.add_format(port_header_border)
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
    format_aantal_pos = wb.add_format(aantal_pos)
    format_aantal_neg = wb.add_format(aantal_neg)
    format_aantal_neutral = wb.add_format(aantal_neutral)
    format_waarde = wb.add_format(waarde)
    format_procent_pos = wb.add_format(procent_pos)
    format_procent_neg = wb.add_format(procent_neg)
    format_procent_neutral = wb.add_format(procent_neutral)

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


def format_totals(wb, ws, totals: pd.DataFrame, start_row: int):
    """
    Format overview totals and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param totals: Dataframe with totals overview.
    :param start_row: Integer to start writing at.
    """

    format_totaal_font = wb.add_format(totaal_font)
    format_totaal_num = wb.add_format(totaal_num)
    totals = totals.reset_index()
    format_procent_pos = wb.add_format(procent_pos)
    format_procent_neg = wb.add_format(procent_neg)
    format_procent_neutral = wb.add_format(procent_neutral)

    for i in range(len(totals.index)):
        for j in range(len(totals.columns)):
            if j > 0:
                if 'waarde' in totals.columns[j]:
                    if i == 0:
                        ws.write(start_row + i + 1, j, totals.iloc[i, j], format_totaal_num)
                    else:
                        ws.write(start_row + i + 1, j, totals.iloc[i, j])
                elif 'procent' in totals.columns[j]:
                    if j > 4 and i == 0:
                        value = (totals.iloc[i, j - 1] - totals.iloc[i, j - 4]) / totals.iloc[i, j - 4]
                        if value > 0:
                            ws.write(start_row + i + 1, j, value, format_procent_pos)
                        elif value < 0:
                            ws.write(start_row + i + 1, j, value, format_procent_neg)
                        else:
                            ws.write(start_row + i + 1, j, value, format_procent_neutral)
            else:
                if i == 0:
                    ws.write(start_row + i + 1, j, totals.iloc[i, j], format_totaal_font)
                else:
                    ws.write(start_row + i + 1, j, totals.iloc[i, j])


def format_winstverlies(wb, ws, totals: pd.DataFrame, start_row: int):
    """
    Format wint verlies row and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param totals: Dataframe with totals overview.
    :param start_row: Integer to start writing at.
    """

    winstverlies = totals.loc['Verschil t.o.v. vorige maand', :] - totals.loc['Aankopen', :]

    format_procent_pos = wb.add_format(procent_pos)
    format_procent_neg = wb.add_format(procent_neg)
    format_procent_neutral = wb.add_format(procent_neutral)
    format_winstverlies_font = wb.add_format(winstverlies_font)
    format_winstverlies_num = wb.add_format(winstverlies_num)

    ws.write(start_row + 1, 0, "Winst/Verlies", format_winstverlies_font)
    for i in range(len(winstverlies.index)):
        if 'waarde' in winstverlies.index[i]:
            ws.write(start_row + 1, i + 1, winstverlies.iloc[i], format_winstverlies_num)
        elif 'aantal' in winstverlies.index[i]:
            ws.write(start_row + 1, i + 1, "", format_winstverlies_num)
        else:
            if i > 2:
                value = winstverlies.iloc[i - 1] / totals.iloc[0, i - 4]
                if value > 0:
                    ws.write(start_row + 1, i + 1, value, format_procent_pos)
                elif value < 0:
                    ws.write(start_row + 1, i + 1, value, format_procent_neg)
                else:
                    ws.write(start_row + 1, i + 1, value, format_procent_neutral)
            else:
                ws.write(start_row + 1, i + 1, "", format_winstverlies_num)


def write_portefeuille(portefeuille: pd.DataFrame, totals: pd.DataFrame, output_path: str = 'results'):
    """
    Write portefeuille info to Excel.

    :param portefeuille: DataFrame with portefeuille info.
    :param totals: Dataframe with totals info.
    :param output_path: String where to write the output to.
    """

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
    start_row += len(portefeuille.index) + 2
    format_totals(wb, ws, totals, start_row)
    start_row += len(totals.index) + 1
    format_winstverlies(wb, ws, totals, start_row)
    format_header(wb, ws, 2020)
    set_cell_widths(ws, len(portefeuille.columns))
    ws.freeze_panes(2, 1)
    writer.save()
