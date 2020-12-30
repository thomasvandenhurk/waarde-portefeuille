import os

import pandas as pd
import matplotlib.pyplot as plt

from settings import header_r, header_l, port_header, port_header_border, aantal_pos, aantal_neg, aantal_neutral, \
    waarde, procent_pos, procent_neg, procent_neutral, totaal_font, totaal_num, winstverlies_font, winstverlies_num, \
    jaaroverzicht_font, jaaroverzicht_num, header_color, positive_color


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

    ws.set_column(0, 0, 35)
    ws.set_column(1, colnum, 17)
    ws.set_row(0, 32)


def format_portefeuille(wb, ws, df: pd.DataFrame, start_row: int, port_prev: pd.DataFrame = None):
    """
    Format portefeuille overview and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param df: Dataframe with portefeuille overview.
    :param start_row: Integer to start writing at.
    :param port_prev: Dataframe with portefeuille data from previous year.
    """

    df.drop('Symbool/ISIN', axis=1, inplace=True)

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
                        if totals.iloc[i, j] > 0:
                            ws.write(start_row + i + 1, j, totals.iloc[i, j], format_procent_pos)
                        elif totals.iloc[i, j] < 0:
                            ws.write(start_row + i + 1, j, totals.iloc[i, j], format_procent_neg)
                        else:
                            ws.write(start_row + i + 1, j, totals.iloc[i, j], format_procent_neutral)
            else:
                if i == 0:
                    ws.write(start_row + i + 1, j, totals.iloc[i, j], format_totaal_font)
                else:
                    ws.write(start_row + i + 1, j, totals.iloc[i, j])


def format_winstverlies(wb, ws, winstverlies: pd.DataFrame, start_row: int):
    """
    Format wint verlies row and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param winstverlies: Series with winstverlies overview.
    :param start_row: Integer to start writing at.
    """

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
                if winstverlies.iloc[i] > 0:
                    ws.write(start_row + 1, i + 1, winstverlies.iloc[i], format_procent_pos)
                elif winstverlies.iloc[i] < 0:
                    ws.write(start_row + 1, i + 1, winstverlies.iloc[i], format_procent_neg)
                else:
                    ws.write(start_row + 1, i + 1, winstverlies.iloc[i], format_procent_neutral)
            else:
                ws.write(start_row + 1, i + 1, "", format_winstverlies_num)


def format_jaaroverzicht(wb, ws, totals: pd.DataFrame, start_row: int, year: str, total_invested: float):
    """
    Format jaaroverzicht and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param totals: Dataframe with totals overview.
    :param start_row: Integer to start writing at.
    :param year: String with current year of writing.
    :param total_invested: Float number of total investments.
    :return total_invested: Float updated number of total investments.
    """

    format_winstverlies_font = wb.add_format(winstverlies_font)
    format_jaaroverzicht_font = wb.add_format(jaaroverzicht_font)
    format_jaaroverzicht_num = wb.add_format(jaaroverzicht_num)

    totals_waarde = totals.loc[:, totals.columns.str.contains('waarde')]
    totals_waarde = totals_waarde.append(pd.Series(
        totals_waarde.loc['Verschil t.o.v. vorige maand', :] - totals_waarde.loc['Inleg', :], name='Winst/Verlies'
    ), ignore_index=False).transpose()
    totals_waarde.index = totals_waarde.index.str[:10]
    totals_waarde = totals_waarde[['Totaal portefeuille', 'Inleg', 'Winst/Verlies']].reset_index()
    totals_waarde['Inleg'] = totals_waarde['Inleg'].cumsum() + total_invested
    total_invested = totals_waarde['Inleg'].iloc[-1]
    totals_waarde = totals_waarde.rename(columns={'Totaal portefeuille': 'Portefeuille'})

    for i in range(len(totals.columns) + 1):
        if 0 < i < len(totals_waarde.columns):
            ws.write(start_row, i, totals_waarde.columns[i], format_jaaroverzicht_font)
        elif i >= len(totals_waarde.columns):
            ws.write(start_row, i, '', format_jaaroverzicht_font)
        else:
            ws.write(start_row, 0, 'Jaaroverzicht', format_winstverlies_font)

    for i in range(len(totals_waarde.index)):
        for j in range(len(totals_waarde.columns)):
            if j > 0:
                ws.write(start_row + i + 1, j, totals_waarde.iloc[i, j], format_jaaroverzicht_num)
            else:
                ws.write(start_row + i + 1, j, totals_waarde.iloc[i, j])

    add_jaaroverzicht_plot(ws, totals_waarde, start_row, year)

    return total_invested


def add_jaaroverzicht_plot(ws, totals_waarde: pd.DataFrame, start_row: int, year: str):
    """
    Capture jaaroverzicht in image and write to Excel.

    :param ws: xlsxwriter Worksheet object.
    :param totals_waarde: Dataframe with totals info.
    :param start_row: Integer to start writing at.
    :param year: String with current year of writing.
    """

    plt.plot(totals_waarde['Portefeuille'], color=header_color, label='Portefeuille')
    plt.plot(totals_waarde['Inleg'], 'k', label='Inleg')
    totals_waarde['Winst/Verlies'].plot(kind='bar', color=positive_color, label='Winst/Verlies')

    ax = plt.gca()
    ax.set_xticklabels(totals_waarde['index'])
    plt.xticks(rotation=45)
    plt.title('Portefeuille ontwikkeling ' + year)
    plt.legend()
    plt.grid(True)
    plt.savefig('portefeuille_ontwikkeling_' + year + '.png', bbox_inches='tight', dpi=100)
    ws.insert_image(start_row + 2, 5, 'portefeuille_ontwikkeling_' + year + '.png')
    plt.close()


def write_portefeuille(portefeuille_dict: dict, totals_dict: dict, winstverlies_dict: dict,
                       output_path: str = 'results'):
    """
    Write portefeuille info to Excel.

    :param portefeuille_dict: Dict with portefeuille info.
    :param totals_dict: Dict with totals info.
    :param winstverlies_dict: Dict with winstverlies.
    :param output_path: String where to write the output to.
    """

    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    wb = writer.book

    port_prev = None
    total_invested = 0
    for key in portefeuille_dict:
        portefeuille = portefeuille_dict[key]
        totals = totals_dict[key]
        winstverlies = winstverlies_dict[key]

        ws = wb.add_worksheet(key)
        writer.sheets[key] = ws

        start_row = 1

        # format sheet
        format_portefeuille(wb, ws, portefeuille, start_row, port_prev)
        start_row += len(portefeuille.index) + 2
        format_totals(wb, ws, totals, start_row)
        start_row += len(totals.index) + 1
        format_winstverlies(wb, ws, winstverlies, start_row)
        start_row += 4
        total_invested = format_jaaroverzicht(wb, ws, totals, start_row, key, total_invested)
        format_header(wb, ws, key)
        set_cell_widths(ws, len(portefeuille.columns))
        ws.freeze_panes(2, 1)

        port_prev = portefeuille

    # save the file
    writer.save()
