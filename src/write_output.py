import calendar

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from src.read_data import read_dividends, read_costs
from settings import header_r, header_l, port_header, port_header_border, aantal_pos, aantal_neg, aantal_neutral, \
    waarde, procent_pos, procent_neg, procent_neutral, totaal_font, totaal_num, winstverlies_font, winstverlies_num, \
    jaaroverzicht_font, jaaroverzicht_num, header_color, positive_color, months_translation, rekeningoverzicht_filename


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

    # make vrije ruimte first entry
    df['Product'] = df['Product'].str.replace('CASH', 'AAACASH')
    df.sort_values('Product', inplace=True)
    df['Product'] = df['Product'].str.replace('AAACASH', 'CASH')
    df.drop('Symbool/ISIN', axis=1, inplace=True)

    keeps = (df.drop('Product', axis=1) != 0).any(axis=1)
    df = df.loc[keeps].reset_index(drop=True).copy()
    if port_prev is not None:
        port_prev = port_prev.loc[keeps].reset_index(drop=True).copy()
        port_prev = port_prev[port_prev.columns[-1][:10] + ' (aantal)']

    # portefeuille header
    format_port_header = wb.add_format(port_header)
    format_port_header_border = wb.add_format(port_header_border)
    for col_num, value in enumerate(df.columns.values):
        if 'waarde' in value:
            value = value.replace(' (waarde)', '')
            value = str(int(value[8:10])) + ' ' + months_translation[calendar.month_name[int(value[5:7])]]
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

    # count number of rows with cash positions (need different style below)
    rows_cash_position = df['Product'].str.contains('CASH').sum() - 1

    for i in range(len(df.index)):
        for j in range(len(df.columns)):
            # writing aantallen
            if 'aantal' in df.columns[j]:
                if j > 2:
                    # we can determine the aantal coloring based on previous entry
                    if df.iloc[i, j] > df.iloc[i, j - 3] and i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_pos)
                    elif df.iloc[i, j] < df.iloc[i, j - 3] and i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neg)
                    elif i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neutral)
                    else:
                        ws.write(start_row + i + 1, j, '', format_aantal_neutral)
                elif i > rows_cash_position and port_prev is None:
                    # first entry of the year and there is not previous data
                    ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neutral)
                elif i > rows_cash_position:
                    # first entry of the year but we have a previous year
                    if df.iloc[i, j] > port_prev[i] and i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_pos)
                    elif df.iloc[i, j] < port_prev[i] and i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neg)
                    elif i > rows_cash_position:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_aantal_neutral)
                    else:
                        ws.write(start_row + i + 1, j, '', format_aantal_neutral)

            elif 'waarde' in df.columns[j]:
                ws.write(start_row + i + 1, j, df.iloc[i, j], format_waarde)
            elif 'procent' in df.columns[j]:
                if i > rows_cash_position:
                    if df.iloc[i, j] > 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_pos)
                    elif df.iloc[i, j] < 0:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_neg)
                    else:
                        ws.write(start_row + i + 1, j, df.iloc[i, j], format_procent_neutral)
            else:
                # product names
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


def format_jaaroverzicht(
        wb, ws, totals: pd.DataFrame, start_row: int, year: str, total_invested: float, total_waarde_full: list,
        deposits: pd.DataFrame, benchmark: pd.DataFrame
):
    """
    Format jaaroverzicht and write to sheet.

    :param wb: xlsxwriter Workbook object.
    :param ws: xlsxwriter Worksheet object.
    :param totals: Dataframe with totals overview.
    :param start_row: Integer to start writing at.
    :param year: String with current year of writing.
    :param total_invested: Float number of total investments.
    :param total_waarde_full: List with total waarde to pass to overview.
    :param deposits: Dataframe with deposits made.
    :param benchmark: Dataframe with benchmark pricing data.

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

    add_jaaroverzicht_plot(ws, totals_waarde, deposits, benchmark, start_row, year)

    total_waarde_full.append(totals_waarde)  # pass to overview over the years

    return total_invested, total_waarde_full


def simulate_benchmark_values(totals_waarde: pd.DataFrame, deposits: pd.DataFrame, benchmark: pd.DataFrame):
    """
    Simulate value of benchmark file with the deposits and starting value that was made.

    :param totals_waarde: List with total waarde to pass to overview.
    :param deposits: Dataframe with deposits made.
    :param benchmark: Dataframe with benchmark pricing data.

    :return totals_waarde: Dataframe with total waarde.
    """

    deposits_period = deposits[
        (deposits['Datum'] > min(totals_waarde['index'])) &
        (deposits['Datum'] <= max(totals_waarde['index']))
    ]
    deposits_period = pd.concat([
        deposits_period,
        pd.DataFrame({
                'Datum': [pd.to_datetime(totals_waarde['index'].iloc[0])],
                'Storting': [totals_waarde['Portefeuille'].iloc[0]]
        })
    ]).reset_index(drop=True)
    deposits_period = deposits_period.merge(benchmark, left_on='Datum', right_on='date').drop('date', axis=1)
    deposits_period = deposits_period.sort_values('Datum').reset_index(drop=True)

    for col in deposits_period:
        if 'price_' in col:
            name = col.replace('price_', '')
            deposits_period[f"{name}"] = (deposits_period["Storting"] / deposits_period[col]).cumsum() * deposits_period[col]
            deposits_period = deposits_period.drop(col, axis=1)

    deposits_period['Datum'] = deposits_period['Datum'].astype(str)
    totals_waarde = totals_waarde.merge(deposits_period, left_on='index', right_on='Datum', how='left')
    totals_waarde = totals_waarde.drop(['Storting', 'Datum'], axis=1)
    totals_waarde = totals_waarde.ffill().bfill()

    return totals_waarde


def add_jaaroverzicht_plot(
        ws, totals_waarde: pd.DataFrame, deposits: pd.DataFrame, benchmark: pd.DataFrame, start_row: int, year: str
):
    """
    Capture jaaroverzicht in image and write to Excel.

    :param ws: xlsxwriter Worksheet object.
    :param totals_waarde: Dataframe with totals info.

    :param benchmark: Dataframe with benchmark pricing data.
    :param start_row: Integer to start writing at.
    :param year: String with current year of writing.
    """

    totals_waarde = simulate_benchmark_values(totals_waarde, deposits, benchmark)

    plt.plot(totals_waarde['Portefeuille'], color=header_color, label='Portefeuille')
    plt.plot(totals_waarde['Inleg'], 'k', label='Inleg')
    plt.plot(totals_waarde['SP500'], 'b', label='SP500')
    plt.plot(totals_waarde['MSCI_World'], 'c', label='MSCI_World')
    plt.plot(totals_waarde['AEX'], 'm', label='AEX')
    totals_waarde['Winst/Verlies'].plot(kind='bar', color=positive_color, label='Winst/Verlies')

    ax = plt.gca()
    ax.set_xticklabels(totals_waarde['index'])
    for p in ax.patches:
        if p.get_height() > 0:
            ax.annotate(
                str(int(round(p.get_height(), 0))),
                (p.get_x() * (1 + 0.003*(12-totals_waarde.shape[0])), max(p.get_height() * 1.007, 20))
            )
        else:
            ax.annotate('(' + str(int(round(abs(p.get_height()), 0))) + ')', (p.get_x() * 0.999, 20))

    plt.xticks(rotation=45)
    plt.title('Portefeuille ontwikkeling ' + year)
    plt.legend()
    plt.grid(True)

    if year == '':
        figure = plt.gcf()  # get current figure
        figure.set_size_inches(len(totals_waarde)/3, 6)

    plt.savefig('portefeuille_ontwikkeling_' + year + '.png', bbox_inches='tight', dpi=100)
    ws.insert_image(start_row + 2, 5 if year != '' else 1, 'portefeuille_ontwikkeling_' + year + '.png')
    plt.close()


def write_portefeuille(
        portefeuille_dict: dict, totals_dict: dict, winstverlies_dict: dict, deposits: pd.DataFrame,
        benchmark: pd.DataFrame, wb, writer
):
    """
    Write portefeuille info to Excel.

    :param portefeuille_dict: Dict with portefeuille info.
    :param totals_dict: Dict with totals info.
    :param winstverlies_dict: Dict with winstverlies.
    :param deposits: Dataframe with deposits made.
    :param benchmark: Dataframe with benchmark pricing data.
    :param wb: xlsxwriter Workbook object.
    :param writer: xlsxwriter Writer object.
    :return writer: xlsxwriter Writer object.
    """

    port_prev = None
    total_invested = 0
    totals_waarde_full = []
    for key in portefeuille_dict:
        portefeuille = portefeuille_dict[key]
        totals = totals_dict[key]
        winstverlies = winstverlies_dict[key]

        ws = wb.get_worksheet_by_name(key)
        writer.sheets[key] = ws

        start_row = 1

        # format sheet
        # get the length of the portefeuille that will be written to the sheet
        len_port = sum((portefeuille.drop(['Product', 'Symbool/ISIN'], axis=1) != 0).any(axis=1))
        format_portefeuille(wb, ws, portefeuille, start_row, port_prev)
        start_row += len_port + 2

        # write totals, winstverlies
        format_totals(wb, ws, totals, start_row)
        start_row += len(totals.index) + 1
        format_winstverlies(wb, ws, winstverlies, start_row)
        start_row += 4

        # format jaaroverzicht. keep the total invest amount to take to the next year total inleg
        total_invested, totals_waarde_full = format_jaaroverzicht(
            wb, ws, totals, start_row, key, total_invested, totals_waarde_full, deposits, benchmark
        )
        format_header(wb, ws, key)
        set_cell_widths(ws, len(portefeuille.columns))
        ws.freeze_panes(2, 1)

        port_prev = portefeuille

    return writer, totals_waarde_full


def plot_total_dividend(totals: pd.DataFrame):
    """
    Create a stacked barplot with quarterly dividends.

    :param totals: dataframe with dividends.
    """

    x = np.arange(0, len(totals))
    fig, ax = plt.subplots()

    incr = list(np.linspace(start=-0.2, stop=0.2, num=len(totals.columns)))
    color = ['#1D2F6F', '#FAC748', '#6EAF46', '#8390FA']
    i = 0

    if len(totals.columns):
        NotImplementedError('More than four different currencies dividend is currently not supported.')

    for col in totals.columns:
        plt.bar(x+incr[i], totals[col], width=0.3, color=color[i])
        i += 1

    # remove spines
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.tick_params(axis='x', labelrotation=90)

    # x y details
    plt.ylabel('Dividend')
    plt.xticks(x, totals.index)

    plt.title('Dividend overview')
    plt.legend(totals.columns, loc='upper left')
    plt.savefig('dividend_ontwikkeling.png', bbox_inches='tight', dpi=100)
    plt.close()


def write_dividend_overview(writer: pd.ExcelWriter, wb) -> pd.ExcelWriter:
    """
    Write dividend overview to separate sheets. This creates two tables and a stacked barplot.

    :param writer: Exelwriter object.
    :param wb: Excelwriter workbook.
    :return overview of deposits over time.
    """
    sheet_name = 'Dividends'

    # read dividends
    dividends = read_dividends()

    # get totals
    total_overview = dividends.groupby(['Quarter', 'Mutatie']).sum(['Dividend']).reset_index()
    wide_totals = total_overview.pivot(index='Quarter', columns='Mutatie')
    wide_totals.columns = [col[-1] for col in wide_totals.columns.values]
    plot_total_dividend(wide_totals)
    wide_totals['zzzTotal'] = wide_totals.sum(axis=1)  # zzz added for sorting later
    wide_totals = wide_totals.reset_index()

    # agg to year view
    wide_totals['Year'] = wide_totals['Quarter'].astype(str).str[:4].astype(int)
    wide_totals['Quarter'] = wide_totals['Quarter'].astype(str).str[4:]
    year_view = wide_totals.pivot(index='Year', columns='Quarter', values=list(wide_totals.columns[~wide_totals.columns.isin(['Year', 'Quarter'])]))
    year_view = year_view.swaplevel(axis=1).sort_index(axis=1)
    year_view['FY', 'Total'] = list(wide_totals.groupby('Year')['zzzTotal'].sum())
    year_view = year_view.rename(columns={'zzzTotal': 'Total'})
    year_view.to_excel(writer, sheet_name=sheet_name)

    div_per_company = dividends.groupby('Product')['Dividend'].sum().sort_values(ascending=False)
    div_per_company.to_excel(writer, startrow=len(year_view)+6, sheet_name=sheet_name)

    ws = wb.get_worksheet_by_name(sheet_name)
    ws.insert_image(len(year_view)+6, 5, 'dividend_ontwikkeling.png')
    ws.set_column(0, 0, 30)
    ws.set_column(1, len(year_view.columns), 10)

    return writer


def write_costs_overview(writer: pd.ExcelWriter, wb) -> pd.ExcelWriter:
    """
    Write cost overview to separate sheet. This creates a table.

    :param writer: Exelwriter object.
    :param wb: Excelwriter workbook.
    :return overview of costs over time.
    """
    sheet_name = 'Costs'

    # read dividends
    costs = read_costs()

    # get totals
    total_overview = costs.groupby(['Jaar', 'Omschrijving']).sum(['Kosten']).reset_index()
    total_overview['Kosten'] = abs(total_overview['Kosten'])
    wide_totals = total_overview.pivot(index='Jaar', columns='Omschrijving')
    wide_totals.columns = [col[-1] for col in wide_totals.columns.values]
    wide_totals['Total'] = wide_totals.sum(axis=1)
    wide_totals = wide_totals.reset_index()
    wide_totals.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = wb.get_worksheet_by_name(sheet_name)
    ws.set_column(1, len(wide_totals.columns) - 1, 20)

    return writer


def write_returns_overview(
        totals_dict: dict, totals_waarde_full: list, deposits: pd.DataFrame, benchmark: pd.DataFrame,
        writer: pd.ExcelWriter, wb
) -> pd.ExcelWriter:
    """
    Write yearly returns overview to separate sheet. This creates a table. The yearly return is calculated by
    calculating a weighted 'inleg' (perc of the year the money could generate) returns multiplied by the amount.

    :param totals_dict: dictionary with totals per year to calculate returns.
    :param totals_waarde_full: list with portfolio value over time.
    :param deposits: Dataframe with deposits made.
    :param benchmark: Dataframe with benchmark pricing data.
    :param writer: Exelwriter object.
    :param wb: Excelwriter workbook.
    :return overview of costs over time.
    """

    sheet_name = 'Yearly returns'

    returns_overview = pd.DataFrame()
    for key, value in totals_dict.items():
        if key == list(totals_dict.keys())[-1]:
            # last year (not complete) so pass
            continue

        year = int(key)

        df_waarde = value[[x for x in value.columns if '(waarde)' in x]]
        df_waarde.columns = [x.replace(' (waarde)', '') for x in df_waarde.columns]
        value_start = round(df_waarde.loc['Totaal portefeuille'][0], 2) if key != list(totals_dict.keys())[0] else 0
        value_end = round(df_waarde.loc['Totaal portefeuille'][-1], 2)
        total_returns = value_end-value_start-df_waarde.loc['Inleg'].sum()

        # for weight deposit, we count the inleg * the percentage of the year that that amount could generate returns
        # the amount at the start of the year will be fully accounted for (i.e. weight 1)
        weighted_deposit = round(df_waarde.loc['Totaal portefeuille'][0], 2)
        ny = pd.to_datetime(str(year+1))
        for col in df_waarde.columns:
            perc_year = (ny-pd.to_datetime(col)).days/(365+calendar.isleap(year))
            weighted_deposit += perc_year*df_waarde[col]['Inleg']

        returns_overview = returns_overview.append(pd.DataFrame(data={
            'Jaar': [year],
            'Start bedrag': [value_start],
            'Eind bedrag': [value_end],
            'Winst/verlies': [total_returns],
            'Gewogen rendement': [round(100*(total_returns/weighted_deposit), 2)]
        }))

    returns_overview['Winst/verliest (cumulatief)'] = returns_overview['Winst/verlies'].cumsum()
    returns_overview.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = wb.get_worksheet_by_name(sheet_name)
    ws.set_column(1, len(returns_overview.columns), 20)

    totals_waarde_full = pd.concat(totals_waarde_full).reset_index(drop=True)
    add_jaaroverzicht_plot(ws, totals_waarde_full, deposits, benchmark, start_row=len(returns_overview)+3, year='')

    return writer
