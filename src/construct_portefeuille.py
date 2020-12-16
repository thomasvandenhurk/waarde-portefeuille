import pandas as pd

from src.read_data import read_portefeuille, read_deposits


def add_percentages(portefeuille):
    # init first column
    col_date = portefeuille.columns[0].replace('(aantal)', '').replace('(waarde)', '')
    portefeuille.insert(2, col_date + '(procent)', 0)

    # insert remaining columns
    for i in range(5, int(1.5 * len(portefeuille.columns)), 3):
        col_date_prev = col_date
        col_date = portefeuille.columns[i - 1].replace('(aantal)', '').replace('(waarde)', '')

        # calculate percentage change per share (if now and prev date #shares > 0)
        values_old = portefeuille[col_date_prev + '(waarde)'] / portefeuille[col_date_prev + '(aantal)']
        values_new = portefeuille[col_date + '(waarde)'] / portefeuille[col_date + '(aantal)']
        percent_change = (values_new - values_old) / values_old
        percent_change = percent_change.fillna(0).round(2)

        # insert columns
        portefeuille.insert(i, col_date + '(procent)', percent_change)

    return portefeuille


def calculate_totals(portefeuille, deposits):
    # create total of each column
    sum_portefeuille = portefeuille.sum(axis=0)
    sum_portefeuille.loc[~sum_portefeuille.index.str.contains('(waarde)')] = 0

    # calculate difference with previous date
    # TODO look at userwarning
    difference = [sum_portefeuille.iloc[i] - sum_portefeuille.iloc[i - 3] for i in range(3, len(portefeuille.columns))]
    # add three leading zeros
    difference = [0, 0, 0] + difference

    # add deposits

    # combine info
    totals = pd.DataFrame([sum_portefeuille.values, difference], columns=portefeuille.columns,
                          index=['Totaal portefeuille', 'Verschil t.o.v. vorige maand'])

    return totals


def construct_portefeuille():
    portefeuille = read_portefeuille()
    portefeuille = add_percentages(portefeuille=portefeuille)
    deposits = read_deposits()
    totals = calculate_totals(portefeuille=portefeuille, deposits=deposits)

    return portefeuille, deposits
