import warnings
from typing import Tuple

import datetime as dt
import pandas as pd

from src.read_data import read_portefeuille, read_deposits


def add_percentages(portefeuille: pd.DataFrame) -> pd.DataFrame:
    """
    Add percentages to portefeuille dataframe.

    :param portefeuille: Dataframe with portefeuille data.
    :return portefeuille: Dataframe with percentage data appended.
    """

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


def calculate_totals(portefeuille: pd.DataFrame, deposits: pd.DataFrame) -> pd.DataFrame:
    """
    Create total overview of portefeuille. The following information is created:
        - Totaal portefeuille
        - Verschil t.o.v. vorige maand
        - Aankopen

    :param portefeuille: Dataframe with portefeuille data.
    :param deposits: Dataframe with deposits data.
    :return totals: Dataframe with the respective totals.
    """

    # create total of each column
    sum_portefeuille = portefeuille.sum(axis=0)
    sum_portefeuille.loc[~sum_portefeuille.index.str.contains('waarde')] = 0

    # calculate difference with previous date
    difference = [sum_portefeuille.iloc[i] - sum_portefeuille.iloc[i - 3] for i in range(3, len(portefeuille.columns))]
    difference = [0, sum_portefeuille.iloc[1], 0] + difference  # make first entry as totaal portefeuille

    # add deposits
    dates = list(set([x[:10] for x in sum_portefeuille.index.tolist()]))
    dates = [dt.datetime.strptime(date, '%Y-%m-%d') for date in dates]
    dates.sort()

    # match to nearest date in the future
    date_new = []
    for index, row in deposits.iterrows():
        diff = [(pd.Timestamp.to_pydatetime(row['Datum'])-date).days for date in dates]
        maxmin = [i for i in diff if i <= 0]
        try:
            minpos = diff.index(max(maxmin))
            date_new.append(dates[minpos])
        except ValueError:
            date_new.append(dates[-1])
            warnings.warn('A "Stortingsdatum is after the newest export. '
                          'It is set to equal the latest date of the export."', UserWarning)

    deposits['Datum'] = date_new
    deposits = deposits.groupby('Datum').sum().reset_index()

    # map to all dates
    deposits_full = pd.Series(0, index=sum_portefeuille.index)
    for index, row in deposits.iterrows():
        deposits_full.loc[deposits_full.index.str.contains(str(row['Datum'])[:10])] = row['Storting']
    deposits_full.loc[~deposits_full.index.str.contains('waarde')] = 0

    # combine info
    totals = pd.DataFrame([sum_portefeuille.values, difference, deposits_full.values],
                          columns=portefeuille.columns,
                          index=['Totaal portefeuille', 'Verschil t.o.v. vorige maand', 'Aankopen'])

    return totals


def construct_portefeuille() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Wrapper to construct portefeuille data. The portefeuille data is collected and a total overview is generated.

    :return portefeuille: Dataframe with percentage data appended.
    :return totals: Dataframe with the respective totals.
    """

    portefeuille = read_portefeuille()
    portefeuille = add_percentages(portefeuille=portefeuille)
    deposits = read_deposits()
    totals = calculate_totals(portefeuille=portefeuille, deposits=deposits)

    return portefeuille, totals
