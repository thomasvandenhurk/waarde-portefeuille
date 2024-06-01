import os
from functools import reduce
from os import listdir

import pandas as pd

from settings import rekeningoverzicht_filename


def read_costs() -> pd.DataFrame:
    """
    Read costs file in folder data/deposits.

    :return overview of costs over time.
    """
    costs = pd.read_csv(os.path.join('data', 'deposits', rekeningoverzicht_filename))
    costs = costs[costs['Omschrijving'].str.contains('kosten').fillna(False)].copy()
    costs = costs.rename(columns={'Unnamed: 8': 'Kosten'})
    #costs['Kosten'] = costs['Kosten'].str.replace(',', '.').astype(float)
    # change omschrijving
    costs['Omschrijving'] = costs['Omschrijving'].str.replace('DEGIRO transactiekosten', 'Transactiekosten')
    costs['Omschrijving'] = costs['Omschrijving'].str.replace('.*Aansluitingskosten.*', 'Aansluitingskosten')
    costs['Jaar'] = pd.DatetimeIndex(pd.to_datetime(costs['Datum'], format='%d-%m-%Y')).year
    costs = costs[['Jaar', 'Omschrijving', 'Kosten']]

    return costs


def read_dividends() -> pd.DataFrame:
    """
    Read dividends file in folder data/deposits.

    :return overview of dividends over time.
    """
    dividend = pd.read_csv(os.path.join('data', 'deposits', rekeningoverzicht_filename))
    dividend = dividend.loc[
        (dividend['Omschrijving'] == 'Dividend') | (dividend['Omschrijving'] == 'Dividendbelasting')
        ].copy()
    dividend = dividend.rename(columns={'Unnamed: 8': 'Dividend'})
    #dividend['Dividend'] = dividend['Dividend'].str.replace(',', '.').astype(float)
    dividend['Datum'] = pd.to_datetime(dividend['Datum'], format='%d-%m-%Y')
    dividend['Quarter'] = pd.PeriodIndex(dividend['Datum'], freq='Q')
    dividend = dividend[['Quarter', 'Product', 'Dividend', 'Mutatie']]

    return dividend


def read_deposits() -> pd.DataFrame:
    """
    Read deposit file in folder data/deposits.

    :return overview of deposits over time.
    """

    deposits = pd.read_csv(os.path.join('data', 'deposits', rekeningoverzicht_filename))
    deposits = deposits.loc[
        (deposits['Omschrijving'] == 'iDEAL Deposit') |
        (deposits['Omschrijving'] == 'iDEAL storting') |
        (deposits['Omschrijving'] == 'flatex terugstorting')
    ].copy()
    deposits = deposits.rename(columns={'Unnamed: 8': 'Storting'})
    deposits = deposits[['Datum', 'Storting']]
    #deposits['Storting'] = deposits['Storting'].str.replace(',', '.').astype(float)
    deposits['Datum'] = pd.to_datetime(deposits['Datum'], format='%d-%m-%Y')

    return deposits


def read_portefeuille(path_to_portefeuille_dir: str = os.path.join('data', 'exports'), suffix: str = '.csv') \
        -> pd.DataFrame:
    """
    Create a list of dataframes with portefeuille holdings.

    :param path_to_portefeuille_dir: path to data folder.
    :param suffix: suffix in which file names must end.
    :return portefeuille: overview of portefeuille combined.
    """

    # obtain all filesnames
    filenames = list_filenames(path_to_portefeuille_dir, suffix)

    # read and prep files
    portefeuille = []
    for file in filenames:
        date = file.replace('.csv', '')
        data_month = pd.read_csv(os.path.join(path_to_portefeuille_dir, file))
        data_month.drop(columns=['Slotkoers', 'Lokale waarde'], axis=1, inplace=True)
        data_month['Waarde in EUR'] = pd.to_numeric(data_month['Waarde in EUR'].str.replace(',', '.'))
        data_month = data_month.rename(columns={'Aantal': date + ' (aantal)', 'Waarde in EUR': date + ' (waarde)'})
        portefeuille.append(data_month)

    # get newest name on ISIN code
    product_isin = pd.concat(portefeuille)
    product_isin = product_isin[product_isin['Symbool/ISIN'].notnull()][['Product', 'Symbool/ISIN']]
    product_isin = product_isin.drop_duplicates(subset='Symbool/ISIN', keep='last')

    portefeuille_updated = []
    for port in portefeuille:
        merged_df = port.merge(product_isin, on='Symbool/ISIN', how='left', suffixes=('', '_new'))
        port['Product'] = merged_df['Product_new'].combine_first(merged_df['Product'])
        portefeuille_updated.append(port)

    # merge all files and set nan to zero
    portefeuille = reduce(lambda x, y: pd.merge(x, y, on=['Product', 'Symbool/ISIN'], how='outer'), portefeuille_updated)
    portefeuille = portefeuille.replace('CASH & CASH FUND & FTX CASH(EUR)', 'VRIJE RUIMTE')
    portefeuille.set_index(['Product', 'Symbool/ISIN'], inplace=True)
    portefeuille = portefeuille.fillna(0)
    return portefeuille


def list_filenames(path_to_portefeuille_dir: str, suffix: str) -> list:
    """
    List all file names in path_to_data_dir with a certain suffix.

    :param path_to_portefeuille_dir: path to data folder.
    :param suffix: suffix in which file names must end.
    :return: all file names existing in path_to_data_dir with suffix
    """

    filenames = listdir(path_to_portefeuille_dir)
    return [filename for filename in filenames if filename.endswith(suffix)]
