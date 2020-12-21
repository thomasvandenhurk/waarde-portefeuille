import os
from functools import reduce
from os import listdir

import pandas as pd


def read_deposits() -> pd.DataFrame:
    """
    Read deposit file in folder data/deposits.

    :return overview of deposits over time.
    """

    deposits = pd.read_csv(os.path.join('data', 'deposits', 'Account.csv'))
    deposits = deposits.loc[
        (deposits['Omschrijving'] == 'iDEAL Deposit') | (deposits['Omschrijving'] == 'iDEAL storting')
    ].copy()
    deposits = deposits.rename(columns={'Unnamed: 8': 'Storting'})
    deposits = deposits[['Datum', 'Storting']]
    deposits['Storting'] = deposits['Storting'].str.replace(',', '.').astype(float)
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

    # merge all files and set nan to zero
    portefeuille = reduce(lambda x, y: pd.merge(x, y, on=['Product', 'Symbool/ISIN'], how='outer'), portefeuille)
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
