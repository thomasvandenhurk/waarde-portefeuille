import requests
import pandas as pd
import numpy as np
from functools import reduce
import os
from datetime import datetime as dt, timedelta


def get_data(performance_id):
    """
    Call the morningstar api for a specific entry.

    :param performance_id: string performance_id to extract the information for.
    :return response: response fromm the morningstar api with the information.
    """

    url = "https://morning-star.p.rapidapi.com/stock/get-histories"
    querystring = {"PerformanceId": f"{performance_id}"}
    headers = {
        "x-rapidapi-key": os.environ.get('x-rapidapi-key'),
        "x-rapidapi-host": "morning-star.p.rapidapi.com"
    }
    response = requests.get(url, headers=headers, params=querystring)
    return response


def extract_data(response):
    """
    Extract data from the response for the last 1Y and 5Y.

    :param response: response info to extract the information from.
    :return res: Dataframe with the price information for the index.
    """

    res = []
    for val in ['1Y', '5Y']:
        df = pd.DataFrame(response.json()[0][val])[['DateTime', 'Price']].rename(
            columns={'DateTime': 'date', 'Price': 'price'})
        df['date'] = df['date'].str[0:10]
        df['date'] = pd.to_datetime(df['date'])
        df = df.drop_duplicates(keep='last', subset='date')
        res.append(df)
    return res


def create_index_df():
    """
    Create index df, for three indices. The API is called and the file is constructed for the last 250 days.
    """

    performance_ids = {
        'SP500': '0P00013VX5',  # S&P500 Acc EUR
        'MSCI_World': '0P0000MLIH',  # MSCI Acc EUR
        'AEX': '0P0001KGO0',  # AEX Acc EUR
    }

    dates = pd.DataFrame(
        data={'date': pd.date_range(start=dt.now().date()-timedelta(weeks=262), end=dt.now().date(), freq='1D').date}
    )
    dates['date'] = pd.to_datetime(dates['date'])
    df_full = []

    for name, performance_id in performance_ids.items():
        response = get_data(performance_id=performance_id)
        dfs = extract_data(response=response)
        dates_full = pd.merge(dates, dfs[0], on='date', how='left')
        mask = dates_full['price'].isna()
        dates_full.loc[mask, ['price']] = pd.merge(dates[mask], dfs[1], on='date', how='left')['price']

        dates_full['price'] = np.interp(
            x=np.arange(len(dates_full)),         # The index positions
            xp=np.where(dates_full['price'].notna())[0],  # Index positions where values are not NaN
            fp=dates_full['price'].dropna()         # The actual (non-NaN) values at those positions
        )
        dates_full = dates_full.rename(columns={'price': f'price_{name}'})
        df_full.append(dates_full)

    merged_df = reduce(lambda left, right: pd.merge(left, right, on="date"), df_full)
    return merged_df


def update_benchmarks(benchmark_path):
    """
    Update/create benchmark file. If the file already exists, new entries will be appended to the existing file.

    :param benchmark_path: String path to where the benchmark file is stored.
    :return res: Dataframe with the full benchmark data.
    """

    res = create_index_df()
    if os.path.exists(benchmark_path):
        # if file already exists, take existing benchmark and append missing part
        benchmark = pd.read_csv(benchmark_path)
        benchmark['date'] = pd.to_datetime(benchmark['date'])
        res = pd.concat([benchmark, res[res['date'] > benchmark['date'].max()]]).reset_index(drop=True)

    res.to_csv(benchmark_path, index=False)
    return res


def load_benchmarks():
    """
    Load benchmarks. This functions will update (or create if not exists) the benchmark file.
    """

    benchmark_path = os.path.join('data', 'benchmark', 'benchmarks.csv')
    benchmark = update_benchmarks(benchmark_path)

    return benchmark
