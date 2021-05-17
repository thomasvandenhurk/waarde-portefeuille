import json
import os

import pandas as pd
import requests
from progressbar import progressbar


def get_stock_get_detail(performance_id: str):
    """
    Retrieve information from stock/get-detail.

    :param performance_id: String with performance ID.
    :return res: Dict with results.
    """

    url = "https://morning-star.p.rapidapi.com/stock/get-detail"
    querystring = {"PerformanceId": performance_id}

    headers = {
        'x-rapidapi-key': os.environ.get('x-rapidapi-key'),
        'x-rapidapi-host': "morning-star.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)
    response = json.loads(response.text)[0]

    # to fix 0 dividend
    try:
        forward_dividend_yield = response['Detail']['ForwardDividendYield']
    except KeyError:
        forward_dividend_yield = 0

    # to fix ETFs
    try:
        sector = response['Sector']
        industry = response['Industry']
        eps = response['Detail']['EarningsPerShare']['TrailingTwelveMonths']
    except KeyError:
        sector = None
        industry = None
        eps = None

    res = {
        'Currency': response['Currency'],
        'Exchange': response['Exchange'],
        'TypeName': response['TypeName'],
        'Sector': sector,
        'Industry': industry,
        'ForwardDividendYield (%)': forward_dividend_yield,
        'EPS (TTM)': eps,
        'RegionAndTicker': response['RegionAndTicker']
    }

    return res


def get_stock_v2_get_competitors(performance_id: str):
    """
    Retrieve information from stock/v2/get-competitors.

    :param performance_id: String with performance ID.
    :return res: Dict with results.
    """

    url = "https://morning-star.p.rapidapi.com/stock/v2/get-competitors"
    querystring = {"performanceId": performance_id}

    headers = {
        'x-rapidapi-key': os.environ.get('x-rapidapi-key'),
        'x-rapidapi-host': "morning-star.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)
    response = json.loads(response.text)

    res = {
        'PriceEarnings': response['main']['priceEarnings'],
        'DividendYield (%)': response['main']['dividendYield'],
        'Competitors': ", ".join([comp['name'] for comp in response['competitors']])
    }

    return res


def get_stock_v2_get_key_stats(performance_id: str):
    """
    Retrieve information from stock/v2/get-key-stats.

    :param performance_id: String with performance ID.
    :return res: Dict with results.
    """

    url = "https://morning-star.p.rapidapi.com/stock/v2/get-key-stats"
    querystring = {"performanceId": performance_id}

    headers = {
        'x-rapidapi-key': os.environ.get('x-rapidapi-key'),
        'x-rapidapi-host': "morning-star.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)
    response = json.loads(response.text)

    res = {
        'Revenue3YearGrowth': response['revenue3YearGrowth']['stockValue'],
        'netIncome3YearGrowth': response['netIncome3YearGrowth']['stockValue'],
        'operatingMarginTTM': response['operatingMarginTTM']['stockValue'],
        'netMarginTTM': response['netMarginTTM']['stockValue'],
        'roeTTM': response['roeTTM']['stockValue'],
        'debitToEquity': response['debitToEquity']['stockValue']
    }

    return res


def get_stock_v2_get_realtime_data(performance_id: str):
    """
    Retrieve information from stock/v2/get-realtime-data.

    :param performance_id: String with performance ID.
    :return res: Dict with results.
    """

    url = "https://morning-star.p.rapidapi.com/stock/v2/get-realtime-data"
    querystring = {"performanceId": performance_id}

    headers = {
        'x-rapidapi-key': os.environ.get('x-rapidapi-key'),
        'x-rapidapi-host': "morning-star.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)
    try:
        response = json.loads(response.text)
    except:
        print(response.text)

    res = {
        'LastPrice': response['lastPrice'],
        'DividendYield (%)': response['dividendYield'],
        'YearRangeHigh': response['yearRangeHigh'],
        'YearRangeLow': response['yearRangeLow'],
        'Type': response['type']
    }

    return res


def create_stock_overview(writer: pd.ExcelWriter, stocks_dict: dict, portefeuille_dict: dict):
    """
    Create sheet with stock overview with stock fundamentals.

    :param writer: ExcelWriter object to write the sheet to.
    :param stocks_dict: Dict with the performanceIDs to retrieve information from.
    :param portefeuille_dict: Dict with portefeuille.
    :return res: Dict with results.
    """

    sheet_name = 'Stock details'

    # get current holdings
    last_year = portefeuille_dict[max(portefeuille_dict.keys())]
    all_holdings = last_year[last_year.iloc[:, -3] > 0]['Product']

    if all(key in stocks_dict for key in set(all_holdings)):
        res = []
        for key, value in progressbar(stocks_dict.items()):
            if key in set(all_holdings) and value is not None:
                # get all dict results
                tmp = {'Holding': key}
                tmp.update(get_stock_get_detail(value))
                tmp.update(get_stock_v2_get_competitors(value))
                tmp.update(get_stock_v2_get_key_stats(value))
                tmp.update(get_stock_v2_get_realtime_data(value))
                res.append(tmp)

        # get dict to DF and write to Excel
        res = pd.DataFrame.from_dict(res)
        res.to_excel(writer, sheet_name, index=False)
    else:
        # some stocks info is missing, print which
        missings = [holding for holding in all_holdings if holding not in stocks_dict]

        wb = writer.book
        ws = wb.add_worksheet(sheet_name)
        ws.write(0, 0, 'Stocks are missing, add performanceIDs in portefeuille_dict.py or set to None')
        ws.write(1, 0, 'Missing keys: ' + ", ".join(missings))

    return writer
