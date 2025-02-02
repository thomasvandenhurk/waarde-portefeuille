import pandas as pd
from datetime import datetime as dt
import yfinance as yf


def load_benchmarks():
    """
    Load historic prices of a few ETFs.
    """

    ticker_symbols = {
        'CSPX.AS': 'price_SP500',  # S&P500 Acc EUR
        'IWDA.AS': 'price_MSCI_World',  # MSCI Acc EUR
        'IAEA.AS': 'price_AEX',  # AEX Acc EUR
    }
    history = yf.Tickers(' '.join(ticker_symbols.keys())).history(period='max')
    history = history.loc[:, history.columns.get_level_values(0) == 'Close']
    history.columns = history.columns.get_level_values(1)
    history = history.reset_index()
    history = history.rename(columns={'Date': 'date', **ticker_symbols})

    dates = pd.DataFrame(
        data={'date': pd.date_range(start=history['date'][0], end=dt.now().date(), freq='1D').date}
    )
    dates['date'] = pd.to_datetime(dates['date'])
    dates = dates.merge(history, on='date', how='left')
    dates = dates.fillna(method='bfill').fillna(method='ffill')

    return dates
