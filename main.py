import collections
import json
import os
import requests

import pandas as pd

from src.construct_portefeuille import construct_portefeuille
from src.write_output import write_portefeuille
from dotenv import load_dotenv
load_dotenv()


def main(output_path='results'):
    url = "https://morning-star.p.rapidapi.com/stock/v2/get-financials"

    querystring = {"performanceId": "0P000003RE", "interval": "annual", "reportType": "A"}

    headers = {
        'x-rapidapi-key': os.environ.get('x-rapidapi-key'),
        'x-rapidapi-host': "morning-star.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)

    response2 = json.loads(response.text)

    portefeuille_dict, totals_dict, winstverlies_dict = construct_portefeuille()

    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    wb = writer.book
    for key in sorted(portefeuille_dict.keys(), reverse=True):
        wb.add_worksheet(key)  # newer years first
    writer = write_portefeuille(portefeuille_dict, totals_dict, winstverlies_dict, wb, writer)
    writer.sheets = collections.OrderedDict(sorted(writer.sheets.items(), reverse=True))
    writer.save()


if __name__ == '__main__':
    main()
