import os

import pandas as pd


def write_portefeuille(portefeuille, deposits, output_path='results'):
    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    portefeuille.to_excel(writer, sheet_name='Portefeuille', index=True)

    writer.save()
