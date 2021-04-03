from src.construct_portefeuille import construct_portefeuille
from src.rapidapi_data import *
from src.write_output import write_portefeuille
from portefeuille_dict import stock_input
from dotenv import load_dotenv
load_dotenv()


def main(output_path='results'):
    portefeuille_dict, totals_dict, winstverlies_dict = construct_portefeuille()

    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    wb = writer.book
    for key in sorted(portefeuille_dict.keys(), reverse=True):
        wb.add_worksheet(key)  # newer years first
    writer = write_portefeuille(portefeuille_dict, totals_dict, winstverlies_dict, wb, writer)
    writer = create_stock_overview(writer, stock_input)
    writer.save()


if __name__ == '__main__':
    main()
