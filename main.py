import logging

from src.construct_portefeuille import construct_portefeuille
from src.rapidapi_data import *
from src.write_output import write_portefeuille, write_dividend_overview, write_costs_overview
from portefeuille_dict import stock_input
from dotenv import load_dotenv
load_dotenv()


def main(output_path='results', use_rapid_api=True):
    logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.INFO)
    logging.info('Start program')
    logging.info('Constructing portfolio from exports')
    portefeuille_dict, totals_dict, winstverlies_dict = construct_portefeuille()

    writer = pd.ExcelWriter(os.path.join(output_path, 'portefeuille.xlsx'), engine='xlsxwriter')
    wb = writer.book

    # add portefeuille overview per year
    for key in sorted(portefeuille_dict.keys(), reverse=True):
        wb.add_worksheet(key)  # newer years first
    logging.info('Writing portfolio overview to Excel')
    writer = write_portefeuille(portefeuille_dict, totals_dict, winstverlies_dict, wb, writer)

    if use_rapid_api:
        # create stock overview
        logging.info('Retrieve current portfolio holdings fundamentals')
        writer = create_stock_overview(writer, stock_input, portefeuille_dict)

    # add dividend overview
    logging.info('Writing dividend overview to Excel')
    writer = write_dividend_overview(writer, wb)

    # add costs overview
    logging.info('Writing costs overview to Excel')
    writer = write_costs_overview(writer, wb)

    logging.info('Save Excel file to ' + output_path + ' folder')
    writer.save()
    logging.info('Finished!')


if __name__ == '__main__':
    main()
