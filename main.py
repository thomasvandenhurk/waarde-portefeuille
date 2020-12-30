from src.construct_portefeuille import construct_portefeuille
from src.write_output import write_portefeuille


def main():
    portefeuille_dict, totals_dict, winstverlies_dict = construct_portefeuille()
    write_portefeuille(portefeuille_dict, totals_dict, winstverlies_dict)


if __name__ == '__main__':
    main()
