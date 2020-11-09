from src.construct_portefeuille import construct_portefeuille
from src.write_output import write_portefeuille


def main():
    portefeuille, deposits = construct_portefeuille()
    write_portefeuille(portefeuille, deposits)


if __name__ == '__main__':
    main()
