from src.read_data import read_portefeuille, read_deposits


def construct_portefeuille():
    portefeuille = read_portefeuille()
    deposits = read_deposits()

    return portefeuille, deposits
