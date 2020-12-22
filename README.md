# Waarde portefeuille

This code can be used to keep an overview of the value of your portefeuille over time. It is designed to work with DEGIRO exports.

# Setup and use

Clone the repository. Each time you want to add a new export to the output, follows these steps:
1. Go to the `Portefeuille` page in your account and click on `export`. Set the desired date to add to the output and export as csv.
2. Move the export to the folder `data/exports` in this repo and name it in the format `<%Y-%m-%d>.csv`, e.g. `2020-12-01.csv`.
3. Go to the `Overzichten-Rekeningoverzicht` page in your account and click on `export`. Set the start date before creating your account and the end date after the latest date of your exports.
4. Move the export to the folder `data/deposits` in this repo and name it `Account.csv`
5. Run `main.py`.
6. Find your export in the `results` folder.

Note, the contents of the `data` folder and `results` folder are in the gitignore are will not be pushed to github.