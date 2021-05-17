# Waarde portefeuille

This code can be used to keep an overview of the value of your portefeuille over time. It is designed to work with DEGIRO exports.

# Setup and use

Clone the repository. Each time you want to add a new export to the output, follows these steps:
1. Go to the `Portefeuille` page in your account and click on `export`. Set the desired date to add to the output and export as csv.
2. Move the export to the folder `data/exports` in this repo and name it in the format `<%Y-%m-%d>.csv`, e.g. `2020-12-01.csv`.
3. Go to the `Overzichten-Rekeningoverzicht` page in your account and click on `export`. Set the start date before creating your account and the end date after the latest date of your exports.
4. Move the export to the folder `data/deposits` in this repo and name it `Account.csv`
5. If you want to obtain the stock fundamentals of the positions of your last export, following the steps in the `setup rapid-api` section. Otherwise, set the `use_rapid_api` to `False` in `main.py`.
6. Run `main.py`.
7. Find your export in the `results` folder.

Note, the contents of the `data` folder and `results` folder are in the gitignore are will not be pushed to github.

# Setup rapid-api
1. Get a rapid-api key via `https://rapidapi.com/apidojo/api/morning-star` and save in `.env`
2. Find the performanceIDs of your stocks:
    - Query stock on `https://www.morningstar.co.uk/uk/`. It works best to click a result in the dropdown
    - Validate that you have the right stock.
    - In the webpage address, look for the part `..&id=<FIELD_VALUE>&..`
    - Fill the `<FIELD_VALUE>` in `portefeuille_dict.py`.
3. Repeat the above steps for all stocks. 