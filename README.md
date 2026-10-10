# Alter-millet
A Python-based food substitution and nutrition recommendation tool that maps conventional dishes to millet alternatives and compares nutritional values using condition-specific rules.

The project focuses on the recommendation engine and data pipeline. A
graphical frontend is not yet implemented.

## What it does

-   **Finds mapped alternatives:** Uses a curated mapping between
    conventional dishes and millet-based recipes.
-   **Compares nutrition:** Compares available values per 100 g for
    energy, carbohydrates, protein, fat, and dietary fibre.
-   **Considers a selected health condition:** Applies configurable
    nutrient preferences for supported conditions to evaluate the
    nutritional differences.
-   **Explains the result:** Reports whether a mapped millet option is
    recommended under the current rules, whether the original may be
    preferable, or whether there is no clear gain.
-   **Handles limitations:** Flags cases where a replacement ingredient
    cannot be matched to the nutrition dataset or where the required
    nutrient information is unavailable.

## How it works

1.  Loads the nutrition CSV and parses numeric values, including ranges
    where present.
2.  Loads the millet recipe dataset and the curated mapping workbook.
3.  Finds candidate millet recipes associated with the selected
    conventional dish.
4.  Retrieves nutrition values for the marked replacement ingredient and
    the millet type.
5.  Compares the values using condition-specific nutrient preferences.
6.  Prints the candidate recipes, nutrient differences, and a verdict in
    the terminal.

The current implementation uses a **rule-based comparison**, not a
trained machine-learning model. Its results depend on the supplied
mappings, nutritional data, ingredient-name matching, and the encoded
rules.

## Data files

Place the required data files in the same directory as `millet_cli.py`.
The script searches for filenames matching these patterns:

-   A mapping workbook matching `*Mapping*Final*.xlsx` or
    `*Millet_Mapping*.xlsx`
-   A nutrition CSV matching `*Nutrition*100*.csv`, `*Nutrition*.csv`,
    or `*utrition*.csv`
-   A millet recipe workbook matching `*155*.xlsx` or `*Final*.xlsx`

For the files used during development, these correspond to:

-   `Millet_Mapping_Final (1).xlsx`
-   `Final Nutrition per 100g - Full Dataset.csv`
-   `155 Final(1).xlsx`

If you rename the files, make sure they still match the patterns above.
Keep the mapping workbook's `Top 3 Candidates` sheet and the expected
column names intact.

## Requirements

-   Python 3
-   pandas
-   NumPy
-   openpyxl

Install the dependencies:

``` bash
python -m pip install pandas numpy openpyxl
```

## Run the project

Run commands from the directory containing the script and data files.

### Interactive mode

``` bash
python millet_cli.py
```

The program lists supported health conditions and asks you to select a
condition and a dish.

### List available dishes

``` bash
python millet_cli.py --list-foods
```

### List supported conditions

``` bash
python millet_cli.py --list-diseases
```

### Query a dish and condition directly

``` bash
python millet_cli.py --food "idli" --disease "Type 2 Diabetes"
```

Use a dish name that appears in the mapping data. Partial dish-name
matching is supported.

## Supported conditions

The current rule set includes:

-   Type 2 Diabetes
-   High Cholesterol
-   Cardiovascular Disease
-   Hypertension
-   Chronic Kidney Disease
-   MASLD / Fatty Liver
-   Obesity / Overweight
-   Gout
-   Osteoporosis
-   Dental Caries

The current implementation marks Dental Caries as not computable because
the relevant sugar and calcium information is not present in the
nutrient comparison data. Other conditions use the nutrient preferences
encoded in the script.

## Current scope and limitations

-   This is a research/academic prototype; the command-line interface is
    implemented, while a graphical frontend remains future work.
-   Nutrition comparisons are based on ingredient values per 100 g. They
    should not be interpreted as complete, portion-adjusted nutrition
    estimates for an entire prepared recipe.
-   The condition-specific rules are simplified heuristics, not clinical
    guidelines or a validated medical decision system.
-   The current implementation uses ingredient-name similarity to look
    up nutrition values, so matches should be reviewed when accuracy is
    important.
-   A positive verdict means that an option scores better under the
    rules encoded in this prototype. It does **not** establish that the
    food is medically suitable for an individual.

## Disclaimer

This project is intended for educational and research purposes only. It
does not provide medical advice, diagnose conditions, or replace
guidance from a qualified healthcare professional. People with medical
conditions---especially those requiring dietary restrictions---should
consult a qualified clinician or registered dietitian before making
dietary changes.

## Future work

-   Build a simple frontend for selecting foods and conditions.
-   Improve ingredient normalization and nutrition-data matching.
-   Add portion-aware and full-recipe nutritional calculations.
-   Validate the scoring rules with domain experts and reliable dietary
    guidelines.
-   Add tests and representative example outputs.

## Project status

The core data integration and command-line recommendation logic are
under development. The frontend and further validation are planned next
steps.
