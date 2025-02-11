# BRIEF-MPI Calculator

This Python script processes an Excel file to calculate BRIEF-MPI scores based on various health and demographic fields. The resulting scores are saved in a new Excel file.

## Requirements

- Python 3.x
- pandas
- tkinter

You can install the required Python packages using pip:

```sh
pip install pandas
```

## Usage
To run the script, use the following command:

```sh
python excel_brief_mpi.py [file_path]
```

- file_path: Path to the Excel file containing the data. If not provided, the default is data.xlsx.

## Columns of Interest
The script expects the following columns in the Excel file:

- Anagraphic Fields: Codice Accesso, Codice Fiscale
- ADL Fields: ADL Mangiare, ADL Vestirsi, ADL Controllo
- IADL Fields: IADL Telefono, IADL Farmaci, IADL Acquisti
- Barthel Fields: Barthel Alzarsi, Barthel Camminare, Barthel Scale
- SPMSQ Fields: SPMSQ Data, SPMSQ Anni, SPMSQ Calcolo
- MNA Fields: MNA BMI, MNA Perdita Appetito, MNA Perdita Peso
- Comorbidity Field: Numero Comorb
- Drugs Field: Numero Farmaci
- Cohabitation Field: Stato Coabitativo


## Output
The script generates an output Excel file with the following columns:

- Codice Accesso
- Codice Fiscale
- BRIEF-MPI
- Rischio

## Example
To process an Excel file named data.xlsx, run:

```sh
python excel_brief_mpi.py data.xlsx
```

The script will prompt you to save the output file as an Excel file.

## References
This script was made referring to the BRIEF-MPI as described in https://doi.org/10.2147/CIA.S355801, with the tool available at https://multiplat-age.it/index.php/en/tools being used as a model for the calculation.
