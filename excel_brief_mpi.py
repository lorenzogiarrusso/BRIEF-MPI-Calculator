import pandas as pd
import tkinter.filedialog
import argparse
import os

# Definisci i nomi delle colonne di interesse
ANAGRAPHIC_FIELDS = ['Codice Accesso', 'Codice Fiscale']
ADL_FIELDS = ['ADL Mangiare', 'ADL Vestirsi', 'ADL Controllo']
IADL_FIELDS = ['IADL Telefono', 'IADL Farmaci', 'IADL Acquisti']
BARTHEL_FIELDS = ['Barthel Alzarsi', 'Barthel Camminare', 'Barthel Scale']
SPMSQ_FIELDS = ['SPMSQ Data', 'SPMSQ Anni', 'SPMSQ Calcolo']
MNA_FIELDS = ['MNA BMI', 'MNA Perdita Appetito', 'MNA Perdita Peso']
COMORB_FIELD = 'Numero Comorb'
DRUGS_FIELD = 'Numero Farmaci'
COHABIT_FIELD = 'Stato Coabitativo'

def calc_adl_score(row):
    sum = row[ADL_FIELDS].sum()
    if sum == 3:
        return 0
    elif sum == 1 or sum == 2:
        return 0.5
    elif sum == 0: 
        return 1
    else:
        raise ValueError('Valore invalido nei campi ADL')

def calc_iadl_score(row):
    sum = row[IADL_FIELDS].sum()
    if sum == 3:
        return 0
    elif sum == 1 or sum == 2:
        return 0.5
    elif sum == 0: 
        return 1
    else:
        raise ValueError('Valore invalido nei campi IADL')

def calc_barthel_score(row):
    sum = row[BARTHEL_FIELDS].sum()
    if sum == 3 or sum == 2:
        return 0
    elif sum == 1:
        return 0.5
    elif sum == 0: 
        return 1
    else:
        raise ValueError('Valore invalido nei campi Barthel')
    
def calc_spmsq_score(row):
    sum = row[SPMSQ_FIELDS].sum()
    if sum == 0:
        return 0
    elif sum == 1:
        return 0.5
    elif sum == 2 or sum == 3: 
        return 1
    else:
        raise ValueError('Valore invalido nei campi SPMSQ')
    
def calc_mna_score(row):
    sum = row[MNA_FIELDS].sum()
    if sum == 0:
        return 0
    elif sum == 1:
        return 0.5
    elif sum == 2 or sum == 3: 
        return 1
    else:
        raise ValueError('Valore invalido nei campi MNA')

def calc_comorb_score(row):
    n = row[COMORB_FIELD]
    if n == 0:
        return 0
    elif n < 3:
        return 0.5
    elif n >= 3:
        return 1
    else:
        raise ValueError('Valore invalido nel campo Numero Comorb')
    
def calc_drugs_score(row):
    n = row[DRUGS_FIELD]
    if n >= 0 and n <= 3:
        return 0
    elif n < 7:
        return 0.5
    elif n >= 7:
        return 1
    else:
        raise ValueError('Valore invalido nel campo Numero Farmaci')
    
def calc_cohabit_score(row):
    value = row[COHABIT_FIELD]
    if value == 0: # Da solo/a
        return 1
    elif value == 1: # Con la famiglia
        return 0
    elif value == 2: # In istituto
        return 0.5
    else:
        raise ValueError('Valore invalido nel campo Stato Coabitativo')

def calc_brief_mpi(row):
    total_score = (
        calc_adl_score(row) + calc_iadl_score(row) + calc_barthel_score(row) +
        calc_spmsq_score(row) + calc_mna_score(row) + calc_comorb_score(row) +
        calc_drugs_score(row) + calc_cohabit_score(row)
    )
    num_scores = 8
    avg = total_score / num_scores
    return round(avg, 2)

def brief_mpi_to_risk(brief_mpi):
    if brief_mpi >= 0 and brief_mpi <= 0.33:
        return 'MPI 1'
    elif brief_mpi <= 0.66:
        return 'MPI 2'
    elif brief_mpi <= 1:
        return 'MPI 3'
    else:
        raise ValueError('Valore invalido per il punteggio BRIEF-MPI')
    
def save_as_xlsx():
    filename = tkinter.filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx')], defaultextension='.xlsx', title='Salva come...', initialfile='output.xlsx')
    return filename

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process an Excel file to calculate BRIEF-MPI scores.')
    parser.add_argument('file_path', nargs='?', default='data.xlsx', help='Path to the Excel file (default: data.xlsx)')
    args = parser.parse_args()

    file_path = args.file_path
    if not os.path.exists(file_path):
        print(f'File non trovato: {file_path}')
        parser.print_help()
        exit(1)
    df = pd.read_excel(file_path)

    # Seleziona solo le colonne di interesse
    columns_of_interest = ANAGRAPHIC_FIELDS + ADL_FIELDS + IADL_FIELDS + BARTHEL_FIELDS + SPMSQ_FIELDS + MNA_FIELDS + [COMORB_FIELD, DRUGS_FIELD, COHABIT_FIELD]

    # Verifica se il DataFrame contiene tutte le colonne di interesse
    missing_columns = [col for col in columns_of_interest if col not in df.columns]
    if missing_columns:
        print(f"Errore: il file '{file_path}' non contiene le seguenti colonne richieste: {', '.join(missing_columns)}")
        exit(1)

    df_selected = df[columns_of_interest]
    df_selected.dropna(inplace=True)

    df_selected['ADL Score'] = df_selected.apply(calc_adl_score, axis=1)
    df_selected['IADL Score'] = df_selected.apply(calc_iadl_score, axis=1)
    df_selected['Barthel Score'] = df_selected.apply(calc_barthel_score, axis=1)
    df_selected['SPMSQ Score'] = df_selected.apply(calc_spmsq_score, axis=1)
    df_selected['MNA Score'] = df_selected.apply(calc_mna_score, axis=1)
    df_selected['Comorb Score'] = df_selected.apply(calc_comorb_score, axis=1)
    df_selected['Drugs Score'] = df_selected.apply(calc_drugs_score, axis=1)
    df_selected['Cohabit Score'] = df_selected.apply(calc_cohabit_score, axis=1)
    df_selected['BRIEF-MPI'] = df_selected.apply(calc_brief_mpi, axis=1)
    df_selected['Risk'] = df_selected['BRIEF-MPI'].apply(brief_mpi_to_risk)

    output = df_selected[['Codice Accesso', 'Codice Fiscale', 'BRIEF-MPI', 'Risk']].rename(columns={'Risk': 'Rischio'})

    filename = save_as_xlsx()
    if filename:
        output.to_excel(filename, index=False)
        print(f'File salvato in {filename}')
    else:
        print('Operazione annullata')