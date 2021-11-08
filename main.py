import openpyxl
import os
from pathlib import Path
from datetime import datetime

def get_template(path) -> str:
    """
    Parameters
    ----------
    path : str
        The file path of the template
    
    Returns
    -------
    result : str
        The textual content of the template file
    """
    # Apro il file in lettura e restituisco l'intero contenuto testuale
    with open(path, "r") as f:
        return f.read()

def write_sql_file(entity
                    , excel_file_path_in
                    , excel_sheet_title
                    , sql_file_path_out                 
                    , header_template
                    , block_template                
                    , tail_template
                    ):
    """
    Write the Sql file based on the input parameters.

    Parameters
    ----------
    entity : 'D' or 'U'
        'D' -> domains, 'U' -> users
    excel_file_path_in : 
        The input Excel file path
    excel_sheet_title : 
        The Excel sheet title where the input data are
    sql_file_path_out : 
        The output Sql file path
    header_template : 
        The header's raw textual content of the output Sql file
    block_template : 
        The body block raw textual content of the output Sql file
    tail_template : 
        The tail's raw textual content of the output Sql file
    """
    # Apro il file Excel
    wb = openpyxl.load_workbook(filename=excel_file_path_in)

    # Seleziono il foglio di input
    ws = wb[excel_sheet_title]
    
    # Apro il file Sql in scrittura
    with open(sql_file_path_out, "w") as f:
        # Scrivo la parte fissa della testa
        f.write(header_template) 

        # Ciclo tra le righe del foglio Excel per poter creare 
        # le varie insert associate a ciascuna entità in esse 
        # presenti (min_row=2 --> escludo la prima riga delle intestazioni)
        for row in ws.iter_rows(min_row=2):
            if entity == "D":
                domain_id = row[0].value
                
                if domain_id == None:
                    continue
                
                domain_desc = row[1].value
                
                current_block = block_template.format(
                                        id_dominio=domain_id
                                        , desc_dominio=domain_desc
                                        , desc_dominio_2=domain_desc.replace("'", "''")
                                        ) 
            elif entity == "U":
                fiscal_code = row[0].value
                
                if fiscal_code == None:
                    continue

                domain_id = row[1].value
                role_id = row[2].value
                tmstmp = str(row[3].value)
                end_date = tmstmp[8:10] + "/" + tmstmp[5:7] + "/" + tmstmp[0:4]
                start_date = "10-11-2021"
                notes = row[4].value
                
                current_block = block_template.format(
                                        cod_fis=fiscal_code
                                        , id_dominio=domain_id
                                        , id_ruolo=role_id
                                        , data_fine=end_date
                                        , data_inizio=start_date
                                        , note=notes
                                        ) 
            else:
                current_block =  "********************************************"
                current_block += "******        Unknown entity!       ********"
                current_block += "******   Know entities: 'D' or 'U'  ********"
                current_block += "********************************************" 
                break

            # Scrivo un tassello del blocco centrale
            f.write(current_block)

        # Scrivo la parte fissa della coda
        f.write(tail_template)  

def run():
    """
    Manage the writing of the sql files: add_domain_AUTORITA_PL.sql, add_utenti_autorita.sql
    """
    # Ricavo path assoluto directory dello script e della sottocartella
    # contenente i file
    cur_path = Path(os.path.realpath(__file__))
    parent_dir = cur_path.parent.absolute()
    autorita_files_path = str(parent_dir) + '/autorita_files/'

    # la coda è identica per entrambi gli script sql di output
    tail_template_path = autorita_files_path + 'in/sql_templates/tail.sql'
    tail_template = get_template(tail_template_path)

    # Scrittura file Sql Domini
    entity = "D" # Domini
    excel_file_path_in = autorita_files_path + 'in/excel/Domini autorità.xlsx'
    excel_sheet_title = "Domini autorità"
    sql_file_path_out = autorita_files_path + 'out/add_domain_AUTORITA_PL.sql'
    header_template_path =  autorita_files_path + 'in/sql_templates/domains_header.sql'
    header_template = get_template(header_template_path)
    block_template_path =  autorita_files_path + 'in/sql_templates/domains_block.sql'
    block_template = get_template(block_template_path)

    write_sql_file(entity
                    , excel_file_path_in
                    , excel_sheet_title
                    , sql_file_path_out
                    , header_template
                    , block_template
                    , tail_template
                    )
    
    # Scrittura file Sql Utenti
    entity = "U" # Utenti
    excel_file_path_in = autorita_files_path + 'in/excel/CF autorità.xlsx'
    excel_sheet_title = "Autorità_primo inserimento"
    sql_file_path_out = autorita_files_path + 'out/add_utenti_autorita.sql'
    header_template_path =  autorita_files_path + 'in/sql_templates/users_header.sql'
    header_template = get_template(header_template_path)
    block_template_path =  autorita_files_path + 'in/sql_templates/users_block.sql'
    block_template = get_template(block_template_path)
    
    write_sql_file(entity
                , excel_file_path_in
                , excel_sheet_title
                , sql_file_path_out
                , header_template
                , block_template
                , tail_template
                )

if __name__ == "__main__":
    print("Inizio elaorazione:", datetime.today().strftime('%d-%m-%Y %H:%M'))
    run()
    print("Fine elaorazione:", datetime.today().strftime('%d-%m-%Y %H:%M'))