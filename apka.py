import pandas as pd
import os
import shutil
from docx import Document
from python_docx_replace import docx_replace

def excel_to_list_of_dictionaries(file_path):
    df = pd.read_excel(file_path)
    list_of_dicts = df.to_dict(orient='records')
    for dictionary in list_of_dicts:
        for key, value in dictionary.items():
            dictionary[str(key)] = str(value)
    return list_of_dicts


def process_documents_in_folder(folder_path, list_of_dicts):
    for index, data_dict in enumerate(list_of_dicts, start=1):
        new_doc_name = f"{index}.docx"
        original_doc_path = os.path.join(folder_path, "<base_document>.docx")
        new_doc_path = os.path.join(folder_path, new_doc_name)
        os.makedirs(folder_path, exist_ok=True)
        if not os.path.exists(new_doc_path):
            shutil.copy(original_doc_path, new_doc_path)
        doc = Document(new_doc_path)    
        docx_replace(doc, data_dict)
        doc.save(new_doc_path)

def change_keys_in_list_of_dicts(list_of_dicts, keys_to_replace):
    for data_dict in list_of_dicts:
        for old_key, new_key in keys_to_replace.items():
            if old_key in data_dict:
                data_dict[new_key] = data_dict.pop(old_key)

if __name__=='__main__':
    file_path_excel = './<excel_file_path>'
    original_doc_path = './<base_ducument_path>'
    target_folder = './<done_documents_path>'
    data_as_list_of_dicts = excel_to_list_of_dictionaries(file_path_excel)
    keys_to_replace = {'Imie i Nazwisko': '[-imie]','Data':'[-data]', 'Pan/Pani':'[-propl]','Adres':'[-adres]',
                       'PESEL':'[-pesel]'} 
    change_keys_in_list_of_dicts(data_as_list_of_dicts, keys_to_replace)
    print(data_as_list_of_dicts)
    process_documents_in_folder(original_doc_path,data_as_list_of_dicts)
    