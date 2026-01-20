import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import json
import os
import glob


help = """Help page of the Wordlist application: \n
To add another file path in the options, type "FILE" after they are shown.
To check the length of the working list, type "LENGTH" or "LEN" in the Word key.
To end word inputs, either press ENTER or type "DONE" or "BREAK" in the Word key.
To see this page once again while using the program, just type "HELP" in the Word key.
To get a JSON and a TXT file from the table, type "EXPORT" in the Word key.
To delete your last input, type "UNDO" in the Word key.
To replace your last input, type "REPLACELAST" in the Word key.
To delete a specific row, type "DELETE" in the Word key.
To replace a specific row, type "REPLACE" in the Word key.
To get a specific row, type "GETROW" in the Word key".
"""

directory = r"C:/Users/matti/Coding/PY/wordlist"
excel_files = []

excel_files = glob.glob(os.path.join(directory, "*.xlsx")) + glob.glob(os.path.join(directory, "*.xls"))

options = excel_files

def export(df, base_filename):
    try:
        # JSON
        json_filename = os.path.splitext(base_filename)[0] + '.json'
        
        df.to_json(json_filename, orient='records', indent=2, force_ascii=False)
        
        print(f"\nJSON exported: {json_filename}")

        # TXT
        txt_filename = os.path.splitext(base_filename)[0] + '.txt'
        
        with open(txt_filename, 'w', encoding='utf-8') as f:            
            for index, row in df.iterrows():
                f.write(f"{row['Word']} --> ")
                f.write(f"{row['Meaning']}\n")
        
        print(f"TXT exported: {txt_filename}")

        return True
    except Exception as e:
        print(f"Error exporting to JSON: {e}")
        return False
    
def create_file():
    filename = str(input('Select a name for your new table: '))
    
    df = pd.DataFrame(columns=['Word', 'Meaning', 'Example'])

    file_path = os.path.join(directory, filename + '.xlsx')
    df.to_excel(file_path, index=False)
    print("New Excel file: " + file_path)
    
    return file_path  # Return the file path

def delete(df, row_num):
    if row_num >= len(df) or row_num < 0:
        print(f'Invalid row number. Please provide a valid number or delete manually on the Excel table.\n')
        return df

    df = df.drop(df.index[row_num])
    df = df.reset_index(drop=True)

    print(f'Deleted row {row_num}\n')
    return df

def replace(df, row_num, word, meaning, example):
    row = df.iloc[row_num]
    if row_num >= len(df) or row_num < 0:
        print(f'Invalid row number. Please provide a valid number or delete manually on the Excel table.\n')
        return df

    if 'Word' in df.columns:
        pass
    else:
        print(f'Please provide a valid table.\n')
        return df

    if word != 'same':
        row['Word'] = word
    if meaning != 'same':
        row['Meaning'] = meaning
    if example != 'same':
        row['Example'] = '' if example == 'none' else example

    print(f'Updated row {row_num}.\n')
    return df

def save(df, path):
    df.to_excel(path, index=False)

def get_row(df, row_num):
    if 0 <= row_num < len(df):
        row = df.iloc[row_num]
        print(f"Word: {row['Word']}")
        print(f"Meaning: {row['Meaning']}")
        print(f"Example: {row['Example']}")
    else:
        print(f"Invalid row number. Table has {len(df)} rows.")

while True:
    welcome = input('Welcome to the Wordlist application. Please choose a file to work with from the options shown below. Type "HELP" for help page. Otherwise press enter. ')
    welcome = welcome.lower()
    print("\n")
    if welcome == 'help':
        print(help + "\n")
    else:
        for i in range(len(options)):
            print(i, " ", options[i])
        
        choice = input('Choose a file by its index. If you want to create a new one type "file": ')
        if choice == 'file':
            file = create_file()
            break 

        choice = int(choice)
        file = options[choice]
        print("\n")
        break

table = pd.read_excel(file)

element = {
    'Word' : '',
    'Meaning' : '',
    'Example' : ''
}

while True:
    word = str(input('Word: '))
    word = word.lower()

    if word == '' or word == 'done' or word == 'break':
        break

    elif word == 'help':
        print(help)

    elif word == 'export':
        export(table, file)
        break

    elif word == 'length' or word == 'len':
        print('This table has ' + str(len(table)) + ' items. \n')
        break

    elif word == 'delete':
        dinput = input('Please type the key (word) of the row you would like to delete: ')
        if dinput.isdigit():
            num = int(dinput)

            table = delete(table, num)
            break
        else:
            key = dinput
            key = key.lower()

            found = False
            for i in range(len(table)):
                if table.iloc[i, 0] == key:
                    found = True
                    num = i
                    break

            if found == False:
                print(f'No row with the key ' + key + ' has been found.\n')
                break
            else:
                table = delete(table, num)
                break

    elif word == 'undo':
        if len(table) <= 0:
            print(f'Table is empty.\n')
            break
        table = delete(table, len(table) - 1)
        print('Your last input has been deleted.\n')
        break

    elif word == 'replace':
        rinput = input('Please type the key (word) or the number of the row you would like to replace: ')
        rp_word = str(input('New word: '))
        rp_word = rp_word.lower()
        rp_meaning = str(input('New meaning: '))
        rp_meaning = rp_meaning.lower()
        rp_example= str(input('New example: '))
        rp_example = rp_example.lower()

        if rinput.isdigit():
            num = int(rinput)
            table = replace(table, num, rp_word, rp_meaning, rp_example)
            break
        else:
            key = rinput
            key = key.lower()
            found = False
            for i in range(len(table)):
                if table.iloc[i, 0] == key:
                    found = True
                    num = i
                    break

            if found == False:
                print(f'No row with the key ' + key + ' has been found.\n')
                break
            else:
                table = replace(table, num, rp_word, rp_meaning, rp_example)
                break

    elif word == 'replacelast':
        rp_word = str(input('New word: '))
        rp_word = rp_word.lower()
        rp_meaning = str(input('New meaning: '))
        rp_meaning = rp_meaning.lower()
        rp_example= str(input('New example: '))
        rp_example = rp_example.lower()

        table = replace(table, len(table) - 1, rp_word, rp_meaning, rp_example)
        break
    
    elif word == 'getrow':
        if len(table) == 0:
            print("Table is empty.\n")
            break
        num = int(input('Please select the row you would like to get: '))
        get_row(table, num)
        break

    meaning = str(input('Meaning: '))
    meaning = meaning.lower()

    example = str(input('Example: '))
    example = example.lower()

    print("\n")

    element['Word'] = word
    element['Meaning'] = meaning
    element['Example'] = example

    new_row = pd.DataFrame([element])
    table = pd.concat([table, new_row], ignore_index=True)

    element['Word'] = ''
    element['Meaning'] = ''
    element['Example'] = ''

    if len(table) > 50:
        print(f"List exceeds 50 items. We recommend you start a new one.\n")


save(table, file)
print("Excel file saved successfully!\n")