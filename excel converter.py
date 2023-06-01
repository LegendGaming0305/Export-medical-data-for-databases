import pandas as pd 
import sqlite3
import re


DB_PATH = 'path to db'
EXCEL_PATH = 'path to excel file'

# Create the table in SQL
def create_table():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS mnn_excel ("Trade_name_rus" TEXT, "Registrator_tran" TEXT, "Registrator_country" TEXT,
    "Producer_tran" TEXT, "Producer_country" TEXT, "Dosage_form_full_name" TEXT, "Dose" TEXT, "Sc_name" TEXT,
    "Recipe_status" TEXT, \"As_name_rus" TEXT, ID INT)''')
    conn.close()

# -------------------------------------------------------------------------
# Information input structure

# ВEntering numbers to database
def insert_id():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_id = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='A')
    for row in data_id.itertuples():
        cursor.execute("INSERT INTO mnn_excel (ID) VALUES (?)", (row[1],))
    conn.commit()
    conn.close()

# Information input about dosage of drug
def insert_dosage():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_dosa = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='B')
    for index, row in data_dosa.iterrows():
        drug_dose = []
        # Set a template for searching for the name of the drug.
        pattern = r"\d+(?:[.,]\d+)?\s*(?:мг|мл|мкг|анти-ХА МЕ|г)(?:/мл|/мг)?(?:\s*(?:/|\+)\s*\d+(?:[.,]\d+)?\s*(?:мг|мл|мкг|анти-ХА МЕ|г)(?:/мл|/мг)?)?"
        # Searching for matches to the template in the text.
        matches = re.findall(pattern, str(row[1]))
        # Return the first match found.
        if matches:
            drug_dose.append(matches[0])
        cursor.execute('UPDATE mnn_excel SET "Dose" = ? WHERE ID = ?', (', '.join(drug_dose) if drug_dose else None, index+1))
    conn.commit()
    conn.close()
    print('Дозировки внесены')

# Trade name of a medicine.
def insert_trade_name():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_torgname = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='B')
    # List to avoid some errors.
    unwanted_spisok_words = {' раствор для', ' р', ' лиофилизат для', ' порошок для', 'капсулы', 'комп', 'порош', ' пор'}
    pattern = r"[A-Za-zА-Яа-я®]+(?:-[A-Za-zА-Яа-я®]+)?(?: [A-Za-zА-Яа-я®]+(?: [A-Za-zА-Яа-я®]+)?)?"

    # Adding information to a table.
    for index, row in data_torgname.iterrows():
        drug_name = ''
        matches = re.findall(pattern, str(row[1]), re.I)
        if matches:
            drug_name += matches[0]
        for unwanted_words in unwanted_spisok_words:
            if unwanted_words in drug_name:
                drug_name = drug_name.replace(unwanted_words, '')
        if drug_name == 'Ульприксошок для':
            drug_name = 'Ульприкс'
        if drug_name == 'Дексаметазонаствор для':
            drug_name = 'Дексаметазон'
        if drug_name == 'Дифлюкан®аствор для':
            drug_name = 'Дифлюкан®'
        cursor.execute('UPDATE mnn_excel SET "Trade_name_rus" = ? WHERE ID = ?', (drug_name, index+1))
        drug_name = ''
    print('Названия препаратов внесены')
    conn.commit()
    conn.close()

# Drug administration form
def insert_form_of_usage():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_form = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='B')
    for index, row in data_form.iterrows():
        drug_form = ''
        pattern = r"(?:таблетки жевательные|таблетки пролонгированного действия|таблетки|табл.п.п.о.|табл.шип.|таб.п.о|табл.п.о.|табл.жев.|табл.ваг.|табл.п.кишечн.о.|табл\.|таб.п.пл.об.|таб.п.киш.раств.об.|таб\.|капсулы|сироп|гранулы для приготовления суспензии|суспензия|пор.д/приг.р-ра|пор.д/приг. р-ра|пор. д/ингал.доз.|раствор|гран.д/р-ра|гран.д/сусп.местн.|р-р|сусп.д/приема|сусп.д/пр. вн.|сусп.|аэроз.д/ингал.доз.|лиоф.порош.для|пор.лиоф.д/ин.амп.|пор.лиоф.д/ин.во фл.|комп. капс.|капс.|порошок для приготовления суспензии|порош.для|порошок для приготовления раствора|пор.для|таблеток|капли|таб|к/р|спрей сублингв.доз.|лиофилизат для пр. ра-ра для в/в с раст.|лиофилизат для приготовления раствора)[\s,]*(?:для приема внутрь и ингаляций|для приема внутрь|д/приема внутрь|для инъекций|для ингаляций|д/ин.амп.|д/ин.|внутрь|пр.р-ра для ин.с раств.|для внутримышечного введения|для внутривенного введения|пр.р-ра для ин.во фл.|пр.р-ра для ин.|пр.сусп.для пр.вн.|пр.р-ра для в/в и в/м введ.во фл.|д/в/в и в/м введ|для внутримышечных и внутривенных инъекций|для инфузий|д./инф.|местн.|к/р|для в/в введения|для в/в введ.)?[\s,]*(?:с модифицированным высвобождением|с пролонгированным высвобождением|с пролонгированным высвобождение|с модиф.высвоб.|с пролонг. выс.|диспергируемые)?[\s,]*(?:покрытые оболочкой|покрытые пленочной оболочкой|покрытые кишечнорастворимой оболочкой|обл\.|оболочка|обл\.|об\.|кишечнорастворимые|п.о.)?"        
        matches = re.findall(pattern, str(row[1]))
        if matches:
            drug_form += matches[0]
        cursor.execute('UPDATE mnn_excel SET "Dosage_form_full_name" = ? WHERE ID = ?', (drug_form if drug_form else None, index+1))
        drug_form = ''
    print('Формы применения внесены')
    conn.commit()
    conn.close()

# Active substance
def insert_mnn():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_mnn = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='C')
    for index, row in data_mnn.iterrows():
        cursor.execute('UPDATE mnn_excel SET "As_name_rus" = ? WHERE ID = ?', (row[2], index+1))
    conn.commit()
    conn.close()
    print('Мнн внесены')

# Name of registrator
def insert_registartor_tran():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_registartor = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='D')
    for index, row in data_registartor.iterrows():
        cursor.execute('UPDATE mnn_excel SET "Registrator_tran" = ? WHERE ID = ?', (row[3], index+1))
    conn.commit()
    conn.close()
    print('Фирма регистратора внесена')

# Manufacturer's name
def insert_producer_tran():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_producer = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='E')
    for index, row in data_producer.iterrows():
        cursor.execute('UPDATE mnn_excel SET "Producer_tran" = ? WHERE ID = ?', (row[4], index+1))
    conn.commit()
    conn.close()
    print('Название фирмы производителя внесена')

# Country of the manufacturer company.
def insert_producer_country():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    data_producer = pd.read_excel(EXCEL_PATH, header=None, skiprows=4, usecols='F')
    for index, row in data_producer.iterrows():
        cursor.execute('UPDATE mnn_excel SET "Producer_country" = ? WHERE ID = ?', (row[5], index+1))
    conn.commit()
    conn.close()
    print('Страна фирмы производителя внесена')

def main():
    create_table()
    insert_id()
    insert_dosage()
    insert_trade_name()
    insert_form_of_usage()
    insert_mnn()
    insert_registartor_tran()
    insert_producer_tran()
    insert_producer_country()
    print('Перенос закончен')

if __name__ == "__main__":
    main()