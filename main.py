import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import os


def make_content_dir():
    path = "content/"
    if not os.path.exists(path):
        print("Criando diretório 'content'")
        os.makedirs(path)
    else:
        print("Diretório content/ já existe\n\n")


def get_main_spreadsheet():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)

    CCEE_URL = "https://www.ccee.org.br/dados-e-analises/dados-mercado-mensal"

    driver.get(CCEE_URL)

    cards = driver.find_elements(By.CLASS_NAME, "card")

    individual_data_card = None

    for card in cards:
        if "Dados Individuais" in card.text:
            individual_data_card = card
            break

    DOWNLOAD_URL = individual_data_card.find_element(By.TAG_NAME, "a").get_attribute(
        "href"
    )

    print("Baixando 'InfoMercado_Dados_Individuais.xlsx'")
    r = requests.get(DOWNLOAD_URL)

    open("content/InfoMercado_Dados_Individuais.xlsx", "wb").write(r.content)
    print("Arquivo salvo:\t" + "content/InfoMercado_Dados_Individuais.xlsx")


def proccess_spreadsheet(sheet_name, 
                         header_row_value, 
                         footer_row_value):
    
    print("Extraindo tabela:\t" + sheet_name)
    input_path = "content/InfoMercado_Dados_Individuais.xlsx"
    output_path = "content/" + sheet_name + ".csv"
    xls = pd.ExcelFile(input_path)

    df = pd.read_excel(xls, sheet_name, dtype=str)

    header_row = None
    for index, row in df.iterrows():
        for value in row.values:
            if isinstance(value, str) and header_row_value in value:
                header_row = index + 1
                break

    footer_row = None
    for index, row in df.iterrows():
        if footer_row_value in row.values:
            footer_row = index
            break

    if header_row and footer_row is not None:
        header = df.iloc[header_row]
        table = df.iloc[header_row + 1 : footer_row].reset_index(drop=True)
    else:
        print("Não foi possível encontrar a linha de início")
        return

    header = header.dropna(how="all")
    header = header.to_numpy()

    table = table.dropna(how="all")
    table = table.dropna(axis=1, how="all")

    table.to_csv(output_path, header=header, index=False)
    print("Tabela extraída e salva em \t" + output_path)

if __name__ == "__main__":
    make_content_dir()
    get_main_spreadsheet()
    proccess_spreadsheet(
        "007 Lista de Perfis",
        "Tabela 001",
        "Topo"
    )
