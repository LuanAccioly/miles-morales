import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import openpyxl
import os
from openpyxl import load_workbook


def make_content_dir():
    path = "content/"
    if not os.path.exists(path):
        print("Criando diretório content/")
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
    print("Arquivo salvo em content/InfoMercado_Dados_Individuais.xlsx")


def proccess_spreadsheet():
    spreadsheet_path = "content/InfoMercado_Dados_Individuais.xlsx"
    caminho_entrada = "content/InfoMercado_Dados_Individuais.xlsx"
    caminho_saida = "content/007.xlsx"
    nome_planilha = "007 perfis"

    xls = pd.ExcelFile(spreadsheet_path)

    df = pd.read_excel(xls, "007 Lista de Perfis", dtype=str)

    linha_inicio = None
    for index, row in df.iterrows():
        if "Cód. Agente" in row.values:
            linha_inicio = index
            break

    linha_final = None
    for index, row in df.iterrows():
        if "Topo" in row.values:
            linha_final = index
            break

    if linha_inicio is not None:
        header = df.iloc[linha_inicio]
        tabela = df.iloc[linha_inicio + 1 : linha_final].reset_index(drop=True)
    else:
        print("Não foi possível encontrar a linha de início")

    header = header.dropna(how="all")
    header = header.to_numpy()

    tabela = tabela.dropna(how="all")
    tabela = tabela.dropna(axis=1, how="all")

    tabela.to_csv("content/main_table.csv", header=header, index=False)


if __name__ == "__main__":
    make_content_dir()
    get_main_spreadsheet()
    proccess_spreadsheet()
