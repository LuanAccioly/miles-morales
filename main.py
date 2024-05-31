import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import os
from tqdm import tqdm


def make_content_dir():
    content_path = "content/"
    filtered_path = "content/filtered/"
    if not os.path.exists(content_path):
        print("Criando diretório 'content'")
        os.makedirs(content_path)
        os.makedirs(filtered_path)
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

    if not individual_data_card:
        raise ValueError("Card 'Dados Individuais' não encontrado.")

    DOWNLOAD_URL = individual_data_card.find_element(By.TAG_NAME, "a").get_attribute(
        "href"
    )

    print("Baixando 'InfoMercado_Dados_Individuais.xlsx'")

    response = requests.get(DOWNLOAD_URL, stream=True)

    with open("content/InfoMercado_Dados_Individuais.xlsx", "wb") as file, tqdm(
        desc="Baixando",
        unit="B",
        unit_scale=True,
        unit_divisor=1024,
    ) as bar:
        for data in response.iter_content(chunk_size=1024):
            if data:
                file.write(data)
                bar.update(len(data))

    print("Arquivo salvo:\t" + "content/InfoMercado_Dados_Individuais.xlsx")


def extract_csv_files():
    tabs_info = [
        {
            "sheet_name": "007 Lista de Perfis",
            "table_numeration": "Tabela 001",
            "footer_row_value": "Topo",
            "output_name": "perfis.csv",
        },
        {
            "sheet_name": "001 Contratos",
            "table_numeration": "Tabela 001",
            "footer_row_value": "Tabela 002",
            "output_name": "contratos_venda.csv",
        },
        {
            "sheet_name": "001 Contratos",
            "table_numeration": "Tabela 002",
            "footer_row_value": "Topo",
            "output_name": "contratos_compra.csv",
        },
        {
            "sheet_name": "002 Usinas",
            "table_numeration": "Tabela 001",
            "footer_row_value": "Topo",
            "output_name": "usinas.csv",
        },
        {
            "sheet_name": "003 Consumo",
            "table_numeration": "Tabela 002",
            "footer_row_value": "Tabela 003",
            "output_name": "consumo.csv",
        },
    ]
    for tab in tabs_info:
        proccess_spreadsheet(
            tab["sheet_name"],
            tab["table_numeration"],
            tab["footer_row_value"],
            tab["output_name"],
        )


def proccess_spreadsheet(sheet_name, header_row_value, footer_row_value, output_name):

    print("\n\nExtraindo tabela:\t\t" + sheet_name)
    input_path = "content/InfoMercado_Dados_Individuais.xlsx"
    output_path = "content/" + output_name
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
        for value in row.values:
            if isinstance(value, str) and footer_row_value in value:
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

    # table = table.dropna(how="all")
    # table = table.dropna(axis=1, how="all")

    table = table.drop(table.columns[0], axis=1)
    table = table.drop(table.columns[header.size :], axis=1)

    table.to_csv(output_path, header=header, index=False)
    print("Tabela salva em:\t\t" + output_path)


def compare_with_profiles(csv_file):
    csv_name = csv_file.split("/")[-1]
    profiles_df = pd.read_csv("content/perfis.csv")
    valid_cod_profile = profiles_df["Cód. Perfil de Agente"]

    csv_df = pd.read_csv(csv_file)
    filtered_csv = csv_df[csv_df["Cód. Perfil"].isin(valid_cod_profile)]

    filtered_csv = filtered_csv.dropna(subset=["Cód. Perfil"])

    # Exibição dos perfis removidos ==========
    csv_df = csv_df.dropna(subset=["Cód. Perfil"])
    profile_id = csv_df["Cód. Perfil"]

    print("====================================")
    print("\n\nPerfis removidos em: " + csv_file)
    for id in profile_id:
        if id not in valid_cod_profile.values:
            print(int(id), end=" ")
    print("\n====================================\n\n")
    #  =======================================

    filtered_csv.to_csv(f"content/filtered/{csv_name}_filtrado.csv", index=False)


def filter_csv_files():
    csv_files = [
        "content/contratos_venda.csv",
        "content/contratos_compra.csv",
        "content/usinas.csv",
        "content/consumo.csv",
    ]

    for csv_file in csv_files:
        compare_with_profiles(csv_file)


if __name__ == "__main__":
    make_content_dir()
    get_main_spreadsheet()
    extract_csv_files()
    filter_csv_files()
