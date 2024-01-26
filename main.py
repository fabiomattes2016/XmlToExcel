import xmltodict
import os
import pandas as pd


def pegar_infos(nome_arquivo, lista_valores):
    # print(f"Pegou informações do arquivo {nome_arquivo}")

    with open(f"nfs/{arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if "NFe" in dic_arquivo:
            infos_nfe = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nfe = dic_arquivo["nfeProc"]["NFe"]["infNFe"]

        numero_nfe = infos_nfe["@Id"]
        empresa_emissora = infos_nfe["emit"]["xNome"]
        nome_cliente = infos_nfe["dest"]["xNome"]
        rua = infos_nfe["dest"]["enderDest"]["xLgr"]
        numero = infos_nfe["dest"]["enderDest"]["nro"]
        municipio = infos_nfe["dest"]["enderDest"]["xMun"]
        endereco_completo = f"{rua}, {numero} - {municipio}"

        if "vol" in infos_nfe["transp"]:
            peso_bruto = infos_nfe["transp"]["vol"]["pesoB"]
        else:
            peso_bruto = "SEM PESO INFORMADO"

        if float(peso_bruto) <= 0:
            peso_bruto = "SEM PESO INFORMADO"

        lista_valores.append([numero_nfe, empresa_emissora, nome_cliente, endereco_completo, peso_bruto])


lista_arquivos = os.listdir("nfs")
colunas = ["NFe Nº ", "Emissor", "Cliente", "Endereço Completo", "Peso"]
valores = []

for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

try:
    tabela = pd.DataFrame(columns=colunas, data=valores)
    tabela.to_excel("output/NotasFicais.xlsx", index=False)
    print(tabela.head())
except Exception as e:
    print("Ocorreu um erro ao gerar a planilha", e)