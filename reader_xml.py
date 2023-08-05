import os
import xmltodict
import pandas as pd

def ler_xml(nome_arquivo, valores):
    with open(f'notas/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        info_nfe = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        num_nfe = info_nfe["ide"]["nNF"]
        data_emissao = info_nfe["ide"]["dhEmi"]
        serie_nfe = info_nfe["ide"]["serie"]
        emissor_nfe = info_nfe["emit"]["xNome"]
        cnpj_emissor = info_nfe["emit"]["CNPJ"]
        dest_nfe = info_nfe["dest"]["xNome"]
        cnpj_dest = info_nfe["dest"].get("CNPJ") or info_nfe["dest"].get("CPF")  # Corrected line
        cep_dest = info_nfe["dest"]["enderDest"]["CEP"]
        rua_dest = info_nfe["dest"]["enderDest"]["xLgr"]
        num_casa_dest = info_nfe["dest"]["enderDest"]["nro"]
        bairro_dest = info_nfe["dest"]["enderDest"]["xBairro"]
        uf_dest = info_nfe["dest"]["enderDest"]["UF"]
        pais_dest = info_nfe["dest"]["enderDest"]["xPais"]

        if isinstance(info_nfe["det"], list):
            # If there are multiple <det> elements, iterate through them
            for det_item in info_nfe["det"]:
                cod_produto = det_item["prod"]["cProd"]
                nome_produto = det_item["prod"]["xProd"]
                quant_produto = det_item["prod"]["qCom"]
                valor_unit_produto = det_item["prod"]["vProd"]
                valores.append([num_nfe, serie_nfe, data_emissao, emissor_nfe, cnpj_emissor, dest_nfe, cnpj_dest, cep_dest, rua_dest, num_casa_dest, bairro_dest, uf_dest, pais_dest, cod_produto, nome_produto, quant_produto, valor_unit_produto])
        else:
            # If there's only one <det> element, directly read its attributes
            cod_produto = info_nfe["det"]["prod"]["cProd"]
            nome_produto = info_nfe["det"]["prod"]["xProd"]
            quant_produto = info_nfe["det"]["prod"]["qCom"]
            valor_unit_produto = info_nfe["det"]["prod"]["vProd"]
            valores.append([num_nfe, serie_nfe, data_emissao, emissor_nfe, cnpj_emissor, dest_nfe, cnpj_dest, cep_dest, rua_dest, num_casa_dest, bairro_dest, uf_dest, pais_dest, cod_produto, nome_produto, quant_produto, valor_unit_produto])

colunas = ["num_nfe", "serie_nfe", "data_emissao", "emissor_nfe", "cnpj_emissor", "dest_nfe", "cnpj_dest", "cep_dest", "rua_dest", "num_casa_dest", "bairro_dest", "uf_dest", "pais_dest", "cod_produto", "nome_produto", "quant_produto", "valor_unit_produto"]
valores = []
listar_arquivos = os.listdir("notas")

for arquivo in listar_arquivos:
    ler_xml(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("lista_notas.xlsx", index=False)