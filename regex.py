import re
import pandas as pd

def extrair_atributos(descricao):
    # Verifica se a descrição é uma string válida
    if not isinstance(descricao, str):
        return {"item": "Descrição inválida", "peso": "não informado", "unidade_medida": "não informado"}

    # Pré-processa a descrição para normalizar números e formatos
    descricao = descricao.lower()
    descricao = re.sub(r"(\d+)[\.,](\d+)", r"\1.\2", descricao)  # Normaliza números decimais (01,5 -> 1.5)
    descricao = re.sub(r"\b(um|uma|dois|duas|meio|metade)\b", "1", descricao)  # Converte palavras numéricas para números
    descricao = re.sub(r"quilo[s]?", "kg", descricao)  # Uniformiza "quilo" para "kg"
    descricao = re.sub(r"gramas?", "g", descricao)  # Uniformiza "gramas" para "g"

    # Regex para capturar o nome do item (pega as primeiras palavras até o primeiro número ou palavra-chave de peso)
    item_pattern = r"([A-Z\s]+)(?=\W|$)"
    
    # Regex para capturar o peso e a unidade de medida
    peso_pattern = r"(\d+(?:\.\d+)?)\s?(kg|g|litro|ml)"
    
    # Busca o item na descrição
    item_match = re.search(item_pattern, descricao.upper())
    item = item_match.group(1).strip() if item_match else "Item não encontrado"
    
    # Busca o peso na descrição
    peso_match = re.search(peso_pattern, descricao)
    if peso_match:
        peso = peso_match.group(1)
        unidade_medida = peso_match.group(2)
    else:
        peso = "não informado"
        unidade_medida = "não informado"

    return {"item": item, "peso": peso, "unidade_medida": unidade_medida}

def processar_excel(arquivo_excel):
    # Carrega o arquivo Excel
    df = pd.read_excel(arquivo_excel)

    # Aplica a função de extração de atributos para preencher novas colunas
    atributos = df['DESCRICAO'].apply(extrair_atributos)
    df['ITEM'] = atributos.apply(lambda x: x['item'])
    df['PESO'] = atributos.apply(lambda x: x['peso'])
    df['UNIDADE_MEDIDA'] = atributos.apply(lambda x: x['unidade_medida'])

    # Salva o arquivo Excel atualizado
    df.to_excel('resultado_regex.xlsx', index=False)
    print("Arquivo processado e salvo como 'resultado_regex.xlsx'.")

# Processa o arquivo '200Produtos.xlsx' -> este arquivo contém as 200 descrições distintas
processar_excel('/home/usr/tcc-structured-extraction/200Produtos.xlsx')