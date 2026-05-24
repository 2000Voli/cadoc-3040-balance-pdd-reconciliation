# -*- coding: utf-8 -*-
"""
Created on Thu Apr 17 17:18:19 2025

@author: ov0006
"""


# Importação das bibliotecas necessárias

import xml.etree.ElementTree as ET
import pandas as pd
import os
import chardet


# Funções para detectar o encoding do XML

def detectar_encoding(caminho_arquivo):
    with open(caminho_arquivo, 'rb') as f:
        rawdata = f.read(10000)
    return chardet.detect(rawdata)['encoding']

def ler_xml_com_encoding(caminho_xml):
    encoding_detectado = detectar_encoding(caminho_xml)
    try:
        with open(caminho_xml, "r", encoding=encoding_detectado) as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(caminho_xml, "r", encoding="ISO-8859-1") as f:
                return f.read()
        except UnicodeDecodeError:
            with open(caminho_xml, "r", encoding="utf-8-sig") as f:
                return f.read()

# Processamento do XML de maneira incremental para lidar com arquivos grandes

def processar_xml_incremental(caminho_xml):
    if os.path.getsize(caminho_xml) == 0:
        print(f"Aviso: O arquivo {caminho_xml} está vazio e será ignorado.")
        return None

    saldo_total = pdd_total = prejuizo_total = 0
    saldo_cli = pdd_cli = prejuizo_cli = 0
    saldo_agreg = pdd_agreg = prejuizo_agreg = 0
    classop_totals = {}

    try:
        for event, elem in ET.iterparse(caminho_xml, events=("start", "end")):
            if event == "end" and elem.tag == "Cli":
                for operacao in elem.findall('Op'):
                    classop = operacao.attrib.get('ClassOp', elem.attrib.get('ClassOp', 'N/A'))
                    classop_saldo = classop_pdd = classop_prejuizo = 0

                    for vencimento in operacao.findall('Venc'):
                        for vcod, valor in vencimento.attrib.items():
                            valor_f = float(valor)
                            if vcod.startswith(("v310", "v320", "v330")):
                                prejuizo_cli += valor_f
                                classop_prejuizo += valor_f
                            else:
                                saldo_cli += valor_f
                                classop_saldo += valor_f

                    pdd_value = float(operacao.attrib.get('ProvConsttd', 0))
                    pdd_cli += pdd_value
                    classop_pdd += pdd_value

                    if classop not in classop_totals:
                        classop_totals[classop] = {
                            'Saldo_Cli': 0, 'PDD_Cli': 0, 'Prejuizo_Cli': 0,
                            'Saldo_Agreg': 0, 'PDD_Agreg': 0, 'Prejuizo_Agreg': 0
                        }

                    classop_totals[classop]['Saldo_Cli'] += classop_saldo
                    classop_totals[classop]['PDD_Cli'] += classop_pdd
                    classop_totals[classop]['Prejuizo_Cli'] += classop_prejuizo

                elem.clear()

            elif event == "end" and elem.tag == "Agreg":
                classop = elem.attrib.get('ClassOp', 'N/A')
                classop_saldo = classop_pdd = classop_prejuizo = 0

                for vencimento in elem.findall('Venc'):
                    for vcod, valor in vencimento.attrib.items():
                        valor_f = float(valor)
                        if vcod.startswith(("v310", "v320", "v330")):
                            prejuizo_agreg += valor_f
                            classop_prejuizo += valor_f
                        else:
                            saldo_agreg += valor_f
                            classop_saldo += valor_f

                pdd_value = float(elem.attrib.get('ProvConsttd', 0))
                pdd_agreg += pdd_value
                classop_pdd += pdd_value

                if classop not in classop_totals:
                    classop_totals[classop] = {
                        'Saldo_Cli': 0, 'PDD_Cli': 0, 'Prejuizo_Cli': 0,
                        'Saldo_Agreg': 0, 'PDD_Agreg': 0, 'Prejuizo_Agreg': 0
                    }

                classop_totals[classop]['Saldo_Agreg'] += classop_saldo
                classop_totals[classop]['PDD_Agreg'] += classop_pdd
                classop_totals[classop]['Prejuizo_Agreg'] += classop_prejuizo

                elem.clear()

        saldo_total = saldo_cli + saldo_agreg
        pdd_total = pdd_cli + pdd_agreg
        prejuizo_total = prejuizo_cli + prejuizo_agreg

        return saldo_total, pdd_total, prejuizo_total, classop_totals

    except ET.ParseError as e:
        print(f"Erro ao processar {caminho_xml}: {e}")
        return None

# Função para formatar os valores em BR

def formatar_valor(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Função para gerar o relatório com as colunas desejadas (Saldo, Prejuízo e PDD)

def gerar_relatorio_saldo_pdd(diretorio_xml):
    dados_conciliacao = []

    for xml_file in os.listdir(diretorio_xml):
        if xml_file.endswith(".xml"):
            caminho_xml = os.path.join(diretorio_xml, xml_file)

            try:
                resultado = processar_xml_incremental(caminho_xml)
                if resultado:
                    saldo, pdd, prejuizo, classop_totals = resultado
                    saldo_formatado = formatar_valor(saldo)
                    pdd_formatado = formatar_valor(pdd)
                    prejuizo_formatado = formatar_valor(prejuizo)

                    dados_conciliacao.append({
                        "Nome do arquivo XML": xml_file,
                        "Saldo": saldo_formatado,
                        "PDD": pdd_formatado,
                        "Prejuízo": prejuizo_formatado
                    })
            except Exception as e:
                print(f"Erro ao processar {xml_file}: {e}")

    df_conciliacao = pd.DataFrame(dados_conciliacao)

    caminho_saida = os.path.join(
        diretorio_xml,
        r"outputs/saldo_pdd_prejuizo.xlsx" # Alterar para o diretório onde deseja que seja gerado o arquivo excel do saldo e pdd
    )

    with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
        df_conciliacao.to_excel(writer, sheet_name="Resumo", index=False)

    print("Relatório de saldo, PDD e prejuízo gerado!")

# Execução do código
diretorio_xml = r"data/xmls" # Alterar para o diretório onde está o XML que deverá ser lido
gerar_relatorio_saldo_pdd(diretorio_xml)
